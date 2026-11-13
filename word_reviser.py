#!/usr/bin/env python3
"""
Word docx revision tool using lxml for robust XML manipulation.
Generates Word Track Changes (revision mode) with format preservation.

Text model: the searchable text of a paragraph is the concatenation of every
w:t in runs that are not inside an existing revision (w:ins / w:del /
w:moveFrom / w:moveTo). Tabs, breaks, field codes, footnote/comment references
and drawings are zero-width: they are never part of a match, and every edit
keeps them (and bookmarks, comment ranges and other people's revisions) in
place.
"""

from collections import OrderedDict
from copy import deepcopy
from datetime import datetime, timedelta, timezone
from difflib import SequenceMatcher
import glob
import json
import os
import re
import shutil
import sys
import tempfile
import unicodedata
import warnings
import zipfile

from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
RELS = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'

SUPPORTED_PARTS = ('document', 'headers', 'footers', 'footnotes', 'endnotes', 'comments')
PART_PATTERNS = {
    'document': ['word/document.xml'],
    'headers': ['word/header*.xml'],
    'footers': ['word/footer*.xml'],
    'footnotes': ['word/footnotes.xml'],
    'endnotes': ['word/endnotes.xml'],
    'comments': ['word/comments.xml'],
}
REVISION_TAGS = ('ins', 'del', 'moveFrom', 'moveTo')

BEIJING_TZ = timezone(timedelta(hours=8))


def wtag(tag):
    return '{%s}%s' % (W, tag)


REVISION_TAG_SET = frozenset(wtag(tag) for tag in REVISION_TAGS)
DELETED_TAG_SET = frozenset((wtag('del'), wtag('moveFrom')))
P_TAG = wtag('p')
R_TAG = wtag('r')
T_TAG = wtag('t')
RPR_TAG = wtag('rPr')
PPR_TAG = wtag('pPr')
TXBX_TAG = wtag('txbxContent')
# Schema order inside w:pPr is "... w:rPr, w:sectPr, w:pPrChange".
PPR_AFTER_RPR = frozenset((wtag('sectPr'), wtag('pPrChange')))
PARAGRAPH_MARK_REVISIONS = (wtag('ins'), wtag('del'), wtag('moveFrom'), wtag('moveTo'))
ANNOTATION_RUN_CHILDREN = frozenset((wtag('commentReference'), wtag('annotationRef')))
TOC_STYLE_NAME = re.compile(r'^toc\s*\d+$', re.IGNORECASE)
TOC_STYLE_ID = re.compile(r'^toc\d+$', re.IGNORECASE)
# Every one of these in new text is a paragraph mark (^p); the tool never writes w:br.
PARAGRAPH_BREAKS = (('\r\n', '\n'), ('\r', '\n'), ('\x0b', '\n'), ('\x0c', '\n'), (' ', '\n'), (' ', '\n'))
FIND_BREAK_CHARS = '\n\r\x0b\x0c  '
# XML 1.0 forbids C0 controls other than tab, newline and carriage return.
ILLEGAL_XML_CHARS = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1f￾￿\ud800-\udfff]')
NUMBERING_PREFIX = re.compile(
    r'^\s*(?:第[一二三四五六七八九十百千零〇两\d]+[条章节款项篇]'
    r'|[（(][一二三四五六七八九十\d]+[）)]'
    r'|[一二三四五六七八九十]+[、.．]'
    r'|\d+(?:[.．]\d+)*[.、．]?'
    r'|[A-Za-z][.)、])\s*'
)
# Children of w:settings that the schema places after w:trackRevisions.
SETTINGS_AFTER_TRACK_REVISIONS = frozenset('''
    doNotTrackMoves doNotTrackFormatting documentProtection autoFormatOverride styleLockTheme
    styleLockQFSet defaultTabStop autoHyphenation consecutiveHyphenLimit hyphenationZone
    doNotHyphenateCaps showEnvelope summaryLength clickAndTypeStyle defaultTableStyle
    evenAndOddHeaders bookFoldRevPrinting bookFoldPrinting bookFoldPrintingSheets
    drawingGridHorizontalSpacing drawingGridVerticalSpacing displayHorizontalDrawingGridEvery
    displayVerticalDrawingGridEvery doNotUseMarginsForDrawingGridOrigin drawingGridHorizontalOrigin
    drawingGridVerticalOrigin doNotShadeFormData noPunctuationKerning characterSpacingControl
    printTwoOnOne strictFirstAndLastChars noLineBreaksAfter noLineBreaksBefore savePreviewPicture
    doNotValidateAgainstSchema saveInvalidXml ignoreMixedContent alwaysShowPlaceholderText
    doNotDemarcateInvalidXml saveXmlDataOnly useXSLTWhenSaving saveThroughXslt showXMLTags
    alwaysMergeEmptyNamespace updateFields hdrShapeDefaults footnotePr endnotePr compat docVars
    rsids mathPr attachedSchema themeFontLang clrSchemeMapping doNotIncludeSubdocsInStats
    doNotAutoCompressPictures forceUpgrade captions readModeInkLockDown smartTagType schemaLibrary
    shapeDefaults doNotEmbedSmartTags decimalSymbol listSeparator
'''.split())


def _is_under_revision(elem):
    for anc in elem.iterancestors():
        if anc.tag in REVISION_TAG_SET:
            return True
    return False


def _nearest_paragraph(elem):
    for anc in elem.iterancestors(P_TAG):
        return anc
    return None


def _local(tag):
    return tag.split('}')[-1] if isinstance(tag, str) else str(tag)


def _set_text(t, text):
    """Set w:t / w:delText content, adding xml:space=preserve when whitespace would be lost."""
    t.text = text
    if text and (text[0].isspace() or text[-1].isspace() or '  ' in text):
        t.set(XML_SPACE, 'preserve')
    return t


def _content_children(run):
    return [child for child in run if isinstance(child.tag, str) and child.tag != RPR_TAG]


def _find_all(text, needle):
    """Non-overlapping occurrences, left to right (same as Word and str.replace)."""
    positions = []
    start = text.find(needle)
    while start != -1:
        positions.append(start)
        start = text.find(needle, start + len(needle))
    return positions


def _snippet(text, start, length, radius=15):
    a = max(0, start - radius)
    b = min(len(text), start + length + radius)
    return ('…' if a > 0 else '') + text[a:b] + ('…' if b < len(text) else '')


def diff_opcodes(old, new, merge_short_equal=2):
    """Character diff opcodes; equal islands of <= merge_short_equal chars between two
    changes are folded into one replacement so redlines read as whole phrases."""
    ops = [list(op) for op in SequenceMatcher(None, old, new, autojunk=False).get_opcodes()]
    if merge_short_equal and merge_short_equal > 0:
        merged = True
        while merged:
            merged = False
            for i in range(1, len(ops) - 1):
                tag, i1, i2 = ops[i][0], ops[i][1], ops[i][2]
                if tag == 'equal' and i2 - i1 <= merge_short_equal \
                        and ops[i - 1][0] != 'equal' and ops[i + 1][0] != 'equal':
                    prev, nxt = ops[i - 1], ops[i + 1]
                    ops[i - 1:i + 2] = [['replace', prev[1], nxt[2], prev[3], nxt[4]]]
                    merged = True
                    break
    result = []
    for tag, i1, i2, j1, j2 in ops:
        if tag == 'replace' and i1 == i2:
            tag = 'insert'
        elif tag == 'replace' and j1 == j2:
            tag = 'delete'
        result.append((tag, i1, i2, j1, j2))
    return result


class TextNotFoundError(ValueError):
    """Find/anchor text not found; carries closest-match suggestions and hints."""

    def __init__(self, message, suggestions=None, hints=None):
        super().__init__(message)
        self.suggestions = suggestions or []
        self.hints = hints or []


class DocxReviser:
    BATCH_OPS = {
        'replace': 'replace_text',
        'replace_text': 'replace_text',
        'smart_replace': 'replace_text_smart',
        'replace_text_smart': 'replace_text_smart',
        'replace_full': 'replace_text_full',
        'replace_text_full': 'replace_text_full',
        'insert_text_after': 'insert_text_after',
        'insert_after': 'insert_after',
        'insert_multiple_after': 'insert_multiple_after',
        'replace_para': 'replace_paragraph',
        'replace_paragraph': 'replace_paragraph',
        'delete_para': 'delete_paragraph',
        'delete_paragraph': 'delete_paragraph',
    }
    TEXT_OP_NEEDLES = {
        'replace_text': 'find',
        'replace_text_smart': 'old_text',
        'replace_text_full': 'find',
        'insert_text_after': 'anchor',
    }

    def __init__(self, input_path, author='Agent', parts=None, include_toc=False, merge_short_equal=2):
        self.input_path = input_path
        self.author = author
        self.include_toc = include_toc
        self.merge_short_equal = merge_short_equal
        self.rev_id = 1000
        self.last_error = None
        self.part_names = self._normalize_parts(parts)
        self.part_trees = []
        self._style_names = {}
        self._zip_order = []

        self.temp_dir = tempfile.mkdtemp(prefix='word_reviser_')
        try:
            with zipfile.ZipFile(input_path, 'r') as z:
                self._zip_order = [name for name in z.namelist() if not name.endswith('/')]
                z.extractall(self.temp_dir)
            self._load_parts()
            self._style_names = self._load_style_names()
            self.rev_id = max(1000, self._max_existing_id() + 1)
        except Exception:
            self.close()
            raise

    @property
    def revision_date(self):
        return datetime.now(BEIJING_TZ).replace(microsecond=0).isoformat()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        self.close()
        return False

    def close(self):
        """Remove the extracted working copy. Called automatically by save()."""
        temp_dir = getattr(self, 'temp_dir', None)
        if temp_dir and os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)
        self.temp_dir = None

    # ------------------------------------------------------------------ loading

    def _normalize_parts(self, parts):
        if parts is None:
            return ['document']
        if parts == 'all':
            return list(SUPPORTED_PARTS)
        if isinstance(parts, str):
            parts = [item.strip() for item in parts.split(',') if item.strip()]
        unknown = [part for part in parts if part not in SUPPORTED_PARTS]
        if unknown:
            raise ValueError("Unsupported parts: %s" % ', '.join(unknown))
        return list(parts)

    def _part_paths_for(self, part_name):
        paths = []
        for pattern in PART_PATTERNS[part_name]:
            full_pattern = os.path.join(self.temp_dir, pattern)
            matches = sorted(glob.glob(full_pattern))
            if matches:
                paths.extend(matches)
            elif '*' not in pattern:
                path = os.path.join(self.temp_dir, pattern)
                if os.path.exists(path):
                    paths.append(path)
        return paths

    def _load_parts(self):
        for part_name in self.part_names:
            for path in self._part_paths_for(part_name):
                tree = etree.parse(path)
                self.part_trees.append({
                    'name': part_name,
                    'path': path,
                    'tree': tree,
                    'root': tree.getroot(),
                })
        if not self.part_trees:
            raise ValueError("No supported Word XML parts found for: %s" % ', '.join(self.part_names))

    def _load_style_names(self):
        names = {}
        path = os.path.join(self.temp_dir, 'word', 'styles.xml')
        if not os.path.exists(path):
            return names
        try:
            root = etree.parse(path).getroot()
        except etree.XMLSyntaxError:
            return names
        for style in root.iter(wtag('style')):
            style_id = style.get(wtag('styleId'))
            name = style.find(wtag('name'))
            if style_id and name is not None:
                names[style_id] = (name.get(wtag('val')) or '').strip().lower()
        return names

    def _max_existing_id(self):
        """Largest numeric w:id in any word/ XML part, so new revision ids never collide."""
        loaded = {os.path.abspath(part['path']): part['root'] for part in self.part_trees}
        max_id = -1
        pattern = os.path.join(self.temp_dir, 'word', '**', '*.xml')
        for path in glob.glob(pattern, recursive=True):
            root = loaded.get(os.path.abspath(path))
            if root is not None:
                values = root.xpath('//@w:id', namespaces={'w': W})
            else:
                with open(path, 'rb') as fh:
                    values = [v.decode('ascii') for v in re.findall(rb'\bw:id="(-?\d+)"', fh.read())]
            for value in values:
                try:
                    max_id = max(max_id, int(value))
                except ValueError:
                    pass
        return max_id

    # ------------------------------------------------------------- paragraphs

    def _all_paragraphs(self):
        """Get all paragraphs in selected parts, including inside tables/text boxes."""
        paragraphs = []
        for part in self.part_trees:
            paragraphs.extend(part['root'].findall('.//' + wtag('p')))
        return paragraphs

    def _target_paragraphs(self, include_nested=True):
        """Paragraphs that find-based operations may edit (TOC skipped unless include_toc)."""
        result = []
        for para in self._all_paragraphs():
            if not include_nested and _nearest_paragraph(para) is not None:
                continue
            if not self.include_toc and self._is_toc_paragraph(para):
                continue
            result.append(para)
        return result

    def _is_toc_paragraph(self, para):
        ppr = para.find(PPR_TAG)
        if ppr is not None:
            pstyle = ppr.find(wtag('pStyle'))
            if pstyle is not None:
                style_id = pstyle.get(wtag('val')) or ''
                if TOC_STYLE_NAME.match(self._style_names.get(style_id, '')) or TOC_STYLE_ID.match(style_id):
                    return True
        for sdt in para.iterancestors(wtag('sdt')):
            gallery = sdt.find('%s/%s/%s' % (wtag('sdtPr'), wtag('docPartObj'), wtag('docPartGallery')))
            if gallery is not None and (gallery.get(wtag('val')) or '').lower().startswith('table of contents'):
                return True
        return False

    def _paragraph_index(self, para):
        for index, candidate in enumerate(self._all_paragraphs()):
            if candidate is para:
                return index
        return -1

    # ------------------------------------------------------------- text model

    def _editable_runs(self, para):
        return [run for run in para.iter(R_TAG) if not _is_under_revision(run)]

    def _own_runs(self, para):
        """Editable runs that belong to this paragraph itself (not to a nested text box)."""
        return [run for run in self._editable_runs(para) if _nearest_paragraph(run) is para]

    def _run_text(self, run):
        return ''.join(child.text or '' for child in run if child.tag == T_TAG)

    def _para_text(self, para):
        """Get editable paragraph text, excluding existing revision text."""
        return ''.join(self._run_text(run) for run in self._editable_runs(para))

    def _own_text(self, para):
        return ''.join(self._run_text(run) for run in self._own_runs(para))

    def _visible_text(self, para, own=False):
        """Text a reader sees with markup shown: includes insertions, excludes deletions."""
        parts = []
        for run in para.iter(R_TAG):
            if own and _nearest_paragraph(run) is not para:
                continue
            if any(anc.tag in DELETED_TAG_SET for anc in run.iterancestors()):
                continue
            parts.append(self._run_text(run))
        return ''.join(parts)

    def _visible_context(self, para, start, length):
        """Snippet of what the paragraph reads now (insertions shown, deletions hidden)
        around the editable-text range [start, start+length)."""
        visible = []
        vis_pos = edit_pos = 0
        vis_start = vis_end = None
        end = start + length
        for run in para.iter(R_TAG):
            if any(anc.tag in DELETED_TAG_SET for anc in run.iterancestors()):
                continue
            text = self._run_text(run)
            if not text:
                continue
            if not _is_under_revision(run):
                if vis_start is None and edit_pos <= start < edit_pos + len(text):
                    vis_start = vis_pos + start - edit_pos
                if vis_end is None and edit_pos < end <= edit_pos + len(text):
                    vis_end = vis_pos + end - edit_pos
                edit_pos += len(text)
            visible.append(text)
            vis_pos += len(text)
        if vis_start is None or vis_end is None:
            return _snippet(self._para_text(para), start, length)
        return _snippet(''.join(visible), vis_start, vis_end - vis_start)

    def _normalize_runs(self, para):
        """Split runs holding several content children (e.g. w:tab + w:t, or several w:t)
        into one-child runs with the same properties, so copies never duplicate content."""
        for run in self._editable_runs(para):
            children = _content_children(run)
            if len(children) <= 1:
                continue
            parent = run.getparent()
            index = parent.index(run)
            rpr = run.find(RPR_TAG)
            parent.remove(run)
            for offset, child in enumerate(children):
                piece = etree.Element(R_TAG, attrib=dict(run.attrib))
                if rpr is not None:
                    piece.append(deepcopy(rpr))
                piece.append(child)
                parent.insert(index + offset, piece)

    def _run_entries(self, para, normalize=True):
        """Return editable runs with text offsets; zero-width runs carry 'zw': True."""
        if normalize:
            self._normalize_runs(para)
        entries = []
        pos = 0
        for run in self._editable_runs(para):
            texts = [child for child in run if child.tag == T_TAG]
            text = ''.join(t.text or '' for t in texts)
            entries.append({
                'run': run,
                't': texts[0] if texts else None,
                'start': pos,
                'end': pos + len(text),
                'text': text,
                'zw': not text,
            })
            pos += len(text)
        return entries

    def _locate_range(self, entries, start_idx, end_idx):
        first = last = None
        for i, entry in enumerate(entries):
            if entry['zw']:
                continue
            if first is None and entry['start'] <= start_idx < entry['end']:
                first = i
            if entry['start'] < end_idx <= entry['end']:
                last = i
                break
        if first is None or last is None or last < first:
            return None
        return first, last

    # ------------------------------------------------------- element builders

    def _clone_ppr(self, para):
        ppr = para.find(PPR_TAG)
        return deepcopy(ppr) if ppr is not None else None

    def _clean_ppr(self, ppr, keep_numbering=True):
        """Copy of pPr fit for a new paragraph mark: no section break, no existing revisions."""
        if ppr is None:
            return None
        clone = deepcopy(ppr)
        for child in list(clone):
            if child.tag in PPR_AFTER_RPR:
                clone.remove(child)
        rpr = clone.find(RPR_TAG)
        if rpr is not None:
            for child in list(rpr):
                if child.tag in PARAGRAPH_MARK_REVISIONS or child.tag == wtag('rPrChange'):
                    rpr.remove(child)
        if not keep_numbering:
            self._remove_numpr(clone)
        return clone

    def _clone_rpr(self, run):
        rpr = run.find(RPR_TAG)
        if rpr is None:
            return None
        clone = deepcopy(rpr)
        for change in clone.findall(wtag('rPrChange')):
            clone.remove(change)
        return clone

    def _set_bold(self, rpr, bold):
        for child in list(rpr):
            if child.tag in (wtag('b'), wtag('bCs')):
                rpr.remove(child)
        index = 0
        for i, child in enumerate(rpr):
            if child.tag in (wtag('rStyle'), wtag('rFonts')):
                index = i + 1
        for offset, tag in enumerate(('b', 'bCs')):
            elem = etree.Element(wtag(tag))
            if not bold:
                elem.set(wtag('val'), '0')
            rpr.insert(index + offset, elem)

    def _run_with_props(self, rpr_elem=None, bold=None):
        r = etree.Element(R_TAG)
        if rpr_elem is not None:
            r.append(deepcopy(rpr_elem))
        if bold is not None:
            rpr = r.find(RPR_TAG)
            if rpr is None:
                rpr = etree.Element(RPR_TAG)
                r.insert(0, rpr)
            self._set_bold(rpr, bold)
        return r

    def _make_run(self, text, rpr_elem=None, bold=None, deleted=False):
        """bold: None keeps the copied formatting, True forces bold, False forces not bold."""
        r = self._run_with_props(rpr_elem, bold)
        t = etree.SubElement(r, wtag('delText' if deleted else 't'))
        _set_text(t, text)
        return r

    def _make_runs_from_text(self, text, rpr_elem=None, bold=None, deleted=False):
        # Paragraph breaks are built at the API level; flatten any residual newline so
        # this helper can never emit w:br (^l). Tabs become real w:tab elements.
        text = text.replace('\r\n', '\n').replace('\r', '\n').replace('\n', ' ')
        runs = []
        for i, piece in enumerate(text.split('\t')):
            if i:
                tab_run = self._run_with_props(rpr_elem, bold)
                etree.SubElement(tab_run, wtag('tab'))
                runs.append(tab_run)
            if piece:
                runs.append(self._make_run(piece, rpr_elem, bold=bold, deleted=deleted))
        return runs or [self._make_run('', rpr_elem, bold=bold, deleted=deleted)]

    def _split_newline_segments(self, text):
        """Split text into paragraph segments; every newline is a paragraph break."""
        text = text.replace('\r\n', '\n').replace('\r', '\n')
        return text.split('\n')

    def _clean_new_text(self, text):
        if not isinstance(text, str):
            raise TypeError('text must be a string, got %s' % type(text).__name__)
        for source, target in PARAGRAPH_BREAKS:
            text = text.replace(source, target)
        match = ILLEGAL_XML_CHARS.search(text)
        if match:
            raise ValueError('new text contains control character U+%04X at position %d; remove it'
                             % (ord(match.group()), match.start()))
        return text

    def _validate_find(self, text):
        if not isinstance(text, str) or not text:
            raise ValueError('find/anchor text must be a non-empty string')
        if any(ch in text for ch in FIND_BREAK_CHARS):
            raise ValueError('find/anchor text cannot contain a line or paragraph break: matching never '
                             'spans paragraphs; split the edit into one call per paragraph')
        if '\t' in text:
            raise ValueError('find/anchor text cannot contain a tab: tabs are not part of the searchable '
                             'text (they are kept in place); remove the tab from the find text')
        match = ILLEGAL_XML_CHARS.search(text)
        if match:
            raise ValueError('find/anchor text contains control character U+%04X; remove it' % ord(match.group()))

    def _next_rev_id(self):
        rid = self.rev_id
        self.rev_id += 1
        return rid

    def _revision_marker(self, kind, rev_id=None):
        marker = etree.Element(wtag(kind))
        marker.set(wtag('id'), str(self._next_rev_id() if rev_id is None else rev_id))
        marker.set(wtag('author'), self.author)
        marker.set(wtag('date'), self.revision_date)
        return marker

    def _make_ins(self, children, rev_id=None):
        ins = self._revision_marker('ins', rev_id)
        for child in children:
            ins.append(child)
        return ins

    def _convert_run_to_deleted(self, run):
        run_copy = deepcopy(run)
        for node in list(run_copy.iter(T_TAG, wtag('instrText'))):
            owner = next(node.iterancestors(R_TAG), None)
            if owner is not run_copy:
                continue  # text of a nested story (text box) keeps its own markup
            node.tag = wtag('delText') if node.tag == T_TAG else wtag('delInstrText')
        return run_copy

    def _make_del(self, children, rev_id=None):
        d = self._revision_marker('del', rev_id)
        for child in children:
            d.append(self._convert_run_to_deleted(child))
        return d

    def _copy_run_with_text(self, run, text):
        run_copy = deepcopy(run)
        texts = [child for child in run_copy if child.tag == T_TAG]
        if texts:
            _set_text(texts[0], text)
            for extra in texts[1:]:
                run_copy.remove(extra)
        return run_copy

    def _runs_for_text_range(self, entries, start_idx, end_idx):
        """Copy editable runs covering a paragraph text range."""
        runs = []
        for entry in entries:
            if entry.get('zw', entry['t'] is None) or entry['t'] is None:
                continue
            overlap_start = max(start_idx, entry['start'])
            overlap_end = min(end_idx, entry['end'])
            if overlap_start >= overlap_end:
                continue
            text = entry['text'][overlap_start - entry['start']:overlap_end - entry['start']]
            runs.append(self._copy_run_with_text(entry['run'], text))
        return runs

    def _rpr_at_or_before(self, entries, offset):
        fallback = None
        for entry in entries:
            if entry.get('zw', entry['t'] is None):
                continue
            fallback = entry['run']
            if entry['start'] <= offset <= entry['end']:
                return self._clone_rpr(entry['run'])
        if fallback is not None:
            return self._clone_rpr(fallback)
        return None

    def _ensure_ppr(self, para):
        ppr = para.find(PPR_TAG)
        if ppr is None:
            ppr = etree.Element(PPR_TAG)
            para.insert(0, ppr)
        return ppr

    def _ensure_ppr_rpr(self, ppr):
        rpr = ppr.find(RPR_TAG)
        if rpr is None:
            rpr = etree.Element(RPR_TAG)
            for index, child in enumerate(ppr):
                if child.tag in PPR_AFTER_RPR:
                    ppr.insert(index, rpr)
                    break
            else:
                ppr.append(rpr)
        return rpr

    def _mark_paragraph_inserted(self, ppr):
        """Track the paragraph mark itself, so Reject All removes the paragraph entirely."""
        self._ensure_ppr_rpr(ppr).insert(0, self._revision_marker('ins'))
        return ppr

    def _make_inserted_paragraph(self, text, ppr=None, rpr=None, bold=None):
        p = etree.Element(P_TAG)
        p_ppr = self._clean_ppr(ppr) if ppr is not None else etree.Element(PPR_TAG)
        self._mark_paragraph_inserted(p_ppr)
        p.append(p_ppr)
        if text:
            p.append(self._make_ins(self._make_runs_from_text(text, rpr, bold=bold)))
        return p

    def _remove_numpr(self, ppr):
        if ppr is not None:
            numpr = ppr.find(wtag('numPr'))
            if numpr is not None:
                ppr.remove(numpr)
        return ppr

    def _first_editable_run(self, para):
        runs = self._own_runs(para)
        return runs[0] if runs else None

    def _first_text_run(self, para):
        for run in self._own_runs(para):
            if self._run_text(run):
                return run
        return None

    def _is_annotation_run(self, run):
        children = _content_children(run)
        return bool(children) and all(child.tag in ANNOTATION_RUN_CHILDREN for child in children)

    def _has_revisions(self, para):
        return bool(para.xpath('.//w:ins | .//w:del | .//w:moveFrom | .//w:moveTo', namespaces={'w': W}))

    # ----------------------------------------------------------- range edits

    def _range_problem(self, para, start_idx, old_text, new_text=''):
        """Why [start_idx, start_idx+len(old_text)) cannot be edited, or None."""
        if not old_text:
            return 'empty text range'
        entries = self._run_entries(para, normalize=False)
        text = ''.join(entry['text'] for entry in entries)
        end_idx = start_idx + len(old_text)
        if start_idx < 0 or text[start_idx:end_idx] != old_text:
            return 'paragraph text at offset %d does not match %r' % (start_idx, old_text)
        located = self._locate_range(entries, start_idx, end_idx)
        if located is None:
            return 'could not map the text range onto runs'
        first_i, last_i = located
        parents = []
        for entry in entries[first_i:last_i + 1]:
            parent = entry['run'].getparent()
            if all(parent is not seen for seen in parents):
                parents.append(parent)
        if len(parents) > 1:
            containers = sorted({_local(p.tag) for p in parents if p.tag != P_TAG})
            where = '/'.join('w:' + name for name in containers) if containers else 'text box (nested paragraph)'
            return 'match crosses a %s boundary; edit the text inside and outside it separately' % where
        if '\n' in new_text and parents[0].tag != P_TAG:
            return ('paragraph breaks cannot be inserted inside a w:%s; edit without line breaks or use '
                    'insert_after' % _local(parents[0].tag))
        return None

    def _render_ops(self, entries, start_idx, old_text, new_text, ops, fixed):
        """Build the replacement children for a range from diff opcodes.

        `fixed` holds (relative offset, element) for everything between the first and
        last run that is not text (zero-width runs, bookmarks, other people's
        revisions); each is re-emitted at its offset, after any insertion there.
        """
        pieces = []
        pending = list(fixed)
        cursor = [0]

        def flush(upto):
            while cursor[0] < len(pending) and pending[cursor[0]][0] <= upto:
                pieces.append(pending[cursor[0]][1])
                cursor[0] += 1

        for tag, i1, i2, j1, j2 in ops:
            inserted = new_text[j1:j2]
            if tag == 'insert':
                if inserted:
                    rpr = self._rpr_at_or_before(entries, start_idx + i1)
                    pieces.append(self._make_ins(self._make_runs_from_text(inserted, rpr)))
                continue
            flush(i1)
            cuts = [i1] + sorted({off for off, _ in pending[cursor[0]:] if i1 < off < i2}) + [i2]
            for a, b in zip(cuts, cuts[1:]):
                if a > i1:
                    flush(a)
                runs = self._runs_for_text_range(entries, start_idx + a, start_idx + b)
                if not runs:
                    continue
                if tag == 'equal':
                    pieces.extend(runs)
                else:
                    pieces.append(self._make_del(runs))
            if tag == 'replace' and inserted:
                rpr = self._rpr_at_or_before(entries, start_idx + i1)
                pieces.append(self._make_ins(self._make_runs_from_text(inserted, rpr)))
        flush(len(old_text))
        return pieces

    def _edit_range(self, para, start_idx, old_text, new_text, mode='smart'):
        """Replace a text range. mode 'smart' diffs old/new so shared text stays normal;
        'full' marks one whole deletion and one insertion. Newlines in new_text split
        the paragraph (Word ^p). Returns False (reason in self.last_error) if refused."""
        self.last_error = None
        new_text = self._clean_new_text(new_text)
        problem = self._range_problem(para, start_idx, old_text, new_text)
        if problem:
            self.last_error = problem
            return False

        entries = self._run_entries(para)
        end_idx = start_idx + len(old_text)
        first_i, last_i = self._locate_range(entries, start_idx, end_idx)
        first, last = entries[first_i], entries[last_i]
        parent = first['run'].getparent()
        children = list(parent)
        first_pos = children.index(first['run'])
        last_pos = children.index(last['run'])

        text_runs = {entry['run']: entry for entry in entries if not entry['zw']}
        fixed = []
        offset = first['end']
        for child in children[first_pos + 1:last_pos]:
            entry = text_runs.get(child)
            if entry is not None:
                offset = entry['end']
            else:
                fixed.append((offset - start_idx, child))

        rpr = self._rpr_at_or_before(entries, start_idx)
        before = self._runs_for_text_range(entries, first['start'], start_idx)
        after = self._runs_for_text_range(entries, end_idx, last['end'])
        segments = self._split_newline_segments(new_text)
        line = segments[0]
        if mode == 'full' or len(segments) > 1:
            ops = [('replace' if line else 'delete', 0, len(old_text), 0, len(line))]
        else:
            ops = diff_opcodes(old_text, line, self.merge_short_equal)
        pieces = self._render_ops(entries, start_idx, old_text, line, ops, fixed)

        for child in children[first_pos:last_pos + 1]:
            parent.remove(child)
        position = first_pos
        for child in before + pieces + after:
            parent.insert(position, child)
            position += 1

        if len(segments) > 1:
            self._split_paragraph(parent, first_pos + len(before) + len(pieces), segments[1:], rpr)
        return True

    def _split_paragraph(self, host_p, split_index, segments, rpr, bold=None):
        """Split host_p before child index `split_index` (Word ^p semantics).

        Each segment becomes a new paragraph. Content after the split point moves to
        the last one, which also receives host_p's original paragraph mark (pPr, incl.
        any section break); host_p and intermediate paragraphs get new, tracked marks,
        so Reject All re-joins everything into the original paragraph."""
        trailing = [child for child in list(host_p)[split_index:] if child.tag != PPR_TAG]
        for child in trailing:
            host_p.remove(child)
        original_ppr = host_p.find(PPR_TAG)
        clean = self._clean_ppr(original_ppr)
        if original_ppr is not None:
            host_p.remove(original_ppr)

        def tracked_mark():
            ppr = deepcopy(clean) if clean is not None else etree.Element(PPR_TAG)
            return self._mark_paragraph_inserted(ppr)

        host_p.insert(0, tracked_mark())
        outer = host_p.getparent()
        position = outer.index(host_p) + 1
        last_p = None
        for i, segment in enumerate(segments):
            is_last = i == len(segments) - 1
            new_p = etree.Element(P_TAG)
            if not is_last:
                new_p.append(tracked_mark())
            elif original_ppr is not None:
                new_p.append(original_ppr)
            if segment:
                new_p.append(self._make_ins(self._make_runs_from_text(segment, rpr, bold=bold)))
            if is_last:
                for child in trailing:
                    new_p.append(child)
            outer.insert(position, new_p)
            position += 1
            last_p = new_p
        return last_p

    def _replace_range_smart(self, para, start_idx, old_text, new_text):
        """Replace a range using a text diff so unchanged text stays unmarked."""
        return self._edit_range(para, start_idx, old_text, new_text, 'smart')

    def _replace_range_full(self, para, start_idx, old_text, new_text):
        """Replace a range as one whole deletion and one whole insertion."""
        return self._edit_range(para, start_idx, old_text, new_text, 'full')

    def _replace_range_with_paragraphs(self, para, start_idx, old_text, new_text):
        return self._edit_range(para, start_idx, old_text, new_text, 'full')

    def _insert_problem(self, para, offset, text):
        for entry in self._run_entries(para, normalize=False):
            if entry['zw']:
                continue
            if entry['start'] <= offset <= entry['end']:
                parent = entry['run'].getparent()
                if '\n' in text and parent.tag != P_TAG:
                    return 'paragraph breaks cannot be inserted inside a w:%s' % _local(parent.tag)
                return None
        return 'offset %d is outside the paragraph text' % offset

    def _insert_text_at(self, para, offset, text):
        self.last_error = None
        text = self._clean_new_text(text)
        problem = self._insert_problem(para, offset, text)
        if problem:
            self.last_error = problem
            return False
        entries = self._run_entries(para)
        target = next((e for e in entries if not e['zw'] and e['start'] <= offset <= e['end']), None)
        if target is None:
            self.last_error = 'offset %d is outside the paragraph text' % offset
            return False

        run = target['run']
        parent = run.getparent()
        rpr = self._clone_rpr(run)
        cut = offset - target['start']
        position = parent.index(run)
        parent.remove(run)
        segments = self._split_newline_segments(text)

        new_children = []
        if cut > 0:
            new_children.append(self._copy_run_with_text(run, target['text'][:cut]))
        if segments[0]:
            new_children.append(self._make_ins(self._make_runs_from_text(segments[0], rpr)))
        split_index = position + len(new_children)
        if cut < len(target['text']):
            new_children.append(self._copy_run_with_text(run, target['text'][cut:]))
        for child in new_children:
            parent.insert(position, child)
            position += 1

        if len(segments) > 1:
            self._split_paragraph(parent, split_index, segments[1:], rpr)
        return True

    def _insert_text_with_paragraphs(self, para, offset, text):
        return self._insert_text_at(para, offset, text)

    # --------------------------------------------------------------- matching

    def _text_matches(self, needle):
        matches = []
        for para in self._target_paragraphs(include_nested=False):
            text = self._para_text(para)
            if needle in text:
                matches.extend((para, start) for start in _find_all(text, needle))
        return matches

    def _select_matches(self, matches, needle, replace_all, occurrence, describe):
        if replace_all:
            return list(matches)
        if occurrence is not None:
            if isinstance(occurrence, bool) or not isinstance(occurrence, int) \
                    or not 1 <= occurrence <= len(matches):
                raise ValueError('occurrence=%r is out of range: %r matched %d place(s)'
                                 % (occurrence, needle, len(matches)))
            return [matches[occurrence - 1]]
        if len(matches) > 1:
            warnings.warn(
                '%r matched %d places; only the first (%s) was changed. Pass occurrence=N to choose '
                'another, or replace_all=True to change all.' % (needle, len(matches), describe(matches[0])),
                UserWarning, stacklevel=4)
        return list(matches[:1])

    def _describe_text_match(self, needle):
        def describe(match):
            para, start = match
            return 'paragraph #%d: "%s"' % (self._paragraph_index(para),
                                            self._visible_context(para, start, len(needle)))
        return describe

    def _apply_text_edit(self, needle, replace_all, occurrence, check, apply):
        """Find every match once, validate all, then apply right-to-left per paragraph.

        Never re-searches after an edit, so it always terminates, hits every original
        occurrence and never re-matches text produced by its own edits."""
        self._validate_find(needle)
        if replace_all and occurrence is not None:
            raise ValueError('occurrence cannot be combined with replace_all=True')
        matches = self._text_matches(needle)
        if not matches:
            raise self._not_found_error(needle)
        selected = self._select_matches(matches, needle, replace_all, occurrence, self._describe_text_match(needle))
        for para, start in selected:
            problem = check(para, start)
            if problem:
                raise ValueError('Cannot edit %r at paragraph #%d: %s' % (needle, self._paragraph_index(para), problem))
        grouped = OrderedDict()
        for para, start in selected:
            grouped.setdefault(para, []).append(start)
        count = 0
        for para, starts in grouped.items():
            for start in sorted(starts, reverse=True):
                if not apply(para, start):
                    raise RuntimeError('Failed to apply edit for %r at paragraph #%d: %s'
                                       % (needle, self._paragraph_index(para), self.last_error))
                count += 1
        return count

    def _paragraph_matches(self, find, startswith=False):
        self._validate_find(find)
        matches = []
        for para in self._target_paragraphs(include_nested=True):
            text = self._own_text(para)
            if (text.startswith(find) if startswith else find in text):
                matches.append(para)
        return matches

    def _select_paragraph(self, find, startswith=False, occurrence=None):
        matches = self._paragraph_matches(find, startswith)
        if not matches:
            raise self._not_found_error(find, label='Paragraph', own=True, startswith=startswith)

        def describe(para):
            return 'paragraph #%d: "%s"' % (self._paragraph_index(para), _snippet(self._own_text(para), 0, len(find)))

        return self._select_matches(matches, find, False, occurrence, describe)[0]

    def _find_para(self, text, startswith=False):
        """Find first paragraph containing or starting with text (own text, TOC skipped)."""
        try:
            matches = self._paragraph_matches(text, startswith)
        except ValueError:
            return None
        return matches[0] if matches else None

    def _not_found_error(self, needle, label='Text', own=False, startswith=False):
        text_of = self._own_text if own else self._para_text
        norm_needle = unicodedata.normalize('NFKC', needle)
        numbering = NUMBERING_PREFIX.match(needle)
        remainder = needle[numbering.end():] if numbering else ''
        matcher = SequenceMatcher(None, autojunk=False)
        matcher.set_seq2(needle)
        scorer = SequenceMatcher(None, autojunk=False)
        scorer.set_seq2(needle)
        min_overlap = max(2, len(needle) // 3)
        scored = []
        hints = OrderedDict()
        for index, para in enumerate(self._all_paragraphs()):
            text = text_of(para)
            if not text:
                continue
            if needle in text:
                if not self.include_toc and self._is_toc_paragraph(para):
                    hints.setdefault('toc', 'Only found in a table-of-contents paragraph (#%d), which is skipped '
                                            'by default; pass include_toc=True to edit it.' % index)
                elif startswith:
                    hints.setdefault('startswith', 'Found inside paragraph #%d but not at its start; drop '
                                                   'startswith=True.' % index)
            else:
                pos = unicodedata.normalize('NFKC', text).find(norm_needle)
                if pos >= 0:
                    hints.setdefault('width', 'The document has this text with different full-width/half-width '
                                              'characters or spaces (paragraph #%d): "%s"; copy the exact '
                                              'characters from the document.' % (index, _snippet(text, pos, len(needle), 0)))
                if 'revision' not in hints and needle in self._visible_text(para, own):
                    hints['revision'] = ('The text is inside an existing tracked change, or was inserted earlier in '
                                         'this session (paragraph #%d); text inside revisions cannot be matched. Plan '
                                         'edits against the original text, or cover all changes to one sentence with '
                                         'a single replace_text_smart call.' % index)
            if len(remainder) >= 2 and remainder in text and 'numbering' not in hints:
                hints['numbering'] = ('Leading "%s" may be automatic list numbering, which is not part of the '
                                      'paragraph text; try "%s".' % (numbering.group().strip(), remainder))
            matcher.set_seq1(text)
            best = matcher.find_longest_match(0, len(text), 0, len(needle))
            if best.size >= min_overlap:
                start = max(0, best.a - best.b)
                scorer.set_seq1(text[start:start + len(needle)])
                score = scorer.ratio()
                if score >= 0.5:
                    scored.append((score, index, _snippet(text, start, len(needle))))
        scored.sort(key=lambda item: -item[0])
        suggestions = [{'paragraph': index, 'text': snippet, 'score': round(score, 2)}
                       for score, index, snippet in scored[:3]]
        lines = ['%s not found: %s' % (label, needle)]
        if suggestions:
            lines.append('Closest paragraph text:')
            lines.extend('  [#%d] %s' % (item['paragraph'], item['text']) for item in suggestions)
        if hints:
            lines.append('Hints:')
            lines.extend('  - ' + hint for hint in hints.values())
        return TextNotFoundError('\n'.join(lines), suggestions, list(hints.values()))

    # ------------------------------------------------------- text operations

    def replace_text(self, find, replace, replace_all=False, occurrence=None):
        """Replace exact text with fine-grained revision marks. Returns places changed."""
        replace = self._clean_new_text(replace)
        return self._apply_text_edit(
            find, replace_all, occurrence,
            check=lambda para, start: self._range_problem(para, start, find, replace),
            apply=lambda para, start: self._edit_range(para, start, find, replace, 'smart'),
        )

    def replace_text_smart(self, old_text, new_text, replace_all=False, occurrence=None):
        """Replace text using a fine-grained diff (same as replace_text)."""
        return self.replace_text(old_text, new_text, replace_all=replace_all, occurrence=occurrence)

    def replace_text_full(self, find, replace, replace_all=False, occurrence=None):
        """Replace exact text as one whole deletion and one whole insertion."""
        replace = self._clean_new_text(replace)
        return self._apply_text_edit(
            find, replace_all, occurrence,
            check=lambda para, start: self._range_problem(para, start, find, replace),
            apply=lambda para, start: self._edit_range(para, start, find, replace, 'full'),
        )

    def insert_text_after(self, anchor, insert, replace_all=False, occurrence=None):
        """Insert tracked text right after anchor. Returns places changed."""
        insert = self._clean_new_text(insert)
        if not insert:
            raise ValueError('insert text must not be empty')
        size = len(anchor) if isinstance(anchor, str) else 0
        return self._apply_text_edit(
            anchor, replace_all, occurrence,
            check=lambda para, start: self._insert_problem(para, start + size, insert),
            apply=lambda para, start: self._insert_text_at(para, start + size, insert),
        )

    # -------------------------------------------------- paragraph operations

    def _delete_runs_in_place(self, para, runs):
        """Wrap runs in w:del where they stand (adjacent siblings share one w:del)."""
        groups = []
        for run in runs:
            if groups and groups[-1][-1].getnext() is run:
                groups[-1].append(run)
            else:
                groups.append([run])
        deletions = []
        for group in groups:
            parent = group[0].getparent()
            index = parent.index(group[0])
            deletion = self._make_del(group)
            for run in group:
                parent.remove(run)
            parent.insert(index, deletion)
            deletions.append(deletion)
        return deletions

    def replace_paragraph(self, find, new_text, startswith=False, bold=None, occurrence=None):
        """Rewrite a whole paragraph in place: old runs become w:del, new text one w:ins.

        Bookmarks, comment anchors and hyperlink containers stay where they are."""
        new_text = self._clean_new_text(new_text)
        para = self._select_paragraph(find, startswith, occurrence)
        if self._has_revisions(para):
            raise ValueError("Paragraph contains existing revision marks; refusing full paragraph replacement")

        segments = self._split_newline_segments(new_text)
        text_run = self._first_text_run(para)
        rpr = self._clone_rpr(text_run) if text_run is not None else None
        removable = [run for run in self._own_runs(para) if not self._is_annotation_run(run)]
        deletions = self._delete_runs_in_place(para, removable)
        if segments[0]:
            ins = self._make_ins(self._make_runs_from_text(segments[0], rpr, bold=bold))
            if deletions:
                top = deletions[-1]
                while top.getparent() is not para:
                    top = top.getparent()
                top.addnext(ins)
            else:
                ppr = para.find(PPR_TAG)
                para.insert(para.index(ppr) + 1 if ppr is not None else 0, ins)
        if len(segments) > 1:
            self._split_paragraph(para, len(para), segments[1:], rpr, bold=bold)
        return 1

    def delete_paragraph(self, find, startswith=False, occurrence=None):
        """Delete entire paragraph as a revision: deleted runs + deleted paragraph mark.

        Word removes the paragraph entirely when the change is accepted,
        so no empty line is left behind.
        """
        para = self._select_paragraph(find, startswith, occurrence)
        if self._has_revisions(para):
            raise ValueError("Paragraph contains existing revision marks; refusing deletion")
        removable = [run for run in self._own_runs(para) if not self._is_annotation_run(run)]
        self._delete_runs_in_place(para, removable)
        rpr = self._ensure_ppr_rpr(self._ensure_ppr(para))
        rpr.insert(0, self._revision_marker('del'))
        return 1

    def _insert_paragraphs_after(self, para, items):
        """Insert tracked paragraphs after para; items are (text, bold) pairs."""
        source_ppr = para.find(PPR_TAG)
        if source_ppr is not None and source_ppr.find(wtag('sectPr')) is not None:
            warnings.warn('The anchor paragraph ends a section; the inserted paragraph(s) start the next section.',
                          UserWarning, stacklevel=3)
        ppr = self._clean_ppr(source_ppr, keep_numbering=False)
        text_run = self._first_text_run(para)
        rpr = self._clone_rpr(text_run) if text_run is not None else None
        parent = para.getparent()
        index = parent.index(para) + 1
        created = []
        for text, bold in items:
            new_p = self._make_inserted_paragraph(text, ppr, rpr, bold=bold)
            parent.insert(index, new_p)
            index += 1
            created.append(new_p)
        return created

    def insert_paragraphs_after_para(self, para, text, bold=None):
        """Insert tracked paragraphs after para; each newline is a paragraph break."""
        text = self._clean_new_text(text)
        return self._insert_paragraphs_after(para, [(segment, bold) for segment in self._split_newline_segments(text)])

    def insert_after(self, find, new_text, startswith=False, bold=None, occurrence=None):
        new_text = self._clean_new_text(new_text)
        para = self._select_paragraph(find, startswith, occurrence)
        self.insert_paragraphs_after_para(para, new_text, bold=bold)
        return 1

    def insert_multiple_after(self, find, texts, startswith=False, bolds=None, occurrence=None):
        texts = [self._clean_new_text(text) for text in texts]
        if bolds is None:
            bolds = [None] * len(texts)
        elif len(bolds) != len(texts):
            raise ValueError("bolds must have same length as texts")
        para = self._select_paragraph(find, startswith, occurrence)
        items = [(segment, bold) for text, bold in zip(texts, bolds) for segment in self._split_newline_segments(text)]
        self._insert_paragraphs_after(para, items)
        return 1

    # ------------------------------------------------------------------ batch

    def _preview_matches(self, method_name, kwargs, limit=20):
        index_of = {para: i for i, para in enumerate(self._all_paragraphs())}
        found = []
        try:
            needle_key = self.TEXT_OP_NEEDLES.get(method_name)
            if needle_key:
                needle = kwargs.get(needle_key)
                self._validate_find(needle)
                for para, start in self._text_matches(needle)[:limit]:
                    found.append({'paragraph': index_of.get(para, -1), 'offset': start,
                                  'context': self._visible_context(para, start, len(needle))})
            else:
                find = kwargs.get('find')
                for para in self._paragraph_matches(find, bool(kwargs.get('startswith')))[:limit]:
                    text = self._own_text(para)
                    found.append({'paragraph': index_of.get(para, -1), 'context': _snippet(text, 0, min(len(text), 40), 0)})
        except (ValueError, TypeError):
            pass
        return found

    def apply_edits(self, edits, dry_run=False):
        """Apply a list of {"op": ..., **method kwargs} edits; one result dict per edit.

        A failed edit leaves the document unchanged for that edit and does not stop
        the others. With dry_run=True each result also lists the matches it found."""
        results = []
        for index, edit in enumerate(edits):
            item = {'index': index, 'op': edit.get('op') if isinstance(edit, dict) else None,
                    'status': 'failed', 'count': 0}
            try:
                if not isinstance(edit, dict) or not edit.get('op'):
                    raise ValueError('each edit must be an object with an "op" field')
                method_name = self.BATCH_OPS.get(edit['op'])
                if method_name is None:
                    raise ValueError('unknown op %r (expected one of: %s)' % (edit['op'], ', '.join(sorted(self.BATCH_OPS))))
                kwargs = {key: value for key, value in edit.items() if key != 'op'}
                if dry_run:
                    item['matches'] = self._preview_matches(method_name, kwargs)
                with warnings.catch_warnings(record=True) as caught:
                    warnings.simplefilter('always')
                    result = getattr(self, method_name)(**kwargs)
                item['status'] = 'applied'
                item['count'] = result if isinstance(result, int) and not isinstance(result, bool) else 1
                if caught:
                    item['warnings'] = [str(w.message) for w in caught]
            except TextNotFoundError as exc:
                item['error'] = str(exc)
                item['suggestions'] = exc.suggestions
            except (ValueError, TypeError, RuntimeError) as exc:
                item['error'] = str(exc)
            results.append(item)
        return results

    # ------------------------------------------------------------------- save

    def _ensure_settings(self):
        """Ensure settings.xml exists and trackRevisions is enabled."""
        settings_path = os.path.join(self.temp_dir, 'word', 'settings.xml')

        if os.path.exists(settings_path):
            tree = etree.parse(settings_path)
            root = tree.getroot()
        else:
            nsmap = {'w': W}
            root = etree.Element(wtag('settings'), nsmap=nsmap)
            tree = etree.ElementTree(root)

            ct_path = os.path.join(self.temp_dir, '[Content_Types].xml')
            ct_tree = etree.parse(ct_path)
            ct_root = ct_tree.getroot()
            exists = any(override.get('PartName') == '/word/settings.xml'
                         for override in ct_root.findall('{%s}Override' % CT))
            if not exists:
                override = etree.SubElement(ct_root, '{%s}Override' % CT)
                override.set('PartName', '/word/settings.xml')
                override.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml')
                ct_tree.write(ct_path, xml_declaration=True, encoding='UTF-8', standalone=True)

            rels_path = os.path.join(self.temp_dir, 'word', '_rels', 'document.xml.rels')
            if os.path.exists(rels_path):
                rels_tree = etree.parse(rels_path)
                rels_root = rels_tree.getroot()
            else:
                os.makedirs(os.path.dirname(rels_path), exist_ok=True)
                rels_root = etree.Element('{%s}Relationships' % RELS)
                rels_tree = etree.ElementTree(rels_root)

            exists = any('settings' in rel.get('Type', '') for rel in rels_root.findall('{%s}Relationship' % RELS))
            if not exists:
                existing_ids = [r.get('Id') for r in rels_root.findall('{%s}Relationship' % RELS)]
                max_num = 0
                for eid in existing_ids:
                    if eid and eid.startswith('rId'):
                        try:
                            max_num = max(max_num, int(eid[3:]))
                        except ValueError:
                            pass
                rel = etree.SubElement(rels_root, '{%s}Relationship' % RELS)
                rel.set('Id', 'rId%d' % (max_num + 1))
                rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings')
                rel.set('Target', 'settings.xml')
                rels_tree.write(rels_path, xml_declaration=True, encoding='UTF-8', standalone=True)

        track = root.find(wtag('trackRevisions'))
        if track is None:
            track = etree.Element(wtag('trackRevisions'))
            for index, child in enumerate(root):
                if isinstance(child.tag, str) and _local(child.tag) in SETTINGS_AFTER_TRACK_REVISIONS:
                    root.insert(index, track)
                    break
            else:
                root.append(track)
        track.set(wtag('val'), 'true')
        tree.write(settings_path, xml_declaration=True, encoding='UTF-8', standalone=True)

    def _bad_deleted_text_count(self, root):
        """w:t inside a deleted run of the same story (deleted text must be w:delText)."""
        count = 0
        for t in root.iter(T_TAG):
            for anc in t.iterancestors():
                if anc.tag == TXBX_TAG:
                    break
                if anc.tag == wtag('del'):
                    count += 1
                    break
        return count

    def _revision_counts(self):
        ins_count = del_count = bad_del_text = del_text_count = 0
        for part in self.part_trees:
            root = part['root']
            ins_count += sum(1 for _ in root.iter(wtag('ins')))
            del_count += sum(1 for _ in root.iter(wtag('del')))
            bad_del_text += self._bad_deleted_text_count(root)
            del_text_count += sum(1 for _ in root.iter(wtag('delText')))
        return ins_count, del_count, bad_del_text, del_text_count

    def _zip_entries(self):
        """[Content_Types].xml first, then the original zip order, then new files."""
        on_disk = []
        for root_dir, dirs, files in os.walk(self.temp_dir):
            dirs.sort()
            for name in sorted(files):
                full_path = os.path.join(root_dir, name)
                on_disk.append(os.path.relpath(full_path, self.temp_dir).replace(os.sep, '/'))
        present = set(on_disk)
        ordered = ['[Content_Types].xml'] if '[Content_Types].xml' in present else []
        for name in self._zip_order + on_disk:
            if name in present and name not in ordered:
                ordered.append(name)
        return ordered

    def save(self, output_path):
        """Save modified docx to output_path (atomic replace). Returns revision mark count."""
        if not self.temp_dir:
            raise RuntimeError('This DocxReviser was already saved or closed; create a new instance to edit again')
        self._ensure_settings()

        for part in self.part_trees:
            part['tree'].write(part['path'], xml_declaration=True, encoding='UTF-8', standalone=True)

        ins_count, del_count, bad_del_text, _ = self._revision_counts()
        total_revs = ins_count + del_count
        if bad_del_text:
            warnings.warn(
                f"Found {bad_del_text} w:t nodes inside w:del. Deleted revision text should use w:delText.",
                UserWarning
            )
        if total_revs == 0:
            warnings.warn(
                "No revision marks (w:ins/w:del) found in selected document parts. "
                "Track changes may not be visible in Word.",
                UserWarning
            )

        output_path = os.path.abspath(output_path)
        out_dir = os.path.dirname(output_path)
        os.makedirs(out_dir, exist_ok=True)
        fd, tmp_path = tempfile.mkstemp(prefix='.%s.' % os.path.basename(output_path), suffix='.tmp', dir=out_dir)
        os.close(fd)
        try:
            with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for arcname in self._zip_entries():
                    zf.write(os.path.join(self.temp_dir, *arcname.split('/')), arcname)
            if os.path.exists(output_path):
                shutil.copymode(output_path, tmp_path)
            else:
                umask = os.umask(0)
                os.umask(umask)
                os.chmod(tmp_path, 0o666 & ~umask)
            os.replace(tmp_path, output_path)
        except BaseException:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
            raise

        self.close()
        return total_revs

    def __del__(self):
        try:
            self.close()
        except Exception:
            pass


USAGE = """Usage: word_reviser.py [--author NAME] [--parts PARTS] [--include-toc] [--merge-short-equal N] <input.docx> <output.docx> <command> [args...]
Parts: document,headers,footers,footnotes,endnotes,comments or all
Commands:
  replace <find> <replace> [--replace-all|--all] [--occurrence N]
  smart_replace <old_text> <new_text> [--replace-all|--all] [--occurrence N]
  replace_full <find> <replace> [--replace-all|--all] [--occurrence N]
  insert_text_after <anchor> <insert> [--replace-all|--all] [--occurrence N]
  replace_para <find> <new_text> [--startswith] [--bold] [--occurrence N]
  delete_para <find> [--startswith] [--occurrence N]
  insert_after <find> <new_text> [--startswith] [--bold] [--occurrence N]
  insert_multiple_after <find> <text1> [text2 ...] [--startswith] [--bold] [--bolds true,false] [--occurrence N]
  apply <edits.json> [--dry-run] [--keep-going]
Exit codes: 0 ok; 1 error (nothing written); 2 apply --keep-going saved with some failed edits."""


class _UsageError(Exception):
    pass


def _pop_option(args, flag, convert=str):
    if flag not in args:
        return None
    index = args.index(flag)
    if index + 1 >= len(args):
        raise _UsageError('Missing value for %s' % flag)
    value = args[index + 1]
    del args[index:index + 2]
    try:
        return convert(value)
    except ValueError:
        raise _UsageError('Invalid value for %s: %s' % (flag, value))


def _pop_flag(args, *flags):
    found = False
    for flag in flags:
        while flag in args:
            args.remove(flag)
            found = True
    return found


def _run_apply(reviser, positional, output_path, dry_run, keep_going):
    if len(positional) != 1:
        print(USAGE, file=sys.stderr)
        return 1
    try:
        with open(positional[0], encoding='utf-8') as fh:
            edits = json.load(fh)
    except (OSError, ValueError) as exc:
        print('Error: cannot read edits file: %s' % exc, file=sys.stderr)
        return 1
    if isinstance(edits, dict):
        edits = edits.get('edits')
    if not isinstance(edits, list):
        print('Error: edits file must be a JSON list of edits (or {"edits": [...]})', file=sys.stderr)
        return 1

    results = reviser.apply_edits(edits, dry_run=dry_run)
    ok = all(item['status'] == 'applied' for item in results)
    report = {'ok': ok, 'dry_run': dry_run, 'output': None, 'revisions': None, 'results': results}
    if dry_run:
        ins_count, del_count, _, _ = reviser._revision_counts()
        report['revisions'] = ins_count + del_count
        code = 0 if ok else 1
    elif ok or keep_going:
        try:
            report['revisions'] = reviser.save(output_path)
            report['output'] = output_path
            code = 0 if ok else 2
        except OSError as exc:
            report['ok'] = False
            report['error'] = 'save failed: %s' % exc
            code = 1
    else:
        code = 1
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return code


def main(argv=None):
    args = list(sys.argv[1:] if argv is None else argv)
    if '-h' in args or '--help' in args:
        print(USAGE)
        return 0
    try:
        author = _pop_option(args, '--author')
        raw_parts = _pop_option(args, '--parts')
        merge_short_equal = _pop_option(args, '--merge-short-equal', int)
        occurrence = _pop_option(args, '--occurrence', int)
        raw_bolds = _pop_option(args, '--bolds')
        include_toc = _pop_flag(args, '--include-toc')
        replace_all = _pop_flag(args, '--replace-all', '--all')
        startswith = _pop_flag(args, '--startswith')
        bold_flag = _pop_flag(args, '--bold')
        dry_run = _pop_flag(args, '--dry-run')
        keep_going = _pop_flag(args, '--keep-going')
    except _UsageError as exc:
        print(exc, file=sys.stderr)
        print(USAGE, file=sys.stderr)
        return 1
    if len(args) < 3:
        print(USAGE, file=sys.stderr)
        return 1

    input_path, output_path, command, positional = args[0], args[1], args[2], args[3:]
    options = {'author': author or 'Agent', 'include_toc': include_toc}
    if raw_parts is not None:
        options['parts'] = 'all' if raw_parts == 'all' else raw_parts.split(',')
    if merge_short_equal is not None:
        options['merge_short_equal'] = merge_short_equal
    try:
        reviser = DocxReviser(input_path, **options)
    except (OSError, ValueError, zipfile.BadZipFile) as exc:
        print('Error: %s' % exc, file=sys.stderr)
        return 1

    bold = True if bold_flag else None

    def need(count, at_least=False):
        if (len(positional) < count) if at_least else (len(positional) != count):
            raise _UsageError('Wrong number of arguments for %s' % command)

    try:
        if command == 'apply':
            return _run_apply(reviser, positional, output_path, dry_run, keep_going)
        if command in ('replace', 'smart_replace', 'replace_full', 'insert_text_after'):
            need(2)
            method = {
                'replace': reviser.replace_text,
                'smart_replace': reviser.replace_text_smart,
                'replace_full': reviser.replace_text_full,
                'insert_text_after': reviser.insert_text_after,
            }[command]
            method(positional[0], positional[1], replace_all=replace_all, occurrence=occurrence)
        elif command in ('replace_para', 'replace_paragraph'):
            need(2)
            reviser.replace_paragraph(positional[0], positional[1], startswith=startswith, bold=bold,
                                      occurrence=occurrence)
        elif command in ('delete_para', 'delete_paragraph'):
            need(1)
            reviser.delete_paragraph(positional[0], startswith=startswith, occurrence=occurrence)
        elif command == 'insert_after':
            need(2)
            reviser.insert_after(positional[0], positional[1], startswith=startswith, bold=bold,
                                 occurrence=occurrence)
        elif command == 'insert_multiple_after':
            need(2, at_least=True)
            texts = positional[1:]
            bolds = None
            if raw_bolds is not None:
                bolds = [item.strip().lower() in ('1', 'true', 'yes', 'y') for item in raw_bolds.split(',')]
            elif bold_flag:
                bolds = [True] * len(texts)
            reviser.insert_multiple_after(positional[0], texts, startswith=startswith, bolds=bolds,
                                          occurrence=occurrence)
        else:
            print('Unknown command: %s' % command, file=sys.stderr)
            print(USAGE, file=sys.stderr)
            return 1
        total = reviser.save(output_path)
    except _UsageError as exc:
        print(exc, file=sys.stderr)
        print(USAGE, file=sys.stderr)
        return 1
    except (ValueError, TypeError, RuntimeError, OSError) as exc:
        print('Error: %s' % exc, file=sys.stderr)
        return 1
    finally:
        reviser.close()
    print(f"Saved to {output_path} ({total} revision marks)")
    return 0


if __name__ == '__main__':
    sys.exit(main())
