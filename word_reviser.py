#!/usr/bin/env python3
"""
Word docx revision tool using lxml for robust XML manipulation.
Generates Word Track Changes (revision mode) with format preservation.
"""

from difflib import SequenceMatcher
import glob
import os
import shutil
import tempfile
import warnings
import zipfile
from copy import deepcopy
from datetime import datetime, timezone
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
RELS = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'

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


def wtag(tag):
    return '{%s}%s' % (W, tag)


def _is_under_revision(elem):
    parent = elem.getparent()
    revision_tags = {wtag(tag) for tag in REVISION_TAGS}
    while parent is not None:
        if parent.tag in revision_tags:
            return True
        parent = parent.getparent()
    return False


class DocxReviser:
    def __init__(self, input_path, author='Agent', parts=None):
        self.input_path = input_path
        self.author = author
        self.rev_id = 1000
        self.revision_date = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace('+00:00', 'Z')
        self.part_names = self._normalize_parts(parts)
        self.part_trees = []

        self.temp_dir = tempfile.mkdtemp()
        with zipfile.ZipFile(input_path, 'r') as z:
            z.extractall(self.temp_dir)

        self._load_parts()

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

    def _para_text(self, para):
        """Get editable paragraph text, excluding existing revision text."""
        texts = []
        for run in para.xpath('.//w:r', namespaces={'w': W}):
            if _is_under_revision(run):
                continue
            t = run.find(wtag('t'))
            if t is not None:
                texts.append(t.text or '')
        return ''.join(texts)

    def _all_paragraphs(self):
        """Get all paragraphs in selected parts, including inside tables/text boxes."""
        paragraphs = []
        for part in self.part_trees:
            paragraphs.extend(part['root'].findall('.//' + wtag('p')))
        return paragraphs

    def _find_para(self, text, startswith=False):
        """Find first paragraph containing or starting with text."""
        for p in self._all_paragraphs():
            pt = self._para_text(p)
            if startswith and pt.startswith(text):
                return p
            if not startswith and text in pt:
                return p
        return None

    def _clone_ppr(self, para):
        ppr = para.find(wtag('pPr'))
        return deepcopy(ppr) if ppr is not None else None

    def _clone_rpr(self, run):
        rpr = run.find(wtag('rPr'))
        return deepcopy(rpr) if rpr is not None else None

    def _make_run(self, text, rpr_elem=None, bold=False, deleted=False):
        r = etree.Element(wtag('r'))
        if rpr_elem is not None:
            r.append(deepcopy(rpr_elem))
        if bold:
            rpr = r.find(wtag('rPr'))
            if rpr is None:
                rpr = etree.SubElement(r, wtag('rPr'))
            if rpr.find(wtag('b')) is None:
                etree.SubElement(rpr, wtag('b'))
        t = etree.SubElement(r, wtag('delText' if deleted else 't'))
        t.text = text
        if text and (text.startswith(' ') or text.endswith(' ')):
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        return r

    def _make_br_run(self, rpr_elem=None, bold=False):
        r = etree.Element(wtag('r'))
        if rpr_elem is not None:
            r.append(deepcopy(rpr_elem))
        if bold:
            rpr = r.find(wtag('rPr'))
            if rpr is None:
                rpr = etree.SubElement(r, wtag('rPr'))
            if rpr.find(wtag('b')) is None:
                etree.SubElement(rpr, wtag('b'))
        etree.SubElement(r, wtag('br'))
        return r

    def _make_runs_from_text(self, text, rpr_elem=None, bold=False, deleted=False):
        if '\n' not in text:
            return [self._make_run(text, rpr_elem, bold=bold, deleted=deleted)]

        parts = text.split('\n')
        runs = []
        for i, part in enumerate(parts):
            if part:
                runs.append(self._make_run(part, rpr_elem, bold=bold, deleted=deleted))
            if i < len(parts) - 1:
                runs.append(self._make_br_run(rpr_elem, bold=bold))
        return runs

    def _make_ins(self, children, rev_id=None):
        ins = etree.Element(wtag('ins'))
        ins.set(wtag('id'), str(rev_id if rev_id is not None else self.rev_id))
        ins.set(wtag('author'), self.author)
        ins.set(wtag('date'), self.revision_date)
        for child in children:
            ins.append(child)
        return ins

    def _convert_run_to_deleted(self, run):
        run_copy = deepcopy(run)
        for t in run_copy.findall('.//' + wtag('t')):
            t.tag = wtag('delText')
        return run_copy

    def _make_del(self, children, rev_id=None):
        d = etree.Element(wtag('del'))
        d.set(wtag('id'), str(rev_id if rev_id is not None else self.rev_id))
        d.set(wtag('author'), self.author)
        d.set(wtag('date'), self.revision_date)
        for child in children:
            d.append(self._convert_run_to_deleted(child))
        return d

    def _next_rev_id(self):
        rid = self.rev_id
        self.rev_id += 1
        return rid

    def _run_entries(self, para):
        """Return editable text runs in a paragraph, excluding existing revisions."""
        entries = []
        pos = 0
        for run in para.xpath('.//w:r', namespaces={'w': W}):
            if _is_under_revision(run):
                continue
            t = run.find(wtag('t'))
            text = t.text if t is not None and t.text is not None else ''
            entries.append({
                'run': run,
                't': t,
                'start': pos,
                'end': pos + len(text),
                'text': text,
            })
            pos += len(text)
        return entries

    def _runs_for_text_range(self, entries, start_idx, end_idx):
        """Copy editable runs covering a paragraph text range."""
        runs = []
        for entry in entries:
            if entry['t'] is None:
                continue
            overlap_start = max(start_idx, entry['start'])
            overlap_end = min(end_idx, entry['end'])
            if overlap_start >= overlap_end:
                continue

            run_copy = deepcopy(entry['run'])
            t_copy = run_copy.find(wtag('t'))
            if t_copy is None:
                continue
            text_start = overlap_start - entry['start']
            text_end = overlap_end - entry['start']
            t_copy.text = entry['text'][text_start:text_end]
            runs.append(run_copy)
        return runs

    def _rpr_at_or_before(self, entries, offset):
        fallback = None
        for entry in entries:
            if entry['t'] is None:
                continue
            fallback = entry['run']
            if entry['start'] <= offset <= entry['end']:
                return self._clone_rpr(entry['run'])
        if fallback is not None:
            return self._clone_rpr(fallback)
        return None

    def _replace_range_smart(self, para, start_idx, old_text, new_text):
        """Replace a range using a text diff so unchanged text stays unmarked."""
        end_idx = start_idx + len(old_text)
        entries = self._run_entries(para)
        if not entries:
            return False

        start_entry = end_entry = None
        for i, entry in enumerate(entries):
            if entry['t'] is not None and entry['start'] <= start_idx < entry['end']:
                start_entry = i
            if entry['t'] is not None and entry['start'] < end_idx <= entry['end']:
                end_entry = i
                break
            if entry['t'] is not None and end_idx == entry['end'] and entry['start'] < entry['end']:
                end_entry = i
                break

        if start_entry is None or end_entry is None:
            return False

        first = entries[start_entry]
        last = entries[end_entry]
        parent = first['run'].getparent()
        insert_at = list(parent).index(first['run'])

        before = self._runs_for_text_range(entries, first['start'], start_idx)
        after = self._runs_for_text_range(entries, end_idx, last['end'])

        for i in range(end_entry, start_entry - 1, -1):
            run = entries[i]['run']
            run.getparent().remove(run)

        for run in before:
            parent.insert(insert_at, run)
            insert_at += 1

        matcher = SequenceMatcher(None, old_text, new_text, autojunk=False)
        for tag, old_start, old_end, new_start, new_end in matcher.get_opcodes():
            abs_old_start = start_idx + old_start
            abs_old_end = start_idx + old_end
            inserted_text = new_text[new_start:new_end]

            if tag == 'equal':
                for run in self._runs_for_text_range(entries, abs_old_start, abs_old_end):
                    parent.insert(insert_at, run)
                    insert_at += 1
            elif tag == 'delete':
                del_runs = self._runs_for_text_range(entries, abs_old_start, abs_old_end)
                if del_runs:
                    parent.insert(insert_at, self._make_del(del_runs, self._next_rev_id()))
                    insert_at += 1
            elif tag == 'insert':
                rpr = self._rpr_at_or_before(entries, abs_old_start)
                ins_runs = self._make_runs_from_text(inserted_text, rpr)
                parent.insert(insert_at, self._make_ins(ins_runs, self._next_rev_id()))
                insert_at += 1
            elif tag == 'replace':
                del_runs = self._runs_for_text_range(entries, abs_old_start, abs_old_end)
                if del_runs:
                    parent.insert(insert_at, self._make_del(del_runs, self._next_rev_id()))
                    insert_at += 1
                if inserted_text:
                    rpr = self._rpr_at_or_before(entries, abs_old_start)
                    ins_runs = self._make_runs_from_text(inserted_text, rpr)
                    parent.insert(insert_at, self._make_ins(ins_runs, self._next_rev_id()))
                    insert_at += 1

        for run in after:
            parent.insert(insert_at, run)
            insert_at += 1

        return True

    def _replace_range_full(self, para, start_idx, old_text, new_text):
        """Replace a range as one whole deletion and one whole insertion."""
        end_idx = start_idx + len(old_text)
        entries = self._run_entries(para)
        if not entries:
            return False

        start_entry = end_entry = None
        for i, entry in enumerate(entries):
            if entry['t'] is not None and entry['start'] <= start_idx < entry['end']:
                start_entry = i
            if entry['t'] is not None and entry['start'] < end_idx <= entry['end']:
                end_entry = i
                break
            if entry['t'] is not None and end_idx == entry['end'] and entry['start'] < entry['end']:
                end_entry = i
                break

        if start_entry is None or end_entry is None:
            return False

        first = entries[start_entry]
        last = entries[end_entry]
        parent = first['run'].getparent()
        insert_at = list(parent).index(first['run'])

        before = self._runs_for_text_range(entries, first['start'], start_idx)
        after = self._runs_for_text_range(entries, end_idx, last['end'])

        # Collect runs for the whole old_text range, preserving original formatting.
        old_runs = self._runs_for_text_range(entries, start_idx, end_idx)

        for i in range(end_entry, start_entry - 1, -1):
            run = entries[i]['run']
            run.getparent().remove(run)

        for run in before:
            parent.insert(insert_at, run)
            insert_at += 1

        if old_runs:
            parent.insert(insert_at, self._make_del(old_runs, self._next_rev_id()))
            insert_at += 1

        if new_text:
            rpr = self._rpr_at_or_before(entries, start_idx)
            ins_runs = self._make_runs_from_text(new_text, rpr)
            parent.insert(insert_at, self._make_ins(ins_runs, self._next_rev_id()))
            insert_at += 1

        for run in after:
            parent.insert(insert_at, run)
            insert_at += 1

        return True

    def _insert_text_at(self, para, offset, text):
        entries = self._run_entries(para)
        if not entries:
            return False

        for entry in entries:
            if entry['t'] is None:
                continue
            if entry['start'] <= offset <= entry['end']:
                rpr = self._clone_rpr(entry['run'])
                before = entry['text'][:offset - entry['start']]
                after = entry['text'][offset - entry['start']:]
                parent = entry['run'].getparent()
                insert_at = list(parent).index(entry['run'])
                parent.remove(entry['run'])

                if before:
                    parent.insert(insert_at, self._make_run(before, rpr))
                    insert_at += 1

                ins_runs = self._make_runs_from_text(text, rpr)
                parent.insert(insert_at, self._make_ins(ins_runs, self._next_rev_id()))
                insert_at += 1

                if after:
                    parent.insert(insert_at, self._make_run(after, rpr))

                return True

        return False

    def replace_text(self, find, replace, replace_all=False):
        """Replace exact text with fine-grained revision marks."""
        found_any = False
        for para in self._all_paragraphs():
            while True:
                text = self._para_text(para)
                if find not in text:
                    break
                start_idx = text.index(find)
                result = self._replace_range_smart(para, start_idx, find, replace)
                if not result:
                    break
                found_any = True
                if not replace_all:
                    return True
                if find in replace:
                    break
        if not found_any:
            raise ValueError(f"Text not found: {find}")
        return True

    def replace_text_smart(self, old_text, new_text, replace_all=False):
        """Replace text using a fine-grained diff."""
        found_any = False
        for para in self._all_paragraphs():
            while True:
                text = self._para_text(para)
                if old_text not in text:
                    break
                start_idx = text.index(old_text)
                result = self._replace_range_smart(para, start_idx, old_text, new_text)
                if not result:
                    break
                found_any = True
                if not replace_all or old_text in new_text:
                    return True
        if not found_any:
            raise ValueError(f"Text not found: {old_text}")
        return True

    def replace_text_full(self, find, replace, replace_all=False):
        """Replace exact text as one whole deletion and one whole insertion."""
        found_any = False
        for para in self._all_paragraphs():
            while True:
                text = self._para_text(para)
                if find not in text:
                    break
                start_idx = text.index(find)
                result = self._replace_range_full(para, start_idx, find, replace)
                if not result:
                    break
                found_any = True
                if not replace_all:
                    return True
                if find in replace:
                    break
        if not found_any:
            raise ValueError(f"Text not found: {find}")
        return True

    def insert_text_after(self, anchor, insert, replace_all=False):
        found_any = False
        for para in self._all_paragraphs():
            while True:
                text = self._para_text(para)
                if anchor not in text:
                    break
                offset = text.index(anchor) + len(anchor)
                result = self._insert_text_at(para, offset, insert)
                if not result:
                    break
                found_any = True
                if not replace_all:
                    return True
                if anchor in insert:
                    break
        if not found_any:
            raise ValueError(f"Text not found: {anchor}")
        return True

    def replace_paragraph(self, find, new_text, startswith=False, bold=False):
        """Replace entire paragraph: del of old + ins of new, keeping pPr."""
        para = self._find_para(find, startswith)
        if para is None:
            raise ValueError(f"Paragraph not found: {find}")
        if para.xpath('.//w:ins | .//w:del | .//w:moveFrom | .//w:moveTo', namespaces={'w': W}):
            raise ValueError("Paragraph contains existing revision marks; refusing full paragraph replacement")

        ppr = self._clone_ppr(para)
        rpr = None
        first_run = None
        for run in para.findall('.//' + wtag('r')):
            if not _is_under_revision(run):
                first_run = run
                break
        if first_run is not None:
            rpr = self._clone_rpr(first_run)

        new_p = etree.Element(wtag('p'))
        if ppr is not None:
            new_p.append(ppr)

        old_runs = [deepcopy(r) for r in para.findall('.//' + wtag('r')) if not _is_under_revision(r)]
        if old_runs:
            new_p.append(self._make_del(old_runs, self._next_rev_id()))

        ins_runs = self._make_runs_from_text(new_text, rpr, bold=bold)
        new_p.append(self._make_ins(ins_runs, self._next_rev_id()))

        parent = para.getparent()
        idx = list(parent).index(para)
        parent.remove(para)
        parent.insert(idx, new_p)

    def delete_paragraph(self, find, startswith=False):
        """Delete entire paragraph as a revision: deleted runs + deleted paragraph mark.

        Word removes the paragraph entirely when the change is accepted,
        so no empty line is left behind.
        """
        para = self._find_para(find, startswith)
        if para is None:
            raise ValueError(f"Paragraph not found: {find}")
        if para.xpath('.//w:ins | .//w:del | .//w:moveFrom | .//w:moveTo', namespaces={'w': W}):
            raise ValueError("Paragraph contains existing revision marks; refusing deletion")

        rev_id = self._next_rev_id()
        editable_runs = [r for r in para.findall('.//' + wtag('r')) if not _is_under_revision(r)]
        old_runs = [deepcopy(r) for r in editable_runs]
        for r in editable_runs:
            r.getparent().remove(r)
        if old_runs:
            para.append(self._make_del(old_runs, rev_id))

        # Mark the paragraph mark itself as deleted (w:pPr/w:rPr/w:del).
        ppr = para.find(wtag('pPr'))
        if ppr is None:
            ppr = etree.Element(wtag('pPr'))
            para.insert(0, ppr)
        rpr = ppr.find(wtag('rPr'))
        if rpr is None:
            rpr = etree.SubElement(ppr, wtag('rPr'))
        d = etree.Element(wtag('del'))
        d.set(wtag('id'), str(rev_id))
        d.set(wtag('author'), self.author)
        d.set(wtag('date'), self.revision_date)
        rpr.insert(0, d)

    def _remove_numpr(self, ppr):
        if ppr is not None:
            numpr = ppr.find(wtag('numPr'))
            if numpr is not None:
                ppr.remove(numpr)
        return ppr

    def _first_editable_run(self, para):
        for run in para.findall('.//' + wtag('r')):
            if not _is_under_revision(run):
                return run
        return None

    def insert_after(self, find, new_text, startswith=False, bold=False):
        para = self._find_para(find, startswith)
        if para is None:
            raise ValueError(f"Paragraph not found: {find}")

        ppr = self._remove_numpr(self._clone_ppr(para))
        first_run = self._first_editable_run(para)
        rpr = self._clone_rpr(first_run) if first_run is not None else None

        new_p = etree.Element(wtag('p'))
        if ppr is not None:
            new_p.append(ppr)

        new_p.append(self._make_ins(self._make_runs_from_text(new_text, rpr, bold=bold), self._next_rev_id()))

        parent = para.getparent()
        idx = list(parent).index(para)
        parent.insert(idx + 1, new_p)

    def insert_multiple_after(self, find, texts, startswith=False, bolds=None):
        para = self._find_para(find, startswith)
        if para is None:
            raise ValueError(f"Paragraph not found: {find}")

        if bolds is None:
            bolds = [False] * len(texts)
        elif len(bolds) != len(texts):
            raise ValueError("bolds must have same length as texts")

        parent = para.getparent()
        idx = list(parent).index(para) + 1
        first_run = self._first_editable_run(para)
        rpr = self._clone_rpr(first_run) if first_run is not None else None

        for text, bold in zip(texts, bolds):
            ppr = self._remove_numpr(self._clone_ppr(para))
            new_p = etree.Element(wtag('p'))
            if ppr is not None:
                new_p.append(ppr)
            new_p.append(self._make_ins(self._make_runs_from_text(text, rpr, bold=bold), self._next_rev_id()))
            parent.insert(idx, new_p)
            idx += 1

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
            track = etree.SubElement(root, wtag('trackRevisions'))
        track.set(wtag('val'), 'true')
        tree.write(settings_path, xml_declaration=True, encoding='UTF-8', standalone=True)

    def _revision_counts(self):
        ins_count = del_count = bad_del_text = del_text_count = 0
        for part in self.part_trees:
            root = part['root']
            ins_count += len(root.findall('.//%s' % wtag('ins')))
            del_count += len(root.findall('.//%s' % wtag('del')))
            bad_del_text += len(root.findall('.//%s//%s' % (wtag('del'), wtag('t'))))
            del_text_count += len(root.findall('.//%s//%s' % (wtag('del'), wtag('delText'))))
        return ins_count, del_count, bad_del_text, del_text_count

    def save(self, output_path):
        """Save modified docx to output_path."""
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

        os.makedirs(os.path.dirname(os.path.abspath(output_path)) or '.', exist_ok=True)
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root_dir, dirs, files in os.walk(self.temp_dir):
                for file in files:
                    full_path = os.path.join(root_dir, file)
                    arcname = os.path.relpath(full_path, self.temp_dir)
                    zf.write(full_path, arcname)

        shutil.rmtree(self.temp_dir)
        self.temp_dir = None

        return total_revs

    def __del__(self):
        if hasattr(self, 'temp_dir') and self.temp_dir and os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)


def main():
    import sys

    def usage(exit_code=1):
        print("Usage: word_reviser.py [--author NAME] [--parts PARTS] <input.docx> <output.docx> <command> [args...]")
        print("Parts: document,headers,footers,footnotes,endnotes,comments or all")
        print("Commands:")
        print("  replace <find> <replace> [--replace-all|--all]")
        print("  smart_replace <old_text> <new_text> [--replace-all|--all]")
        print("  insert_text_after <anchor> <insert> [--replace-all|--all]")
        print("  replace_para <find> <new_text> [--startswith] [--bold]")
        print("  insert_after <find> <new_text> [--startswith] [--bold]")
        print("  insert_multiple_after <find> <text1> [text2 ...] [--startswith] [--bold] [--bolds true,false]")
        sys.exit(exit_code)

    args = sys.argv[1:]
    if "-h" in args or "--help" in args:
        usage(0)

    author = "Agent"
    if "--author" in args:
        idx = args.index("--author")
        if idx + 1 >= len(args):
            print("Missing value for --author", file=sys.stderr)
            sys.exit(1)
        author = args[idx + 1]
        del args[idx:idx + 2]

    parts = None
    if "--parts" in args:
        idx = args.index("--parts")
        if idx + 1 >= len(args):
            print("Missing value for --parts", file=sys.stderr)
            sys.exit(1)
        raw_parts = args[idx + 1]
        parts = "all" if raw_parts == "all" else raw_parts.split(',')
        del args[idx:idx + 2]

    if len(args) < 3:
        usage()

    input_path = args[0]
    output_path = args[1]
    command = args[2]
    command_args = args[3:]

    reviser = DocxReviser(input_path, author=author, parts=parts)

    if command == 'replace':
        if len(command_args) < 2:
            usage()
        replace_all = '--replace-all' in command_args or '--all' in command_args
        positional = [arg for arg in command_args if arg not in ('--replace-all', '--all')]
        if len(positional) != 2:
            usage()
        reviser.replace_text(positional[0], positional[1], replace_all=replace_all)
    elif command == 'smart_replace':
        if len(command_args) < 2:
            usage()
        replace_all = '--replace-all' in command_args or '--all' in command_args
        positional = [arg for arg in command_args if arg not in ('--replace-all', '--all')]
        if len(positional) != 2:
            usage()
        reviser.replace_text_smart(positional[0], positional[1], replace_all=replace_all)
    elif command == 'insert_text_after':
        if len(command_args) < 2:
            usage()
        replace_all = '--replace-all' in command_args or '--all' in command_args
        positional = [arg for arg in command_args if arg not in ('--replace-all', '--all')]
        if len(positional) != 2:
            usage()
        reviser.insert_text_after(positional[0], positional[1], replace_all=replace_all)
    elif command in ('replace_para', 'replace_paragraph'):
        startswith = '--startswith' in command_args
        bold = '--bold' in command_args
        positional = [arg for arg in command_args if arg not in ('--startswith', '--bold')]
        if len(positional) != 2:
            usage()
        reviser.replace_paragraph(positional[0], positional[1], startswith=startswith, bold=bold)
    elif command == 'insert_after':
        startswith = '--startswith' in command_args
        bold = '--bold' in command_args
        positional = [arg for arg in command_args if arg not in ('--startswith', '--bold')]
        if len(positional) != 2:
            usage()
        reviser.insert_after(positional[0], positional[1], startswith=startswith, bold=bold)
    elif command == 'insert_multiple_after':
        startswith = '--startswith' in command_args
        bold_all = '--bold' in command_args
        bolds = None
        if '--bolds' in command_args:
            idx = command_args.index('--bolds')
            if idx + 1 >= len(command_args):
                print("Missing value for --bolds", file=sys.stderr)
                sys.exit(1)
            raw_bolds = command_args[idx + 1]
            bolds = [item.strip().lower() in ('1', 'true', 'yes', 'y') for item in raw_bolds.split(',')]
            del command_args[idx:idx + 2]
        positional = [arg for arg in command_args if arg not in ('--startswith', '--bold')]
        if len(positional) < 2:
            usage()
        find = positional[0]
        texts = positional[1:]
        if bolds is None and bold_all:
            bolds = [True] * len(texts)
        reviser.insert_multiple_after(find, texts, startswith=startswith, bolds=bolds)
    else:
        print(f"Unknown command: {command}")
        sys.exit(1)

    total = reviser.save(output_path)
    print(f"Saved to {output_path} ({total} revision marks)")


if __name__ == '__main__':
    main()
