---
name: word-revision
description: Use when the user wants changes applied to an existing .docx as Word Track Changes (修订模式、留痕修改、修订痕迹、红线、track changes, redline) — replacing, inserting or deleting text or paragraphs in a contract or other Word document while keeping its formatting and any existing revisions. Not for drafting a new document from a template (contract-draft) or for a full contract risk review (contract-review).
---

# word-revision — Word Track Changes for DOCX

Apply edits to `.docx` files as Word Track Changes without rebuilding the document. Shared text stays normal; only real insertions and deletions become revision marks.

- `<skill-dir>` below = the "Base directory for this skill" shown when this skill loads. Always use that absolute path; run the script with `python3`.
- Default revision author is `Agent` (`--author NAME` / `author=`). Revision dates are the current Beijing time with an explicit `+08:00` offset.
- `.doc` is not supported; save it as `.docx` first.

## Workflow

1. Confirm the source is `.docx` and collect the exact edits. Copy the target text from the document itself (full-width punctuation, spaces); do not retype it.
2. Put all edits in one `edits.json` and preview them:
   `python3 <skill-dir>/word_reviser.py in.docx out.docx apply edits.json --dry-run`
   Each result lists the matches it found; check they are the intended places.
3. Apply with the same command without `--dry-run`. The JSON report gives a status per edit.
   - exit 0: all edits applied and saved.
   - exit 1: at least one edit failed and nothing was written. Fix the failed edits using their `error` / `suggestions`, then rerun.
   - `--keep-going`: save the successful edits anyway; exit 2 if any failed.
4. Validate (below) and tell the user what changed.

## Choosing the operation

| Intent | batch `op` / Python method | Notes |
|---|---|---|
| Change words, dates, amounts, names | `replace` / `replace_text(find, replace)` | Character diff; shared text stays unmarked |
| Old clause → new clause | `smart_replace` / `replace_text_smart(old_text, new_text)` | Same diff; one call per sentence |
| One whole deletion + one insertion | `replace_full` / `replace_text_full(find, replace)` | |
| Add words after a phrase | `insert_text_after` / `insert_text_after(anchor, insert)` | |
| New paragraph(s) after a paragraph | `insert_after` / `insert_after(find, new_text, startswith=True)` | Copies the paragraph style, drops list numbering |
| Heading + body paragraphs | `insert_multiple_after` / `insert_multiple_after(find, [title, body], bolds=[True, False])` | `bold`: None inherits, True / False force it |
| Rewrite a whole paragraph | `replace_paragraph` / `replace_paragraph(find, new_text, startswith=True)` | Only for genuine rewrites; keeps bookmarks and comments |
| Delete a paragraph | `delete_paragraph` / `delete_paragraph(find, startswith=True)` | Accepting leaves no empty line |

- Every occurrence: `"replace_all": true` (non-overlapping, left to right, never re-matches its own output; methods return the count).
- One specific occurrence: `"occurrence": N` (1-based, document order). Without either, only the first match changes and a warning reports the total.
- `\n` in new text is a paragraph break (^p); `\t` becomes a real tab. The tool never writes manual line breaks (`w:br`).

```json
[
  {"op": "replace", "find": "3日内", "replace": "5个工作日内"},
  {"op": "replace", "find": "甲方", "replace": "甲方（买方）", "replace_all": true},
  {"op": "insert_after", "find": "1.1", "new_text": "1.2 新增条款……", "startswith": true},
  {"op": "delete_paragraph", "find": "第十条", "startswith": true, "occurrence": 2}
]
```

Edit fields are the Python parameter names. Single edits also work as CLI commands; see `python3 <skill-dir>/word_reviser.py --help`.

## Python API

```python
import sys
sys.path.insert(0, "<skill-dir>")  # the absolute skill path, never "."
from word_reviser import DocxReviser

with DocxReviser("input.docx", author="Agent") as r:
    r.replace_text("3日内", "5个工作日内")
    r.replace_text("甲方", "甲方（买方）", replace_all=True)
    r.insert_after("1.1", "1.2 新增条款……", startswith=True)
    results = r.apply_edits([{"op": "replace", "find": "乙方", "replace": "卖方"}])  # same format as edits.json
    total = r.save("input_revised.docx")  # atomic write; returns the revision mark count
```

Constructor options:
- `parts`: `"document"` by default. Also `headers`, `footers`, `footnotes`, `endnotes`, `comments`, or `"all"`.
- `include_toc=False`: table-of-contents paragraphs are skipped by default.
- `merge_short_equal=2`: equal islands of up to 2 characters between two changes are folded into one replacement, so redlines read as whole phrases. `0` keeps character-level marks.

## Matching rules

- Text matches inside one paragraph, against untracked text only. Tabs, field codes, footnote and comment references and images do not count as text, and they stay in place.
- A match cannot cross a hyperlink, content control, field or text-box boundary; the tool refuses it and says why.
- Text inside existing revisions cannot be matched and is never altered. This includes text inserted earlier in the same session, so plan every edit against the original text. For several changes in one sentence, use a single `replace_text_smart` covering the whole sentence.
- Automatic list numbering ("1.1", "第三条" generated by Word) is not part of the text.
- A "not found" error lists the closest paragraph text and likely causes (full/half-width characters, numbering, revisions); act on those rather than guessing.

## Validate

After changing `word_reviser.py`:

```bash
python3 <skill-dir>/tests/test_word_reviser.py
```

`save()` warns about bad deleted text automatically. To inspect an output file:

```bash
python3 - <<'PY'
from collections import Counter
from zipfile import ZipFile
from lxml import etree
docx = "output.docx"
ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
root = etree.fromstring(ZipFile(docx).read("word/document.xml"))
ids = root.xpath("//w:ins/@w:id | //w:del/@w:id", namespaces=ns)
print("w:ins", len(root.xpath("//w:ins", namespaces=ns)), "w:del", len(root.xpath("//w:del", namespaces=ns)))
print("duplicate ids", [i for i, n in Counter(ids).items() if n > 1])
print("bad w:del//w:t", len(root.xpath("//w:del//w:t[not(ancestor::w:txbxContent)]", namespaces=ns)))
PY
```

`duplicate ids` must be `[]` and `bad w:del//w:t` must be `0`.

## Common mistakes

| Symptom | Fix |
|---|---|
| `Text not found` | Copy the exact characters from the document, drop auto-numbering, read the suggestions in the error |
| Edit landed in the table of contents or the wrong clause | Use `occurrence`, a longer anchor, and keep `include_toc=False` |
| A second edit to the same sentence is not found | The first edit's text is now a revision; combine both into one `replace_text_smart` |
| `ModuleNotFoundError: word_reviser` | `sys.path.insert(0, "<skill-dir>")` with the absolute path |
| `match crosses a w:hyperlink boundary` | Edit the text inside and outside the container separately |
| One failure aborted a long script | Use batch `apply`; each edit succeeds or fails on its own |

## Technical constraints

1. **Format preservation**: edits operate on OOXML runs and paragraphs; run properties are copied, nothing is rebuilt from plain text.
2. **Existing revisions** (`w:ins`, `w:del`, `w:moveFrom`, `w:moveTo`) keep their text, author, date and id. `replace_paragraph` and `delete_paragraph` refuse paragraphs that contain them.
3. **Deletions** use `w:delText` (and `w:delInstrText` for field codes), never `w:t`.
4. **Paragraph marks are tracked**: Reject All restores the original document exactly, including paragraphs split by `\n`.
5. **Revision ids** start above every `w:id` already in the package, so they never collide.
6. **New paragraphs** never copy section breaks or other people's revisions. `insert_after` drops list numbering, and inserting after a section's last paragraph starts the next section (with a warning).
7. **Tables, headers, footers, footnotes, endnotes and comments** are supported through `parts`.
8. **Unsupported**: `.doc` files; editing field instructions, images, OLE objects and drawings.
