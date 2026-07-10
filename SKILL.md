---
name: word-revision
description: Add Word Track Changes revision marks to .docx files while preserving existing formatting.
---

# word-revision — Word Track Changes for DOCX

Apply edits to `.docx` files as Word Track Changes without rebuilding the document from plain text.

**Default strategy**: use fine-grained revisions. Keep unchanged text as normal text, and mark only actual insertions, deletions, and replacements. Use full-paragraph replacement only when a paragraph is genuinely rewritten.

**Usage**: `python3 word_reviser.py [--author NAME] [--parts PARTS] <input.docx> <output.docx> <command> [args...]`

**Default revision author**: `Agent`

---

## Workflow

### 1. Confirm Inputs

1. Confirm the source file is `.docx`.
2. Confirm the requested text edits.
3. If the target text may be outside the main document body, identify the needed parts.

### 2. Choose Revision Operation

| User intent | Operation | Method |
|---|---|---|
| Replace short text | Fine-grained replacement | `replace_text(find, replace)` |
| Replace a whole old clause with a whole new clause | Fine-grained diff replacement | `replace_text_smart(old_text, new_text)` |
| Replace every occurrence | Batch replacement | `replace_text(find, replace, replace_all=True)` |
| Add text after a sentence | Inline insertion | `insert_text_after(anchor, insert)` |
| Add a new paragraph after a clause | Paragraph insertion | `insert_after(find, new_text, startswith=True)` |
| Add title + body paragraphs | Multi-paragraph insertion | `insert_multiple_after(find, [title, body], bolds=[True, False])` |
| Rewrite a paragraph | Full paragraph replacement | `replace_paragraph(find, new_text, startswith=True)` |

**Preflight rules**:
- Prefer `replace_text`, `replace_text_smart`, `insert_text_after`, or `insert_after`.
- If the user provides "old whole clause -> new whole clause", use `replace_text_smart`.
- Use `replace_paragraph` only when structure and meaning are rewritten enough that stable local diff is not appropriate.
- Do not use `replace_paragraph` for a few words, dates, amounts, entity names, or a one-sentence supplement.
- Existing revision marks (`w:ins`, `w:del`, `w:moveFrom`, `w:moveTo`) must remain unchanged. Do not alter their text, author, date, or id.

### 3. Select Word Parts

By default, only the main document body is edited:

```python
reviser = DocxReviser("input.docx")
```

To edit other Word parts, pass `parts`:

```python
reviser = DocxReviser("input.docx", parts=["document", "headers", "footers"])
reviser = DocxReviser("input.docx", parts="all")
```

Supported part names:

- `document`
- `headers`
- `footers`
- `footnotes`
- `endnotes`
- `comments`

Missing optional parts are skipped. Unknown part names raise `ValueError`.

### 4. Generate Revision Script

```python
import sys
sys.path.insert(0, ".")
from word_reviser import DocxReviser

reviser = DocxReviser("input.docx", author="Agent")

reviser.replace_text("old text", "old text plus new words")
reviser.replace_text_smart("old whole clause...", "new whole clause...")
reviser.insert_text_after("anchor text", " inserted text")
reviser.insert_after("1.1", "1.2 New clause...", startswith=True)
reviser.insert_multiple_after(
    "2. Existing clause",
    ["New heading", "New body text."],
    bolds=[True, False],
)

# Use only when the entire paragraph is genuinely rewritten.
reviser.replace_paragraph("3.1", "3.1 Completely rewritten text...", startswith=True)

total = reviser.save("input_revised.docx")
print(f"Saved revised file with {total} revision marks")
```

### 5. Validate

Run the regression tests after changing `word_reviser.py`:

```bash
python3 -m py_compile word_reviser.py
python3 tests/test_word_reviser.py
```

For a generated file, inspect XML counts:

```bash
python3 - <<'PY'
from zipfile import ZipFile
from lxml import etree
docx = "output.docx"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ns = {"w": W}
root = etree.fromstring(ZipFile(docx).read("word/document.xml"))
print("w:ins", len(root.xpath(".//w:ins", namespaces=ns)))
print("w:del", len(root.xpath(".//w:del", namespaces=ns)))
print("bad w:del//w:t", len(root.xpath(".//w:del//w:t", namespaces=ns)))
print("w:del//w:delText", len(root.xpath(".//w:del//w:delText", namespaces=ns)))
PY
```

`bad w:del//w:t` must be `0`.

---

## Technical Constraints

1. **Format preservation**: edits operate on OOXML runs/paragraphs, not plain-text document reconstruction.
2. **Fine-grained visual revisions**: shared text remains normal; only actual changes become revision marks.
3. **Existing revision protection**: existing `w:ins`, `w:del`, `w:moveFrom`, and `w:moveTo` are skipped and preserved.
4. **Deletion standard**: deleted text inside `w:del` must use `w:delText`, not `w:t`.
5. **Line breaks**: `\n` becomes `w:br`.
6. **Table support**: table cell paragraphs are supported.
7. **Part support**: main body, headers, footers, footnotes, endnotes, and comments can be selected.
8. **List behavior**: inserted paragraphs remove copied `w:numPr` to avoid unexpected automatic numbering.
9. **Unsupported format**: `.doc` is not supported; save as `.docx` first.
10. **Unsupported objects**: the tool does not intentionally edit field instructions, images, OLE objects, relationship files, or drawing objects beyond ordinary `w:p/w:r/w:t` text available in selected parts.
