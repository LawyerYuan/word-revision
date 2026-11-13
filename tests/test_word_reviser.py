#!/usr/bin/env python3
import json
import os
import re
import shutil
import signal
import subprocess
import sys
import tempfile
import traceback
import warnings
import zipfile
from datetime import datetime, timedelta, timezone
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from word_reviser import DocxReviser  # noqa: E402

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
V = "urn:schemas-microsoft-com:vml"
NS = {"w": W}
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
TEST_TIMEOUT_SECONDS = 15


def wtag(tag):
    return f"{{{W}}}{tag}"


# --------------------------------------------------------------------------
# Document builders
# --------------------------------------------------------------------------

def el(tag, **attrs):
    elem = etree.Element(wtag(tag))
    for key, value in attrs.items():
        elem.set(wtag(key), str(value))
    return elem


def make_run(text, bold=False):
    r = etree.Element(wtag("r"))
    if bold:
        rpr = etree.SubElement(r, wtag("rPr"))
        etree.SubElement(rpr, wtag("b"))
    t = etree.SubElement(r, wtag("t"))
    t.text = text
    return r


def make_run_children(children):
    """Run with arbitrary content children; strings become w:t."""
    r = etree.Element(wtag("r"))
    for child in children:
        if isinstance(child, str):
            etree.SubElement(r, wtag("t")).text = child
        else:
            r.append(child)
    return r


def make_p(*items, numbered=False, ppr=None):
    p = etree.Element(wtag("p"))
    if ppr is not None:
        p.append(ppr)
    elif numbered:
        ppr = etree.SubElement(p, wtag("pPr"))
        numpr = etree.SubElement(ppr, wtag("numPr"))
        etree.SubElement(numpr, wtag("ilvl")).set(wtag("val"), "0")
        etree.SubElement(numpr, wtag("numId")).set(wtag("val"), "1")
    for item in items:
        p.append(make_run(item) if isinstance(item, str) else item)
    return p


def make_textbox_run(*inner_paragraphs):
    r = etree.Element(wtag("r"))
    pict = etree.SubElement(r, wtag("pict"))
    shape = etree.SubElement(pict, f"{{{V}}}shape")
    box = etree.SubElement(shape, f"{{{V}}}textbox")
    content = etree.SubElement(box, wtag("txbxContent"))
    for p in inner_paragraphs:
        content.append(p)
    return r


def make_field_runs(instruction, result):
    instr = etree.Element(wtag("instrText"))
    instr.text = instruction
    instr.set(XML_SPACE, "preserve")
    return [
        make_run_children([el("fldChar", fldCharType="begin")]),
        make_run_children([instr]),
        make_run_children([el("fldChar", fldCharType="separate")]),
        make_run(result),
        make_run_children([el("fldChar", fldCharType="end")]),
    ]


def make_styles(*pairs):
    root = etree.Element(wtag("styles"), nsmap={"w": W})
    for style_id, name in pairs:
        style = etree.SubElement(root, wtag("style"))
        style.set(wtag("type"), "paragraph")
        style.set(wtag("styleId"), style_id)
        etree.SubElement(style, wtag("name")).set(wtag("val"), name)
    return root


def styled_p(items, style_id):
    ppr = el("pPr")
    ppr.append(el("pStyle", val=style_id))
    return make_p(*items, ppr=ppr)


def make_existing_revised_p():
    p = etree.Element(wtag("p"))
    p.append(make_run("保持"))
    ins = etree.SubElement(p, wtag("ins"))
    ins.set(wtag("id"), "9")
    ins.set(wtag("author"), "ExistingAuthor")
    ins.set(wtag("date"), "2020-01-01T00:00:00Z")
    ins.append(make_run("既有插入"))
    d = etree.SubElement(p, wtag("del"))
    d.set(wtag("id"), "10")
    d.set(wtag("author"), "ExistingAuthor")
    d.set(wtag("date"), "2020-01-01T00:00:00Z")
    dr = etree.SubElement(d, wtag("r"))
    dt = etree.SubElement(dr, wtag("delText"))
    dt.text = "既有删除"
    p.append(make_run("尾部"))
    return p


def write_xml(path, elem):
    path.parent.mkdir(parents=True, exist_ok=True)
    etree.ElementTree(elem).write(path, xml_declaration=True, encoding="UTF-8", standalone=True)


def wrap_document(children):
    root = etree.Element(wtag("document"), nsmap={"w": W, "v": V})
    body = etree.SubElement(root, wtag("body"))
    for child in children:
        body.append(child)
    etree.SubElement(body, wtag("sectPr"))
    return root


def wrap_part(root_name, children):
    root = etree.Element(wtag(root_name), nsmap={"w": W})
    for child in children:
        root.append(child)
    return root


def make_docx(path, parts, extra=None):
    work = path.parent / f"src_{path.stem}"
    if work.exists():
        shutil.rmtree(work)
    (work / "word" / "_rels").mkdir(parents=True)
    write_xml(work / "word" / "document.xml", wrap_document(parts.get("document", [])))

    optional_roots = {
        "header1.xml": "hdr",
        "footer1.xml": "ftr",
        "footnotes.xml": "footnotes",
        "endnotes.xml": "endnotes",
        "comments.xml": "comments",
    }
    for filename, root_name in optional_roots.items():
        key = filename[:-4]
        if key in parts:
            write_xml(work / "word" / filename, wrap_part(root_name, parts[key]))
    for relative, elem in (extra or {}).items():
        write_xml(work / relative, elem)

    rels = etree.Element(f"{{{RELS}}}Relationships")
    write_xml(work / "word" / "_rels" / "document.xml.rels", rels)

    types = etree.Element(f"{{{CT}}}Types")
    for part_name in [
        "/word/document.xml",
        "/word/header1.xml",
        "/word/footer1.xml",
        "/word/footnotes.xml",
        "/word/endnotes.xml",
        "/word/comments.xml",
    ]:
        if (work / part_name.lstrip("/")).exists():
            override = etree.SubElement(types, f"{{{CT}}}Override")
            override.set("PartName", part_name)
            override.set("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")
    write_xml(work / "[Content_Types].xml", types)

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for file in sorted(work.rglob("*")):
            if file.is_file():
                zf.write(file, file.relative_to(work))


# --------------------------------------------------------------------------
# Inspection helpers
# --------------------------------------------------------------------------

def parse_part(docx, part):
    with zipfile.ZipFile(docx) as zf:
        return etree.fromstring(zf.read(part))


def accepted_text(root):
    return "".join(t.text or "" for t in root.xpath(".//w:t[not(ancestor::w:del)]", namespaces=NS))


def normal_text(root):
    return "".join(t.text or "" for t in root.xpath(".//w:t[not(ancestor::w:del) and not(ancestor::w:ins)]", namespaces=NS))


def authors(root):
    return set(root.xpath(".//w:ins/@w:author | .//w:del/@w:author", namespaces=NS))


def _nearest(node, tag):
    for anc in node.iterancestors(wtag(tag)):
        return anc
    return None


def _has_ancestor_before(node, tag, stop):
    for anc in node.iterancestors():
        if anc is stop:
            return False
        if anc.tag == wtag(tag):
            return True
    return False


def paragraph_text_view(p, mode):
    """Own text of a paragraph after Accept All ("accept") or Reject All ("reject")."""
    parts = []
    for node in p.iter(wtag("t"), wtag("delText"), wtag("tab")):
        if _nearest(node, "p") is not p:
            continue
        inserted = _has_ancestor_before(node, "ins", p)
        deleted = _has_ancestor_before(node, "del", p)
        if mode == "accept" and deleted:
            continue
        if mode == "reject" and inserted:
            continue
        if node.tag == wtag("tab"):
            parts.append("\t")
        else:
            parts.append(node.text or "")
    return "".join(parts)


def paragraph_mark_state(p):
    rpr = p.find(f"{wtag('pPr')}/{wtag('rPr')}")
    if rpr is None:
        return None
    if rpr.find(wtag("ins")) is not None:
        return "ins"
    if rpr.find(wtag("del")) is not None:
        return "del"
    return None


def paragraph_views(root, mode):
    """Paragraph texts after Accept All / Reject All, merging removed paragraph marks."""
    result, carry = [], ""
    for p in root.iter(wtag("p")):
        if _nearest(p, "p") is not None:
            continue  # text-box paragraph, part of its host's story
        text = carry + paragraph_text_view(p, mode)
        state = paragraph_mark_state(p)
        if (mode == "accept" and state == "del") or (mode == "reject" and state == "ins"):
            carry = text
            continue
        carry = ""
        result.append(text)
    if carry:
        result.append(carry)
    return result


def views(docx, mode="accept", part="word/document.xml"):
    return paragraph_views(parse_part(docx, part), mode)


def bad_deleted_text(root):
    """w:t that belongs to a deleted run of the same story (should be w:delText)."""
    bad = []
    for t in root.iter(wtag("t")):
        for anc in t.iterancestors():
            if anc.tag == wtag("txbxContent"):
                break
            if anc.tag == wtag("del"):
                bad.append(t)
                break
    return bad


def revision_ids(root):
    tags = [wtag(name) for name in ("ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange")]
    return [e.get(wtag("id")) for e in root.iter(*tags)]


def para_containing(root, text):
    for p in root.iter(wtag("p")):
        if text in "".join(p.xpath("./w:r/w:t/text() | ./w:ins/w:r/w:t/text()", namespaces=NS)):
            return p
    raise AssertionError(f"no paragraph contains {text!r}")


def effective_bold(run):
    b = run.find(f"{wtag('rPr')}/{wtag('b')}")
    if b is None:
        return False
    return b.get(wtag("val")) not in ("0", "false", "off")


def expect_error(fn, exc_type=ValueError, contains=None):
    try:
        fn()
    except exc_type as exc:
        if contains is not None:
            assert contains.lower() in str(exc).lower(), f"{contains!r} not in {exc}"
        return exc
    raise AssertionError(f"expected {exc_type.__name__}")


def run_cli(*args):
    return subprocess.run(
        [sys.executable, str(ROOT / "word_reviser.py"), *[str(a) for a in args]],
        capture_output=True,
        text=True,
        env={**os.environ, "PYTHONDONTWRITEBYTECODE": "1"},
    )


def assert_integrity(docx, parts=("word/document.xml",)):
    for part in parts:
        root = parse_part(docx, part)
        assert not bad_deleted_text(root), f"bad w:t in w:del for {part}"
        assert root.xpath(".//w:ins | .//w:del", namespaces=NS), f"missing revisions in {part}"
        ids = revision_ids(root)
        assert len(ids) == len(set(ids)), f"duplicate revision ids in {part}: {ids}"
    settings = parse_part(docx, "word/settings.xml")
    assert settings.xpath(".//w:trackRevisions", namespaces=NS)


# --------------------------------------------------------------------------
# Existing behaviour
# --------------------------------------------------------------------------

def test_core_operations(tmp):
    old = "甲方应在3日内付款，乙方应提供发票，并保证资料真实。"
    new = "甲方应在5个工作日内付款，乙方应提供合法有效发票，并保证资料真实。"
    table = etree.Element(wtag("tbl"))
    tr = etree.SubElement(table, wtag("tr"))
    tc = etree.SubElement(tr, wtag("tc"))
    tc.append(make_p("表格旧文本"))
    tc.append(make_p("表格锚点"))

    input_path = tmp / "core_input.docx"
    output_path = tmp / "core_output.docx"
    make_docx(input_path, {
        "document": [
            make_p("前言：", old[:7], old[7:18], old[18:], "尾部。"),
            make_p("跨", "run", "目标"),
            make_p("请在此处补充"),
            make_p("换行目标"),
            make_p("1. 列表锚点", numbered=True),
            make_p("标题锚点"),
            table,
        ]
    })

    reviser = DocxReviser(str(input_path))
    reviser.replace_text_smart(old, new)
    reviser.replace_text("跨run目标", "跨run替换结果")
    reviser.insert_text_after("请在此处", "追加")
    reviser.replace_text("换行目标", "第一行\n第二行")
    reviser.replace_text("表格旧文本", "表格新文本")
    reviser.insert_after("表格锚点", "表格新增段", startswith=True)
    reviser.insert_after("1. 列表锚点", "列表后新增", startswith=True)
    reviser.insert_multiple_after("标题锚点", ["新增标题", "新增正文"], bolds=[True, False])
    total = reviser.save(str(output_path))
    assert total > 0

    assert_integrity(output_path)
    doc = parse_part(output_path, "word/document.xml")
    assert f"前言：{new}尾部。" in accepted_text(doc)
    assert "乙方应提供" in normal_text(doc)
    assert "跨run替换结果" in accepted_text(doc)
    assert "请在此处追加补充" in accepted_text(doc)
    assert not doc.xpath(".//w:br", namespaces=NS)
    first_line_p = doc.xpath(".//w:t[text()='第一行']/ancestor::w:p[1]", namespaces=NS)[0]
    second_line_p = doc.xpath(".//w:t[text()='第二行']/ancestor::w:p[1]", namespaces=NS)[0]
    assert first_line_p is not second_line_p
    assert first_line_p.getnext() is second_line_p
    assert authors(doc) == {"Agent"}

    table_anchor = doc.xpath(".//w:t[text()='表格锚点']/ancestor::w:p[1]", namespaces=NS)[0]
    table_inserted = doc.xpath(".//w:t[text()='表格新增段']/ancestor::w:p[1]", namespaces=NS)[0]
    assert table_anchor.getparent() is table_inserted.getparent()
    list_inserted = doc.xpath(".//w:t[text()='列表后新增']/ancestor::w:p[1]", namespaces=NS)[0]
    assert not list_inserted.xpath("./w:pPr/w:numPr", namespaces=NS)
    assert doc.xpath(".//w:t[text()='新增标题']/ancestor::w:r[1]/w:rPr/w:b", namespaces=NS)


def test_newline_is_paragraph_break(tmp):
    input_path = tmp / "newline_input.docx"
    output_path = tmp / "newline_output.docx"
    make_docx(input_path, {
        "document": [
            make_p("换行替换旧文"),
            make_p("锚点前文尾部"),
            make_p("整段旧文内容"),
            make_p("新增锚点段落"),
        ]
    })

    reviser = DocxReviser(str(input_path))
    reviser.replace_text("换行替换旧文", "第一段\n第二段")
    reviser.insert_text_after("锚点前文", "插入一\n插入二")
    reviser.replace_paragraph("整段旧文", "新段一\n新段二", startswith=True)
    reviser.insert_after("新增锚点段落", "追加一\n追加二", startswith=True)
    reviser.save(str(output_path))

    assert_integrity(output_path)
    doc = parse_part(output_path, "word/document.xml")
    assert not doc.xpath(".//w:br", namespaces=NS)
    assert not doc.xpath(".//w:cr", namespaces=NS)

    def para_of(text):
        return doc.xpath(f".//w:t[text()='{text}']/ancestor::w:p[1]", namespaces=NS)[0]

    text = accepted_text(doc)
    assert "第一段第二段" in text
    assert "锚点前文插入一插入二尾部" in text
    assert "新段一新段二" in text
    assert "新增锚点段落追加一追加二" in text

    # replace_text / replace_paragraph split into consecutive paragraphs.
    assert para_of("第一段").getnext() is para_of("第二段")
    assert para_of("新段一").getnext() is para_of("新段二")

    # insert_text_after moves the trailing original text into the final new paragraph.
    assert para_of("锚点前文").getnext() is para_of("插入二")
    assert "尾部" in "".join(para_of("插入二").xpath(".//w:t/text()", namespaces=NS))

    # insert_after creates one tracked paragraph per line.
    assert para_of("新增锚点段落").getnext() is para_of("追加一")
    assert para_of("追加一").getnext() is para_of("追加二")

    assert authors(doc) == {"Agent"}


def test_author_and_cli(tmp):
    input_path = tmp / "cli_input.docx"
    output_path = tmp / "cli_output.docx"
    make_docx(input_path, {"document": [make_p("旧文本")]})
    subprocess.run(
        [
            sys.executable,
            str(ROOT / "word_reviser.py"),
            "--author",
            "Reviewer",
            str(input_path),
            str(output_path),
            "replace",
            "旧文本",
            "新文本",
        ],
        check=True,
        capture_output=True,
        env={**os.environ, "PYTHONDONTWRITEBYTECODE": "1"},
    )
    doc = parse_part(output_path, "word/document.xml")
    assert accepted_text(doc) == "新文本"
    assert authors(doc) == {"Reviewer"}


def test_existing_revisions_are_preserved(tmp):
    input_path = tmp / "existing_input.docx"
    output_path = tmp / "existing_output.docx"
    make_docx(input_path, {"document": [make_existing_revised_p(), make_p("普通旧文本")]})
    before = parse_part(input_path, "word/document.xml")
    old_revision_xml = [etree.tostring(e) for e in before.xpath(".//w:ins | .//w:del", namespaces=NS)]

    blocked = DocxReviser(str(input_path), author="Agent")
    try:
        blocked.replace_paragraph("保持尾部", "替换整段")
    except ValueError as exc:
        assert "existing revision marks" in str(exc)
    else:
        raise AssertionError("replace_paragraph must refuse paragraphs with existing revisions")

    reviser = DocxReviser(str(input_path), author="Agent")
    reviser.replace_text("普通旧文本", "普通新文本")
    reviser.save(str(output_path))

    after = parse_part(output_path, "word/document.xml")
    preserved = [etree.tostring(e) for e in after.xpath(".//w:ins[@w:id='9'] | .//w:del[@w:id='10']", namespaces=NS)]
    assert preserved == old_revision_xml
    assert "普通新文本" in accepted_text(after)


def test_parts_selection(tmp):
    input_path = tmp / "parts_input.docx"
    doc_only_output = tmp / "parts_document_only.docx"
    all_output = tmp / "parts_all.docx"
    make_docx(input_path, {
        "document": [make_p("正文旧")],
        "header1": [make_p("页眉旧")],
        "footer1": [make_p("页脚旧")],
        "footnotes": [make_p("脚注旧")],
        "endnotes": [make_p("尾注旧")],
        "comments": [make_p("批注旧")],
    })

    reviser = DocxReviser(str(input_path))
    reviser.replace_text("正文旧", "正文新")
    reviser.save(str(doc_only_output))
    assert "正文新" in accepted_text(parse_part(doc_only_output, "word/document.xml"))
    assert "页眉旧" in accepted_text(parse_part(doc_only_output, "word/header1.xml"))

    reviser = DocxReviser(str(input_path), parts="all")
    reviser.replace_text("旧", "新", replace_all=True)
    reviser.save(str(all_output))
    for part in [
        "word/document.xml",
        "word/header1.xml",
        "word/footer1.xml",
        "word/footnotes.xml",
        "word/endnotes.xml",
        "word/comments.xml",
    ]:
        assert "旧" not in accepted_text(parse_part(all_output, part)), part
        assert_integrity(all_output, (part,))


def test_revision_dates_are_beijing_time(tmp):
    input_path = tmp / "date_input.docx"
    output_path = tmp / "date_output.docx"
    make_docx(input_path, {"document": [make_p("日期旧文本")]})

    reviser = DocxReviser(str(input_path))
    reviser.replace_text("日期旧文本", "日期新文本")
    reviser.save(str(output_path))

    doc = parse_part(output_path, "word/document.xml")
    dates = doc.xpath(".//w:ins/@w:date | .//w:del/@w:date", namespaces=NS)
    assert dates, "missing revision date attributes"
    pattern = re.compile(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\+08:00$")
    now_beijing = datetime.now(timezone(timedelta(hours=8)))
    for value in dates:
        assert pattern.match(value), f"revision date not Beijing time with +08:00 offset: {value}"
        stamp = datetime.fromisoformat(value)
        assert abs(now_beijing - stamp) < timedelta(minutes=5), f"revision date not current: {value}"


# --------------------------------------------------------------------------
# P0: replace_all must terminate, hit every occurrence, never cascade
# --------------------------------------------------------------------------

def test_replace_all_terminates_when_find_survives_diff(tmp):
    input_path = tmp / "in.docx"
    output_path = tmp / "out.docx"
    make_docx(input_path, {"document": [
        make_p("价款为人民币100万元，另付100万元。"),
        make_p("详见第三条及第三条附件。"),
    ]})
    reviser = DocxReviser(str(input_path))
    assert reviser.replace_text("100万元", "100.5万元", replace_all=True) == 2
    assert reviser.replace_text("第三条", "第三十条", replace_all=True) == 2
    reviser.save(str(output_path))
    assert views(output_path) == ["价款为人民币100.5万元，另付100.5万元。", "详见第三十条及第三十条附件。"]
    assert views(output_path, "reject") == ["价款为人民币100万元，另付100万元。", "详见第三条及第三条附件。"]


def test_insert_text_after_replace_all(tmp):
    input_path = tmp / "in.docx"
    output_path = tmp / "out.docx"
    make_docx(input_path, {"document": [make_p("甲方应付款，甲方应开票。")]})
    reviser = DocxReviser(str(input_path))
    assert reviser.insert_text_after("甲方", "（买方）", replace_all=True) == 2
    reviser.save(str(output_path))
    assert views(output_path) == ["甲方（买方）应付款，甲方（买方）应开票。"]


def test_smart_replace_all_covers_every_paragraph(tmp):
    input_path = tmp / "in.docx"
    output_path = tmp / "out.docx"
    make_docx(input_path, {"document": [make_p("甲方A"), make_p("甲方B与甲方C")]})
    reviser = DocxReviser(str(input_path))
    assert reviser.replace_text_smart("甲方", "甲方公司", replace_all=True) == 3
    reviser.save(str(output_path))
    assert views(output_path) == ["甲方公司A", "甲方公司B与甲方公司C"]


def test_replace_all_substring_every_occurrence(tmp):
    input_path = tmp / "in.docx"
    output_path = tmp / "out.docx"
    make_docx(input_path, {"document": [make_p("甲方与甲方签约。")]})
    reviser = DocxReviser(str(input_path))
    assert reviser.replace_text("甲方", "甲方公司", replace_all=True) == 2
    reviser.save(str(output_path))
    assert views(output_path) == ["甲方公司与甲方公司签约。"]


def test_replace_all_is_not_cascading(tmp):
    input_path = tmp / "in.docx"
    output_path = tmp / "out.docx"
    make_docx(input_path, {"document": [make_p("AABB"), make_p("CCD")]})
    reviser = DocxReviser(str(input_path))
    assert reviser.replace_text_full("AB", "X", replace_all=True) == 1
    assert reviser.replace_text("CD", "D", replace_all=True) == 1
    reviser.save(str(output_path))
    assert views(output_path) == ["AXB", "CD"]


def test_invalid_find_text_is_rejected(tmp):
    input_path = tmp / "in.docx"
    make_docx(input_path, {"document": [make_p("甲方应付款")]})
    reviser = DocxReviser(str(input_path))
    expect_error(lambda: reviser.replace_text("", "x"))
    expect_error(lambda: reviser.insert_text_after("", "x", replace_all=True))
    expect_error(lambda: reviser.replace_text("甲方\n应付款", "x"), contains="paragraph")
    expect_error(lambda: reviser.replace_text("甲方\t应付款", "x"), contains="tab")


def test_textbox_text_edited_once(tmp):
    input_path = tmp / "in.docx"
    output_path = tmp / "out.docx"
    make_docx(input_path, {"document": [
        make_p("正文100万元", make_textbox_run(make_p("框内100万元"))),
    ]})
    reviser = DocxReviser(str(input_path))
    assert reviser.replace_text("100万元", "100.5万元", replace_all=True) == 2
    reviser.save(str(output_path))
    root = parse_part(output_path, "word/document.xml")
    assert len(root.xpath(".//w:txbxContent", namespaces=NS)) == 1
    inner = root.xpath(".//w:txbxContent/w:p", namespaces=NS)[0]
    assert paragraph_text_view(inner, "accept") == "框内100.5万元"
    assert views(output_path) == ["正文100.5万元"]


# --------------------------------------------------------------------------
# P1: no silent loss or corruption of document structure
# --------------------------------------------------------------------------

def test_zero_width_runs_inside_range_are_preserved(tmp):
    input_path = tmp / "in.docx"
    output_path = tmp / "out.docx"
    make_docx(input_path, {"document": [
        make_p("合同", make_run_children([el("footnoteReference", id="1")]), "约定"),
        make_p("甲方", make_run_children([el("tab")]), "乙方"),
        make_p("见第", *make_field_runs(" REF _Ref1 \\h ", "3.2"), "条"),
    ]})
    reviser = DocxReviser(str(input_path))
    reviser.replace_text("合同约定", "协议约定")
    reviser.replace_text("甲方乙方", "甲方丙方")
    reviser.replace_text("第3.2条", "第4.1条")
    reviser.save(str(output_path))

    root = parse_part(output_path, "word/document.xml")
    assert views(output_path) == ["协议约定", "甲方\t丙方", "见第4.1条"]
    assert views(output_path, "reject") == ["合同约定", "甲方\t乙方", "见第3.2条"]
    assert len(root.xpath(".//w:footnoteReference[not(ancestor::w:del)]", namespaces=NS)) == 1
    assert len(root.xpath(".//w:tab[not(ancestor::w:del)]", namespaces=NS)) == 1
    assert root.xpath(".//w:fldChar/@w:fldCharType", namespaces=NS) == ["begin", "separate", "end"]
    assert not root.xpath(".//w:fldChar[ancestor::w:del] | .//w:instrText[ancestor::w:del]", namespaces=NS)
    assert len(root.xpath(".//w:instrText", namespaces=NS)) == 1
    assert_integrity(output_path)


def test_multi_content_runs_are_not_duplicated(tmp):
    input_path = tmp / "in.docx"
    output_path = tmp / "out.docx"
    make_docx(input_path, {"document": [
        make_p(make_run_children(["abc", el("tab"), "def"])),
        make_p(make_run_children([el("tab"), "Text"])),
    ]})
    reviser = DocxReviser(str(input_path))
    assert reviser._para_text(reviser._all_paragraphs()[0]) == "abcdef"
    reviser.replace_text("b", "X")
    reviser.replace_text("ex", "EX")
    reviser.save(str(output_path))
    assert views(output_path) == ["aXc\tdef", "\tTEXt"]
    root = parse_part(output_path, "word/document.xml")
    assert len(root.xpath(".//w:tab[not(ancestor::w:del)]", namespaces=NS)) == 2


def test_split_runs_keep_edge_spaces(tmp):
    input_path = tmp / "in.docx"
    output_path = tmp / "out.docx"
    make_docx(input_path, {"document": [make_p("hello world")]})
    reviser = DocxReviser(str(input_path))
    reviser.replace_text("world", "earth")
    reviser.save(str(output_path))
    root = parse_part(output_path, "word/document.xml")
    for node in root.iter(wtag("t"), wtag("delText")):
        text = node.text or ""
        if text != text.strip():
            assert node.get(XML_SPACE) == "preserve", f"missing xml:space=preserve on {text!r}"
    assert views(output_path) == ["hello earth"]


def test_cross_container_match_is_rejected(tmp):
    input_path = tmp / "in.docx"
    output_path = tmp / "out.docx"
    link = el("hyperlink")
    link.append(make_run("官网"))
    make_docx(input_path, {"document": [make_p("请访问", link, "了解")]})
    reviser = DocxReviser(str(input_path))
    expect_error(lambda: reviser.replace_text("访问官网", "浏览网站"), contains="hyperlink")
    para = reviser._all_paragraphs()[0]
    assert reviser._replace_range_smart(para, 1, "访问官网", "浏览网站") is False
    root = reviser.part_trees[0]["root"]
    assert not root.xpath(".//w:ins | .//w:del", namespaces=NS)
    assert root.xpath(".//w:hyperlink//w:t/text()", namespaces=NS) == ["官网"]
    reviser.replace_text("官网", "网站")
    reviser.save(str(output_path))
    out = parse_part(output_path, "word/document.xml")
    link_text = "".join(out.xpath(".//w:hyperlink//w:t[not(ancestor::w:del)]/text()", namespaces=NS))
    assert link_text == "网站"


def test_inserted_paragraphs_are_tracked_and_sanitized(tmp):
    input_path = tmp / "in.docx"
    output_path = tmp / "out.docx"
    ppr = el("pPr")
    rpr = etree.SubElement(ppr, wtag("rPr"))
    existing = etree.SubElement(rpr, wtag("ins"))
    for key, value in (("id", "7"), ("author", "ExistingAuthor"), ("date", "2020-01-01T00:00:00Z")):
        existing.set(wtag(key), value)
    etree.SubElement(ppr, wtag("sectPr"))
    change = etree.SubElement(ppr, wtag("pPrChange"))
    for key, value in (("id", "8"), ("author", "ExistingAuthor"), ("date", "2020-01-01T00:00:00Z")):
        change.set(wtag(key), value)
    etree.SubElement(change, wtag("pPr"))
    make_docx(input_path, {"document": [make_p("节末段", ppr=ppr), make_p("下一节")]})

    reviser = DocxReviser(str(input_path))
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        reviser.insert_after("节末段", "新增一\n新增二", startswith=True)
    assert any("section" in str(w.message).lower() for w in caught)
    reviser.save(str(output_path))

    root = parse_part(output_path, "word/document.xml")
    assert len(root.xpath(".//w:pPr/w:sectPr", namespaces=NS)) == 1
    assert para_containing(root, "节末段").xpath("./w:pPr/w:sectPr", namespaces=NS)
    assert len(root.xpath(".//*[@w:author='ExistingAuthor']", namespaces=NS)) == 2
    for text in ("新增一", "新增二"):
        p = para_containing(root, text)
        marks = p.xpath("./w:pPr/w:rPr/w:ins", namespaces=NS)
        assert len(marks) == 1 and marks[0].get(wtag("author")) == "Agent", text
        assert not p.xpath("./w:pPr/w:pPrChange | ./w:pPr/w:sectPr", namespaces=NS), text
    for ppr_elem in root.iter(wtag("pPr")):
        children = [child.tag for child in ppr_elem]
        if wtag("rPr") in children:
            for later in ("sectPr", "pPrChange"):
                if wtag(later) in children:
                    assert children.index(wtag("rPr")) < children.index(wtag(later)), children
    ids = revision_ids(root)
    assert len(ids) == len(set(ids)), ids


def test_reject_all_restores_original(tmp):
    input_path = tmp / "in.docx"
    output_path = tmp / "out.docx"
    original = ["换行目标尾巴", "锚点前文尾部", "新增锚点", "标题锚点", "整段旧文内容", "待删除段落", "普通旧文本", "结尾"]
    make_docx(input_path, {"document": [
        make_p("换行", "目标", "尾巴"),
        make_p("锚点前文尾部"),
        make_p("新增锚点"),
        make_p("标题锚点"),
        make_p("整段旧文内容"),
        make_p("待删除段落"),
        make_p("普通旧文本"),
        make_p("结尾"),
    ]})
    assert views(input_path) == original

    reviser = DocxReviser(str(input_path))
    reviser.replace_text("换行目标", "第一行\n第二行")
    reviser.insert_text_after("锚点前文", "插入一\n插入二")
    reviser.insert_after("新增锚点", "追加一\n追加二", startswith=True)
    reviser.insert_multiple_after("标题锚点", ["新标题", "新正文"], bolds=[True, False])
    reviser.replace_paragraph("整段旧文", "新段一\n新段二", startswith=True)
    reviser.delete_paragraph("待删除段落", startswith=True)
    reviser.replace_text("普通旧文本", "普通新文本")
    reviser.save(str(output_path))

    assert views(output_path, "reject") == original
    assert views(output_path) == [
        "第一行", "第二行尾巴", "锚点前文插入一", "插入二尾部", "新增锚点", "追加一", "追加二",
        "标题锚点", "新标题", "新正文", "新段一", "新段二", "普通新文本", "结尾",
    ]
    assert_integrity(output_path)


def test_revision_ids_are_unique(tmp):
    input_path = tmp / "in.docx"
    output_path = tmp / "out.docx"
    existing = el("ins", id="1000", author="Other", date="2020-01-01T00:00:00Z")
    existing.append(make_run("他人插入"))
    make_docx(input_path, {"document": [make_p("旧文本", existing), make_p("待删段落"), make_p("保留")]})
    reviser = DocxReviser(str(input_path))
    reviser.replace_text("旧文本", "新文本")
    reviser.delete_paragraph("待删段落")
    reviser.save(str(output_path))
    ids = revision_ids(parse_part(output_path, "word/document.xml"))
    assert len(ids) == len(set(ids)), ids
    assert all(int(value) > 1000 for value in ids if value != "1000"), ids


def test_replace_paragraph_keeps_annotations(tmp):
    input_path = tmp / "in.docx"
    output_path = tmp / "out.docx"
    p = make_p()
    p.append(el("bookmarkStart", id="0", name="_Ref1"))
    p.append(el("commentRangeStart", id="5"))
    p.append(make_run("3.1 原条款"))
    link = el("hyperlink")
    link.append(make_run("（链接）"))
    p.append(link)
    p.append(el("commentRangeEnd", id="5"))
    p.append(make_run_children([el("commentReference", id="5")]))
    p.append(el("bookmarkEnd", id="0"))
    make_docx(input_path, {"document": [p, make_p("下一段")]})

    reviser = DocxReviser(str(input_path))
    reviser.replace_paragraph("3.1", "3.1 新条款", startswith=True)
    reviser.save(str(output_path))

    root = parse_part(output_path, "word/document.xml")
    for tag in ("bookmarkStart", "bookmarkEnd", "commentRangeStart", "commentRangeEnd", "hyperlink"):
        assert len(root.xpath(f".//w:{tag}", namespaces=NS)) == 1, tag
    assert len(root.xpath(".//w:commentReference[not(ancestor::w:del)]", namespaces=NS)) == 1
    first = root.xpath(".//w:body/w:p", namespaces=NS)[0]
    children = list(first)
    ins = first.find(wtag("ins"))
    assert ins is not None
    assert children.index(ins) < children.index(first.find(wtag("bookmarkEnd")))
    assert views(output_path) == ["3.1 新条款", "下一段"]
    assert views(output_path, "reject") == ["3.1 原条款（链接）", "下一段"]


def test_delete_paragraph_keeps_annotations_and_textbox_text(tmp):
    input_path = tmp / "in.docx"
    output_path = tmp / "out.docx"
    p = make_p()
    p.append(el("bookmarkStart", id="0", name="_Ref2"))
    p.append(make_run("待删正文"))
    p.append(make_textbox_run(make_p("框内文字")))
    p.append(el("bookmarkEnd", id="0"))
    make_docx(input_path, {"document": [p, make_p("保留段")]})

    reviser = DocxReviser(str(input_path))
    reviser.delete_paragraph("待删正文", startswith=True)
    reviser.save(str(output_path))

    root = parse_part(output_path, "word/document.xml")
    assert root.xpath(".//w:delText/text()", namespaces=NS) == ["待删正文"]
    assert root.xpath(".//w:txbxContent//w:t/text()", namespaces=NS) == ["框内文字"]
    assert len(root.xpath(".//w:txbxContent", namespaces=NS)) == 1
    assert len(root.xpath(".//w:bookmarkStart", namespaces=NS)) == 1
    assert views(output_path) == ["保留段"]
    assert not bad_deleted_text(root)


def test_new_text_is_sanitized(tmp):
    input_path = tmp / "in.docx"
    output_path = tmp / "out.docx"
    make_docx(input_path, {"document": [make_p("甲"), make_p("乙"), make_p("丙")]})
    reviser = DocxReviser(str(input_path))
    reviser.replace_text("甲", "甲\t方")
    reviser.replace_text("乙", "乙一\x0b乙二")
    expect_error(lambda: reviser.replace_text("丙", "丙\x01"), contains="control character")
    reviser.save(str(output_path))
    root = parse_part(output_path, "word/document.xml")
    assert not any("\t" in (t.text or "") for t in root.iter(wtag("t")))
    assert root.xpath(".//w:ins//w:tab", namespaces=NS)
    assert not root.xpath(".//w:br", namespaces=NS)
    assert views(output_path) == ["甲\t方", "乙一", "乙二", "丙"]


# --------------------------------------------------------------------------
# P2: fewer agent retries, cleaner redlines
# --------------------------------------------------------------------------

def test_occurrence_selects_nth_match(tmp):
    input_path = tmp / "in.docx"
    output_path = tmp / "out.docx"
    make_docx(input_path, {"document": [make_p("甲方、甲方、甲方"), make_p("1.1 第一处"), make_p("1.1 第二处")]})
    reviser = DocxReviser(str(input_path))
    assert reviser.replace_text("甲方", "乙方", occurrence=2) == 1
    reviser.insert_after("1.1", "新增段", startswith=True, occurrence=2)
    expect_error(lambda: reviser.replace_text("甲方", "丙方", occurrence=5), contains="occurrence")
    expect_error(lambda: reviser.replace_text("甲方", "丙方", replace_all=True, occurrence=1), contains="occurrence")
    reviser.save(str(output_path))
    assert views(output_path) == ["甲方、乙方、甲方", "1.1 第一处", "1.1 第二处", "新增段"]


def test_ambiguous_match_warns(tmp):
    input_path = tmp / "in.docx"
    make_docx(input_path, {"document": [make_p("甲方与甲方")]})
    reviser = DocxReviser(str(input_path))
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        reviser.replace_text("甲方", "乙方")
    messages = [str(w.message) for w in caught]
    assert any("2" in m and "occurrence" in m for m in messages), messages


def test_toc_paragraphs_are_skipped(tmp):
    input_path = tmp / "in.docx"
    sdt = el("sdt")
    sdt_pr = etree.SubElement(sdt, wtag("sdtPr"))
    gallery = etree.SubElement(etree.SubElement(sdt_pr, wtag("docPartObj")), wtag("docPartGallery"))
    gallery.set(wtag("val"), "Table of Contents")
    etree.SubElement(sdt, wtag("sdtContent")).append(make_p("第五条 违约责任 9"))
    make_docx(
        input_path,
        {"document": [
            styled_p([make_run_children(["第三条 付款方式", el("tab"), "5"])], "10"),
            sdt,
            make_p("第三条 付款方式：买方应付款。"),
            make_p("第五条 违约责任：违约方赔偿。"),
        ]},
        extra={"word/styles.xml": make_styles(("10", "toc 1"))},
    )

    output_path = tmp / "out.docx"
    reviser = DocxReviser(str(input_path))
    reviser.insert_after("第三条", "新增款项", startswith=True)
    reviser.replace_text("第五条", "第六条")
    reviser.save(str(output_path))
    assert views(output_path) == [
        "第三条 付款方式\t5", "第五条 违约责任 9", "第三条 付款方式：买方应付款。", "新增款项", "第六条 违约责任：违约方赔偿。",
    ]

    toc_output = tmp / "toc.docx"
    reviser = DocxReviser(str(input_path), include_toc=True)
    reviser.insert_after("第三条", "新增款项", startswith=True, occurrence=1)
    reviser.save(str(toc_output))
    assert views(toc_output)[:2] == ["第三条 付款方式\t5", "新增款项"]


def test_not_found_error_suggests_closest_text(tmp):
    input_path = tmp / "in.docx"
    make_docx(input_path, {"document": [make_p("甲方（买方）应于三日内付款。"), make_p("其他内容")]})
    reviser = DocxReviser(str(input_path))
    exc = expect_error(lambda: reviser.replace_text("甲方(买方)应于三日内付款", "x"))
    message = str(exc)
    assert "甲方（买方）应于三日内付款" in message, message
    assert "width" in message.lower(), message


def test_smart_diff_merges_short_equal_islands(tmp):
    old = "本协议自双方签字盖章之日起生效"
    new = "本协议经各方签署并经审批机关批准后生效"
    input_path = tmp / "in.docx"
    make_docx(input_path, {"document": [make_p(old)]})

    merged_path = tmp / "merged.docx"
    reviser = DocxReviser(str(input_path))
    reviser.replace_text_smart(old, new)
    reviser.save(str(merged_path))
    root = parse_part(merged_path, "word/document.xml")
    assert len(root.xpath(".//w:ins", namespaces=NS)) == 1
    assert len(root.xpath(".//w:del", namespaces=NS)) == 1
    assert normal_text(root) == "本协议生效"
    assert views(merged_path) == [new]

    char_path = tmp / "char.docx"
    reviser = DocxReviser(str(input_path), merge_short_equal=0)
    reviser.replace_text_smart(old, new)
    reviser.save(str(char_path))
    root = parse_part(char_path, "word/document.xml")
    assert len(root.xpath(".//w:ins", namespaces=NS)) == 2
    assert len(root.xpath(".//w:del", namespaces=NS)) == 2
    assert views(char_path) == [new]


def test_cli_apply_batch(tmp):
    input_path = tmp / "in.docx"
    make_docx(input_path, {"document": [make_p("甲方应付款。"), make_p("乙方应开票。")]})
    good = [
        {"op": "replace", "find": "甲方", "replace": "买方"},
        {"op": "insert_text_after", "anchor": "乙方", "insert": "（卖方）"},
    ]
    bad = [good[0], {"op": "replace", "find": "不存在的文字", "replace": "x"}]
    edits_ok = tmp / "ok.json"
    edits_bad = tmp / "bad.json"
    edits_ok.write_text(json.dumps(good, ensure_ascii=False), encoding="utf-8")
    edits_bad.write_text(json.dumps(bad, ensure_ascii=False), encoding="utf-8")

    out_ok = tmp / "ok.docx"
    res = run_cli(input_path, out_ok, "apply", edits_ok)
    assert res.returncode == 0, res.stdout + res.stderr
    report = json.loads(res.stdout)
    assert report["ok"] is True
    assert [item["status"] for item in report["results"]] == ["applied", "applied"]
    assert report["results"][0]["count"] == 1
    assert views(out_ok) == ["买方应付款。", "乙方（卖方）应开票。"]

    out_bad = tmp / "bad.docx"
    res = run_cli(input_path, out_bad, "apply", edits_bad)
    assert res.returncode == 1, res.stdout + res.stderr
    report = json.loads(res.stdout)
    assert report["ok"] is False
    assert report["results"][1]["status"] == "failed"
    assert "not found" in report["results"][1]["error"].lower()
    assert not out_bad.exists()

    out_keep = tmp / "keep.docx"
    res = run_cli(input_path, out_keep, "apply", edits_bad, "--keep-going")
    assert res.returncode == 2, res.stdout + res.stderr
    assert out_keep.exists()
    assert views(out_keep) == ["买方应付款。", "乙方应开票。"]

    out_dry = tmp / "dry.docx"
    res = run_cli(input_path, out_dry, "apply", edits_ok, "--dry-run")
    assert res.returncode == 0, res.stdout + res.stderr
    report = json.loads(res.stdout)
    assert report["dry_run"] is True
    assert not out_dry.exists()
    assert "甲方" in report["results"][0]["matches"][0]["context"]


def test_dry_run_context_shows_current_text(tmp):
    input_path = tmp / "in.docx"
    make_docx(input_path, {"document": [make_p("甲方应在3日内付款。")]})
    reviser = DocxReviser(str(input_path))
    results = reviser.apply_edits([
        {"op": "replace", "find": "3日内", "replace": "5个工作日内"},
        {"op": "replace", "find": "甲方", "replace": "买方"},
    ], dry_run=True)
    assert [item["status"] for item in results] == ["applied", "applied"]
    context = results[1]["matches"][0]["context"]
    assert "甲方应在5个工作日内付款" in context, context


def test_cli_delete_para_replace_full_and_occurrence(tmp):
    input_path = tmp / "in.docx"
    make_docx(input_path, {"document": [make_p("甲方与甲方"), make_p("待删除"), make_p("全量旧文")]})

    out1 = tmp / "o1.docx"
    res = run_cli(input_path, out1, "replace", "甲方", "乙方", "--occurrence", "2")
    assert res.returncode == 0, res.stdout + res.stderr
    assert views(out1)[0] == "甲方与乙方"

    out2 = tmp / "o2.docx"
    res = run_cli(input_path, out2, "delete_para", "待删除", "--startswith")
    assert res.returncode == 0, res.stdout + res.stderr
    assert views(out2) == ["甲方与甲方", "全量旧文"]

    out3 = tmp / "o3.docx"
    res = run_cli(input_path, out3, "replace_full", "全量旧文", "全量新文")
    assert res.returncode == 0, res.stdout + res.stderr
    root = parse_part(out3, "word/document.xml")
    assert root.xpath(".//w:del//w:delText/text()", namespaces=NS) == ["全量旧文"]
    assert root.xpath(".//w:ins//w:t/text()", namespaces=NS) == ["全量新文"]


def test_bold_none_inherits_and_false_clears(tmp):
    input_path = tmp / "in.docx"
    output_path = tmp / "out.docx"
    make_docx(input_path, {"document": [make_p(make_run("加粗锚点", bold=True))]})
    reviser = DocxReviser(str(input_path))
    reviser.insert_after("加粗锚点", "继承加粗", startswith=True)
    reviser.insert_after("加粗锚点", "不加粗", startswith=True, bold=False)
    reviser.save(str(output_path))
    root = parse_part(output_path, "word/document.xml")
    inherit = root.xpath(".//w:t[text()='继承加粗']/ancestor::w:r[1]", namespaces=NS)[0]
    cleared = root.xpath(".//w:t[text()='不加粗']/ancestor::w:r[1]", namespaces=NS)[0]
    assert effective_bold(inherit)
    assert not effective_bold(cleared)


def test_context_manager_and_atomic_save(tmp):
    input_path = tmp / "in.docx"
    make_docx(input_path, {"document": [make_p("旧文本")]})
    with zipfile.ZipFile(input_path) as zf:
        original_names = zf.namelist()
    with DocxReviser(str(input_path)) as reviser:
        temp_dir = reviser.temp_dir
        reviser.replace_text("旧文本", "新文本")
        reviser.save(str(input_path))
    assert not os.path.exists(temp_dir)
    with zipfile.ZipFile(input_path) as zf:
        names = zf.namelist()
    assert names[0] == "[Content_Types].xml", names
    expected = ["[Content_Types].xml"] + [n for n in original_names if n != "[Content_Types].xml"]
    assert [n for n in names if n in original_names] == expected, names
    assert views(input_path) == ["新文本"]
    assert not list(tmp.glob("*.tmp*"))


TESTS = [
    test_core_operations,
    test_newline_is_paragraph_break,
    test_author_and_cli,
    test_existing_revisions_are_preserved,
    test_parts_selection,
    test_revision_dates_are_beijing_time,
    test_replace_all_terminates_when_find_survives_diff,
    test_insert_text_after_replace_all,
    test_smart_replace_all_covers_every_paragraph,
    test_replace_all_substring_every_occurrence,
    test_replace_all_is_not_cascading,
    test_invalid_find_text_is_rejected,
    test_textbox_text_edited_once,
    test_zero_width_runs_inside_range_are_preserved,
    test_multi_content_runs_are_not_duplicated,
    test_split_runs_keep_edge_spaces,
    test_cross_container_match_is_rejected,
    test_inserted_paragraphs_are_tracked_and_sanitized,
    test_reject_all_restores_original,
    test_revision_ids_are_unique,
    test_replace_paragraph_keeps_annotations,
    test_delete_paragraph_keeps_annotations_and_textbox_text,
    test_new_text_is_sanitized,
    test_occurrence_selects_nth_match,
    test_ambiguous_match_warns,
    test_toc_paragraphs_are_skipped,
    test_not_found_error_suggests_closest_text,
    test_smart_diff_merges_short_equal_islands,
    test_cli_apply_batch,
    test_dry_run_context_shows_current_text,
    test_cli_delete_para_replace_full_and_occurrence,
    test_bold_none_inherits_and_false_clears,
    test_context_manager_and_atomic_save,
]


def _on_timeout(signum, frame):
    raise TimeoutError(f"test exceeded {TEST_TIMEOUT_SECONDS}s (possible infinite loop)")


def main():
    selected = [t for t in TESTS if not sys.argv[1:] or t.__name__ in sys.argv[1:]]
    tmp = Path(tempfile.mkdtemp(prefix="word_reviser_test_"))
    has_alarm = hasattr(signal, "SIGALRM")
    if has_alarm:
        signal.signal(signal.SIGALRM, _on_timeout)
    failures = []
    try:
        for test in selected:
            case_dir = tmp / test.__name__
            case_dir.mkdir()
            if has_alarm:
                signal.alarm(TEST_TIMEOUT_SECONDS)
            try:
                test(case_dir)
                print(f"PASS {test.__name__}")
            except Exception:
                failures.append(test.__name__)
                print(f"FAIL {test.__name__}")
                traceback.print_exc(limit=3)
            finally:
                if has_alarm:
                    signal.alarm(0)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    if failures:
        print(f"{len(failures)} of {len(selected)} test(s) failed: {', '.join(failures)}")
        sys.exit(1)
    print(f"word_reviser tests passed ({len(selected)} tests)")


if __name__ == "__main__":
    main()
