#!/usr/bin/env python3
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from word_reviser import DocxReviser  # noqa: E402

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"w": W}


def wtag(tag):
    return f"{{{W}}}{tag}"


def make_run(text, bold=False):
    r = etree.Element(wtag("r"))
    if bold:
        rpr = etree.SubElement(r, wtag("rPr"))
        etree.SubElement(rpr, wtag("b"))
    t = etree.SubElement(r, wtag("t"))
    t.text = text
    return r


def make_p(*texts, numbered=False):
    p = etree.Element(wtag("p"))
    if numbered:
        ppr = etree.SubElement(p, wtag("pPr"))
        numpr = etree.SubElement(ppr, wtag("numPr"))
        etree.SubElement(numpr, wtag("ilvl")).set(wtag("val"), "0")
        etree.SubElement(numpr, wtag("numId")).set(wtag("val"), "1")
    for text in texts:
        p.append(make_run(text))
    return p


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
    root = etree.Element(wtag("document"), nsmap={"w": W})
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


def make_docx(path, parts):
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
        for file in work.rglob("*"):
            if file.is_file():
                zf.write(file, file.relative_to(work))


def parse_part(docx, part):
    with zipfile.ZipFile(docx) as zf:
        return etree.fromstring(zf.read(part))


def accepted_text(root):
    return "".join(t.text or "" for t in root.xpath(".//w:t[not(ancestor::w:del)]", namespaces=NS))


def normal_text(root):
    return "".join(t.text or "" for t in root.xpath(".//w:t[not(ancestor::w:del) and not(ancestor::w:ins)]", namespaces=NS))


def authors(root):
    return set(root.xpath(".//w:ins/@w:author | .//w:del/@w:author", namespaces=NS))


def assert_integrity(docx, parts=("word/document.xml",)):
    for part in parts:
        root = parse_part(docx, part)
        assert not root.xpath(".//w:del//w:t", namespaces=NS), f"bad w:t in w:del for {part}"
        assert root.xpath(".//w:ins | .//w:del", namespaces=NS), f"missing revisions in {part}"
    settings = parse_part(docx, "word/settings.xml")
    assert settings.xpath(".//w:trackRevisions", namespaces=NS)


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
    assert doc.xpath(".//w:br", namespaces=NS)
    assert authors(doc) == {"Agent"}

    table_anchor = doc.xpath(".//w:t[text()='表格锚点']/ancestor::w:p[1]", namespaces=NS)[0]
    table_inserted = doc.xpath(".//w:t[text()='表格新增段']/ancestor::w:p[1]", namespaces=NS)[0]
    assert table_anchor.getparent() is table_inserted.getparent()
    list_inserted = doc.xpath(".//w:t[text()='列表后新增']/ancestor::w:p[1]", namespaces=NS)[0]
    assert not list_inserted.xpath("./w:pPr/w:numPr", namespaces=NS)
    assert doc.xpath(".//w:t[text()='新增标题']/ancestor::w:r[1]/w:rPr/w:b", namespaces=NS)


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


def main():
    tmp = Path(tempfile.mkdtemp(prefix="word_reviser_test_"))
    try:
        test_core_operations(tmp)
        test_author_and_cli(tmp)
        test_existing_revisions_are_preserved(tmp)
        test_parts_selection(tmp)
    finally:
        shutil.rmtree(tmp)
    print("word_reviser tests passed")


if __name__ == "__main__":
    main()
