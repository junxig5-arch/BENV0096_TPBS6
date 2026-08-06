# -*- coding: utf-8 -*-
from pathlib import Path
import argparse
import re
import shutil
import tempfile
import zipfile

from lxml import etree
from pypdf import PdfReader


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()


def paragraph_text(p):
    return "".join(t.text or "" for t in p.xpath(".//w:t", namespaces=NS))


def paragraph_style(p):
    node = p.find("./w:pPr/w:pStyle", namespaces=NS)
    return node.get(f"{{{NS['w']}}}val") if node is not None else ""


def strip_page_number(text):
    return re.sub(r"\s*\d+\s*$", "", text).strip()


def patch_last_page_text(p, page):
    texts = p.xpath(".//w:t", namespaces=NS)
    for t in reversed(texts):
        value = t.text or ""
        if re.search(r"\d+\s*$", value):
            t.text = re.sub(r"\d+\s*$", str(page), value)
            return True
    if texts:
        texts[-1].text = (texts[-1].text or "") + f"\t{page}"
        return True
    return False


def pdf_page_texts(pdf):
    reader = PdfReader(str(pdf))
    return [norm(page.extract_text() or "") for page in reader.pages]


def find_page(texts, entry, start_page=1):
    entry_n = norm(entry)
    m = re.match(r"^(Figure|Table)\s+\d+\.\d+\.", entry_n)
    if m:
        entry_n = m.group(0)
    for idx, page_text in enumerate(texts, start=1):
        if idx >= start_page and entry_n and entry_n in page_text:
            return idx
    short = entry_n[:70]
    for idx, page_text in enumerate(texts, start=1):
        if idx >= start_page and short and short in page_text:
            return idx
    return None


def patch(docx, pdf):
    texts = pdf_page_texts(pdf)
    tmp = Path(tempfile.mkdtemp(prefix="docx_pagecache_"))
    with zipfile.ZipFile(docx, "r") as zin:
        zin.extractall(tmp)
    doc_xml = tmp / "word" / "document.xml"
    tree = etree.parse(str(doc_xml))
    root = tree.getroot()
    toc_updates = {}
    list_updates = {}
    in_fig_table_list = False
    for p in root.xpath(".//w:p", namespaces=NS):
        text = paragraph_text(p)
        style = paragraph_style(p)
        clean = strip_page_number(text)
        if not clean:
            continue
        if clean == "List of Figures and Tables":
            in_fig_table_list = True
        elif clean == "Abbreviations":
            in_fig_table_list = False

        if style in {"TOC1", "TOC2", "toc 1", "toc 2"}:
            if clean in {"Declaration", "Ethics and Data Protection Statement", "Acknowledgements", "Abstract"}:
                page = find_page(texts, clean, start_page=2)
            elif clean == "List of Figures and Tables":
                page = find_page(texts, clean, start_page=4)
            elif clean == "Abbreviations":
                page = find_page(texts, clean, start_page=5)
            else:
                page = find_page(texts, clean, start_page=6)
            if page:
                toc_updates[clean] = page
                patch_last_page_text(p, page)
        elif in_fig_table_list and (clean.startswith("Figure ") or clean.startswith("Table ")):
            page = find_page(texts, clean, start_page=6)
            if page:
                list_updates[clean[:90]] = page
                patch_last_page_text(p, page)

    tree.write(str(doc_xml), encoding="UTF-8", xml_declaration=True, standalone=True)
    backup = docx.with_suffix(".pagecache.bak.docx")
    shutil.copy2(docx, backup)
    with zipfile.ZipFile(docx, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for file in tmp.rglob("*"):
            if file.is_file():
                zout.write(file, file.relative_to(tmp).as_posix())
    return toc_updates, list_updates, backup


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx", type=Path)
    ap.add_argument("pdf", type=Path)
    args = ap.parse_args()
    toc, lists, backup = patch(args.docx, args.pdf)
    print(f"patched={args.docx}")
    print(f"backup={backup}")
    print(f"toc_updates={len(toc)}")
    print(f"list_updates={len(lists)}")
    for k, v in list(toc.items())[:80]:
        print(f"TOC {k} -> {v}")
    for k, v in list(lists.items())[:80]:
        print(f"LIST {k} -> {v}")


if __name__ == "__main__":
    main()
