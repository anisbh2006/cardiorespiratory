#!/usr/bin/env python3
"""Inventory pass over the supplied course material (read-only analysis)."""
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

import fitz  # PyMuPDF

ROOT = "/Users/macbook/Desktop/cardioeng"
OUT = "/tmp/uei01-analysis/inventory.json"

A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def clean(t: str) -> str:
    return re.sub(r"\s+", " ", t).strip()


def analyze_pdf(path: str):
    doc = fitz.open(path)
    pages = []
    for pno, page in enumerate(doc):
        text = clean(page.get_text("text"))
        lines = [clean(l) for l in page.get_text("text").splitlines() if clean(l)]
        imgs = page.get_images(full=True)
        tables = []
        try:
            found = page.find_tables()
            tables = [len(t.extract()) for t in found.tables]
        except Exception:
            pass
        pages.append(
            {
                "page": pno + 1,
                "chars": len(text),
                "first_lines": lines[:6],
                "images": len(imgs),
                "tables": tables,
            }
        )
    info = {
        "kind": "pdf",
        "pages": len(doc),
        "total_images": sum(p["images"] for p in pages),
        "total_tables": sum(len(p["tables"]) for p in pages),
        "pages_detail": pages,
    }
    doc.close()
    return info


def analyze_pptx(path: str):
    z = zipfile.ZipFile(path)
    names = z.namelist()
    slides = sorted(
        [n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)],
        key=lambda n: int(re.search(r"(\d+)", n).group(1)),
    )
    media = [n for n in names if n.startswith("ppt/media/")]
    slide_infos = []
    for s in slides:
        root = ET.fromstring(z.read(s))
        texts = [clean(t.text or "") for t in root.iter(f"{A_NS}t") if (t.text or "").strip()]
        # images referenced by this slide
        rels_name = s.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels"
        n_img = 0
        if rels_name in names:
            rels = ET.fromstring(z.read(rels_name))
            for rel in rels:
                tgt = rel.get("Target", "")
                if "../media/" in tgt:
                    n_img += 1
        # tables
        n_tbl = len(list(root.iter(f"{A_NS}tbl")))
        slide_infos.append(
            {
                "slide": int(re.search(r"(\d+)", s).group(1)),
                "first_lines": texts[:6],
                "text_blocks": len(texts),
                "images": n_img,
                "tables": n_tbl,
            }
        )
    z.close()
    return {
        "kind": "pptx",
        "pages": len(slides),
        "total_images": len(media),
        "media_files": [os.path.basename(m) for m in media],
        "total_tables": sum(s["tables"] for s in slide_infos),
        "pages_detail": slide_infos,
    }


def main():
    inv = {}
    for dirpath, _dirnames, filenames in os.walk(ROOT):
        for fn in sorted(filenames):
            if fn == ".DS_Store":
                continue
            full = os.path.join(dirpath, fn)
            rel = os.path.relpath(full, ROOT)
            try:
                if fn.lower().endswith(".pdf"):
                    inv[rel] = analyze_pdf(full)
                elif fn.lower().endswith(".pptx"):
                    inv[rel] = analyze_pptx(full)
            except Exception as e:  # noqa: BLE001
                inv[rel] = {"kind": "error", "error": str(e)}
            print(f"done {rel}", file=sys.stderr)
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w") as f:
        json.dump(inv, f, ensure_ascii=False, indent=1)
    # summary
    for rel, info in inv.items():
        print(
            f"{rel} | {info.get('kind')} | pages={info.get('pages')} imgs={info.get('total_images')} tables={info.get('total_tables')}"
        )


if __name__ == "__main__":
    main()
