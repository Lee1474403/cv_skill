#!/usr/bin/env python3
"""Check searchable PDF text against paragraph-based DOCX content, not layout."""
from __future__ import annotations

import argparse
import json
import re
import unicodedata
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from pypdf import PdfReader


def windows_acl_inheritance_disabled(path: Path) -> bool | None:
    """Read SE_DACL_PROTECTED without changing file or directory permissions."""
    import os
    if os.name != "nt":
        return None
    import ctypes
    from ctypes import wintypes
    api = ctypes.WinDLL("advapi32", use_last_error=True)
    api.GetNamedSecurityInfoW.argtypes = [wintypes.LPWSTR, wintypes.DWORD, wintypes.DWORD,
                                         wintypes.LPVOID, wintypes.LPVOID, wintypes.LPVOID,
                                         wintypes.LPVOID, ctypes.POINTER(wintypes.LPVOID)]
    api.GetNamedSecurityInfoW.restype = wintypes.DWORD
    api.GetSecurityDescriptorControl.argtypes = [wintypes.LPVOID, ctypes.POINTER(wintypes.WORD),
                                                 ctypes.POINTER(wintypes.DWORD)]
    api.GetSecurityDescriptorControl.restype = wintypes.BOOL
    descriptor = wintypes.LPVOID()
    # GetFileSecurityW can omit the protection flag; use the named-object API.
    result = api.GetNamedSecurityInfoW(str(path.resolve()), 1, 4, None, None, None, None, ctypes.byref(descriptor))
    if result:
        raise ctypes.WinError(result)
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.LocalFree.argtypes = [wintypes.LPVOID]
    kernel.LocalFree.restype = wintypes.LPVOID
    try:
        control, revision = wintypes.WORD(), wintypes.DWORD()
        if not api.GetSecurityDescriptorControl(descriptor, ctypes.byref(control), ctypes.byref(revision)):
            raise ctypes.WinError(ctypes.get_last_error())
        return bool(control.value & 0x1000)
    finally:
        kernel.LocalFree(descriptor)


def normalize(text: str) -> str:
    return re.sub(r"\s+", "", unicodedata.normalize("NFC", text))


def verify(docx_path: Path, pdf_path: Path) -> dict:
    document = Document(docx_path)
    if document.tables:
        raise ValueError("This checker expects paragraph-based resumes; tables need a separate check.")
    # Read w:t nodes so hyperlink text is included even with python-docx 1.1.
    paragraphs = ["".join(node.text or "" for node in p._p.iter(qn("w:t"))).strip()
                  for p in document.paragraphs]
    paragraphs = [text for text in paragraphs if text]
    reader = PdfReader(pdf_path)
    texts = [page.extract_text() or "" for page in reader.pages]
    combined = normalize("\n".join(texts))
    source_text = normalize("".join(paragraphs)).replace("•", "")
    full_match = source_text == combined.replace("•", "")
    missing = []
    out_of_order = []
    cursor = 0
    for index, paragraph in enumerate(paragraphs, 1):
        needle = normalize(paragraph)
        position = combined.find(needle, cursor)
        if position >= 0:
            cursor = position + len(needle)
        elif needle not in combined:
            missing.append(index)
        else:
            out_of_order.append(index)
    return {
        "pdf_pages": len(reader.pages),
        "docx_paragraphs": len(paragraphs),
        "pdf_text_characters": sum(len(t) for t in texts),
        "missing_paragraph_indices": missing,
        "out_of_order_paragraph_indices": out_of_order,
        "blank_pdf_page_indices": [i for i, t in enumerate(texts, 1) if not t.strip()],
        "full_text_matches_ignoring_whitespace_and_bullet_markers": full_match,
        "text_consistent": bool(paragraphs) and bool(texts) and not missing and not out_of_order and full_match,
        "visual_layout_checked": False,
        "facts_checked": False,
        "windows_acl_inheritance_disabled": {
            "docx": windows_acl_inheritance_disabled(docx_path),
            "pdf": windows_acl_inheritance_disabled(pdf_path),
        },
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--docx", type=Path, required=True)
    parser.add_argument("--pdf", type=Path, required=True)
    args = parser.parse_args()
    try:
        report = verify(args.docx, args.pdf)
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return 0 if report["text_consistent"] and not report["blank_pdf_page_indices"] else 1
    except Exception as exc:
        parser.exit(1, f"Verification failed: {exc}\n")


if __name__ == "__main__":
    raise SystemExit(main())
