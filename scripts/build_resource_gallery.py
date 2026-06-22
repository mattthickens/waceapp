from __future__ import annotations

import hashlib
import io
import json
import re
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional

import fitz  # PyMuPDF
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = ROOT / "generated" / "resource-previews"
MANIFEST_PATH = ROOT / "generated" / "resource-manifest.js"
STUDY_MANIFEST_PATH = ROOT / "generated" / "pdf-study-manifest.js"

TOPICS = [
    {
        "id": "t11",
        "code": "1.1",
        "name": "Counting & Probability",
        "keywords": [
            "probability",
            "combination",
            "permutation",
            "conditional",
            "independence",
            "sample space",
            "event",
            "set notation",
            "union",
            "intersection",
            "complement",
            "mutually exclusive",
            "pascal",
            "binomial",
            "counting",
            "ncr",
        ],
    },
    {
        "id": "t12",
        "code": "1.2",
        "name": "Functions & Graphs",
        "keywords": [
            "function",
            "domain",
            "range",
            "linear",
            "quadratic",
            "parabola",
            "hyperbola",
            "gradient",
            "slope",
            "intercept",
            "polynomial",
            "dilation",
            "translation",
            "transformation",
            "turning point",
            "roots",
            "cubic",
            "asymptote",
            "circle",
        ],
    },
    {
        "id": "t13",
        "code": "1.3",
        "name": "Trigonometric Functions",
        "keywords": [
            "sin",
            "cos",
            "tan",
            "trig",
            "radian",
            "degree",
            "unit circle",
            "amplitude",
            "period",
            "phase",
            "cosine rule",
            "sine rule",
            "angle",
            "exact values",
            "pi",
            "arcsin",
            "arccos",
            "arctan",
            "sector",
            "arc length",
        ],
    },
    {
        "id": "t21",
        "code": "2.1",
        "name": "Exponential Functions",
        "keywords": [
            "exponential",
            "index",
            "indices",
            "power",
            "base",
            "scientific notation",
            "growth",
            "decay",
            "index laws",
            "fractional index",
            "asymptote",
        ],
    },
    {
        "id": "t22",
        "code": "2.2",
        "name": "Sequences & Series",
        "keywords": [
            "sequence",
            "series",
            "arithmetic",
            "geometric",
            "common difference",
            "common ratio",
            "sum",
            "nth term",
            "recursive",
            "compound interest",
            "simple interest",
            "converge",
            "limiting",
            "t_n",
            "s_n",
            "first term",
        ],
    },
    {
        "id": "t23",
        "code": "2.3",
        "name": "Differential Calculus",
        "keywords": [
            "derivative",
            "differentiate",
            "differentiation",
            "dy/dx",
            "tangent",
            "instantaneous",
            "rate of change",
            "stationary",
            "maximum",
            "minimum",
            "calculus",
            "limit",
            "velocity",
            "displacement",
            "slope of tangent",
        ],
    },
]


def slugify(text: str) -> str:
    text = text.lower().strip()
    text = re.sub(r"\.(pdf|docx?|pptx?)$", "", text)
    text = re.sub(r"[\[\]()+_=,.'!@#$%^&*{}|\\/?;:~`-]", " ", text)
    text = re.sub(r"\b(solutions?|solns?|solution key|marking guide|marking|answers?|answer key|mk|mg|final|copy|v\d+\.\d+|v\d+|generated)\b", " ", text)
    text = re.sub(r"\s+", "-", text)
    text = re.sub(r"-+", "-", text)
    return text.strip("-") or "resource"


def display_title(name: str) -> str:
    title = re.sub(r"\.(pdf|docx?|pptx?)$", "", name, flags=re.I)
    title = re.sub(r"\b(solutions?|solns?|solution key|marking guide|marking|answers?|answer key|mk|mg|final|copy)\b", "", title, flags=re.I)
    title = re.sub(r"\s+", " ", title)
    return title.strip() or name


def is_solution(name: str) -> bool:
    return bool(re.search(r"soln|solution|marking|answer|answers|\bmk\b|\bmg\b", name, flags=re.I))


def calculator_label(text: str) -> str:
    lower = text.lower()
    if re.search(r"calc\s*-?\s*free|calculator\s*-?\s*free|\bcf\b", lower):
        return "Calculator-free"
    if re.search(r"calc\s*-?\s*assumed|calculator\s*-?\s*assumed|\bca\b", lower):
        return "Calculator-assumed"
    return "Unlabelled"


def infer_topic(text: str) -> Dict[str, str]:
    lower = text.lower()
    best = None
    best_score = 0
    for topic in TOPICS:
        score = sum(1 for kw in topic["keywords"] if kw.lower() in lower)
        if score > best_score:
            best_score = score
            best = topic
    if best is None:
        return {"id": "", "code": "", "name": "Topic not inferred"}
    return {"id": best["id"], "code": best["code"], "name": best["name"]}


def render_preview(pdf_path: Path, out_path: Path) -> None:
    doc = fitz.open(pdf_path)
    if doc.page_count == 0:
        raise ValueError(f"No pages found in {pdf_path}")
    page = doc.load_page(0)
    pix = page.get_pixmap(matrix=fitz.Matrix(1.6, 1.6), alpha=False)
    image = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    image.save(out_path, format="JPEG", quality=85, optimize=True, progressive=True)
    doc.close()


def extract_first_page_text(pdf_path: Path) -> str:
    doc = fitz.open(pdf_path)
    if doc.page_count == 0:
        doc.close()
        return ""
    text = doc.load_page(0).get_text("text")
    doc.close()
    return text or ""


def stable_name(relative_path: str, role: str) -> str:
    stem = slugify(relative_path)
    digest = hashlib.sha1(relative_path.encode("utf-8")).hexdigest()[:10]
    return f"{stem}-{role}-{digest}.jpg"


def stable_page_name(relative_path: str, role: str, page_number: int) -> str:
    stem = slugify(relative_path)
    digest = hashlib.sha1(f"{relative_path}:{page_number}".encode("utf-8")).hexdigest()[:10]
    return f"{stem}-p{page_number:03d}-{role}-{digest}.jpg"


def extract_page_text(pdf_path: Path, page_index: int) -> str:
    doc = fitz.open(pdf_path)
    try:
        if page_index < 0 or page_index >= doc.page_count:
            return ""
        return doc.load_page(page_index).get_text("text") or ""
    finally:
        doc.close()


def render_page_preview(pdf_path: Path, page_index: int, out_path: Path) -> None:
    doc = fitz.open(pdf_path)
    try:
        if page_index < 0 or page_index >= doc.page_count:
            raise ValueError(f"Page {page_index} out of range for {pdf_path}")
        page = doc.load_page(page_index)
        pix = page.get_pixmap(matrix=fitz.Matrix(1.55, 1.55), alpha=False)
        image = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
        out_path.parent.mkdir(parents=True, exist_ok=True)
        image.save(out_path, format="JPEG", quality=84, optimize=True, progressive=True)
    finally:
        doc.close()


def clean_page_text(text: str) -> str:
    text = re.sub(r"\s+", " ", text or "").strip()
    return text


def page_title_from_text(text: str, page_number: int) -> str:
    match = re.search(r"Question\s+(\d+[A-Za-z]?)", text, flags=re.I)
    if match:
        return f"Question {match.group(1)}"
    match = re.search(r"Q\.?\s*(\d+[A-Za-z]?)", text, flags=re.I)
    if match:
        return f"Question {match.group(1)}"
    return f"Page {page_number}"


def parse_marks(text: str) -> int:
    patterns = [
        r"(\d+)\s*marks?",
        r"=\s*(\d+)\s*marks?",
        r"\((?:\d+[,\s-]*)*(\d+)\s*marks?\)",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.I)
        if match:
            try:
                return int(match.group(1))
            except Exception:
                continue
    return 0


def should_include_page(text: str, page_index: int) -> bool:
    cleaned = (text or "").strip()
    if not cleaned:
        return False
    if page_index == 0 and not re.search(r"question\s+\d+", cleaned, flags=re.I):
        return False
    if len(cleaned) < 60 and page_index == 0:
        return False
    return True


def pick_solution_page_index(question_page_index: int, question_pages: int, solution_pages: int) -> int:
    if solution_pages <= 0:
        return -1
    if solution_pages == question_pages:
        return min(question_page_index, solution_pages - 1)
    if solution_pages == question_pages + 1:
        return min(question_page_index + 1, solution_pages - 1)
    return min(question_page_index, solution_pages - 1)


def build_pdf_study_cards(groups: Dict[str, Dict[str, object]]) -> List[Dict[str, object]]:
    study_cards: List[Dict[str, object]] = []
    for key in sorted(groups):
        entry = groups[key]
        question = entry["question"]
        solution = entry["solution"]
        if question is None:
            continue

        question_doc = fitz.open(question)
        solution_doc = fitz.open(solution) if solution is not None else None
        try:
            question_rel = question.relative_to(ROOT).as_posix()
            solution_rel = solution.relative_to(ROOT).as_posix() if solution is not None else None
            folder = question.parent.relative_to(ROOT).as_posix()
            source_display = display_title(question.name)
            question_pages = question_doc.page_count
            solution_pages = solution_doc.page_count if solution_doc is not None else 0

            for page_index in range(question_pages):
                page_text = extract_page_text(question, page_index)
                if not should_include_page(page_text, page_index):
                    continue
                topic = infer_topic(f"{source_display} {folder} {page_text}")
                calculator = calculator_label(f"{source_display} {folder} {page_text}")
                page_number = page_index + 1
                question_image_rel = (OUTPUT_DIR / stable_page_name(question_rel, "question", page_number)).relative_to(ROOT).as_posix()
                render_page_preview(question, page_index, ROOT / question_image_rel)

                solution_image_rel = None
                solution_text = ""
                if solution_doc is not None and solution_pages > 0:
                    solution_page_index = pick_solution_page_index(page_index, question_pages, solution_pages)
                    if solution_page_index >= 0:
                        solution_image_rel = (OUTPUT_DIR / stable_page_name(solution_rel or question_rel, "solution", page_number)).relative_to(ROOT).as_posix()
                        render_page_preview(solution, solution_page_index, ROOT / solution_image_rel)
                        solution_text = clean_page_text(extract_page_text(solution, solution_page_index))

                page_excerpt = clean_page_text(page_text)
                page_title = page_title_from_text(page_text, page_number)
                study_cards.append(
                    {
                        "id": f"{key}::p{page_number}",
                        "sourceId": key,
                        "title": page_title,
                        "folder": folder,
                        "sourceTitle": source_display,
                        "pageNumber": page_number,
                        "topicId": topic["id"],
                        "topicCode": topic["code"],
                        "topicName": topic["name"],
                        "topicLabel": f"Topic {topic['code']} · {topic['name']}" if topic["code"] else "Topic not inferred",
                        "calculator": calculator,
                        "questionText": page_excerpt[:1200],
                        "answerText": solution_text[:1200] if solution_text else "See solution screenshot",
                        "questionImage": question_image_rel,
                        "solutionImage": solution_image_rel,
                        "questionPdf": question_rel,
                        "solutionPdf": solution_rel,
                        "marks": parse_marks(page_text),
                    }
                )
        finally:
            question_doc.close()
            if solution_doc is not None:
                solution_doc.close()

    return study_cards


def main() -> None:
    pdfs = sorted(p for p in ROOT.rglob("*.pdf") if "generated" not in p.parts)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    groups = defaultdict(lambda: {"question": None, "solution": None, "files": []})
    for pdf in pdfs:
        relative = pdf.relative_to(ROOT).as_posix()
        folder = pdf.parent.relative_to(ROOT).as_posix()
        base = pdf.name
        base_key = slugify(base)
        if is_solution(base):
            base_key = re.sub(r"-(solutions?|solns?|solution|marking-guide|marking|answers?|answer-key|mk|mg|final|copy)(-v\d+\.\d+|-v\d+)?$", "", base_key)
        else:
            base_key = re.sub(r"-(final|copy)(-v\d+\.\d+|-v\d+)?$", "", base_key)
        key = f"{folder}::{base_key}"
        entry = groups[key]
        entry["files"].append(pdf)
        if is_solution(base):
            if entry["solution"] is None:
                entry["solution"] = pdf
        else:
            if entry["question"] is None:
                entry["question"] = pdf

    manifest: List[Dict[str, object]] = []
    rendered = 0
    for key in sorted(groups):
        entry = groups[key]
        question = entry["question"]
        solution = entry["solution"]
        any_pdf = question or solution or entry["files"][0]
        display = display_title(any_pdf.name)
        folder = any_pdf.parent.relative_to(ROOT).as_posix()
        page_text = extract_first_page_text(any_pdf)
        topic = infer_topic(f"{display} {folder} {page_text}")
        calculator = calculator_label(f"{display} {folder} {page_text}")

        question_img = None
        solution_img = None
        question_pdf = None
        solution_pdf = None

        if question is not None:
            question_rel = question.relative_to(ROOT).as_posix()
            question_pdf = question_rel
            question_img = (OUTPUT_DIR / stable_name(question_rel, "question")).relative_to(ROOT).as_posix()
            render_preview(question, ROOT / question_img)
            rendered += 1
        if solution is not None:
            solution_rel = solution.relative_to(ROOT).as_posix()
            solution_pdf = solution_rel
            solution_img = (OUTPUT_DIR / stable_name(solution_rel, "solution")).relative_to(ROOT).as_posix()
            render_preview(solution, ROOT / solution_img)
            rendered += 1

        manifest.append(
            {
                "id": key,
                "title": display,
                "folder": folder,
                "topicId": topic["id"],
                "topicCode": topic["code"],
                "topicName": topic["name"],
                "topicLabel": f"Topic {topic['code']} · {topic['name']}" if topic["code"] else "Topic not inferred",
                "calculator": calculator,
                "questionPdf": question_pdf,
                "solutionPdf": solution_pdf,
                "questionImage": question_img,
                "solutionImage": solution_img,
                "fileCount": len(entry["files"]),
            }
        )

    study_cards = build_pdf_study_cards(groups)
    payload = json.dumps(manifest, ensure_ascii=False, indent=2)
    study_payload = json.dumps(study_cards, ensure_ascii=False, indent=2)
    MANIFEST_PATH.parent.mkdir(parents=True, exist_ok=True)
    MANIFEST_PATH.write_text(f"window.RESOURCE_MANIFEST = {payload};\nwindow.RESOURCE_MANIFEST_UPDATED = {json.dumps('2026-06-22')};\n", encoding="utf-8")
    STUDY_MANIFEST_PATH.write_text(f"window.PDF_STUDY_MANIFEST = {study_payload};\nwindow.PDF_STUDY_MANIFEST_UPDATED = {json.dumps('2026-06-22')};\n", encoding="utf-8")
    print(f"Wrote {len(manifest)} resource entries, {len(study_cards)} study cards, and rendered {rendered} previews to {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
