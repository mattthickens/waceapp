from __future__ import annotations

import hashlib
import io
import json
import re
from collections import defaultdict
from pathlib import Path
from typing import Dict, List, Optional

import fitz  # PyMuPDF
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = ROOT / "generated" / "resource-previews"
DOCX_CACHE_DIR = ROOT / "generated" / "docx-pdf-cache"
MANIFEST_PATH = ROOT / "generated" / "resource-manifest.js"
STUDY_MANIFEST_PATH = ROOT / "generated" / "pdf-study-manifest.js"

# applications.html lives one directory above ROOT, so every web-facing path
# (image/pdf src referenced from the manifest) needs this prefix.
WEB_PREFIX = ROOT.name


def web_path(rel: Optional[str]) -> Optional[str]:
    return f"{WEB_PREFIX}/{rel}" if rel else rel


EXCLUDE_NAME_RE = re.compile(r"syllabus", re.I)

TOPICS = [
    {
        "id": "a11",
        "code": "1.1",
        "name": "Consumer arithmetic",
        "keywords": [
            "percentage", "percent", "salary", "wage", "wages", "overtime", "commission",
            "piecework", "simple interest", "compound interest", "inflation", "cpi",
            "consumer price index", "mark-up", "markup", "discount", "gst", "profit", "loss",
            "currency", "exchange rate", "dividend", "share", "price-to-earnings", "p/e ratio",
            "spreadsheet", "budget", "wage-sheet", "taxable income", "net pay", "gross pay",
            "allowance", "pension", "earn", "earns", "earning", "income", "rate of pay",
            "per hour", "hourly rate", "timesheet", "annual salary", "weekly", "fortnightly",
            "cost price", "selling price", "retail price", "sale price", "interest rate",
            "invest", "investment", "loan", "repayment", "deposit", "withdraw", "bank account",
        ],
    },
    {
        "id": "a12",
        "code": "1.2",
        "name": "Algebra and matrices",
        "keywords": [
            "substitute", "formula", "subject of the formula", "transposition", "matrix",
            "matrices", "row matrix", "column matrix", "identity matrix", "leading diagonal",
            "scalar multiplication", "matrix multiplication", "order of a matrix", "zero matrix",
            "table of values", "algebraic expression", "expand the expression",
            "simplify the expression", "pronumeral",
        ],
    },
    {
        "id": "a13",
        "code": "1.3",
        "name": "Shape and measurement",
        "keywords": [
            "pythagoras", "pythagorean", "hypotenuse", "perimeter", "area of a circle",
            "area of a sector", "composite shape", "volume", "surface area", "cylinder", "cone",
            "sphere", "prism", "pyramid", "similar figures", "similarity", "scale factor",
            "scale drawing", "similar triangles", "sector", "radius", "diameter",
            "circumference", "composite figure", "net of a solid", "capacity", "litres",
            "millimetres", "centimetres", "kilometres", "rectangle", "parallelogram",
            "trapezium", "right-angled", "right angled triangle",
        ],
    },
    {
        "id": "a21",
        "code": "2.1",
        "name": "Univariate data analysis",
        "keywords": [
            "categorical", "numerical variable", "discrete", "continuous", "dot plot",
            "stem plot", "bar chart", "histogram", "modality", "skewed", "mean",
            "standard deviation", "standard score", "z-score", "normal distribution", "68%",
            "95%", "99.7%", "quantile", "box plot", "boxplot", "iqr", "interquartile range",
            "outlier", "five-number summary", "median", "statistical investigation",
            "data set", "survey", "sample", "population", "frequency", "frequency table",
        ],
    },
    {
        "id": "a22",
        "code": "2.2",
        "name": "Applications of trigonometry",
        "keywords": [
            "sine rule", "cosine rule", "heron's rule", "herons rule", "angle of elevation",
            "angle of depression", "bearing", "non-right-angled", "right-angled triangle",
            "triangulation", "area of a triangle", "trigonometric ratio", "elevation",
            "depression", "angle of inclination",
        ],
    },
    {
        "id": "a23",
        "code": "2.3",
        "name": "Linear equations and their graphs",
        "keywords": [
            "linear equation", "straight-line graph", "gradient", "slope", "intercept",
            "simultaneous", "break-even", "piece-wise", "piecewise", "step graph",
            "linear relationship", "x-intercept", "y-intercept", "line graph", "rearrange",
            "solve for", "solve the equation",
        ],
    },
]


def strip_duplicate_marker(text: str) -> str:
    """Strip a trailing download-duplicate marker, e.g. 'Exam (2)' -> 'Exam'."""
    return re.sub(r"\s*\(\d+\)\s*$", "", text).strip()


def strip_trailing_year(text: str) -> str:
    """Strip a lone trailing year, e.g. '...Exam CA 2017' -> '...Exam CA'."""
    return re.sub(r"\s+(19|20)\d{2}\s*$", "", text).strip()


def slugify(text: str) -> str:
    text = text.lower().strip()
    text = re.sub(r"\.(pdf|docx?|pptx?)$", "", text)
    text = strip_duplicate_marker(text)
    text = strip_trailing_year(text)
    text = re.sub(r"[\[\]()+_=,.'!@#$%^&*{}|\\/?;:~`-]", " ", text)
    text = re.sub(
        r"\b(solutions?|solns?|sols?|solution key|marking guide|marking|answers?|answer key|"
        r"questions?|mk|mg|final|copy|v\d+\.\d+|v\d+|generated)\b",
        " ",
        text,
    )
    text = re.sub(r"\s+", "-", text)
    text = re.sub(r"-+", "-", text)
    return text.strip("-") or "resource"


def display_title(name: str) -> str:
    title = re.sub(r"\.(pdf|docx?|pptx?)$", "", name, flags=re.I)
    title = re.sub(
        r"\b(solutions?|solns?|sols?|solution key|marking guide|marking|answers?|answer key|"
        r"questions?|mk|mg|final|copy)\b",
        "",
        title,
        flags=re.I,
    )
    title = re.sub(r"\s+", " ", title)
    return title.strip() or name


def is_solution(name: str) -> bool:
    return bool(re.search(r"soln|solution|marking|answer|answers|\bsol\b|\bsols\b|\bmk\b|\bmg\b", name, flags=re.I))


def calculator_label(text: str) -> str:
    lower = text.lower()
    if re.search(r"non[\s-]?calc|calc\s*-?\s*free|calculator\s*-?\s*free|\bcf\b", lower):
        return "Calculator-free"
    if re.search(r"calc\s*-?\s*assumed|calculator\s*-?\s*assumed|\bca\b", lower):
        return "Calculator-assumed"
    return "Unlabelled"


def extract_section_number(text: str) -> Optional[int]:
    """WA exams split into Section One (calculator-free) and Section Two
    (calculator-assumed); filenames spell this out very inconsistently."""
    lower = text.lower()
    if re.search(r"\bs1\b|\bsec(tion)?\s*1\b|\bsection\s*one\b|\bpart\s*1\b", lower):
        return 1
    if re.search(r"\bs2\b|\bsec(tion)?\s*2\b|\bsection\s*two\b|\bpart\s*2\b", lower):
        return 2
    return None


def infer_topic(text: str) -> Dict[str, str]:
    lower = text.lower()
    best = None
    best_score = 0.0
    for topic in TOPICS:
        # Multi-word phrases ("sample space") are far more topic-specific than
        # bare single words ("event", "circle", "minimum"), which are prone to
        # matching exam boilerplate (e.g. "Circle your teacher's name"); weight
        # phrase hits higher so a couple of real signals beat incidental noise.
        score = sum((2 if " " in kw else 1) for kw in topic["keywords"] if kw.lower() in lower)
        if score > best_score:
            best_score = score
            best = topic
    if best is None:
        return {"id": "", "code": "", "name": "Topic not inferred"}
    return {"id": best["id"], "code": best["code"], "name": best["name"]}


def is_locked_office_temp(name: str) -> bool:
    return name.startswith("~$")


def convert_docx_to_pdf(docx_path: Path) -> Optional[Path]:
    rel = docx_path.relative_to(ROOT).as_posix()
    digest = hashlib.sha1(rel.encode("utf-8")).hexdigest()[:10]
    stem = slugify(docx_path.name)[:60].strip("-") or "resource"
    out_path = DOCX_CACHE_DIR / f"{stem}-{digest}.pdf"
    if out_path.exists() and out_path.stat().st_mtime >= docx_path.stat().st_mtime:
        return out_path
    DOCX_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    from docx2pdf import convert  # imported lazily; only needed when .docx files are present

    try:
        convert(str(docx_path), str(out_path))
    except Exception as exc:  # pragma: no cover - depends on local Word install
        print(f"  WARN: could not convert {rel} to PDF ({exc})")
        return None
    return out_path if out_path.exists() else None


def discover_documents() -> List[Dict[str, Path]]:
    """Returns a list of {"original": Path, "pdf": Path} for every exam/test document."""
    docs: List[Dict[str, Path]] = []
    for path in sorted(ROOT.rglob("*")):
        if not path.is_file():
            continue
        if "generated" in path.parts:
            continue
        if is_locked_office_temp(path.name):
            continue
        if EXCLUDE_NAME_RE.search(path.name):
            continue
        suffix = path.suffix.lower()
        if suffix == ".pdf":
            docs.append({"original": path, "pdf": path})
        elif suffix == ".docx":
            pdf_path = convert_docx_to_pdf(path)
            if pdf_path is not None:
                docs.append({"original": path, "pdf": pdf_path})
    return docs


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


def render_preview(pdf_path: Path, out_path: Path) -> None:
    render_page_preview(pdf_path, 0, out_path)


def extract_page_text(pdf_path: Path, page_index: int) -> str:
    doc = fitz.open(pdf_path)
    try:
        if page_index < 0 or page_index >= doc.page_count:
            return ""
        return doc.load_page(page_index).get_text("text") or ""
    finally:
        doc.close()


def extract_first_page_text(pdf_path: Path) -> str:
    return extract_page_text(pdf_path, 0)


def clean_page_text(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


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


BOILERPLATE_ONLY_RE = re.compile(
    r"supplementary page|left blank intentionally|structure of this paper|this page has been left blank",
    re.I,
)


def should_include_page(text: str, page_index: int) -> bool:
    cleaned = (text or "").strip()
    if not cleaned:
        return False
    has_question_marker = bool(re.search(r"question\s+\d+", cleaned, flags=re.I))
    if BOILERPLATE_ONLY_RE.search(cleaned) and not has_question_marker:
        return False
    if page_index == 0 and not has_question_marker:
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


def extract_question_number(text: str) -> Optional[int]:
    match = re.search(r"Question\s+(\d{1,3})\b", text, flags=re.I)
    if match:
        return int(match.group(1))
    return None


def build_solution_page_map(solution_doc) -> Dict[int, int]:
    """Solutions booklets are often far denser than the question booklet
    (several worked answers per page where the questions had one per page),
    so naive page-index alignment drifts badly. Scan the solution doc once
    for "Question N" / "N. (a)" style markers and map question number ->
    the solution page that actually answers it."""
    page_map: Dict[int, int] = {}
    for page_index in range(solution_doc.page_count):
        text = solution_doc.load_page(page_index).get_text("text") or ""
        for match in re.finditer(r"Question\s+(\d{1,3})\b", text, flags=re.I):
            num = int(match.group(1))
            if num not in page_map:
                page_map[num] = page_index
        for match in re.finditer(r"(?m)^\s*(\d{1,3})\.\s", text):
            num = int(match.group(1))
            if num not in page_map:
                page_map[num] = page_index
    return page_map


MAX_SLUG_LEN = 60


def stable_name(relative_path: str, role: str) -> str:
    stem = slugify(relative_path)[:MAX_SLUG_LEN].strip("-") or "resource"
    digest = hashlib.sha1(relative_path.encode("utf-8")).hexdigest()[:10]
    return f"{stem}-{role}-{digest}.jpg"


def stable_page_name(relative_path: str, role: str, page_number: int) -> str:
    stem = slugify(relative_path)[:MAX_SLUG_LEN].strip("-") or "resource"
    digest = hashlib.sha1(f"{relative_path}:{page_number}".encode("utf-8")).hexdigest()[:10]
    return f"{stem}-p{page_number:03d}-{role}-{digest}.jpg"


def slug_tokens(text: str) -> set:
    return set(t for t in slugify(text).split("-") if t)


def find_fallback_pairs(groups: Dict[str, Dict[str, object]]) -> None:
    """Same-folder fallback pass: pair leftover question/solution orphans that
    exact-key matching missed (genuinely different naming conventions), using
    folder + calculator-label + section-number + token-overlap as a
    conservative heuristic."""
    by_folder: Dict[str, List[str]] = defaultdict(list)
    for key, entry in groups.items():
        folder = key.rsplit("::", 1)[0]
        by_folder[folder].append(key)

    for folder, keys in by_folder.items():
        orphan_q_keys = [k for k in keys if groups[k]["question"] is not None and groups[k]["solution"] is None]
        orphan_s_keys = [k for k in keys if groups[k]["solution"] is not None and groups[k]["question"] is None]
        if not orphan_q_keys or not orphan_s_keys:
            continue

        used_solution_keys: set = set()
        for q_key in orphan_q_keys:
            q_doc = groups[q_key]["question"]
            q_name = q_doc["original"].name
            q_calc = calculator_label(q_name)
            q_section = extract_section_number(q_name)
            q_tokens = slug_tokens(q_name)

            best_key = None
            best_score = 0.0
            for s_key in orphan_s_keys:
                if s_key in used_solution_keys:
                    continue
                s_doc = groups[s_key]["solution"]
                s_name = s_doc["original"].name
                s_calc = calculator_label(s_name)
                if q_calc != "Unlabelled" and s_calc != "Unlabelled" and q_calc != s_calc:
                    continue
                s_section = extract_section_number(s_name)
                if q_section is not None and s_section is not None and q_section != s_section:
                    continue
                s_tokens = slug_tokens(s_name)
                if not q_tokens or not s_tokens:
                    continue
                overlap = len(q_tokens & s_tokens) / len(q_tokens | s_tokens)
                # By-elimination bonus: if this is the only candidate left in the
                # folder, allow a much lower bar (e.g. "2018 Test 3" vs "Test 3
                # solutions" sharing every token except the year).
                if len(orphan_q_keys) == 1 and len(orphan_s_keys) == 1:
                    overlap = max(overlap, 0.5)
                # Matching section numbers (Section One/Two) are a strong signal
                # even when the rest of the filename uses a totally different
                # naming convention.
                if q_section is not None and s_section is not None and q_section == s_section:
                    overlap = max(overlap, 0.5)
                if overlap > best_score:
                    best_score = overlap
                    best_key = s_key

            if best_key is not None and best_score >= 0.45:
                groups[q_key]["solution"] = groups[best_key]["solution"]
                groups[q_key]["files"].extend(groups[best_key]["files"])
                used_solution_keys.add(best_key)

        for s_key in used_solution_keys:
            del groups[s_key]


def build_pdf_study_cards(groups: Dict[str, Dict[str, object]]) -> List[Dict[str, object]]:
    study_cards: List[Dict[str, object]] = []
    for key in sorted(groups):
        entry = groups[key]
        question = entry["question"]
        solution = entry["solution"]
        if question is None:
            continue

        q_original: Path = question["original"]
        q_pdf: Path = question["pdf"]
        s_original: Optional[Path] = solution["original"] if solution else None
        s_pdf: Optional[Path] = solution["pdf"] if solution else None

        question_doc = fitz.open(q_pdf)
        solution_doc = fitz.open(s_pdf) if s_pdf is not None else None
        try:
            question_rel = q_pdf.relative_to(ROOT).as_posix()
            solution_rel = s_pdf.relative_to(ROOT).as_posix() if s_pdf is not None else None
            folder = q_original.parent.relative_to(ROOT).as_posix()
            source_display = display_title(q_original.name)
            question_pages = question_doc.page_count
            solution_pages = solution_doc.page_count if solution_doc is not None else 0
            solution_page_map = build_solution_page_map(solution_doc) if solution_doc is not None else {}

            current_card: Optional[Dict[str, object]] = None
            for page_index in range(question_pages):
                page_text = extract_page_text(q_pdf, page_index)
                if not should_include_page(page_text, page_index):
                    continue

                # A page with no "Question N" marker is a continuation of the
                # previous question (e.g. part (ii)/(iii) printed on the next
                # page) rather than a new standalone question. Showing it on
                # its own loses the context needed to even understand what's
                # being asked, so fold its text into the card we already have
                # instead of creating an orphan fragment.
                has_marker = bool(re.search(r"question\s+\d+", page_text, flags=re.I))
                if not has_marker and current_card is not None:
                    merged_text = f"{current_card['_rawText']} {clean_page_text(page_text)}"
                    current_card["_rawText"] = merged_text
                    topic = infer_topic(f"{source_display} {folder} {merged_text}")
                    current_card["topicId"] = topic["id"]
                    current_card["topicCode"] = topic["code"]
                    current_card["topicName"] = topic["name"]
                    current_card["topicLabel"] = f"Topic {topic['code']} · {topic['name']}" if topic["code"] else "Topic not inferred"
                    current_card["questionText"] = merged_text[:2000]
                    current_card["marks"] = current_card["marks"] + parse_marks(page_text)
                    continue

                if current_card is not None:
                    study_cards.append(current_card)
                    current_card = None

                topic = infer_topic(f"{source_display} {folder} {page_text}")
                calculator = calculator_label(f"{source_display} {folder} {page_text}")
                page_number = page_index + 1
                question_image_rel = (OUTPUT_DIR / stable_page_name(question_rel, "question", page_number)).relative_to(ROOT).as_posix()
                render_page_preview(q_pdf, page_index, ROOT / question_image_rel)

                solution_image_rel = None
                solution_text = ""
                if solution_doc is not None and solution_pages > 0:
                    question_number = extract_question_number(page_text)
                    if question_number is not None and question_number in solution_page_map:
                        solution_page_index = solution_page_map[question_number]
                    else:
                        solution_page_index = pick_solution_page_index(page_index, question_pages, solution_pages)
                    if solution_page_index >= 0:
                        solution_image_rel = (OUTPUT_DIR / stable_page_name(solution_rel or question_rel, "solution", page_number)).relative_to(ROOT).as_posix()
                        render_page_preview(s_pdf, solution_page_index, ROOT / solution_image_rel)
                        solution_text = clean_page_text(extract_page_text(s_pdf, solution_page_index))

                page_excerpt = clean_page_text(page_text)
                page_title = page_title_from_text(page_text, page_number)
                current_card = {
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
                    "_rawText": page_excerpt,
                    "questionText": page_excerpt[:2000],
                    "answerText": solution_text[:1200] if solution_text else "See solution screenshot",
                    "questionImage": web_path(question_image_rel),
                    "solutionImage": web_path(solution_image_rel),
                    "questionPdf": web_path(question_rel),
                    "solutionPdf": web_path(solution_rel),
                    "marks": parse_marks(page_text),
                }

            if current_card is not None:
                study_cards.append(current_card)
        finally:
            question_doc.close()
            if solution_doc is not None:
                solution_doc.close()

    for card in study_cards:
        card.pop("_rawText", None)
    return study_cards


def main() -> None:
    documents = discover_documents()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    groups: Dict[str, Dict[str, object]] = defaultdict(lambda: {"question": None, "solution": None, "files": []})
    for doc in documents:
        original = doc["original"]
        relative = original.relative_to(ROOT).as_posix()
        folder = original.parent.relative_to(ROOT).as_posix()
        base = original.name
        base_key = slugify(base)
        key = f"{folder}::{base_key}"
        entry = groups[key]
        entry["files"].append(doc)
        if is_solution(base):
            if entry["solution"] is None:
                entry["solution"] = doc
        else:
            if entry["question"] is None:
                entry["question"] = doc

    find_fallback_pairs(groups)

    manifest: List[Dict[str, object]] = []
    rendered = 0
    for key in sorted(groups):
        entry = groups[key]
        question = entry["question"]
        solution = entry["solution"]
        any_doc = question or solution or entry["files"][0]
        any_original: Path = any_doc["original"]
        any_pdf: Path = any_doc["pdf"]
        display = display_title(any_original.name)
        folder = any_original.parent.relative_to(ROOT).as_posix()
        page_text = extract_first_page_text(any_pdf)
        topic = infer_topic(f"{display} {folder} {page_text}")
        calculator = calculator_label(f"{display} {folder} {page_text}")

        question_img = None
        solution_img = None
        question_pdf = None
        solution_pdf = None

        if question is not None:
            q_pdf: Path = question["pdf"]
            question_rel = q_pdf.relative_to(ROOT).as_posix()
            question_pdf = question_rel
            question_img = (OUTPUT_DIR / stable_name(question_rel, "question")).relative_to(ROOT).as_posix()
            render_preview(q_pdf, ROOT / question_img)
            rendered += 1
        if solution is not None:
            s_pdf: Path = solution["pdf"]
            solution_rel = s_pdf.relative_to(ROOT).as_posix()
            solution_pdf = solution_rel
            solution_img = (OUTPUT_DIR / stable_name(solution_rel, "solution")).relative_to(ROOT).as_posix()
            render_preview(s_pdf, ROOT / solution_img)
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
                "questionPdf": web_path(question_pdf),
                "solutionPdf": web_path(solution_pdf),
                "questionImage": web_path(question_img),
                "solutionImage": web_path(solution_img),
                "fileCount": len(entry["files"]),
            }
        )

    study_cards = build_pdf_study_cards(groups)
    payload = json.dumps(manifest, ensure_ascii=False, indent=2)
    study_payload = json.dumps(study_cards, ensure_ascii=False, indent=2)
    MANIFEST_PATH.parent.mkdir(parents=True, exist_ok=True)
    MANIFEST_PATH.write_text(f"window.APPS_RESOURCE_MANIFEST = {payload};\nwindow.APPS_RESOURCE_MANIFEST_UPDATED = {json.dumps('2026-06-25')};\n", encoding="utf-8")
    STUDY_MANIFEST_PATH.write_text(f"window.APPS_PDF_STUDY_MANIFEST = {study_payload};\nwindow.APPS_PDF_STUDY_MANIFEST_UPDATED = {json.dumps('2026-06-25')};\n", encoding="utf-8")
    print(f"Wrote {len(manifest)} resource entries, {len(study_cards)} study cards, and rendered {rendered} previews to {OUTPUT_DIR}")
    missing = sum(1 for e in groups.values() if e["question"] is not None and e["solution"] is None)
    print(f"Question groups still missing a solution: {missing} / {sum(1 for e in groups.values() if e['question'] is not None)}")


if __name__ == "__main__":
    main()
