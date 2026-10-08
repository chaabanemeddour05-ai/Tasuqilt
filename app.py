import os
import io
import csv
import json
import re
import difflib
from collections import Counter, defaultdict
from typing import List, Dict, Tuple
from urllib.parse import quote

import requests
import streamlit as st
from docx import Document
from openpyxl import load_workbook


# ============================================================
# TASUQILT DZ - FREE HYBRID TRANSLATION ENGINE
#
# GitHub (tm.xlsx / tm.docx / data/...)  ->  Translation Memory index
#   -> exact TM match?  YES: use it   NO: Gemini Free + retrieved context
#   -> Terminology validator -> final translation
#
# tm.xlsx format (sheet 1): Column A = French, Column B = Tamazight
# ============================================================


# ============================================================
# 1. PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Tasuqilt DZ 🇩🇿",
    page_icon="🇩🇿",
    layout="wide"
)


# ============================================================
# 2. DEFAULT CONFIGURATION
# ============================================================

DEFAULT_GITHUB_OWNER = "chaabanemeddour05-ai"
DEFAULT_GITHUB_REPO = "Tasuqilt"
DEFAULT_GITHUB_BRANCH = "main"

GITHUB_OWNER = os.getenv("GITHUB_OWNER", DEFAULT_GITHUB_OWNER)
GITHUB_REPO = os.getenv("GITHUB_REPO", DEFAULT_GITHUB_REPO)
GITHUB_BRANCH = os.getenv("GITHUB_BRANCH", DEFAULT_GITHUB_BRANCH)

GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-flash-latest")

GITHUB_CACHE_TTL = 900

# tm.xlsx is ~5.5 MB, so the old 8 MB limit was too close.
MAX_FILE_BYTES = 25 * 1024 * 1024

# Terminology = only SHORT entries. Full paragraphs belong to the TM only.
TERM_MAX_WORDS = 5
TERM_MAX_CHARS = 60


# ============================================================
# 3. SAFE SECRETS
# ============================================================

def get_secret(name: str, default: str = "") -> str:
    try:
        value = st.secrets.get(name, "")
        if value:
            return str(value)
    except Exception:
        pass

    return os.getenv(name, default)


ADMIN_PASSWORD = get_secret("ADMIN_PASSWORD", "")
GEMINI_API_KEY = get_secret("GEMINI_API_KEY", "")
GEMINI_MODEL = get_secret("GEMINI_MODEL", GEMINI_MODEL)


# ============================================================
# 4. SESSION STATE
# ============================================================

if "custom_system_instruction" not in st.session_state:
    st.session_state["custom_system_instruction"] = """
You are the official translation assistant for Tasuqilt DZ.

Your task is to translate news and institutional texts accurately.

ABSOLUTE RULES:

1. The supplied Tasuqilt terminology and translation memory have priority.
2. Do not invent an official Tamazight term when an approved term is supplied.
3. Do not replace an approved term with a different synonym.
4. Preserve names, numbers, dates, institutions and factual information.
5. Do not add explanations, comments or notes.
6. Return only the requested translation.
7. When the supplied database does not contain an exact equivalent,
   use the closest supplied terminology and examples.
8. Never pretend that a term is official if it was not supplied by Tasuqilt.
9. If a paragraph of the source text is identical to a SOURCE in the
   supplied translation memory, reuse its TAMAZIGHT text verbatim.
""".strip()

if "last_translation" not in st.session_state:
    st.session_state["last_translation"] = ""

if "last_retrieval" not in st.session_state:
    st.session_state["last_retrieval"] = {}

if "admin_mode" not in st.session_state:
    st.session_state["admin_mode"] = False


# ============================================================
# 5. NORMALIZATION
# ============================================================

def normalize_text(text: str) -> str:
    if not text:
        return ""

    text = text.replace("\u00A0", " ")
    text = text.replace("’", "'")
    text = text.replace("“", '"')
    text = text.replace("”", '"')
    text = text.replace("«", '"')
    text = text.replace("»", '"')
    text = re.sub(r"\s+", " ", text)

    return text.strip().lower()


def tokenize(text: str) -> List[str]:
    text = normalize_text(text)

    return re.findall(
        r"[a-zA-ZÀ-ÖØ-öø-ÿɛƐɣƔʷṭḍṛṣẓḥčžǧ']+|[\u0600-\u06FF]+|\d+",
        text,
        flags=re.UNICODE
    )


def split_paragraphs(text: str) -> List[str]:
    return [p.strip() for p in re.split(r"\n+", text) if p.strip()]


# ============================================================
# 6. BUILT-IN OFFICIAL TERMINOLOGY
# ============================================================

DEFAULT_TERMINOLOGY = [
    {"source": "président de la république", "target": "Aselway n Tegduda", "priority": 100},
    {"source": "le président", "target": "Aselway", "priority": 90},
    {"source": "président", "target": "Aselway", "priority": 90},
    {"source": "رئيس الجمهورية", "target": "Aselway n Tegduda", "priority": 100},
    {"source": "رئيس", "target": "Aselway", "priority": 90},
    {"source": "alger", "target": "DZAYER TAMANEƔT", "priority": 100},
    {"source": "الجزائر", "target": "DZAYER TAMANEƔT", "priority": 100},
    {"source": "conseil des ministres", "target": "Aseqqamu n Yineɣlaf", "priority": 100},
    {"source": "مجلس الوزراء", "target": "Aseqqamu n Yineɣlaf", "priority": 100},
    {"source": "réunion", "target": "Timlilt", "priority": 90},
    {"source": "اجتماع", "target": "Timlilt", "priority": 90},
    {"source": "gouvernement", "target": "Anabaḍ", "priority": 90},
    {"source": "الحكومة", "target": "Anabaḍ", "priority": 90},
]

for _t in DEFAULT_TERMINOLOGY:
    _t.setdefault("origin", "Built-in")


# ============================================================
# 7. GITHUB ACCESS
# ============================================================

def github_headers() -> Dict[str, str]:
    headers = {
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2026-03-10",
        "User-Agent": "Tasuqilt-DZ"
    }

    github_token = get_secret("GITHUB_TOKEN", "")

    if github_token:
        headers["Authorization"] = f"Bearer {github_token}"

    return headers


@st.cache_data(ttl=120, max_entries=1)
def get_latest_commit_sha() -> str:
    """
    Latest commit of the branch. Every add / delete / edit of a file in
    GitHub creates a new commit, so using this SHA as a cache key makes
    the app notice changes (e.g. a deleted tm.docx) within ~2 minutes
    instead of waiting for the old 15-minute cache to expire.
    Falls back to the branch name if the API call fails.
    """

    try:
        url = (
            f"https://api.github.com/repos/"
            f"{GITHUB_OWNER}/{GITHUB_REPO}/commits/{GITHUB_BRANCH}"
        )

        response = requests.get(url, headers=github_headers(), timeout=20)
        response.raise_for_status()

        return response.json().get("sha", "") or GITHUB_BRANCH

    except Exception:
        return GITHUB_BRANCH


@st.cache_data(ttl=GITHUB_CACHE_TTL, max_entries=2)
def get_github_tree(commit_sha: str) -> List[Dict]:
    url = (
        f"https://api.github.com/repos/"
        f"{GITHUB_OWNER}/{GITHUB_REPO}/git/trees/"
        f"{GITHUB_BRANCH}?recursive=1"
    )

    response = requests.get(url, headers=github_headers(), timeout=20)
    response.raise_for_status()

    data = response.json()

    if data.get("truncated"):
        raise RuntimeError(
            "GitHub returned a truncated file tree. "
            "The repository contains too many files."
        )

    return data.get("tree", [])


def github_raw_url(path: str, ref: str) -> str:
    # ref = commit SHA (immutable, never served stale) or branch name
    return (
        f"https://raw.githubusercontent.com/"
        f"{GITHUB_OWNER}/{GITHUB_REPO}/"
        f"{ref}/{quote(path)}"
    )


# ============================================================
# 8. WHICH FILES ARE DATA FILES?
# ============================================================

ALLOWED_EXTENSIONS = {".txt", ".md", ".csv", ".json", ".docx", ".xlsx"}

IGNORED_DIRECTORIES = {
    ".git", ".github", "__pycache__", ".streamlit", "venv", ".venv"
}

IGNORED_FILES = {"app.py", "requirements.txt", "README.md"}

ROOT_DATA_FILENAMES = {
    # "tm.docx" was deleted from the repository (replaced by tm.xlsx).
    "tm.xlsx",
    "translation_memory.csv",
    "translation_memory.xlsx",
    "lexicon.csv",
    "terminology.csv",
    "corpus.txt",
    "corpus.md",
    "terminology.json",
    "lexicon.json",
}


def is_data_file(path: str) -> bool:
    normalized_path = path.replace("\\", "/")
    parts = normalized_path.split("/")

    if any(part in IGNORED_DIRECTORIES for part in parts):
        return False

    filename = parts[-1]

    if filename in IGNORED_FILES:
        return False

    extension = os.path.splitext(filename)[1].lower()

    if extension not in ALLOWED_EXTENSIONS:
        return False

    if "data" in parts:
        return True

    return filename.lower() in ROOT_DATA_FILENAMES


# ============================================================
# 9. READ DATA FILES
# ============================================================

def decode_text(content: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-8", "cp1256", "latin-1"):
        try:
            return content.decode(encoding)
        except Exception:
            continue

    return content.decode("utf-8", errors="ignore")


def read_xlsx(content: bytes) -> Tuple[str, List[Dict]]:
    """
    tm.xlsx: FIRST sheet only.
        Column A = French (source)
        Column B = Tamazight (target)
    Other sheets (review / report) are intentionally ignored.
    A header row is skipped automatically if present.
    """

    pairs = []

    workbook = load_workbook(
        io.BytesIO(content),
        read_only=True,
        data_only=True
    )

    try:
        sheet = workbook[workbook.sheetnames[0]]

        header_words = {
            "source", "french", "français", "francais", "fr",
            "foreign", "original", "arabe", "arabic"
        }

        for row_number, row in enumerate(
            sheet.iter_rows(values_only=True),
            start=1
        ):

            if not row or len(row) < 2:
                continue

            left, right = row[0], row[1]

            if left is None or right is None:
                continue

            source = str(left).strip()
            target = str(right).strip()

            if not source or not target:
                continue

            if row_number == 1 and source.lower() in header_words:
                continue

            pairs.append({
                "source": source,
                "target": target,
                "origin": "GitHub XLSX"
            })

    finally:
        workbook.close()

    # The pairs are not added to the plain text blob (22k rows would
    # only waste memory); they live in the TM index.
    return "", pairs


def read_docx(content: bytes) -> Tuple[str, List[Dict]]:
    """
    Supports paragraphs in the '@' format (Tamazight @ Foreign)
    and simple two-column tables (Foreign | Tamazight).
    """

    doc = Document(io.BytesIO(content))

    paragraphs = []
    pairs = []

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()

        if not text:
            continue

        paragraphs.append(text)

        if "@" in text:
            left, right = text.split("@", 1)
            left = left.strip()
            right = right.strip()

            if left and right:
                pairs.append({
                    "source": right,
                    "target": left,
                    "origin": "GitHub DOCX"
                })

    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]

            if len(cells) >= 2 and cells[0] and cells[1]:
                pairs.append({
                    "source": cells[0],
                    "target": cells[1],
                    "origin": "GitHub DOCX table"
                })

    return "\n".join(paragraphs), pairs


def read_csv_file(content: bytes) -> Tuple[str, List[Dict]]:
    text = decode_text(content)
    rows = []

    try:
        reader = csv.DictReader(io.StringIO(text))

        if reader.fieldnames:
            fields = {
                field.strip().lower(): field
                for field in reader.fieldnames
                if field
            }

            source_field = None
            target_field = None

            for candidate in [
                "source", "foreign", "original",
                "français", "french", "arabe", "arabic"
            ]:
                if candidate in fields:
                    source_field = fields[candidate]
                    break

            for candidate in [
                "target", "tamazight", "amazigh", "translation"
            ]:
                if candidate in fields:
                    target_field = fields[candidate]
                    break

            if source_field and target_field:
                for row in reader:
                    source = (row.get(source_field) or "").strip()
                    target = (row.get(target_field) or "").strip()

                    if source and target:
                        rows.append({
                            "source": source,
                            "target": target,
                            "origin": "GitHub CSV"
                        })

    except Exception:
        pass

    return text, rows


def read_json_file(content: bytes) -> Tuple[str, List[Dict]]:
    text = decode_text(content)
    pairs = []

    try:
        data = json.loads(text)

        if isinstance(data, list):
            for item in data:
                if not isinstance(item, dict):
                    continue

                source = (
                    item.get("source") or item.get("foreign")
                    or item.get("original") or item.get("fr")
                )
                target = (
                    item.get("target") or item.get("tamazight")
                    or item.get("amazigh") or item.get("translation")
                )

                if source and target:
                    pairs.append({
                        "source": str(source).strip(),
                        "target": str(target).strip(),
                        "origin": "GitHub JSON"
                    })

        elif isinstance(data, dict):
            for source, target in data.items():
                if isinstance(target, str):
                    pairs.append({
                        "source": str(source).strip(),
                        "target": target.strip(),
                        "origin": "GitHub JSON"
                    })

    except Exception:
        pass

    return text, pairs


def read_data_file(path: str, content: bytes) -> Tuple[str, List[Dict]]:
    extension = os.path.splitext(path)[1].lower()

    if extension == ".xlsx":
        return read_xlsx(content)

    if extension == ".docx":
        return read_docx(content)

    if extension == ".csv":
        return read_csv_file(content)

    if extension == ".json":
        return read_json_file(content)

    return decode_text(content), []


# cache_resource: no pickling/copying of 22k pairs on every rerun.
@st.cache_resource(ttl=GITHUB_CACHE_TTL, max_entries=2)
def load_github_knowledge(commit_sha: str) -> Dict:

    tree = get_github_tree(commit_sha)

    data_files = [
        item for item in tree
        if item.get("type") == "blob"
        and is_data_file(item.get("path", ""))
    ]

    all_text = []
    translation_memory = []
    loaded_files = []
    errors = []

    data_files = data_files[:100]

    for item in data_files:

        path = item["path"]

        try:
            response = requests.get(
                github_raw_url(path, commit_sha),
                headers={"User-Agent": "Tasuqilt-DZ"},
                timeout=60
            )

            response.raise_for_status()

            content = response.content

            if len(content) > MAX_FILE_BYTES:
                errors.append(f"{path}: file larger than 25 MB skipped.")
                continue

            text, pairs = read_data_file(path, content)

            if text:
                all_text.append(f"\n===== FILE: {path} =====\n{text}")

            translation_memory.extend(pairs)
            loaded_files.append(f"{path} ({len(pairs)} pairs)")

        except Exception as error:
            errors.append(f"{path}: {str(error)}")

    return {
        "files": loaded_files,
        "text": "\n".join(all_text),
        "tm": translation_memory,
        "errors": errors
    }


# ============================================================
# 10. TRANSLATION-MEMORY INDEX (fast retrieval for 20k+ pairs)
# ============================================================

def empty_index() -> Dict:
    return {
        "entries": [],
        "exact": {"src": {}, "tgt": {}},
        "inv": {"src": defaultdict(list), "tgt": defaultdict(list)},
    }


@st.cache_resource(ttl=GITHUB_CACHE_TTL, max_entries=2)
def build_tm_index(commit_sha: str) -> Dict:
    """
    Builds, once per cache period:
      - exact lookup tables (normalized text -> entry)
      - inverted word index for both languages
    """

    knowledge = load_github_knowledge(commit_sha)

    index = empty_index()
    seen = set()

    for pair in knowledge.get("tm", []):

        source = pair.get("source", "").strip()
        target = pair.get("target", "").strip()

        if not source or not target:
            continue

        src_norm = normalize_text(source)
        tgt_norm = normalize_text(target)

        key = (src_norm, tgt_norm)

        if key in seen:
            continue

        seen.add(key)

        entry = {
            "source": source,
            "target": target,
            "origin": pair.get("origin", ""),
            "src_norm": src_norm,
            "tgt_norm": tgt_norm,
            "src_tok": frozenset(tokenize(source)),
            "tgt_tok": frozenset(tokenize(target)),
        }

        idx = len(index["entries"])
        index["entries"].append(entry)

        index["exact"]["src"].setdefault(src_norm, idx)
        index["exact"]["tgt"].setdefault(tgt_norm, idx)

        for tok in entry["src_tok"]:
            index["inv"]["src"][tok].append(idx)

        for tok in entry["tgt_tok"]:
            index["inv"]["tgt"][tok].append(idx)

    return index


def side_of(direction: str) -> str:
    """'src' = query is French/Arabic; 'tgt' = query is Tamazight."""
    return "src" if "Auto-Detect" in direction.split("➔")[0] else "tgt"


def entry_view(entry: Dict, side: str, score: float = 1.0) -> Dict:
    """Return the entry oriented as query-language -> answer-language."""

    if side == "src":
        return {
            "source": entry["source"],
            "target": entry["target"],
            "origin": entry.get("origin", ""),
            "score": score
        }

    return {
        "source": entry["target"],
        "target": entry["source"],
        "origin": entry.get("origin", ""),
        "score": score
    }


def tm_exact(query: str, index: Dict, side: str):
    idx = index["exact"][side].get(normalize_text(query))

    if idx is None:
        return None

    return entry_view(index["entries"][idx], side, 1.0)


def search_tm(
    query: str,
    index: Dict,
    side: str,
    limit: int = 3
) -> List[Dict]:

    entries = index["entries"]

    if not entries:
        return []

    query_norm = normalize_text(query)
    query_tokens = set(tokenize(query_norm))

    if not query_tokens:
        return []

    inv = index["inv"][side]
    cap = max(50, int(len(entries) * 0.05))   # ignore very common words

    votes = Counter()

    for tok in query_tokens:
        ids = inv.get(tok)

        if not ids or len(ids) > cap:
            continue

        for i in ids:
            votes[i] += 1

    norm_key = "src_norm" if side == "src" else "tgt_norm"
    tok_key = "src_tok" if side == "src" else "tgt_tok"

    results = []

    for i, _ in votes.most_common(25):

        entry = entries[i]
        entry_norm = entry[norm_key]
        entry_tokens = entry[tok_key]

        if not entry_tokens:
            continue

        common = len(entry_tokens & query_tokens)

        dice = 2 * common / (len(entry_tokens) + len(query_tokens))
        overlap = common / len(entry_tokens)

        if len(entry_norm) >= 12 and entry_norm in query_norm:
            score = 1.0
        elif len(entry_norm) <= 400 and len(query_norm) <= 400:
            score = max(
                difflib.SequenceMatcher(None, query_norm, entry_norm).ratio(),
                dice
            )
        else:
            score = max(dice, overlap * 0.8)

        if score >= 0.55:
            results.append(entry_view(entry, side, score))

    results.sort(key=lambda x: (x["score"], len(x["source"])), reverse=True)

    return results[:limit]


# ============================================================
# 11. TERMINOLOGY (short entries only)
# ============================================================

def build_terminology(knowledge: Dict) -> List[Dict]:

    terminology = list(DEFAULT_TERMINOLOGY)

    for pair in knowledge.get("tm", []):

        source = pair.get("source", "").strip()
        target = pair.get("target", "").strip()

        if not source or not target:
            continue

        if (
            len(source) > TERM_MAX_CHARS
            or len(source.split()) > TERM_MAX_WORDS
            or len(target.split()) > TERM_MAX_WORDS + 3
        ):
            continue

        terminology.append({
            "source": source,
            "target": target,
            "priority": 80,
            "origin": pair.get("origin", "GitHub")
        })

    unique = {}

    for item in terminology:
        key = (
            normalize_text(item["source"]),
            normalize_text(item["target"])
        )
        unique[key] = item

    result = list(unique.values())

    result.sort(
        key=lambda x: (x.get("priority", 0), len(x.get("source", ""))),
        reverse=True
    )

    return result


def contains_phrase(haystack_norm: str, phrase_norm: str) -> bool:
    """Whole-word containment (so 'alger' does not match 'algerien')."""

    if not phrase_norm:
        return False

    return re.search(
        r"(?<!\w)" + re.escape(phrase_norm) + r"(?!\w)",
        haystack_norm
    ) is not None


# ============================================================
# 12. INTELLIGENT RETRIEVAL
# ============================================================

def retrieve_context(
    query: str,
    index: Dict,
    terminology: List[Dict],
    direction: str,
    max_pairs: int = 12
) -> Dict:

    side = side_of(direction)

    query_norm = normalize_text(query)
    query_tokens = set(tokenize(query))

    # --- 12.1 Terminology -------------------------------------

    matched_terms = []

    for term in terminology:

        if side == "src":
            source, target = term["source"], term["target"]
        else:
            source, target = term["target"], term["source"]

        source_norm = normalize_text(source)

        if not source_norm:
            continue

        if contains_phrase(query_norm, source_norm):
            matched_terms.append({
                "source": source, "target": target,
                "priority": term.get("priority", 0),
                "origin": term.get("origin", ""), "score": 1.0
            })
            continue

        source_tokens = set(tokenize(source_norm))

        if len(source_tokens) >= 2 and query_tokens:
            overlap = len(source_tokens & query_tokens) / len(source_tokens)

            if overlap >= 0.75:
                matched_terms.append({
                    "source": source, "target": target,
                    "priority": term.get("priority", 0),
                    "origin": term.get("origin", ""), "score": overlap
                })

    matched_terms.sort(
        key=lambda x: (x["score"], x["priority"], len(x["source"])),
        reverse=True
    )

    matched_terms = matched_terms[:20]

    # --- 12.2 Translation memory (paragraph by paragraph) -----

    exact_match = tm_exact(query, index, side)

    merged = {}

    for paragraph in split_paragraphs(query):

        for hit in search_tm(paragraph, index, side, limit=3):
            key = (hit["source"], hit["target"])

            if key not in merged or hit["score"] > merged[key]["score"]:
                merged[key] = hit

    tm_matches = sorted(
        merged.values(),
        key=lambda x: (x["score"], len(x["source"])),
        reverse=True
    )[:max_pairs]

    return {
        "exact_match": exact_match,
        "terms": matched_terms,
        "tm_matches": tm_matches
    }


# ============================================================
# 13. CONFIDENCE
# ============================================================

def calculate_confidence(retrieval: Dict) -> int:

    if retrieval.get("exact_match"):
        return 100

    terms = retrieval.get("terms", [])
    matches = retrieval.get("tm_matches", [])

    if terms and matches:
        return 90

    if terms:
        return 80

    if matches:
        best = matches[0].get("score", 0)
        return int(min(75, max(50, best * 100)))

    return 0


# ============================================================
# 14. BUILD AI CONTEXT
# ============================================================

def build_ai_context(retrieval: Dict, direction: str) -> str:

    sections = ["TASUQILT OFFICIAL TERMINOLOGY:\n"]

    terms = retrieval.get("terms", [])

    if terms:
        for term in terms:
            sections.append(f"- {term['source']} = {term['target']}")
    else:
        sections.append("- No matching official term was found.")

    sections.append("\nTASUQILT TRANSLATION MEMORY EXAMPLES:\n")

    matches = retrieval.get("tm_matches", [])

    if matches:
        for number, pair in enumerate(matches, 1):
            sections.append(
                f"{number}. SOURCE: {pair['source']}\n"
                f"   TRANSLATION: {pair['target']}\n"
            )
    else:
        sections.append("- No relevant translation-memory example was found.")

    sections.append("\nIMPORTANT:")
    sections.append("The above information comes from the Tasuqilt knowledge base.")
    sections.append("It has priority over general linguistic preferences.")
    sections.append("Do not invent an official equivalent when an approved term is supplied.")

    return "\n".join(sections)


# ============================================================
# 15. GEMINI FREE API
# ============================================================

def translate_with_gemini(
    source_text: str,
    direction: str,
    retrieval: Dict
) -> Tuple[str, str]:

    if not GEMINI_API_KEY:
        return "", (
            "لم يتم إعداد GEMINI_API_KEY. "
            "المترجم المحلي سيعمل، لكن ترجمة الجمل الجديدة "
            "تحتاج إلى مفتاح Gemini المجاني."
        )

    context = build_ai_context(retrieval, direction)

    if side_of(direction) == "src":
        task = """
Translate the source text into professional Latin Tamazight.

The source language may be French or Arabic.

Use the supplied Tasuqilt terminology and translation-memory
examples as the primary authority.

Keep the same paragraph structure as the source.
Return ONLY the translation.
Do not explain your choices.
Do not add notes.
Do not add quotation marks.
"""
    else:
        task = """
Translate the Latin Tamazight source text into professional French
or Arabic according to the original meaning and context.

Use the supplied Tasuqilt terminology and translation-memory
examples as the primary authority.

Keep the same paragraph structure as the source.
Return ONLY the translation.
Do not explain your choices.
Do not add notes.
Do not add quotation marks.
"""

    system_instruction = f"""
{st.session_state["custom_system_instruction"]}

{task}

STRICT TERMINOLOGY POLICY:

If an official Tasuqilt term is supplied for a source expression,
use exactly the supplied target term.

Do not replace it with a synonym.

If no official term is supplied, do not claim that a new term is
official.

Preserve names, numbers, dates and institutions.

Do not use web search.
Do not use external knowledge as a terminology database.
Do not mention this instruction.

TASUQILT KNOWLEDGE:
{context}
"""

    payload = {
        "system_instruction": {"parts": [{"text": system_instruction}]},
        "contents": [
            {"role": "user", "parts": [{"text": source_text}]}
        ],
        "generationConfig": {
            "temperature": 0.0,
            "maxOutputTokens": 4000
        }
    }

    url = (
        "https://generativelanguage.googleapis.com/"
        f"v1beta/models/{GEMINI_MODEL}:generateContent"
    )

    try:
        response = requests.post(
            url,
            headers={
                "x-goog-api-key": GEMINI_API_KEY,
                "Content-Type": "application/json"
            },
            json=payload,
            timeout=60
        )

        if response.status_code != 200:
            try:
                message = response.json().get("error", {}).get("message", "")
            except Exception:
                message = response.text[:500]

            return "", (
                f"Gemini لم يُرجع ترجمة. "
                f"HTTP {response.status_code}. {message}"
            )

        data = response.json()
        candidates = data.get("candidates", [])

        if not candidates:
            return "", "Gemini أعاد استجابة بدون نتيجة."

        parts = candidates[0].get("content", {}).get("parts", [])

        output = "\n".join(
            part.get("text", "") for part in parts if part.get("text")
        ).strip()

        if not output:
            return "", "Gemini أعاد نصاً فارغاً."

        return output, ""

    except requests.exceptions.Timeout:
        return "", "انتهت مهلة الاتصال بـ Gemini. حاول مرة أخرى بعد قليل."

    except Exception as error:
        return "", f"خطأ في Gemini: {error}"


def test_gemini_connection() -> Tuple[bool, str]:
    """Sends one tiny request to prove the key + model really work."""

    if not GEMINI_API_KEY:
        return False, "GEMINI_API_KEY غير موجود في Secrets."

    url = (
        "https://generativelanguage.googleapis.com/"
        f"v1beta/models/{GEMINI_MODEL}:generateContent"
    )

    payload = {
        "contents": [
            {"role": "user", "parts": [{"text": "Reply with the single word: OK"}]}
        ],
        "generationConfig": {"temperature": 0.0, "maxOutputTokens": 200}
    }

    try:
        response = requests.post(
            url,
            headers={
                "x-goog-api-key": GEMINI_API_KEY,
                "Content-Type": "application/json"
            },
            json=payload,
            timeout=30
        )

        if response.status_code == 200:
            return True, f"الاتصال ناجح ✅ — النموذج: {GEMINI_MODEL}"

        try:
            message = response.json().get("error", {}).get("message", "")
        except Exception:
            message = response.text[:300]

        hints = {
            400: "مفتاح غير صالح أو طلب خاطئ.",
            403: "المفتاح لا يملك صلاحية لهذا النموذج/المشروع.",
            404: f"اسم النموذج غير موجود: {GEMINI_MODEL}. غيّر GEMINI_MODEL في Secrets.",
            429: "تجاوزت الحد المجاني للطلبات. انتظر قليلاً.",
        }

        return False, (
            f"فشل الاتصال ❌ HTTP {response.status_code}. "
            f"{hints.get(response.status_code, '')} {message}"
        )

    except Exception as error:
        return False, f"فشل الاتصال ❌ {error}"


# ============================================================
# 16. TERMINOLOGY VALIDATOR
# ============================================================

def _stems(text: str) -> List[str]:
    """
    Tamazight changes the first vowel of nouns after prepositions
    (Aselway -> uselway, Anabaḍ -> unabaḍ, Tegduda -> tegduda).
    Dropping the first letter of each word makes the check tolerant
    of that alternation instead of rejecting correct translations.
    """

    return [w[1:] if len(w) > 3 else w for w in tokenize(text)]


def output_contains_term(output_text: str, target: str) -> bool:

    out_stems = _stems(output_text)
    tgt_stems = _stems(target)

    if not tgt_stems:
        return True

    n = len(tgt_stems)

    for i in range(len(out_stems) - n + 1):
        if out_stems[i:i + n] == tgt_stems:
            return True

    return False


def validate_translation(
    source_text: str,
    output_text: str,
    retrieval: Dict,
    direction: str
) -> Dict:

    if not output_text:
        return {
            "valid": False,
            "warnings": ["الترجمة فارغة."],
            "checked": 0,
            "missing": []
        }

    # Terminology is only enforced for source -> Tamazight.
    if side_of(direction) != "src":
        return {"valid": True, "warnings": [], "checked": 0, "missing": []}

    missing = []
    source_norm_text = normalize_text(source_text)

    for term in retrieval.get("terms", []):

        if contains_phrase(source_norm_text, normalize_text(term["source"])):
            if not output_contains_term(output_text, term["target"]):
                missing.append({
                    "source": term["source"],
                    "expected": term["target"]
                })

    if missing:
        return {
            "valid": False,
            "warnings": [
                f"المصطلح الرسمي غير موجود في الناتج: "
                f"{m['source']} → {m['expected']}"
                for m in missing
            ],
            "checked": len(retrieval.get("terms", [])),
            "missing": missing
        }

    return {
        "valid": True,
        "warnings": [],
        "checked": len(retrieval.get("terms", [])),
        "missing": []
    }


# ============================================================
# 17. EXACT LOCAL TRANSLATION (paragraph-aware)
# ============================================================

def exact_local_translation(
    source_text: str,
    index: Dict,
    direction: str
) -> str:
    """
    Returns a translation only when EVERY paragraph of the input has an
    exact match in the translation memory (or the whole input is a
    built-in term). Otherwise returns "" so the pipeline goes on.
    """

    side = side_of(direction)

    paragraphs = split_paragraphs(source_text)

    if paragraphs:
        results = []

        for paragraph in paragraphs:
            hit = tm_exact(paragraph, index, side)

            if not hit:
                results = []
                break

            results.append(hit["target"])

        if results:
            return "\n".join(results)

    query_norm = normalize_text(source_text)

    for term in DEFAULT_TERMINOLOGY:
        if side == "src" and normalize_text(term["source"]) == query_norm:
            return term["target"]

        if side == "tgt" and normalize_text(term["target"]) == query_norm:
            return term["source"]

    return ""


# ============================================================
# 18. UI
# ============================================================

st.title("📖 Tasuqilt DZ 🇩🇿")

st.markdown(
    """
    <p style='font-size:1.05rem;color:gray;'>
    نظام ترجمة هجين يعتمد أولاً على قاعدة Tasuqilt الموجودة في GitHub،
    ثم يستخدم Gemini المجاني فقط عند الحاجة.
    </p>
    """,
    unsafe_allow_html=True
)


# ============================================================
# 19. SIDEBAR
# ============================================================

st.sidebar.header("⚙️ Configuration")

st.sidebar.info(
    "هذا الإصدار لا يستخدم OpenAI أو Claude أو DeepSeek أو Grok "
    "لأنك طلبت نظاماً مجانياً فقط."
)

engine_choice = st.sidebar.selectbox(
    "محرك الترجمة:",
    [
        "Tasuqilt Local — بدون API",
        "Gemini Free — عند الحاجة"
    ]
)

direction = st.sidebar.selectbox(
    "اتجاه الترجمة:",
    [
        "Auto-Detect [Français/Arabe] ➔ Tamazight",
        "Tamazight ➔ Auto-Detect [Français/Arabe]"
    ]
)

use_gemini = engine_choice.startswith("Gemini")


# ============================================================
# 20. LOAD GITHUB KNOWLEDGE
# ============================================================

with st.spinner("🔄 قراءة قاعدة Tasuqilt من GitHub..."):

    try:
        commit_sha = get_latest_commit_sha()
        knowledge = load_github_knowledge(commit_sha)
        tm_index = build_tm_index(commit_sha)
        github_ok = True

    except Exception as error:
        commit_sha = ""
        knowledge = {
            "files": [],
            "text": "",
            "tm": [],
            "errors": [str(error)]
        }
        tm_index = empty_index()
        github_ok = False


terminology = build_terminology(knowledge)


# ============================================================
# 21. GITHUB STATUS
# ============================================================

if github_ok:
    st.sidebar.success(
        f"🟢 GitHub متصل — {len(knowledge['files'])} ملف بيانات"
    )
    st.sidebar.caption(
        f"Translation Memory: {len(tm_index['entries']):,} زوج ترجمة"
    )
    st.sidebar.caption(
        f"المصطلحات القصيرة: {len(terminology):,}"
    )

    for filename in knowledge["files"]:
        st.sidebar.caption(f"📄 {filename}")

    if commit_sha and len(commit_sha) >= 7:
        st.sidebar.caption(f"آخر نسخة من GitHub: {commit_sha[:7]}")

    if not knowledge["files"]:
        st.sidebar.warning(
            "لم يُعثر على أي ملف بيانات. تأكد أن tm.xlsx في جذر "
            "المستودع أو داخل مجلد data."
        )
else:
    st.sidebar.error("🔴 تعذر قراءة GitHub")

    for error in knowledge.get("errors", [])[:3]:
        st.sidebar.caption(error)


if GEMINI_API_KEY:
    st.sidebar.success(f"🟢 مفتاح Gemini موجود — النموذج: {GEMINI_MODEL}")
    st.sidebar.caption("وجود المفتاح لا يعني أنه يعمل. استخدم زر الاختبار في وضع الإدارة.")
else:
    st.sidebar.warning(
        "🟡 Gemini غير مفعّل — المترجم المحلي فقط متاح."
    )


# ============================================================
# 22. ADMIN
# ============================================================

st.sidebar.markdown("---")
st.sidebar.subheader("🔐 Administration")

admin_input = st.sidebar.text_input("Code d'accès admin:", type="password")

if ADMIN_PASSWORD and admin_input == ADMIN_PASSWORD:
    st.session_state["admin_mode"] = True
elif admin_input:
    st.sidebar.error("رمز الإدارة غير صحيح.")


if st.session_state["admin_mode"]:

    st.sidebar.success("🔓 وضع الإدارة مفعّل")

    updated_prompt = st.sidebar.text_area(
        "System Instruction:",
        value=st.session_state["custom_system_instruction"],
        height=220
    )

    if st.sidebar.button("💾 حفظ التعليمات"):
        st.session_state["custom_system_instruction"] = updated_prompt
        st.sidebar.success("تم حفظ التعليمات لهذه الجلسة.")

    if st.sidebar.button("🧪 اختبار اتصال Gemini"):
        with st.spinner("اختبار الاتصال..."):
            ok, message = test_gemini_connection()

        if ok:
            st.sidebar.success(message)
        else:
            st.sidebar.error(message)

    st.sidebar.markdown("---")
    st.sidebar.write("ملفات البيانات المقروءة:")

    for filename in knowledge.get("files", []):
        st.sidebar.caption(f"📄 {filename}")

    if knowledge.get("errors"):
        st.sidebar.warning("بعض الملفات تم تجاوزها:")

        for error in knowledge["errors"][:10]:
            st.sidebar.caption(error)

    if st.sidebar.button("🔄 تحديث قاعدة GitHub الآن"):
        get_latest_commit_sha.clear()
        get_github_tree.clear()
        load_github_knowledge.clear()
        build_tm_index.clear()
        st.rerun()


# ============================================================
# 23. EXPRESS DICTIONARY
# ============================================================

st.markdown("### 📖 Dictionnaire Express / القاموس الفوري")

dict_col1, dict_col2 = st.columns(2)

with dict_col1:
    word_to_find = st.text_input(
        "ابحث عن كلمة أو مصطلح:",
        placeholder="président / réunion / gouvernement..."
    )

with dict_col2:
    st.markdown("**النتيجة:**")

    if word_to_find.strip():

        query = normalize_text(word_to_find)

        found = None

        for term in terminology:
            if normalize_text(term["source"]) == query:
                found = term
                break

        if found:
            st.success(found["target"])
            st.caption(f"مصدر المصطلح: {found.get('origin', 'Tasuqilt')}")
        else:
            st.info("هذا المصطلح غير موجود في قاعدة Tasuqilt.")


st.markdown("---")


# ============================================================
# 24. MAIN TRANSLATOR
# ============================================================

st.markdown("### 📰 Traducteur de Dépêches / مترجم البرقيات الإعلامية")

col1, col2 = st.columns(2)

with col1:
    text_to_translate = st.text_area(
        "النص المصدر:",
        height=280,
        placeholder=(
            "أدخل النص الفرنسي أو العربي هنا..."
            if side_of(direction) == "src"
            else "أدخل نص Tamazight باللاتينية هنا..."
        )
    )

    submit_button = st.button(
        "🚀 ترجم",
        type="primary",
        use_container_width=True
    )


# ============================================================
# 25. TRANSLATION PIPELINE
# ============================================================

if submit_button:

    if not text_to_translate.strip():
        st.warning("أدخل نصاً أولاً.")

    else:

        source_text = text_to_translate.strip()

        output_text = ""
        engine_used = ""
        confidence = 0

        # STEP 1 — RETRIEVE
        with st.spinner("🔎 البحث في قاعدة Tasuqilt..."):
            retrieval = retrieve_context(
                source_text, tm_index, terminology, direction
            )

        st.session_state["last_retrieval"] = retrieval
        confidence = calculate_confidence(retrieval)

        # STEP 2 — EXACT LOCAL MATCH
        local_result = exact_local_translation(
            source_text, tm_index, direction
        )

        if local_result:

            output_text = local_result
            engine_used = "Tasuqilt Translation Memory"
            confidence = 100
            st.session_state["last_translation"] = output_text

        # STEP 3 — GEMINI ONLY IF NECESSARY
        elif use_gemini:

            with st.spinner(
                "🤖 لا توجد مطابقة كاملة — "
                "إرسال السياق المرتبط فقط إلى Gemini..."
            ):
                output_text, error_message = translate_with_gemini(
                    source_text, direction, retrieval
                )

            if error_message:
                st.error(error_message)
                output_text = ""
                engine_used = "Gemini"

            else:
                engine_used = "Gemini Free + Tasuqilt Retrieval"

                # STEP 4 — VALIDATE
                validation = validate_translation(
                    source_text, output_text, retrieval, direction
                )

                if not validation["valid"]:

                    st.error("⛔ لم تعتمد Tasuqilt الترجمة.")

                    for warning in validation["warnings"]:
                        st.warning(warning)

                    with st.expander("عرض الترجمة المرفوضة للمراجعة اليدوية"):
                        st.text_area(
                            "ترجمة Gemini (غير معتمدة):",
                            value=output_text,
                            height=200
                        )

                    output_text = ""

        # STEP 5 — LOCAL ONLY
        else:
            st.info(
                "وضع Tasuqilt Local لا يخترع ترجمة. "
                "أضف الجملة إلى Translation Memory "
                "أو فعّل Gemini Free."
            )

        # STEP 6 — DISPLAY
        with col2:

            if output_text:

                st.markdown("**الترجمة المعتمدة:**")
                st.success(output_text)

                st.text_area(
                    "النتيجة القابلة للنسخ:",
                    value=output_text,
                    height=280
                )

                st.caption(f"⚙️ المحرك: {engine_used}")

                if confidence >= 90:
                    st.success(f"🟢 اعتماد قاعدة Tasuqilt: {confidence}%")
                elif confidence >= 60:
                    st.warning(f"🟡 اعتماد قاعدة Tasuqilt: {confidence}%")
                else:
                    st.info(f"🔵 اعتماد قاعدة Tasuqilt: {confidence}%")

            else:
                st.text_area(
                    "النتيجة:",
                    value="",
                    height=280,
                    disabled=True
                )


# ============================================================
# 26. RETRIEVAL INFORMATION
# ============================================================

if st.session_state.get("last_retrieval"):

    retrieval = st.session_state["last_retrieval"]

    with st.expander("🔎 عرض ما وجده Tasuqilt قبل الترجمة"):

        terms = retrieval.get("terms", [])
        matches = retrieval.get("tm_matches", [])

        if terms:
            st.markdown("#### 📚 المصطلحات الرسمية")

            for term in terms:
                st.write(f"**{term['source']}** → **{term['target']}**")
        else:
            st.write("لم يجد مصطلحات رسمية مطابقة.")

        if matches:
            st.markdown("#### 🧠 أمثلة Translation Memory")

            for pair in matches:
                st.write(f"**Original:** {pair['source']}")
                st.write(f"**Tamazight:** {pair['target']}")
                st.caption(f"درجة التشابه: {round(pair['score'] * 100)}%")
                st.markdown("---")
        else:
            st.write("لم يجد أمثلة Translation Memory قريبة.")


# ============================================================
# 27. FOOTER
# ============================================================

st.markdown("---")

st.caption(
    "Tasuqilt DZ — نظام ترجمة يعتمد على قاعدة المصطلحات "
    "وذاكرة الترجمة الخاصة بالمشروع قبل استخدام الذكاء الاصطناعي."
)
