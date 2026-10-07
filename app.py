import os
import io
import csv
import json
import re
import difflib
from typing import List, Dict, Tuple

import requests
import streamlit as st
from docx import Document


# ============================================================
# TASUQILT DZ - FREE HYBRID TRANSLATION ENGINE
# ============================================================
#
# ARCHITECTURE:
#
# GitHub public repository
#        ↓
# GitHub Reader
#        ↓
# Translation Memory + Lexicon + Corpus
#        ↓
# Intelligent Retrieval
#        ↓
# ┌────────────────────────────────────────────┐
# │ Exact TM match?                            │
# │       YES → use database result            │
# │       NO  → Gemini Free with retrieved data│
# └────────────────────────────────────────────┘
#        ↓
# Terminology Validator
#        ↓
# Final Translation
#
# IMPORTANT:
# - No OpenAI paid API
# - No Claude paid API
# - No DeepSeek paid API
# - No Grok paid API
# - No Pollinations dependency
# - Gemini is optional and only used if a free API key exists
# - Without Gemini, the local dictionary/TM still works
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

# Public repository shown in your screenshot.
# You can change these values later through Streamlit Secrets.
DEFAULT_GITHUB_OWNER = "chaabanemeddour05-ai"
DEFAULT_GITHUB_REPO = "Tasuqilt"
DEFAULT_GITHUB_BRANCH = "main"

GITHUB_OWNER = os.getenv("GITHUB_OWNER", DEFAULT_GITHUB_OWNER)
GITHUB_REPO = os.getenv("GITHUB_REPO", DEFAULT_GITHUB_REPO)
GITHUB_BRANCH = os.getenv("GITHUB_BRANCH", DEFAULT_GITHUB_BRANCH)

# Gemini model can be changed without modifying the code.
# The current Google documentation shows Gemini 3.8 Flash.
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.8-flash")

# GitHub cache lifetime.
# 15 minutes means we don't hit GitHub on every button click.
GITHUB_CACHE_TTL = 900


# ============================================================
# 3. SAFE SECRETS
# ============================================================

def get_secret(name: str, default: str = "") -> str:
    """
    Read a secret from Streamlit Secrets first,
    then from environment variables.
    """

    try:
        value = st.secrets.get(name, "")
        if value:
            return str(value)
    except Exception:
        pass

    return os.getenv(name, default)


ADMIN_PASSWORD = get_secret("ADMIN_PASSWORD", "")
GEMINI_API_KEY = get_secret("GEMINI_API_KEY", "")


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
    """
    Normalize text for searching without destroying the original.
    """

    if not text:
        return ""

    text = text.replace("\u00A0", " ")
    text = text.replace("’", "'")
    text = text.replace("“", '"')
    text = text.replace("”", '"')

    # Remove excessive spaces
    text = re.sub(r"\s+", " ", text)

    return text.strip().lower()


def tokenize(text: str) -> List[str]:
    """
    Basic multilingual tokenizer.
    """

    text = normalize_text(text)

    return re.findall(
        r"[a-zA-ZÀ-ÖØ-öø-ÿɛƐɣƔʷṭḍṛṣẓčžǧ']+|[\u0600-\u06FF]+|\d+",
        text,
        flags=re.UNICODE
    )


# ============================================================
# 6. BUILT-IN OFFICIAL TERMINOLOGY
# ============================================================
#
# These are your current mandatory terms.
#
# Later we can move all of them to lexicon.csv.
# ============================================================

DEFAULT_TERMINOLOGY = [
    {
        "source": "président de la république",
        "target": "Aselway n Tegduda",
        "priority": 100
    },
    {
        "source": "président de la République",
        "target": "Aselway n Tegduda",
        "priority": 100
    },
    {
        "source": "le président",
        "target": "Aselway",
        "priority": 90
    },
    {
        "source": "président",
        "target": "Aselway",
        "priority": 90
    },
    {
        "source": "رئيس الجمهورية",
        "target": "Aselway n Tegduda",
        "priority": 100
    },
    {
        "source": "رئيس",
        "target": "Aselway",
        "priority": 90
    },
    {
        "source": "alger",
        "target": "DZAYER TAMANEƔT",
        "priority": 100
    },
    {
        "source": "الجزائر",
        "target": "DZAYER TAMANEƔT",
        "priority": 100
    },
    {
        "source": "conseil des ministres",
        "target": "Aseqqamu n Yineɣlaf",
        "priority": 100
    },
    {
        "source": "مجلس الوزراء",
        "target": "Aseqqamu n Yineɣlaf",
        "priority": 100
    },
    {
        "source": "réunion",
        "target": "Timlilt",
        "priority": 90
    },
    {
        "source": "اجتماع",
        "target": "Timlilt",
        "priority": 90
    },
    {
        "source": "gouvernement",
        "target": "Anabaḍ",
        "priority": 90
    },
    {
        "source": "الحكومة",
        "target": "Anabaḍ",
        "priority": 90
    },
]


# ============================================================
# 7. GITHUB ACCESS
# ============================================================

def github_headers() -> Dict[str, str]:
    """
    Public repository does not require authentication.
    If later you make the repository private, you can add
    GITHUB_TOKEN through Streamlit Secrets.
    """

    headers = {
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2026-03-10",
        "User-Agent": "Tasuqilt-DZ"
    }

    github_token = get_secret("GITHUB_TOKEN", "")

    if github_token:
        headers["Authorization"] = f"Bearer {github_token}"

    return headers


@st.cache_data(ttl=GITHUB_CACHE_TTL, max_entries=2)
def get_github_tree() -> List[Dict]:
    """
    Get the complete repository tree.

    We intentionally inspect the repository tree rather than blindly
    reading every file.
    """

    url = (
        f"https://api.github.com/repos/"
        f"{GITHUB_OWNER}/{GITHUB_REPO}/git/trees/"
        f"{GITHUB_BRANCH}?recursive=1"
    )

    response = requests.get(
        url,
        headers=github_headers(),
        timeout=20
    )

    response.raise_for_status()

    data = response.json()

    if data.get("truncated"):
        raise RuntimeError(
            "GitHub returned a truncated file tree. "
            "The repository contains too many files."
        )

    return data.get("tree", [])


def github_raw_url(path: str) -> str:
    return (
        f"https://raw.githubusercontent.com/"
        f"{GITHUB_OWNER}/{GITHUB_REPO}/"
        f"{GITHUB_BRANCH}/{path}"
    )


# ============================================================
# 8. WHICH FILES ARE DATA FILES?
# ============================================================

ALLOWED_EXTENSIONS = {
    ".txt",
    ".md",
    ".csv",
    ".json",
    ".docx"
}

IGNORED_DIRECTORIES = {
    ".git",
    ".github",
    "__pycache__",
    ".streamlit",
    "venv",
    ".venv"
}

IGNORED_FILES = {
    "app.py",
    "requirements.txt",
    "README.md"
}


def is_data_file(path: str) -> bool:
    """
    We do NOT send arbitrary source code to the AI.

    We only read files intended to contain linguistic data.
    """

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

    # Strong preference for the data directory.
    # But we also allow root-level tm.docx because you currently have it.
    if "data" in parts:
        return True

    if filename.lower() in {
        "tm.docx",
        "translation_memory.csv",
        "lexicon.csv",
        "terminology.csv",
        "corpus.txt",
        "corpus.md",
        "terminology.json",
        "lexicon.json"
    }:
        return True

    return False


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


def read_docx(content: bytes) -> Tuple[str, List[Dict]]:
    """
    Supports:
    - paragraphs
    - simple tables

    Existing @ format:
    Tamazight @ Foreign

    Example:
    Aselway n Tegduda @ Président de la République
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
            parts = text.split("@", 1)

            left = parts[0].strip()
            right = parts[1].strip()

            if left and right:
                pairs.append({
                    "source": right,
                    "target": left,
                    "origin": "GitHub DOCX"
                })

    # Tables
    for table in doc.tables:
        for row in table.rows:

            cells = [
                cell.text.strip()
                for cell in row.cells
            ]

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
                "source",
                "foreign",
                "original",
                "français",
                "french",
                "arabe",
                "arabic"
            ]:
                if candidate in fields:
                    source_field = fields[candidate]
                    break

            for candidate in [
                "target",
                "tamazight",
                "amazigh",
                "translation"
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
                    item.get("source")
                    or item.get("foreign")
                    or item.get("original")
                    or item.get("fr")
                )

                target = (
                    item.get("target")
                    or item.get("tamazight")
                    or item.get("amazigh")
                    or item.get("translation")
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

    if extension == ".docx":
        return read_docx(content)

    if extension == ".csv":
        return read_csv_file(content)

    if extension == ".json":
        return read_json_file(content)

    return decode_text(content), []


@st.cache_data(ttl=GITHUB_CACHE_TTL, max_entries=2)
def load_github_knowledge() -> Dict:

    tree = get_github_tree()

    data_files = [
        item
        for item in tree
        if item.get("type") == "blob"
        and is_data_file(item.get("path", ""))
    ]

    all_text = []
    translation_memory = []
    loaded_files = []
    errors = []

    # Safety limit.
    # This prevents accidentally downloading hundreds of large files.
    data_files = data_files[:100]

    for item in data_files:

        path = item["path"]

        try:

            raw_url = github_raw_url(path)

            response = requests.get(
                raw_url,
                headers={
                    "User-Agent": "Tasuqilt-DZ"
                },
                timeout=20
            )

            response.raise_for_status()

            content = response.content

            # Avoid loading very large files.
            if len(content) > 8 * 1024 * 1024:
                errors.append(
                    f"{path}: file larger than 8 MB skipped."
                )
                continue

            text, pairs = read_data_file(
                path,
                content
            )

            if text:
                all_text.append(
                    f"\n===== FILE: {path} =====\n{text}"
                )

            translation_memory.extend(pairs)

            loaded_files.append(path)

        except Exception as error:

            errors.append(
                f"{path}: {str(error)}"
            )

    return {
        "files": loaded_files,
        "text": "\n".join(all_text),
        "tm": translation_memory,
        "errors": errors
    }


# ============================================================
# 10. MERGE BUILT-IN TERMINOLOGY WITH GITHUB TERMINOLOGY
# ============================================================

def build_terminology(knowledge: Dict) -> List[Dict]:

    terminology = list(DEFAULT_TERMINOLOGY)

    for pair in knowledge.get("tm", []):

        source = pair.get("source", "").strip()
        target = pair.get("target", "").strip()

        if source and target:

            terminology.append({
                "source": source,
                "target": target,
                "priority": 80
            })

    # Remove duplicate normalized pairs
    unique = {}

    for item in terminology:

        key = (
            normalize_text(item["source"]),
            normalize_text(item["target"])
        )

        unique[key] = item

    result = list(unique.values())

    result.sort(
        key=lambda x: (
            x.get("priority", 0),
            len(x.get("source", ""))
        ),
        reverse=True
    )

    return result


# ============================================================
# 11. INTELLIGENT RETRIEVAL
# ============================================================

def similarity(a: str, b: str) -> float:

    a = normalize_text(a)
    b = normalize_text(b)

    if not a or not b:
        return 0.0

    if a in b or b in a:
        return 1.0

    return difflib.SequenceMatcher(
        None,
        a,
        b
    ).ratio()


def retrieve_context(
    query: str,
    knowledge: Dict,
    terminology: List[Dict],
    direction: str,
    max_pairs: int = 8
) -> Dict:

    query_norm = normalize_text(query)
    query_tokens = set(tokenize(query))

    # --------------------------------------------------------
    # 11.1 Relevant terminology
    # --------------------------------------------------------

    matched_terms = []

    for term in terminology:

        source = term["source"]
        target = term["target"]

        source_norm = normalize_text(source)

        if not source_norm:
            continue

        if source_norm in query_norm:

            matched_terms.append({
                **term,
                "score": 1.0
            })

            continue

        # Token overlap
        source_tokens = set(tokenize(source_norm))

        if source_tokens and query_tokens:

            overlap = len(
                source_tokens.intersection(query_tokens)
            ) / max(len(source_tokens), 1)

            if overlap >= 0.50:

                matched_terms.append({
                    **term,
                    "score": overlap
                })

    # Sort
    matched_terms.sort(
        key=lambda x: (
            x["score"],
            x.get("priority", 0),
            len(x["source"])
        ),
        reverse=True
    )

    matched_terms = matched_terms[:20]

    # --------------------------------------------------------
    # 11.2 Translation Memory
    # --------------------------------------------------------

    tm_matches = []

    for pair in knowledge.get("tm", []):

        source = pair.get("source", "")
        target = pair.get("target", "")

        if not source or not target:
            continue

        score = similarity(
            query_norm,
            source
        )

        source_tokens = set(
            tokenize(source)
        )

        overlap = 0.0

        if source_tokens and query_tokens:

            overlap = len(
                source_tokens.intersection(query_tokens)
            ) / max(len(source_tokens), 1)

        final_score = max(
            score,
            overlap
        )

        # Exact / highly relevant
        if (
            normalize_text(source) in query_norm
            or final_score >= 0.55
        ):

            tm_matches.append({
                **pair,
                "score": final_score
            })

    tm_matches.sort(
        key=lambda x: (
            x["score"],
            len(x.get("source", ""))
        ),
        reverse=True
    )

    tm_matches = tm_matches[:max_pairs]

    # --------------------------------------------------------
    # 11.3 Exact Translation
    # --------------------------------------------------------

    exact_match = None

    for pair in knowledge.get("tm", []):

        if normalize_text(pair["source"]) == query_norm:

            exact_match = pair
            break

    # --------------------------------------------------------
    # 11.4 Build context
    # --------------------------------------------------------

    return {
        "exact_match": exact_match,
        "terms": matched_terms,
        "tm_matches": tm_matches
    }


# ============================================================
# 12. DETERMINE CONFIDENCE
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

        return int(
            min(75, max(50, best * 100))
        )

    return 0


# ============================================================
# 13. BUILD AI CONTEXT
# ============================================================

def build_ai_context(
    retrieval: Dict,
    direction: str
) -> str:

    sections = []

    sections.append(
        "TASUQILT OFFICIAL TERMINOLOGY:\n"
    )

    terms = retrieval.get("terms", [])

    if terms:

        for term in terms:

            sections.append(
                f"- {term['source']} = {term['target']}"
            )

    else:

        sections.append(
            "- No matching official term was found."
        )

    sections.append(
        "\nTASUQILT TRANSLATION MEMORY EXAMPLES:\n"
    )

    matches = retrieval.get("tm_matches", [])

    if matches:

        for index, pair in enumerate(matches, 1):

            sections.append(
                f"{index}. SOURCE: {pair['source']}\n"
                f"   TAMAZIGHT: {pair['target']}\n"
            )

    else:

        sections.append(
            "- No relevant translation-memory example was found."
        )

    sections.append(
        "\nIMPORTANT:"
    )

    sections.append(
        "The above information comes from the Tasuqilt knowledge base."
    )

    sections.append(
        "It has priority over general linguistic preferences."
    )

    sections.append(
        "Do not invent an official equivalent when an approved term is supplied."
    )

    return "\n".join(sections)


# ============================================================
# 14. GEMINI FREE API
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

    context = build_ai_context(
        retrieval,
        direction
    )

    if "Auto-Detect" in direction:

        task = """
Translate the source text into professional Latin Tamazight.

The source language may be French or Arabic.

Use the supplied Tasuqilt terminology and translation-memory
examples as the primary authority.

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
        "system_instruction": {
            "parts": [
                {
                    "text": system_instruction
                }
            ]
        },
        "contents": [
            {
                "role": "user",
                "parts": [
                    {
                        "text": source_text
                    }
                ]
            }
        ],
        "generationConfig": {
            "temperature": 0.0,
            "maxOutputTokens": 2000
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
            timeout=40
        )

        if response.status_code != 200:

            try:
                error_data = response.json()

                message = (
                    error_data
                    .get("error", {})
                    .get("message", "")
                )

            except Exception:

                message = response.text[:500]

            return "", (
                f"Gemini لم يُرجع ترجمة. "
                f"HTTP {response.status_code}. "
                f"{message}"
            )

        data = response.json()

        candidates = data.get("candidates", [])

        if not candidates:

            return "", "Gemini أعاد استجابة بدون نتيجة."

        parts = (
            candidates[0]
            .get("content", {})
            .get("parts", [])
        )

        output = "\n".join(
            part.get("text", "")
            for part in parts
            if part.get("text")
        ).strip()

        if not output:

            return "", "Gemini أعاد نصاً فارغاً."

        return output, ""

    except requests.exceptions.Timeout:

        return "", (
            "انتهت مهلة الاتصال بـ Gemini. "
            "حاول مرة أخرى بعد قليل."
        )

    except Exception as error:

        return "", f"خطأ في Gemini: {error}"


# ============================================================
# 15. TERMINOLOGY VALIDATOR
# ============================================================

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

    warnings = []
    missing = []

    # We only enforce source → Tamazight here.
    # For reverse translation we still display terminology
    # information but do not incorrectly reject French/Arabic.
    if "Auto-Detect" not in direction:

        return {
            "valid": True,
            "warnings": [],
            "checked": 0,
            "missing": []
        }

    output_norm = normalize_text(output_text)

    for term in retrieval.get("terms", []):

        source = term["source"]
        target = term["target"]

        source_norm = normalize_text(source)

        # If the source expression appears in input,
        # the official target should normally appear in output.
        if source_norm in normalize_text(source_text):

            target_norm = normalize_text(target)

            if target_norm not in output_norm:

                missing.append({
                    "source": source,
                    "expected": target
                })

    if missing:

        for item in missing:

            warnings.append(
                f"المصطلح الرسمي غير موجود في الناتج: "
                f"{item['source']} → {item['expected']}"
            )

        return {
            "valid": False,
            "warnings": warnings,
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
# 16. SAFE LOCAL TRANSLATION
# ============================================================

def exact_local_translation(
    source_text: str,
    knowledge: Dict,
    direction: str
) -> str:

    query_norm = normalize_text(source_text)

    # Exact TM
    for pair in knowledge.get("tm", []):

        if normalize_text(pair["source"]) == query_norm:

            if "Auto-Detect" in direction:
                return pair["target"]

            # Reverse
            if normalize_text(pair["target"]) == query_norm:
                return pair["source"]

    # Exact terminology
    for term in DEFAULT_TERMINOLOGY:

        if normalize_text(term["source"]) == query_norm:

            if "Auto-Detect" in direction:
                return term["target"]

    return ""


# ============================================================
# 17. UI
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
# 18. SIDEBAR
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


# ============================================================
# 19. LOAD GITHUB KNOWLEDGE
# ============================================================

with st.spinner("🔄 قراءة قاعدة Tasuqilt من GitHub..."):

    try:

        knowledge = load_github_knowledge()

        github_ok = True

    except Exception as error:

        knowledge = {
            "files": [],
            "text": "",
            "tm": [],
            "errors": [str(error)]
        }

        github_ok = False


terminology = build_terminology(
    knowledge
)


# ============================================================
# 20. GITHUB STATUS
# ============================================================

if github_ok:

    st.sidebar.success(
        f"🟢 GitHub متصل — "
        f"{len(knowledge['files'])} ملف بيانات"
    )

    st.sidebar.caption(
        f"Translation Memory: {len(knowledge['tm'])} زوج ترجمة"
    )

else:

    st.sidebar.error(
        "🔴 تعذر قراءة GitHub"
    )

    for error in knowledge.get("errors", [])[:3]:

        st.sidebar.caption(error)


if GEMINI_API_KEY:

    st.sidebar.success(
        "🟢 Gemini API مفعّل"
    )

else:

    st.sidebar.warning(
        "🟡 Gemini غير مفعّل — "
        "المترجم المحلي فقط متاح."
    )


# ============================================================
# 21. ADMIN
# ============================================================

st.sidebar.markdown("---")
st.sidebar.subheader("🔐 Administration")

admin_input = st.sidebar.text_input(
    "Code d'accès admin:",
    type="password"
)

if ADMIN_PASSWORD and admin_input == ADMIN_PASSWORD:

    st.session_state["admin_mode"] = True

elif admin_input:

    st.sidebar.error(
        "رمز الإدارة غير صحيح."
    )


if st.session_state["admin_mode"]:

    st.sidebar.success(
        "🔓 وضع الإدارة مفعّل"
    )

    updated_prompt = st.sidebar.text_area(
        "System Instruction:",
        value=st.session_state["custom_system_instruction"],
        height=220
    )

    if st.sidebar.button(
        "💾 حفظ التعليمات"
    ):

        st.session_state[
            "custom_system_instruction"
        ] = updated_prompt

        st.sidebar.success(
            "تم حفظ التعليمات لهذه الجلسة."
        )

    st.sidebar.markdown("---")

    st.sidebar.write(
        "ملفات البيانات المقروءة:"
    )

    for filename in knowledge.get("files", []):

        st.sidebar.caption(
            f"📄 {filename}"
        )

    if knowledge.get("errors"):

        st.sidebar.warning(
            "بعض الملفات تم تجاوزها:"
        )

        for error in knowledge["errors"][:10]:

            st.sidebar.caption(error)

    if st.sidebar.button(
        "🔄 تحديث قاعدة GitHub الآن"
    ):

        get_github_tree.clear()
        load_github_knowledge.clear()

        st.rerun()


# ============================================================
# 22. EXPRESS DICTIONARY
# ============================================================

st.markdown(
    "### 📖 Dictionnaire Express / القاموس الفوري"
)

dict_col1, dict_col2 = st.columns(2)

with dict_col1:

    word_to_find = st.text_input(
        "ابحث عن كلمة أو مصطلح:",
        placeholder="président / réunion / gouvernement..."
    )

with dict_col2:

    st.markdown("**النتيجة:**")

    if word_to_find.strip():

        query = normalize_text(
            word_to_find
        )

        found = None

        for term in terminology:

            if normalize_text(
                term["source"]
            ) == query:

                found = term
                break

        if found:

            st.success(
                found["target"]
            )

            st.caption(
                f"مصدر المصطلح: {found.get('origin', 'Tasuqilt')}"
            )

        else:

            st.info(
                "هذا المصطلح غير موجود في قاعدة Tasuqilt."
            )


st.markdown("---")


# ============================================================
# 23. MAIN TRANSLATOR
# ============================================================

st.markdown(
    "### 📰 Traducteur de Dépêches / مترجم البرقيات الإعلامية"
)

col1, col2 = st.columns(2)

with col1:

    text_to_translate = st.text_area(
        "النص المصدر:",
        height=280,
        placeholder=(
            "أدخل النص الفرنسي أو العربي هنا..."
            if "Auto-Detect" in direction
            else
            "أدخل نص Tamazight باللاتينية هنا..."
        )
    )

    submit_button = st.button(
        "🚀 ترجم",
        type="primary",
        use_container_width=True
    )


# ============================================================
# 24. TRANSLATION PIPELINE
# ============================================================

if submit_button:

    if not text_to_translate.strip():

        st.warning(
            "أدخل نصاً أولاً."
        )

    else:

        source_text = text_to_translate.strip()

        # ----------------------------------------------------
        # STEP 1 — RETRIEVE FROM Tasuqilt
        # ----------------------------------------------------

        with st.spinner(
            "🔎 البحث في قاعدة Tasuqilt..."
        ):

            retrieval = retrieve_context(
                source_text,
                knowledge,
                terminology,
                direction
            )

        st.session_state[
            "last_retrieval"
        ] = retrieval

        confidence = calculate_confidence(
            retrieval
        )

        # ----------------------------------------------------
        # STEP 2 — EXACT LOCAL MATCH
        # ----------------------------------------------------

        local_result = exact_local_translation(
            source_text,
            knowledge,
            direction
        )

        if local_result:

            output_text = local_result

            validation = {
                "valid": True,
                "warnings": [],
                "checked": 0,
                "missing": []
            }

            engine_used = (
                "Tasuqilt Translation Memory"
            )

            st.session_state[
                "last_translation"
            ] = output_text

        # ----------------------------------------------------
        # STEP 3 — GEMINI ONLY IF NECESSARY
        # ----------------------------------------------------

        elif engine_choice == "Gemini Free":

            with st.spinner(
                "🤖 لا توجد مطابقة كاملة — "
                "إرسال السياق المرتبط فقط إلى Gemini..."
            ):

                output_text, error_message = (
                    translate_with_gemini(
                        source_text,
                        direction,
                        retrieval
                    )
                )

            if error_message:

                st.error(
                    error_message
                )

                output_text = ""

                engine_used = "Gemini"

                validation = {
                    "valid": False,
                    "warnings": [error_message],
                    "checked": 0,
                    "missing": []
                }

            else:

                engine_used = (
                    "Gemini Free + Tasuqilt Retrieval"
                )

                # ------------------------------------------------
                # STEP 4 — VALIDATE
                # ------------------------------------------------

                validation = validate_translation(
                    source_text,
                    output_text,
                    retrieval,
                    direction
                )

                if not validation["valid"]:

                    output_text = ""

                    st.error(
                        "⛔ لم تعتمد Tasuqilt الترجمة."
                    )

                    for warning in validation[
                        "warnings"
                    ]:

                        st.warning(
                            warning
                        )

        # ----------------------------------------------------
        # STEP 5 — LOCAL ONLY WHEN NO GEMINI
        # ----------------------------------------------------

        else:

            output_text = ""

            engine_used = (
                "Tasuqilt Local"
            )

            validation = {
                "valid": False,
                "warnings": [
                    "لا توجد مطابقة كاملة في قاعدة Tasuqilt."
                ],
                "checked": 0,
                "missing": []
            }

            st.info(
                "وضع Tasuqilt Local لا يخترع ترجمة. "
                "أضف الجملة إلى Translation Memory "
                "أو فعّل Gemini Free."
            )


        # ----------------------------------------------------
        # STEP 6 — DISPLAY RESULT
        # ----------------------------------------------------

        with col2:

            if output_text:

                st.markdown(
                    "**الترجمة المعتمدة:**"
                )

                st.success(
                    output_text
                )

                st.text_area(
                    "النتيجة القابلة للنسخ:",
                    value=output_text,
                    height=280
                )

                st.caption(
                    f"⚙️ المحرك: {engine_used}"
                )

                # Confidence
                if confidence >= 90:

                    st.success(
                        f"🟢 اعتماد قاعدة Tasuqilt: "
                        f"{confidence}%"
                    )

                elif confidence >= 60:

                    st.warning(
                        f"🟡 اعتماد قاعدة Tasuqilt: "
                        f"{confidence}%"
                    )

                else:

                    st.info(
                        f"🔵 اعتماد قاعدة Tasuqilt: "
                        f"{confidence}%"
                    )

            else:

                st.text_area(
                    "النتيجة:",
                    value="",
                    height=280,
                    disabled=True
                )


# ============================================================
# 25. RETRIEVAL INFORMATION
# ============================================================

if st.session_state.get(
    "last_retrieval"
):

    retrieval = st.session_state[
        "last_retrieval"
    ]

    with st.expander(
        "🔎 عرض ما وجده Tasuqilt قبل الترجمة"
    ):

        terms = retrieval.get(
            "terms", []
        )

        matches = retrieval.get(
            "tm_matches", []
        )

        if terms:

            st.markdown(
                "#### 📚 المصطلحات الرسمية"
            )

            for term in terms:

                st.write(
                    f"**{term['source']}** → "
                    f"**{term['target']}**"
                )

        else:

            st.write(
                "لم يجد مصطلحات رسمية مطابقة."
            )

        if matches:

            st.markdown(
                "#### 🧠 أمثلة Translation Memory"
            )

            for pair in matches:

                st.write(
                    f"**Original:** {pair['source']}"
                )

                st.write(
                    f"**Tamazight:** {pair['target']}"
                )

                st.caption(
                    f"درجة التشابه: "
                    f"{round(pair['score'] * 100)}%"
                )

                st.markdown("---")

        else:

            st.write(
                "لم يجد أمثلة Translation Memory قريبة."
            )


# ============================================================
# 26. FOOTER
# ============================================================

st.markdown("---")

st.caption(
    "Tasuqilt DZ — نظام ترجمة يعتمد على قاعدة المصطلحات "
    "وذاكرة الترجمة الخاصة بالمشروع قبل استخدام الذكاء الاصطناعي."
)
