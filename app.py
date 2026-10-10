import io
import os
import re
import json
import time
import unicodedata
from difflib import SequenceMatcher

import requests
import pandas as pd
import streamlit as st
from docx import Document
from openpyxl import load_workbook


# ============================================================
# 1. CONFIGURATION
# ============================================================

APP_TITLE = "Tasuqilt DZ"
GITHUB_OWNER = "chaabanemeddour05-ai"
GITHUB_REPO = "Tasuqilt"

GITHUB_API = (
    f"https://api.github.com/repos/"
    f"{GITHUB_OWNER}/{GITHUB_REPO}"
)

DEFAULT_GEMINI_MODEL = "gemini-3.8-flash"

ALLOWED_EXTENSIONS = {
    ".xlsx", ".csv", ".json", ".txt", ".md", ".docx"
}

OFFICIAL_TERMINOLOGY = [
    {
        "source": "Président de la République",
        "target": "Aselway n Tegduda",
    },
    {
        "source": "Président",
        "target": "Aselway",
    },
    {
        "source": "le président",
        "target": "Aselway",
    },
    {
        "source": "Conseil des ministres",
        "target": "Aseqqamu n Yineɣlaf",
    },
    {
        "source": "Réunion",
        "target": "Timlilt",
    },
    {
        "source": "Alger",
        "target": "DZAYER TAMANEƔT",
    },
    {
        "source": "Gouvernement",
        "target": "Anabaḍ",
    },
]

DIRECTIONS = {
    "Français → Tamazight": {
        "source": "fr",
        "target": "tz",
    },
    "العربية → Tamazight": {
        "source": "ar",
        "target": "tz",
    },
    "Tamazight → Français": {
        "source": "tz",
        "target": "fr",
    },
    "Tamazight → العربية": {
        "source": "tz",
        "target": "ar",
    },
}


# ============================================================
# 2. STREAMLIT PAGE
# ============================================================

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="🌿",
    layout="wide",
)

st.title("🌿 Tasuqilt DZ")
st.caption(
    "منصة الترجمة الأمازيغية المعيارية "
    "للخطاب الإعلامي والصحفي"
)


# ============================================================
# 3. SECRETS AND SETTINGS
# ============================================================

def get_setting(name, default=""):
    """Read a setting from Streamlit Secrets or environment."""

    try:
        value = st.secrets.get(name, default)
        if value:
            return str(value).strip()
    except Exception:
        pass

    return str(os.environ.get(name, default)).strip()


GITHUB_TOKEN = get_setting("GITHUB_TOKEN")
GEMINI_API_KEY = get_setting("GEMINI_API_KEY")
GEMINI_MODEL = get_setting(
    "GEMINI_MODEL",
    DEFAULT_GEMINI_MODEL,
)

REQUEST_TIMEOUT = 30
MAX_FILE_SIZE = 8 * 1024 * 1024


# ============================================================
# 4. TEXT NORMALIZATION
# ============================================================

def normalize_text(value):
    """Normalize text for exact matching without removing Amazigh letters."""

    if value is None:
        return ""

    text = unicodedata.normalize("NFKC", str(value))
    text = text.replace("\u00a0", " ")
    text = text.replace("\u200b", "")

    text = re.sub(r"\s+", " ", text).strip()

    return text.casefold()


def normalize_for_search(value):
    """Create a searchable representation while preserving word boundaries."""

    text = normalize_text(value)
    text = re.sub(r"[^\w\s]", " ", text, flags=re.UNICODE)
    text = re.sub(r"\s+", " ", text).strip()

    return text


def is_empty(value):
    return value is None or not str(value).strip()


def clean_cell(value):
    if value is None:
        return ""

    return str(value).strip()


# ============================================================
# 5. GITHUB CONNECTION
# ============================================================

def github_headers():
    headers = {
        "Accept": "application/vnd.github+json",
        "User-Agent": "Tasuqilt-DZ",
        "X-GitHub-Api-Version": "2022-11-28",
    }

    if GITHUB_TOKEN:
        headers["Authorization"] = f"Bearer {GITHUB_TOKEN}"

    return headers


def github_request(url):
    response = requests.get(
        url,
        headers=github_headers(),
        timeout=REQUEST_TIMEOUT,
    )

    response.raise_for_status()
    return response


@st.cache_data(ttl=300, show_spinner=False)
def fetch_repository_files():
    """
    Read the repository tree and download supported data files.
    The Excel translation memory is supported.
    """

    repo_response = github_request(GITHUB_API)
    repo_info = repo_response.json()

    default_branch = repo_info.get("default_branch", "main")

    tree_url = (
        f"{GITHUB_API}/git/trees/"
        f"{default_branch}?recursive=1"
    )

    tree_response = github_request(tree_url)
    tree_data = tree_response.json()

    if tree_data.get("truncated"):
        raise RuntimeError(
            "GitHub أعاد شجرة ملفات غير مكتملة."
        )

    files = {}

    for item in tree_data.get("tree", []):
        if item.get("type") != "blob":
            continue

        path = item.get("path", "")
        filename = path.rsplit("/", 1)[-1]
        extension = os.path.splitext(filename)[1].lower()

        # Only read supported files in the repository root.
        if "/" in path:
            continue

        if extension not in ALLOWED_EXTENSIONS:
            continue

        if filename.startswith("."):
            continue

        if item.get("size", 0) > MAX_FILE_SIZE:
            continue

        raw_url = (
            f"https://raw.githubusercontent.com/"
            f"{GITHUB_OWNER}/{GITHUB_REPO}/"
            f"{default_branch}/{requests.utils.quote(path)}"
        )

        try:
            response = requests.get(
                raw_url,
                headers={
                    "User-Agent": "Tasuqilt-DZ",
                },
                timeout=REQUEST_TIMEOUT,
            )
            response.raise_for_status()

            files[filename] = {
                "content": response.content,
                "path": path,
                "size": item.get("size", 0),
            }

        except Exception as exc:
            files[filename] = {
                "error": str(exc),
                "path": path,
                "size": item.get("size", 0),
            }

    return {
        "branch": default_branch,
        "files": files,
    }


# ============================================================
# 6. TRANSLATION MEMORY PARSERS
# ============================================================

def looks_like_header(first, second):
    """Detect common source/target column headings."""

    a = normalize_text(first)
    b = normalize_text(second)

    source_headers = {
        "français", "francais", "french", "source",
        "texte source", "original", "original text",
        "fr", "arabic", "العربية",
    }

    target_headers = {
        "tamazight", "amazigh", "kabyle", "target",
        "translation", "traduction", "destination",
        "tz", "berbère", "berbere",
    }

    return a in source_headers and b in target_headers


def make_pair(source, target, origin=""):
    source = clean_cell(source)
    target = clean_cell(target)

    if not source or not target:
        return None

    if normalize_text(source) == normalize_text(target):
        return None

    return {
        "source": source,
        "target": target,
        "origin": origin,
    }


def parse_excel(content, filename):
    pairs = []

    workbook = load_workbook(
        io.BytesIO(content),
        read_only=True,
        data_only=True,
    )

    try:
        sheet = None

        for candidate in workbook.worksheets:
            if candidate.title.strip().casefold() == "tm":
                sheet = candidate
                break

        if sheet is None:
            sheet = workbook.worksheets[0]

        for row_number, row in enumerate(
            sheet.iter_rows(min_col=1, max_col=2, values_only=True),
            start=1,
        ):
            if not row or len(row) < 2:
                continue

            source = clean_cell(row[0])
            target = clean_cell(row[1])

            if not source or not target:
                continue

            if row_number == 1 and looks_like_header(
                source, target
            ):
                continue

            pair = make_pair(
                source,
                target,
                origin=filename,
            )

            if pair:
                pairs.append(pair)

    finally:
        workbook.close()

    return pairs


def parse_csv(content, filename):
    pairs = []

    try:
        df = pd.read_csv(
            io.BytesIO(content),
            dtype=str,
            keep_default_na=False,
            encoding="utf-8-sig",
        )
    except Exception:
        df = pd.read_csv(
            io.BytesIO(content),
            dtype=str,
            keep_default_na=False,
            encoding="latin-1",
        )

    if len(df.columns) < 2:
        return pairs

    for index, row in df.iterrows():
        source = clean_cell(row.iloc[0])
        target = clean_cell(row.iloc[1])

        if index == 0 and looks_like_header(source, target):
            continue

        pair = make_pair(source, target, filename)

        if pair:
            pairs.append(pair)

    return pairs


def parse_json_pairs(content, filename):
    pairs = []

    try:
        data = json.loads(content.decode("utf-8-sig"))
    except Exception:
        return pairs

    if isinstance(data, dict):
        if isinstance(data.get("pairs"), list):
            data = data["pairs"]
        elif isinstance(data.get("translations"), list):
            data = data["translations"]
        else:
            data = [data]

    if not isinstance(data, list):
        return pairs

    for item in data:
        if not isinstance(item, dict):
            continue

        source = (
            item.get("source")
            or item.get("fr")
            or item.get("french")
            or item.get("français")
            or item.get("original")
        )

        target = (
            item.get("target")
            or item.get("tz")
            or item.get("tamazight")
            or item.get("amazigh")
            or item.get("translation")
        )

        pair = make_pair(source, target, filename)

        if pair:
            pairs.append(pair)

    return pairs


def parse_docx_pairs(content, filename):
    pairs = []

    document = Document(io.BytesIO(content))

    # Paragraph format: source @ target
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()

        if "@" not in text:
            continue

        source, target = text.split("@", 1)

        pair = make_pair(source, target, filename)

        if pair:
            pairs.append(pair)

    # Table format: first cell = source, second cell = target
    for table in document.tables:
        for row_number, row in enumerate(table.rows):
            if len(row.cells) < 2:
                continue

            source = row.cells[0].text.strip()
            target = row.cells[1].text.strip()

            if row_number == 0 and looks_like_header(
                source, target
            ):
                continue

            pair = make_pair(source, target, filename)

            if pair:
                pairs.append(pair)

    return pairs


def parse_text_pairs(content, filename):
    pairs = []

    text = content.decode("utf-8-sig", errors="replace")

    for line in text.splitlines():
        line = line.strip()

        if "@" not in line:
            continue

        source, target = line.split("@", 1)

        pair = make_pair(source, target, filename)

        if pair:
            pairs.append(pair)

    return pairs


def parse_data_file(filename, content):
    extension = os.path.splitext(filename)[1].lower()

    if extension == ".xlsx":
        return parse_excel(content, filename)

    if extension == ".csv":
        return parse_csv(content, filename)

    if extension == ".json":
        return parse_json_pairs(content, filename)

    if extension == ".docx":
        return parse_docx_pairs(content, filename)

    if extension in {".txt", ".md"}:
        return parse_text_pairs(content, filename)

    return []


# ============================================================
# 7. TERMINOLOGY LOADING
# ============================================================

def parse_terminology_file(filename, content):
    """
    Terminology files should contain short terms or expressions,
    not entire news articles or general TM sentences.
    """

    extension = os.path.splitext(filename)[1].lower()
    terms = []

    if extension == ".csv":
        try:
            df = pd.read_csv(
                io.BytesIO(content),
                dtype=str,
                keep_default_na=False,
                encoding="utf-8-sig",
            )
        except Exception:
            return terms

        if len(df.columns) < 2:
            return terms

        for _, row in df.iterrows():
            source = clean_cell(row.iloc[0])
            target = clean_cell(row.iloc[1])

            if source and target:
                terms.append({
                    "source": source,
                    "target": target,
                })

    elif extension == ".json":
        try:
            data = json.loads(content.decode("utf-8-sig"))
        except Exception:
            return terms

        if isinstance(data, dict):
            data = data.get("terms", data.get("terminology", data))

        if isinstance(data, dict):
            for source, target in data.items():
                if isinstance(target, str):
                    terms.append({
                        "source": str(source).strip(),
                        "target": target.strip(),
                    })

        elif isinstance(data, list):
            for item in data:
                if not isinstance(item, dict):
                    continue

                source = item.get("source") or item.get("fr")
                target = item.get("target") or item.get("tamazight")

                if source and target:
                    terms.append({
                        "source": str(source).strip(),
                        "target": str(target).strip(),
                    })

    return terms


def build_terminology(files):
    """
    Official terminology is separate from the translation memory.
    Never convert every TM sentence into an official term.
    """

    terminology = list(OFFICIAL_TERMINOLOGY)

    allowed_names = {
        "terminology.csv",
        "lexicon.csv",
        "terminology.json",
        "lexicon.json",
    }

    for filename, file_info in files.items():
        if filename.lower() not in allowed_names:
            continue

        if "content" not in file_info:
            continue

        terminology.extend(
            parse_terminology_file(
                filename,
                file_info["content"],
            )
        )

    unique = {}
    for term in terminology:
        source = term.get("source", "").strip()
        target = term.get("target", "").strip()

        if not source or not target:
            continue

        key = normalize_text(source)

        if key not in unique:
            unique[key] = {
                "source": source,
                "target": target,
            }

    return list(unique.values())


# ============================================================
# 8. LOAD AND INDEX KNOWLEDGE
# ============================================================

@st.cache_data(ttl=300, show_spinner=False)
def load_knowledge():
    result = {
        "connected": False,
        "branch": "",
        "files": {},
        "pairs": [],
        "terminology": [],
        "errors": [],
    }

    try:
        repository = fetch_repository_files()

        result["connected"] = True
        result["branch"] = repository["branch"]
        result["files"] = repository["files"]

    except Exception as exc:
        result["errors"].append(
            f"تعذّر الاتصال بمستودع GitHub: {exc}"
        )
        return result

    all_pairs = []

    for filename, file_info in result["files"].items():
        if "error" in file_info:
            result["errors"].append(
                f"تعذّر تحميل {filename}: {file_info['error']}"
            )
            continue

        try:
            pairs = parse_data_file(
                filename,
                file_info["content"],
            )

            all_pairs.extend(pairs)

        except Exception as exc:
            result["errors"].append(
                f"تعذّرت قراءة {filename}: {exc}"
            )

    # Remove duplicate source-target pairs.
    unique_pairs = {}
    for pair in all_pairs:
        key = (
            normalize_text(pair["source"]),
            normalize_text(pair["target"]),
        )

        if key not in unique_pairs:
            unique_pairs[key] = pair

    result["pairs"] = list(unique_pairs.values())
    result["terminology"] = build_terminology(
        result["files"]
    )

    return result


# ============================================================
# 9. EXACT MATCH TRANSLATION
# ============================================================

def exact_local_translation(text, direction, knowledge):
    """
    Exact matches only.
    The local engine never fabricates a translation.
    """

    query = normalize_text(text)

    if not query:
        return None

    pairs = knowledge["pairs"]

    if direction == "Français → Tamazight":
        for pair in pairs:
            if normalize_text(pair["source"]) == query:
                return pair["target"]

        for term in knowledge["terminology"]:
            if normalize_text(term["source"]) == query:
                return term["target"]

    elif direction == "Tamazight → Français":
        for pair in pairs:
            if normalize_text(pair["target"]) == query:
                return pair["source"]

        for term in knowledge["terminology"]:
            if normalize_text(term["target"]) == query:
                return term["source"]

    # No exact Arabic pairs are assumed to exist in a French-Amazigh TM.
    return None


# ============================================================
# 10. SMART RETRIEVAL
# ============================================================

def token_overlap_score(query, candidate):
    query_tokens = set(normalize_for_search(query).split())
    candidate_tokens = set(normalize_for_search(candidate).split())

    if not query_tokens or not candidate_tokens:
        return 0.0

    intersection = len(query_tokens & candidate_tokens)
    union = len(query_tokens | candidate_tokens)

    jaccard = intersection / union if union else 0.0

    sequence = SequenceMatcher(
        None,
        normalize_for_search(query),
        normalize_for_search(candidate),
    ).ratio()

    return max(jaccard, sequence * 0.7)


def retrieve_context(text, direction, knowledge, limit=6):
    """
    Retrieve similar examples from the correct side of the TM.
    Reverse translation searches the Amazigh target column.
    """

    if direction == "Français → Tamazight":
        source_key = "source"
        target_key = "target"

    elif direction == "Tamazight → Français":
        source_key = "target"
        target_key = "source"

    else:
        # The current TM is French-Amazigh.
        # For Arabic input, French/Amazigh examples may help Gemini,
        # but they are not presented as exact Arabic matches.
        source_key = None
        target_key = None

    if source_key is None:
        return []

    scored = []

    for pair in knowledge["pairs"]:
        candidate = pair[source_key]

        score = token_overlap_score(text, candidate)

        if score >= 0.16:
            scored.append({
                "source": candidate,
                "target": pair[target_key],
                "score": score,
            })

    scored.sort(
        key=lambda item: item["score"],
        reverse=True,
    )

    return scored[:limit]


def format_context(examples):
    if not examples:
        return "لا توجد أمثلة قريبة كافية في ذاكرة الترجمة."

    lines = []

    for index, example in enumerate(examples, start=1):
        lines.append(
            f"{index}. المصدر: {example['source']}\n"
            f"   المقابل: {example['target']}"
        )

    return "\n".join(lines)


# ============================================================
# 11. TERM VALIDATION
# ============================================================

def validate_translation(source_text, translated_text, terminology):
    """
    Warn only when a relevant official term was omitted.
    This is a terminology check, not a full linguistic validator.
    """

    source_normalized = normalize_for_search(source_text)
    target_normalized = normalize_for_search(translated_text)

    warnings = []

    for term in terminology:
        source_term = term["source"].strip()
        target_term = term["target"].strip()

        normalized_source_term = normalize_for_search(source_term)
        normalized_target_term = normalize_for_search(target_term)

        if not normalized_source_term:
            continue

        # Avoid treating every long sentence as a term.
        if len(normalized_source_term.split()) > 5:
            continue

        if normalized_source_term in source_normalized:
            if normalized_target_term not in target_normalized:
                warnings.append(
                    f"راجع المصطلح: «{source_term}» "
                    f"والمقابل المعياري «{target_term}»."
                )

    return warnings


# ============================================================
# 12. GEMINI API
# ============================================================

def build_gemini_prompt(
    source_text,
    direction,
    examples,
    terminology,
):
    term_lines = []

    source_language = DIRECTIONS[direction]["source"]

    for term in terminology:
        source_term = term["source"]
        target_term = term["target"]

        if source_language == "fr":
            term_lines.append(
                f"{source_term} → {target_term}"
            )

        elif source_language == "tz":
            term_lines.append(
                f"{target_term} → {source_term}"
            )

    if not term_lines:
        term_block = "لا توجد قائمة مصطلحات إضافية."
    else:
        term_block = "\n".join(term_lines[:100])

    if direction == "Français → Tamazight":
        instruction = (
            "Translate from French into standard written Amazigh "
            "using Latin script."
        )

    elif direction == "العربية → Tamazight":
        instruction = (
            "Translate from Arabic into standard written Amazigh "
            "using Latin script. Preserve the full meaning and "
            "journalistic register."
        )

    elif direction == "Tamazight → Français":
        instruction = (
            "Translate from Amazigh written in Latin script into French."
        )

    else:
        instruction = (
            "Translate from Amazigh written in Latin script into Arabic."
        )

    return f"""
You are the translation engine of Tasuqilt DZ, a platform for
rigorous standard Amazigh translation for Algerian news journalism.

TASK:
{instruction}

EDITORIAL RULES:
- Preserve the source meaning, facts, names, numbers, dates and quotations.
- Do not add facts that are absent from the source.
- Use clear, formal journalistic language.
- Follow the supplied terminology consistently whenever applicable.
- The approved translation of "Président" is "Aselway".
- The approved translation of "Réunion" is "Timlilt".
- Do not use "Anmazul" as a translation of "Président".
- Do not explain your choices.
- Return only the translated text, with no introduction or quotation marks.
- Do not claim that a result is an exact database match unless it is one.

RETRIEVED TRANSLATION MEMORY EXAMPLES:
{format_context(examples)}

APPROVED TERMINOLOGY:
{term_block}

TEXT TO TRANSLATE:
{source_text}
""".strip()


def call_gemini(source_text, direction, knowledge):
    if not GEMINI_API_KEY:
        raise RuntimeError(
            "مفتاح GEMINI_API_KEY غير موجود في إعدادات التطبيق."
        )

    examples = retrieve_context(
        source_text,
        direction,
        knowledge,
    )

    prompt = build_gemini_prompt(
        source_text,
        direction,
        examples,
        knowledge["terminology"],
    )

    endpoint = (
        "https://generativelanguage.googleapis.com/"
        f"v1beta/models/{GEMINI_MODEL}:generateContent"
    )

    payload = {
        "systemInstruction": {
            "parts": [
                {
                    "text": (
                        "You are a careful professional translator. "
                        "Follow the requested language direction and "
                        "return only the translation."
                    )
                }
            ]
        },
        "contents": [
            {
                "role": "user",
                "parts": [
                    {
                        "text": prompt,
                    }
                ],
            }
        ],
        "generationConfig": {
            "temperature": 0.1,
        },
    }

    response = requests.post(
        endpoint,
        params={"key": GEMINI_API_KEY},
        json=payload,
        timeout=60,
    )

    if response.status_code != 200:
        detail = response.text[:1200]

        raise RuntimeError(
            f"خطأ Gemini HTTP {response.status_code}: {detail}"
        )

    data = response.json()

    candidates = data.get("candidates", [])

    if not candidates:
        raise RuntimeError(
            "لم يُرجع Gemini أي ترجمة. "
            "قد يكون الطلب محجوبًا أو لم ينتج محتوى."
        )

    parts = (
        candidates[0]
        .get("content", {})
        .get("parts", [])
    )

    translated = "\n".join(
        part.get("text", "")
        for part in parts
        if part.get("text")
    ).strip()

    if not translated:
        raise RuntimeError(
            "أعاد Gemini استجابة فارغة."
        )

    return translated, examples


# ============================================================
# 13. SIDEBAR AND DATA STATUS
# ============================================================

with st.sidebar:
    st.header("⚙️ إعدادات Tasuqilt")

    engine_choice = st.selectbox(
        "محرك الترجمة",
        [
            "Tasuqilt Local — مطابقة قاعدة البيانات فقط",
            "Gemini — ترجمة بالذكاء الاصطناعي",
        ],
        index=0,
    )

    st.caption(
        "المحرك المحلي لا يخترع ترجمة عند غياب المطابقة."
    )

    st.divider()

    if st.button(
        "🔄 تحديث بيانات GitHub",
        use_container_width=True,
    ):
        fetch_repository_files.clear()
        load_knowledge.clear()
        st.rerun()

    st.divider()

    st.markdown("**حالة الخدمات**")

    st.write(
        "🟢 مفتاح Gemini موجود"
        if GEMINI_API_KEY
        else "⚪ مفتاح Gemini غير مضبوط"
    )

    if GEMINI_API_KEY:
        st.caption(f"النموذج: {GEMINI_MODEL}")


with st.spinner("تحميل ذاكرة الترجمة والمصطلحات..."):
    knowledge = load_knowledge()


col1, col2, col3 = st.columns(3)

with col1:
    if knowledge["connected"]:
        st.success(
            f"GitHub متصل — الفرع: {knowledge['branch']}"
        )
    else:
        st.error("تعذّر الاتصال بـ GitHub")

with col2:
    st.metric(
        "أزواج ذاكرة الترجمة",
        f"{len(knowledge['pairs']):,}",
    )

with col3:
    st.metric(
        "المصطلحات المعيارية",
        f"{len(knowledge['terminology']):,}",
    )


if knowledge["errors"]:
    with st.expander(
        f"تفاصيل التحميل ({len(knowledge['errors'])})",
        expanded=True,
    ):
        for error in knowledge["errors"]:
            st.warning(error)


with st.expander("الملفات التي عثر عليها التطبيق"):
    if not knowledge["files"]:
        st.info(
            "لم يعثر التطبيق على ملفات بيانات مدعومة في جذر المستودع."
        )
    else:
        for filename, info in knowledge["files"].items():
            if "error" in info:
                st.write(f"❌ {filename}: {info['error']}")
            else:
                size_kb = info.get("size", 0) / 1024
                st.write(
                    f"📄 {filename} — {size_kb:.1f} KB"
                )


# ============================================================
# 14. TRANSLATION INTERFACE
# ============================================================

st.divider()

st.subheader("الترجمة")

direction = st.selectbox(
    "اتجاه الترجمة",
    list(DIRECTIONS.keys()),
)

source_text = st.text_area(
    "النص المراد ترجمته",
    height=220,
    placeholder="ألصق النص هنا...",
)

translate_button = st.button(
    "ترجم النص",
    type="primary",
    use_container_width=True,
)


if translate_button:
    if not source_text.strip():
        st.warning("أدخل النص الذي تريد ترجمته أولًا.")

    elif engine_choice.startswith("Tasuqilt Local"):
        local_result = exact_local_translation(
            source_text,
            direction,
            knowledge,
        )

        if local_result:
            st.success("مطابقة تامة في قاعدة البيانات")

            st.text_area(
                "الترجمة",
                value=local_result,
                height=220,
            )

            st.caption(
                "النتيجة مسترجعة من ذاكرة الترجمة أو المصطلحات، "
                "وليست ترجمة مولّدة."
            )

        else:
            st.info(
                "لم نعثر على مطابقة تامة لهذا النص في قاعدة البيانات. "
                "لم يُنشئ الوضع المحلي ترجمة تخمينية."
            )

            if direction in {
                "Français → Tamazight",
                "Tamazight → Français",
            }:
                examples = retrieve_context(
                    source_text,
                    direction,
                    knowledge,
                    limit=3,
                )

                if examples:
                    st.markdown("**أمثلة قريبة من ذاكرة الترجمة**")

                    for example in examples:
                        st.write(
                            f"**{example['source']}**"
                        )
                        st.write(example["target"])
                        st.caption(
                            f"درجة التشابه التقريبية: "
                            f"{example['score']:.2f}"
                        )

            st.caption(
                "يمكنك اختيار محرك Gemini إذا أردت ترجمة مولّدة "
                "بالذكاء الاصطناعي."
            )

    else:
        if not GEMINI_API_KEY:
            st.error(
                "لا يمكن تشغيل Gemini لأن مفتاح API غير مضبوط. "
                "أضف GEMINI_API_KEY إلى Streamlit Secrets."
            )

        else:
            with st.spinner("جارٍ إعداد الترجمة..."):
                try:
                    
                    translated, examples = call_gemini(
                        source_text,
                        direction,
                        knowledge,
                    )

                    st.success("اكتملت الترجمة")

                    st.text_area(
                        "الترجمة",
                        value=translated,
                        height=260,
                    )

                    warnings = validate_translation(
                        source_text,
                        translated,
                        knowledge["terminology"],
                    )

                    if warnings:
                        with st.expander(
                            "مراجعة المصطلحات المعيارية",
                            expanded=True,
                        ):
                            for warning in warnings:
                                st.warning(warning)

                    st.caption(
                        "هذه ترجمة مولّدة بالذكاء الاصطناعي؛ "
                        "يجب مراجعتها قبل اعتمادها في النشر الصحفي."
                    )

                    if examples:
                        with st.expander(
                            "أمثلة ذاكرة الترجمة التي استُخدمت"
                        ):
                            for example in examples:
                                st.write(
                                    f"**{example['source']}**"
                                )
                                st.write(example["target"])

                    st.download_button(
                        "تنزيل الترجمة بصيغة TXT",
                        data=translated.encode("utf-8"),
                        file_name="tasuqilt_translation.txt",
                        mime="text/plain",
                    )

                except Exception as exc:
                    st.error("تعذّرت الترجمة عبر Gemini.")
                    st.code(str(exc))


# ============================================================
# 15. FOOTER
# ============================================================

st.divider()

st.caption(
    "Tasuqilt DZ — مشروع للترجمة الأمازيغية المعيارية. "
    "المطابقة في ذاكرة الترجمة لا تعادل التحقق اللغوي الكامل، "
    "والترجمة المولّدة تحتاج إلى مراجعة تحريرية."
)
