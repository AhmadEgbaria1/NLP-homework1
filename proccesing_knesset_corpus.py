#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Build a JSONL corpus from Knesset .docx protocols.
Usage:
    python processing_knesset_corpus.py <input_dir> <output.jsonl>

Rules:
- Each output line = one JSON object for one sentence.
- Keep only sentences with >= 4 tokens (token = whitespace-split word).
- No morphological splitting (e.g., "לכולם" stays one token).
- Robust to noisy docs: never crash the whole run on a single bad file.
"""

import sys, os, re, json
from pathlib import Path
from typing import List, Tuple, Optional
from docx import Document  # pip install python-docx

# --- מספרים בעברית ---

HEB_UNITS = {
    "אפס":0, "אחד":1, "אחת":1, "שניים":2, "שתיים":2, "שני":2, "שתי":2,
    "שלוש":3, "שלושה":3, "ארבע":4, "ארבעה":4, "חמש":5, "חמישה":5,
    "שש":6, "שישה":6, "שבע":7, "שבעה":7, "שמונה":8, "תשע":9, "תשעה":9,
    # צורות סמיכות (ל־אלפים)
    "שלושת":3, "ארבעת":4, "חמשת":5, "ששת":6, "שבעת":7, "שמונת":8, "תשעת":9
}
HEB_TENS = {
    "עשר":10, "עשרה":10,
    # 11–19 וריאנטים
    "אחת-עשרה":11, "אחת־עשרה":11, "אחת עשרה":11, "אחת-עשר":11, "אחת עשר":11,
    "שתים-עשרה":12, "שתים־עשרה":12, "שתים עשרה":12,
    "שנים-עשרה":12, "שנים־עשרה":12, "שנים עשרה":12,
    "שלוש-עשרה":13, "שלוש עשרה":13,
    "ארבע-עשרה":14, "ארבע עשרה":14,
    "חמש-עשרה":15, "חמש עשרה":15,
    "שש-עשרה":16, "שש עשרה":16,
    "שבע-עשרה":17, "שבע עשרה":17,
    "שמונה-עשרה":18, "שמונה עשרה":18,
    "תשע-עשרה":19, "תשע עשרה":19,
    # עשרות עגולות
    "עשרים":20, "שלושים":30, "ארבעים":40, "חמישים":50,
    "שישים":60, "שבעים":70, "שמונים":80, "תשעים":90
}
HEB_HUNDREDS = {"מאה":100, "מאות":100, "מאתיים":200}
HEB_THOUSANDS = {"אלף":1000, "אלפים":1000, "אלפיים":2000}

# --- Parties & role markers (derived from your JSONL) ---
PARTIES = {
    "הליכוד","העבודה","מרצ","ש\"ס","מפד\"ל","יהדות התורה","חד\"ש","בל\"ד","שינוי","גשר",
    "ישראל ביתנו","התנועה","הבית היהודי","רע\"ם","הרשימה המשותפת","הרשימה הערבית המאוחדת",
    "המחנה הציוני","הליכוד-גשר-צומת","האיחוד הלאומי - ישראל ביתנו","האיחוד הלאומי - מפד\"ל",
    "חד\"ש-בל\"ד","חד\"ש-תע\"ל","התנועה הערבית להתחדשות","מימד","עם אחד","חרות – התנועה הלאומית",
    "אגודת ישראל","דגל התורה","יחד","תע\"ל","עלה ירוק","יהדות התורה המאוחדת","רשימת המדע והטכנולוגיה",
    "העבודה-מימד","חד\"ש-תע\"ל-בל\"ד","אבו","הרשימה הדמוקרטית הערבית"
}

ROLE_HINTS = ("יו\"ר","יו״ר","ועדת","ועדה","השר","שרת","סגן","מ\"מ","מ״מ","זמני","ממשלת")



def _strip_vav(tok: str) -> str:
    # הסר ו'/ה' חיבור בתחילת המילה (למשל: "והתשעים" -> "תשעים", "ושתיים" -> "שתיים")
    while tok.startswith(("ו", "ה")) and len(tok) > 1:
        tok = tok[1:]
    return tok



def hebrew_words_to_int(text: str) -> Optional[int]:
    """
    ממיר צירוף מספרי בעברית (עם מקפים/ו' חיבור) למספר שלם.
    דוגמאות: 'שלוש-מאות-ושישים-ושש', 'אלף-ומאתיים-ושלוש', 'שישים ושמונה'.
    מחזיר None אם לא זוהה.
    """
    if not text or not text.strip():
        return None

    # נירמול מקפים שונים ל־"-"
    t = (text.replace("־", "-")
              .replace("–", "-")
              .replace("—", "-"))
    # פיצול: גם מקפים וגם רווחים
    raw_tokens = []
    for chunk in t.split():
        raw_tokens.extend(chunk.split("-"))

    tokens = [_strip_vav(tok.strip(" ,.:;\"'()[]{}")) for tok in raw_tokens if tok.strip()]
    if not tokens:
        return None

    total = 0
    current = 0

    def flush_current():
        nonlocal total, current
        total += current
        current = 0

    i = 0
    while i < len(tokens):
        tok = tokens[i]

        # אלפים / אלף / אלפיים
        if tok in HEB_THOUSANDS:
            factor = HEB_THOUSANDS[tok]
            # אם יש ערך מצטבר לפני "אלף/אלפים" נכפיל אותו, אחרת נניח 1
            current = max(current, 1) * factor
            flush_current()

        # מאות / מאתיים / מאות
        elif tok in HEB_HUNDREDS:
            factor = HEB_HUNDREDS[tok]
            current = max(current, 1) * factor

        # עשרות (כולל אחת-עשרה וכו')
        elif tok in HEB_TENS:
            current += HEB_TENS[tok]

        else:
            # תבנית "X עשרה" לאחר שפוצל – נסה לחבר חזרה
            if i + 1 < len(tokens) and (tokens[i+1] in ("עשרה", "עשר")):
                pair = f"{tok}-עשרה"
                if pair in HEB_TENS:
                    current += HEB_TENS[pair]
                    i += 1
                elif tok in HEB_UNITS:
                    current += 10 + HEB_UNITS[tok]
                    i += 1
                else:
                    # לא נראה כמספר
                    pass

            # יחידות (כולל צורות סמיכות "שלושת" וכו')
            elif tok in HEB_UNITS:
                current += HEB_UNITS[tok]

            # לא מספרי – נתעלם
            else:
                pass

        i += 1

    flush_current()
    return total if total > 0 else None


# --- זיהוי מספר הישיבה מראש המסמך (ספרות או מילים) ---

# תבניות כותרת שכיחות: "מס' הישיבה", "מספר הישיבה", "פרוטוקול מס'", "ישיבה"
HEADER_NUM_PATTERNS = [
    # ספרות ישירות
    re.compile(
        r"(?:מס(?:'|פר)?\s*ה?ישיבה|מס(?:'|פר)?\s*ישיבה|פרוטוקול(?:\s*מס(?:'|פר)?)?)\s*[:\-]?\s*(\d+)",
        re.UNICODE
    ),
    # מילים לאחר אותן תבניות
    re.compile(
        r"(?:מס(?:'|פר)?\s*ה?ישיבה|מס(?:'|פר)?\s*ישיבה|פרוטוקול(?:\s*מס(?:'|פר)?)?)\s*[:\-]?\s*([^\n]{1,80})",
        re.UNICODE
    ),
    # גיבוי: שורה שמתחילה "הישיבה ..."
    re.compile(r"^\s*ה?ישיבה\s+([^\n]{1,80})$", re.UNICODE),
]

def _candidate_number_span(s: str) -> str:
    s = s.strip()
    # לא חותכים על מקפים או פסיקים
    for sep in (" של ", ":", ";", "("):
        if sep in s:
            s = s.split(sep)[0]
    return s.strip(" :–—-.,;\"'()[]{}")

# --- Normalization helper for noisy DOCX text ---
import re

TAG_RE = re.compile(r"<<\s*[^<>]{1,20}\s*>>")                  # << יור >>, << דובר >> וכו'
ZW_RE  = re.compile(r"[\u200e\u200f\u202a-\u202e\u2066-\u2069\ufeff]")  # סימני כיווניות ובקרה
NIKUD_RE = re.compile(r"[\u0591-\u05C7]")                      # ניקוד וטעמים

def _extract_parenthetical_chunks(name: str):
    return re.findall(r"\(([^)]+)\)", name)

def detect_party_from_name(name: str) -> tuple[str, Optional[str]]:
    """מחזירה (שם_ללא_סוגריים_מפלגתיים, מפלגה|None). מתעלמת מסוגריים שהם תפקיד/ועדה."""
    chunks = _extract_parenthetical_chunks(name)
    party = None
    for ch in chunks:
        if ch in PARTIES and not any(h in ch for h in ROLE_HINTS):
            party = ch
            # הסר רק את סוגריי המפלגה מהשם
            name = re.sub(r"\s*\("+re.escape(ch)+r"\)\s*", " ", name).strip()
            break
    return re.sub(r"\s{2,}", " ", name).strip(), party

ROLE_PREFIXES = re.compile(
    r'^(?:'
    r'היו["״\']?ר|יו["״\']?ר|יושב[\s\-]*ראש|'         # היו"ר, יו"ר, יושב ראש
    r'ח["״]?כ|חבר(?:ת)? הכנסת|'                       # ח"כ, חברת הכנסת
    r'מר|גב(?:\'|")?|'                                # מר, גב'
    r'שר(?:ת)?(?:\s+\S+){0,3}|'                       # שר / שרת + עד 3 מילים (שר החינוך והתרבות)
    r'סגן(?:ית)?\s+שר(?:\s+\S+){0,3}|'                # סגן שר / סגנית שר + עד 3 מילים
    r'ראש\s+הממשלה'                                  # ראש הממשלה
    r')\s+'
)

ROLE_SUFFIXES = re.compile(
    r'\s*-\s*(?:יו["״\']?ר.*|מ["״\']?מ.*|ממלא.*|שר.*|סגן.*)$'
)
#מקפים מופרדים ברווחים 1-10
NOISE_DASHES_RE = re.compile(r'(?:-\s+){1,10}-')
#english letters
LATIN_RE = re.compile(r'[A-Za-z]')



def is_valid_sentence(sent: str) -> bool:
    sent = sent.strip()

    # 1. מכיל אותיות באנגלית → למחוק
    if LATIN_RE.search(sent):
        return False

    # 2. מכיל דפוס "- -", "- - -", " -   -   - " → למחוק
    if NOISE_DASHES_RE.search(sent):
        return False

    return True


def normalize_speaker_name(raw: str) -> str:
    s = _normalize_line(raw)
    s = ROLE_PREFIXES.sub("", s)          # הסר תארים בתחילת השם
    s = ROLE_SUFFIXES.sub("", s)          # הסר תוספות אחרי מקף
    s = re.sub(r'\s*[\(\[\{][^)\]\}]{0,40}[\)\]\}]', '', s)  # תוספות בסוגריים
    s = s.split(',', 1)[0]                # חתוך אחרי פסיק
    s = s.strip(' "\'׳״')
    s = re.sub(r"\s{2,}", " ", s).strip()
    # השאר עד 4 טוקנים עבריים (כולל מקף פנימי/גרשיים בשם)
    parts = [p for p in s.split() if re.match(r'^[\u0590-\u05FF"\'\-]+$', p)]
    return " ".join(parts[:4]) if parts else "לא ידוע"


def _normalize_line(s: str) -> str:
    """Clean invisible/control marks and normalize punctuation before regex matching."""
    # Remove markup tags
    s = TAG_RE.sub("", s)

    # Convert NBSP (non-breaking space) to normal space
    s = s.replace("\u00A0", " ")

    # Remove directionality and control characters
    s = ZW_RE.sub("", s)

    # Remove Hebrew diacritics (nikud)
    s = NIKUD_RE.sub("", s)

    # Normalize quotation marks and dashes
    s = (s.replace("”", '"').replace("“", '"').replace("״", '"')
           .replace("’", "'").replace("׳", "'")
           .replace("\u05BE", "-").replace("–", "-").replace("—", "-"))

    # ---- NEW: remove text-break dashes like -- or --- ----
    s = re.sub(r"\s*[-]{2,}\s*", " ", s)

    # Collapse multiple spaces
    s = re.sub(r"\s+", " ", s)

    return s.strip()



def find_protocol_number(lines: List[str]) -> Optional[int]:
    """מנסה לאתר מספר ישיבה כספרות או כמילים.
    1) מעבר ראשון: עד 200 שורות ראשונות (כמו קודם)
    2) מעבר שני (fallback): סריקה על כל הטקסט המאוחד
    """
    # --- מעבר ראשון: זהה לקודם בערך ---
    for ln in lines[:200]:
        text = ln.strip()
        # ספרות
        m0 = HEADER_NUM_PATTERNS[0].search(text)
        if m0:
            try:
                return int(m0.group(1))
            except ValueError:
                pass
        # מילים אחרי תבניות
        m1 = HEADER_NUM_PATTERNS[1].search(text)
        if m1:
            seg = _candidate_number_span(m1.group(1))
            n = hebrew_words_to_int(seg)
            if n is not None:
                return n
        # גיבוי: "הישיבה ..."
        m2 = HEADER_NUM_PATTERNS[2].match(text)
        if m2:
            seg = _candidate_number_span(m2.group(1))
            n = hebrew_words_to_int(seg)
            if n is not None:
                return n

    # --- מעבר שני: סריקה על כל המסמך (למקרה שהמספר מופיע מאוחר יותר/באמצע שורה) ---
    full = "\n".join(lines)

    # קודם נסה ספרות בכל מקום (לא רק בתחילת שורה)
    m = HEADER_NUM_PATTERNS[0].search(full)
    if m:
        try:
            return int(m.group(1))
        except ValueError:
            pass

    # אחר כך מילים בכל מקום אחרי אותן תבניות
    for m in HEADER_NUM_PATTERNS[1].finditer(full):
        seg = _candidate_number_span(m.group(1))
        n = hebrew_words_to_int(seg)
        if n is not None:
            return n

    # לבסוף: גם "הישיבה ..." אם מופיע באמצע שורה
    for m in re.finditer(r"ה?ישיבה\s+([^\n]{1,80})", full):
        seg = _candidate_number_span(m.group(1))
        n = hebrew_words_to_int(seg)
        if n is not None:
            return n

    return None


# --- Regexes נוספים ---

# הסבר לביטוי הרגולרי:
# 1. (?<!\s[א-ת])  -> Negative Lookbehind: וודא שלפני הנקודה אין "רווח ואות אחת" (מונע חיתוך של ראשי תיבות בשמות, כגון "י. כהן")
# 2. (?<!\sמס)(?<!\sעמ) -> וודא שלפני הנקודה אין קיצורים נפוצים (מס. או עמ.)
# 3. [\.\!\?…]+    -> תפוס אחד או יותר סימני סיום משפט (., !, ?, ...)
# 4. (?=\s|$)      -> Lookahead: וודא שאחרי הסימן יש רווח או סוף שורה
# 5. |:\s+(?=\S)   -> או: נקודתיים שלאחריהן רווח ואז טקסט (עבור דוברים)

SENT_SPLIT = re.compile(
    r'(?:(?<!\s[א-ת])(?<!\sמס)(?<!\sעמ)[\.\!\?…]+(?=\s|$)|:\s+(?=\S))',
    re.UNICODE)
# speaker line: e.g., 'היו"ר X:', 'ח"כ Y:', 'מר Z:', or generic 'שם:'
SPEAKER_RE = re.compile(
    r'^\s*(?:היו"ר|היו״ר|יו״ר|יושב(?:\s*ראש)?|מר|גב(?:\'|")?|ח״כ|ח\"כ|חבר(?:ת)? הכנסת)\s*([^:]{1,60}):\s*$'
    r'|^\s*([^:]{1,60}):\s*$'
)

# chairman line near header
# מקפים אפשריים (רגיל / en / em)
_DASH = r"[–—-]"

# "קדימה": היו"ר ...  (כולל מ"מ, כולל סוגריים משולשים, נקודתיים אופציונליים)
CHAIR_FWD_RE = re.compile(
    r'^\s*<?\s*(?:מ["״\']?מ\s*)?(?:היו["״\']?ר|יו["״\']?ר|יושב[\s\-]*ראש(?:\s*הוועדה)?)[:\s]+([^:<>]{1,60})\s*:?\s*>?\s*$',
    re.UNICODE
)

# keep your _DASH and CHAIR_FWD_RE as-is

# 1–4 Hebrew “name” tokens, allowing quotes/hyphens inside the token
NAME_REV = r'(?:[\u0590-\u05FF"\'\-]{1,20})(?:\s+[\u0590-\u05FF"\'\-]{1,20}){0,3}'

CHAIR_REV_RE = re.compile(
    rf'^\s*<?\s*({NAME_REV})\s*{_DASH}\s*(?:מ["״\']?מ\s*)?(?:היו["״\']?ר|יו["״\']?ר)\s*:?\s*>?\s*$',
    re.UNICODE
)





def parse_filename(name: str) -> Tuple[Optional[int], Optional[str], Optional[int]]:
    """
    Try to infer:
      knesset_number  - first number in name
      protocol_type   - 'plenary' if 'ptm' or pattern 'm_', 'committee' if 'ptv' or 'v_'
      protocol_number - last number in name
    Works with names like: 13_ptm_532058.docx  /  25_ptv_1457545.docx
    """
    base = Path(name).stem
    kns = None
    ptype = None
    pnum = None

    m = re.match(r'^(?P<kns>\d+)_pt(?P<t>[mv])_(?P<pid>\d+)$', base)
    if m:
        kns = int(m.group('kns'))
        ptype = 'plenary' if m.group('t') == 'm' else 'committee'
        pnum = int(m.group('pid'))
        return kns, ptype, pnum

    # fallback: pick first & last numbers if present
    nums = [int(x) for x in re.findall(r'\d+', base)]
    if nums:
        kns = nums[0]
        if len(nums) >= 2:
            pnum = nums[-1]

    if 'ptm' in base.lower() or re.search(r'(^|_)m(_|$)', base.lower()):
        ptype = 'plenary'
    elif 'ptv' in base.lower() or re.search(r'(^|_)v(_|$)', base.lower()):
        ptype = 'committee'

    return kns, ptype, pnum


def split_sentences(text: str) -> List[str]:
    """
    מפצל טקסט למשפטים לפי הביטוי הרגולרי המשופר.
    שומר על ראשי תיבות כמו 'י. כהן' מחוברים.
    """
    # שלב 1: פיצול לפי הביטוי
    # ה-Split של פייתון מוחק את המפריד, אבל אנחנו רוצים לשמור על ההיגיון
    # לכן נשתמש בשיטה של החלפה בסימן מיוחד ואז פיצול, או פשוט split

    # פיצול פשוט (הערה: זה יעלים את הנקודה בסוף המשפט, אבל זה לרוב רצוי ב-NLP)
    # אם חשוב לך לשמור את הנקודה, הלוגיקה צריכה להיות מורכבת טיפה יותר.
    # לשימוש הנוכחי (בניית קורפוס) זה מצוין:

    parts = SENT_SPLIT.split(text)

    # ניקוי רווחים וסינון חלקים ריקים
    final_sentences = [s.strip() for s in parts if s and s.strip()]
    return final_sentences

def token_count(sentence: str) -> int:
    """Token = whitespace-split word (no morphology)."""
    return len([t for t in sentence.split() if t.strip()])


def extract_from_docx(docx_path: Path, protocol_type: Optional[str]) -> Tuple[Optional[str], Optional[int], List[Tuple[str, str, Optional[str], str]]]:
    """
    Returns:
      chairman: Optional[str]
      protocol_number: Optional[int]
      records:  List of (speaker_raw, speaker_norm, speaker_party, sentence_text)
    """
    doc = Document(str(docx_path))

    # נשמור את הפסקאות כדי שנוכל לבדוק bold, אבל גם נבנה lines כמו קודם
    paragraphs = list(doc.paragraphs)
    lines = [p.text.strip() for p in paragraphs if p.text and p.text.strip()]

    # --- protocol number from content ---
    proto_num = find_protocol_number(lines)

    # --- chairman (near top) ---
    chairman = None
    for ln in lines[:300]:
        text = _normalize_line(ln)
        m = CHAIR_FWD_RE.match(text) or CHAIR_REV_RE.match(text)
        if m:
            cand = m.group(1).strip(' "\'׳״')
            if ' ' not in cand or re.search(r'[.,!?]{1}', cand):
                continue
            chairman = cand
            break

    # --- speakers & sentences ---
    records: List[Tuple[str, str, Optional[str], str]] = []  # (raw, norm, party, sentence)
    current_raw: Optional[str] = None
    current_norm: Optional[str] = None
    current_party: Optional[str] = None
    buffer: List[str] = []

    def split_sentences(text: str) -> List[str]:
        parts = [s.strip() for s in SENT_SPLIT.split(text) if s and s.strip()]
        return parts

    def token_count(sentence: str) -> int:
        return len([t for t in sentence.split() if t.strip()])

    def flush():
        nonlocal buffer, current_raw, current_norm, current_party
        if current_norm and buffer:
            text = " ".join(buffer)
            for sent in split_sentences(text):
                if token_count(sent) >= 4 and is_valid_sentence(sent):
                    records.append((current_raw or "לא ידוע", current_norm, current_party, sent))

        buffer.clear()

    # פה מגיע השינוי: לולאה על פסקאות, עם בדיקת bold כשצריך
    for p in paragraphs:
        ln = p.text
        if not ln or not ln.strip():
            # שורה ריקה – לא סוגרת דובר, פשוט מדלגים / ממשיכים לצבור
            continue

        ln_norm = _normalize_line(ln)
        sm = SPEAKER_RE.match(ln_norm)

        if sm:
            # underline condition ONLY for committee (ptv)
            if protocol_type == "committee":
                is_underlined = any(run.underline for run in p.runs if run.text.strip())
                if not is_underlined:
                    if current_norm:
                        buffer.append(ln)
                    continue

        if sm:
            # --- תנאי חדש: אם זה ptv והפסקה BOLD – לא דובר, כנראה כותרת ('נכחו:' וכו') ---
            if protocol_type == "committee":
                is_bold = any(run.bold for run in p.runs if run.text.strip())
                if is_bold:
                    # מתעלמים מהפסקה הזו כדובר; אם כבר יש דובר פעיל – נצרף לטקסט שלו
                    if current_norm:
                        buffer.append(ln)
                    continue

            # אם הגענו לפה – זה באמת דובר חדש
            flush()
            name_raw = (sm.group(1) or sm.group(2) or "").strip(' "\'׳״') or "לא ידוע"
            name_wo_party, party = detect_party_from_name(name_raw)
            current_raw = name_raw
            current_norm = normalize_speaker_name(name_wo_party)
            current_party = party

        else:
            if current_norm:
                buffer.append(ln)

    flush()
    return chairman, proto_num, records


def main():
    if len(sys.argv) != 3:
        print("usage: python processing_knesset_corpus.py <input_dir> <output.jsonl>", file=sys.stderr)
        sys.exit(1)

    in_dir = Path(sys.argv[1])
    out_path = Path(sys.argv[2])
    out_path.parent.mkdir(parents=True, exist_ok=True)

    total_files = 0
    total_sentences = 0

    with open(out_path, "w", encoding="utf-8") as fout:
        for docx_path in sorted(in_dir.glob("*.docx")):
            total_files += 1
            protocol_name = docx_path.name

            # From filename we only take: knesset_number + protocol_type
            knesset_number, protocol_type, _ = parse_filename(protocol_name)

            try:
                protocol_chairman, proto_num_from_doc, recs = extract_from_docx(docx_path, protocol_type)

            except Exception as e:
                sys.stderr.write(f"[WARN] failed to parse {protocol_name}: {e}\n")
                continue

            # If not found inside the document, store -1 (per spec)
            protocol_number = proto_num_from_doc if proto_num_from_doc is not None else -1

            for speaker_raw, speaker_norm, speaker_party, sentence_text in recs:
                obj = {
                    "protocol_name": protocol_name,
                    "knesset_number": knesset_number,
                    "protocol_type": protocol_type,
                    "protocol_number": protocol_number,
                    "protocol_chairman": protocol_chairman,
                    "speaker_name": speaker_norm,  # מנורמל
                    "sentence_text": sentence_text
                }
                json.dump(obj, fout, ensure_ascii=False)
                fout.write("\n")
                total_sentences += 1

    print(f"[OK] processed files: {total_files}, sentences written: {total_sentences}")
    print(f"[OUT] {out_path.resolve()}")



if __name__ == "__main__":
    main()
