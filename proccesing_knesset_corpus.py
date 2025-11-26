#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""

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

# hebrew numbers

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

#Parties & role markers
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
    # Remove Hebrew prefix "ו"/"ה" (conjunction or definite article) from the beginning of the word.
    # Examples: "והתשעים" → "תשעים", "ושתיים"→ "שתיים"
    while tok.startswith(("ו", "ה")) and len(tok) > 1:
        tok = tok[1:]
    return tok



def hebrew_words_to_int(text: str) -> Optional[int]:
    """
    convert hebrew number that written in words to areal number
    examples: 'שלוש-מאות-ושישים-ושש', 'אלף-ומאתיים-ושלוש', 'שישים ושמונה'.
    return none if not recognize
    """
    if not text or not text.strip():
        return None

    # Normalize various dash characters to '-'
    t = (text.replace("־", "-")
              .replace("–", "-")
              .replace("—", "-"))
    # Split using both hyphens and whitespace
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

        # "אלפים / אלף / אלפיים"
        if tok in HEB_THOUSANDS:
            factor = HEB_THOUSANDS[tok]
            # if there is an accumlated number before thousnads we multiply it else assume 1
            current = max(current, 1) * factor
            flush_current()

        # "מאות / מאתיים / מאות"
        elif tok in HEB_HUNDREDS:
            factor = HEB_HUNDREDS[tok]
            current = max(current, 1) * factor

        # "עשרות (כולל אחת-עשרה וכו')"
        elif tok in HEB_TENS:
            current += HEB_TENS[tok]

        else:
            #  # Pattern like "X עשרה" after splitting — try to recombine
            if i + 1 < len(tokens) and (tokens[i+1] in ("עשרה", "עשר")):
                pair = f"{tok}-עשרה"
                if pair in HEB_TENS:
                    current += HEB_TENS[pair]
                    i += 1
                elif tok in HEB_UNITS:
                    current += 10 + HEB_UNITS[tok]
                    i += 1
                else:
                    # not avalid number
                    pass

            #units (including smikhut forms like "שלושת")
            elif tok in HEB_UNITS:
                current += HEB_UNITS[tok]

            # not avalid number- ignore
            else:
                pass

        i += 1

    flush_current()
    return total if total > 0 else None


# detect the session/protocol number from the top of the document (digits or Hebrew words)

# common header patterns e.g.: "מס' הישיבה", "מספר הישיבה", "פרוטוקול מס'", "ישיבה"
HEADER_NUM_PATTERNS = [
    # direct numeric match eg: '40'...
    re.compile(
        r"(?:מס(?:'|פר)?\s*ה?ישיבה|מס(?:'|פר)?\s*ישיבה|פרוטוקול(?:\s*מס(?:'|פר)?)?)\s*[:\-]?\s*(\d+)",
        re.UNICODE
    ),
    # numbers that written in words (hebrew)
    re.compile(
        r"(?:מס(?:'|פר)?\s*ה?ישיבה|מס(?:'|פר)?\s*ישיבה|פרוטוקול(?:\s*מס(?:'|פר)?)?)\s*[:\-]?\s*([^\n]{1,80})",
        re.UNICODE
    ),
    #  # Fallback: a line starting with "הישיבה ..." (e.g., "הישיבה הארבע עשרה")
    re.compile(r"^\s*ה?ישיבה\s+([^\n]{1,80})$", re.UNICODE),
]

def _candidate_number_span(s: str) -> str:
    s = s.strip()
    # Do not split on hyphens or commas — only split on stronger separators
    for sep in (" של ", ":", ";", "("):
        if sep in s:
            s = s.split(sep)[0]
    return s.strip(" :–—-.,;\"'()[]{}")

# Normalization helper for noisy DOCX text
import re

TOKEN_PATTERN = re.compile(
    r'\d{1,2}:\d{2}'                       # 13:20
    r'|\d{1,2}\.\d{1,2}\.\d{4}'            # 1.1.2013
    r'|\d+\.\d+'                           # 3333.3333
    r'|\d+(?:,\d{3})*(?:\.\d+)?%?'         # 23,456  40%  123.45
    r'|[\dא-ת]\.'                          # א.  1.  (סעיפים/קיצורים)
    r'|[\w״"\'א-ת]+'                       # words in hebrew with qoutation marks
    r'|[:.,!?;%–()]'                       # single punctation characters
    r'|\d+\.(?=\s*[\w״"\'א-ת]|[:.,!?;%–()])'  # Dot after a number when followed by word/punctuation
    r'|\.{3}'                              # ...
)

TAG_RE = re.compile(r"<<\s*[^<>]{1,20}\s*>>")                # Matches markup tags such as << יור >> , << דובר >> etc.
ZW_RE  = re.compile(r"[\u200e\u200f\u202a-\u202e\u2066-\u2069\ufeff]")  # Directionality and control characters (Unicode bidi markers)
NIKUD_RE = re.compile(r"[\u0591-\u05C7]")                     # Hebrew diacritics (nikud and cantillation marks)

#helper function that extracts all text that appears inside parentheses
def _extract_parenthetical_chunks(name: str):
    return re.findall(r"\(([^)]+)\)", name)

def detect_party_from_name(name: str) -> tuple[str, Optional[str]]:
    """Returns (name_without_party_parentheses, party_or_None).
        Ignores parentheses that contain roles/committee titles rather than parties.
     """
    chunks = _extract_parenthetical_chunks(name)
    party = None
    for ch in chunks:
        if ch in PARTIES and not any(h in ch for h in ROLE_HINTS):
            party = ch
            # remove only the only parentheses conatins party name
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

# Noisy dash sequences: two or more dashes (any type: -, –, —) with or without spaces
NOISE_DASHES_RE = re.compile(r'([\-–—]\s*){2,}')
ELLIPSIS_RE = re.compile(r'\.{3}|…')  # Ellipsis: either "..." or the single-character ellipsis "…"

LATIN_RE = re.compile(r'[A-Za-z]')

def is_valid_sentence(sent: str) -> bool:
    sent = sent.strip()
    if not sent:
        return False

    # delete english letters
    if LATIN_RE.search(sent):
        return False

    # Reject noisy dash sequences like "--", "- -", "–––", "– – –"
    if NOISE_DASHES_RE.search(sent):
        return False

    # Reject ellipsis sequences ("..." or "…"), if desired
    if ELLIPSIS_RE.search(sent):
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
    """
    Try to detect the session/protocol number, either as digits or as Hebrew words.
    Strategy:
      1) First pass: scan up to the first 200 lines.
      2) Second pass (fallback): scan the entire text as one string.
    """
    # first pass
    for ln in lines[:200]:
        text = ln.strip()
        # try direct numeric path
        m0 = HEADER_NUM_PATTERNS[0].search(text)
        if m0:
            try:
                return int(m0.group(1))
            except ValueError:
                pass
        # words after specifoc patterns
        m1 = HEADER_NUM_PATTERNS[1].search(text)
        if m1:
            seg = _candidate_number_span(m1.group(1))
            n = hebrew_words_to_int(seg)
            if n is not None:
                return n
        # fallback lines starting with a "ישיבה"
        m2 = HEADER_NUM_PATTERNS[2].match(text)
        if m2:
            seg = _candidate_number_span(m2.group(1))
            n = hebrew_words_to_int(seg)
            if n is not None:
                return n

    # second pass on the whole file if we didnt find it in the first 200 lines
    full = "\n".join(lines)

    # Try direct numeric match (digits)
    m = HEADER_NUM_PATTERNS[0].search(full)
    if m:
        try:
            return int(m.group(1))
        except ValueError:
            pass

    # Try Hebrew words after the header patterns
    for m in HEADER_NUM_PATTERNS[1].finditer(full):
        seg = _candidate_number_span(m.group(1))
        n = hebrew_words_to_int(seg)
        if n is not None:
            return n

    # Fallback: lines starting with "הישיבה ..."
    for m in re.finditer(r"ה?ישיבה\s+([^\n]{1,80})", full):
        seg = _candidate_number_span(m.group(1))
        n = hebrew_words_to_int(seg)
        if n is not None:
            return n

    return None


# another Regexes

# Explanation of the sentence-splitting regex:
# 1. (?<!\s[א-ת])
#       Negative lookbehind: ensure the period is NOT preceded by "space + Hebrew letter".
#       Prevents splitting inside initials like "י. כהן".
#
# 2. (?<!\sמס)(?<!\sעמ)
#       Negative lookbehind: ensure the period is not part of common abbreviations
#       such as "מס." or "עמ.".
#
# 3. [\.\!\?…]+
#       Match one or more sentence-ending characters: ".", "!", "?", "…"
#
# 4. (?=\s|$)
#       Lookahead: ensure the character is followed by whitespace or end-of-line.
#
# 5. |:\s+(?=\S)
#       OR: match a colon followed by whitespace and then a non-space —
#       used to detect speaker lines such as "מר כהן: ..."

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

 # Forward-style chairman declaration near the document header.
 # Examples:
# #   "היו"ר משה כהן"
# #   "מ"מ היו"ר דני דנון"
# #   "< היו"ר עליזה לביא >"
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

    Split Hebrew text into sentences with safeguards.
    Protect against false splits in:
      - dates (e.g., 1.1.2013)
      - times (e.g., 13:20.)
      - section markers such as 'א.' / 'ב.' / '1.'

    """
    points = ['א', 'ב', 'ג', 'ד', 'ה', 'ן']
    split_set = ['.', '?', '!']
    sentences: List[str] = []
    current = ""

    text = text.strip()

    for i, ch in enumerate(text):
        current += ch

        if ch in split_set:


            # 1. at the beginning of a number/date → do NOT split
            if i == 1 and len(text) > 1 and text[i-1].isdigit() and ch == '.':
                continue

            # 2."א." / "ב." at the start of a line → section markers, do NOT split
            if i == 1 and len(text) > 1 and text[i-1] in points and ch == '.':
                continue

            # 3) Time format like "13:20." → do NOT split
            if len(text) > i > 3 and ch == '.' and text[i-1].isdigit() and text[i-3] == ':':
                continue

            # 4) "א." after ":" or "." or ";" → part of section numbering
            if len(text) > i > 3 and ch == '.' and text[i-1] in points and text[i-3] in [':', '.', ';']:
                continue

            # 5) Decimal number like "1.2" (digit-dot-digit) → do NOT split
            if len(text) - 1 > i > 0 and ch == '.' and text[i-1].isdigit() and text[i+1].isdigit():
                continue

            # if we rech here this dot is for real sentence boundary
            if ch == '.':
                # Avoid splitting on number-dot-number, although already checked above
                if len(text) - 1 > i > 0 and (not text[i-1].isdigit() or not text[i+1].isdigit()):
                    sent = current.strip()
                    if sent:
                        sentences.append(sent)
                    current = ""
            else:
                # ? או ! → סוף משפט
                sent = current.strip()
                if sent:
                    sentences.append(sent)
                current = ""

    if current.strip():
        sentences.append(current.strip())

    return sentences

def token_count(sentence: str) -> int:
    """Count the number of tokens produced by tokenize(), not just whitespace splits."""
    tokens = tokenize(sentence)
    return len([t for t in tokens if t.strip()])



def tokenize(sentence: str) -> list[str]:
    """
      Return a list of tokens (words, numbers, dates, punctuation marks)
      according to the TOKEN_PATTERN regex.
      """
    return TOKEN_PATTERN.findall(sentence)


def extract_from_docx(docx_path: Path, protocol_type: Optional[str]) -> Tuple[Optional[str], Optional[int], List[Tuple[str, str, Optional[str], str]]]:
    """
    Returns:
      chairman: Optional[str]
      protocol_number: Optional[int]
      records:  List of (speaker_raw, speaker_norm, speaker_party, sentence_text)
    """
    doc = Document(str(docx_path))

    # Keep the full paragraph objects (to inspect bold/underline),
    # but also build a simple list of non-empty text lines.
    paragraphs = list(doc.paragraphs)
    lines = [p.text.strip() for p in paragraphs if p.text and p.text.strip()]

    #  protocol number from content
    proto_num = find_protocol_number(lines)

    # chairman (near top)
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

    # speakers and sentences
    records: List[Tuple[str, str, Optional[str], str]] = []  # (raw, norm, party, sentence)
    current_raw: Optional[str] = None
    current_norm: Optional[str] = None
    current_party: Optional[str] = None
    buffer: List[str] = []

    def flush():
        nonlocal buffer, current_raw, current_norm, current_party
        if current_norm and buffer:
            text = " ".join(buffer)
            for sent in split_sentences(text):
                if token_count(sent) >= 4 and is_valid_sentence(sent):
                    records.append((current_raw or "לא ידוע", current_norm, current_party, sent))

        buffer.clear()

    # Iterate over paragraphs so we can inspect bold/underline when needed
    for p in paragraphs:
        ln = p.text
        if not ln or not ln.strip():
            #empty line: dont close the current speaker, just skip
            continue

        ln_norm = _normalize_line(ln)
        sm = SPEAKER_RE.match(ln_norm)

        if sm:
            if protocol_type == "committee":
                # decide whether we require underline based on knesset number
                kns, _, _ = parse_filename(docx_path.name)
                require_underline = (kns is None) or (kns < 23)

                if require_underline:
                    is_underlined = any(run.underline for run in p.runs if run.text.strip())
                    if not is_underlined:
                        if current_norm:
                            buffer.append(ln)
                        continue
                # if kns>=23 we dont require underline SPEAKER_RE match enough



        if sm:
            # for committee protocols, bold paragraphs like "נכחו:" are headings,
            #             # not actual speaker lines.
            if protocol_type == "committee":
                is_bold = any(run.bold for run in p.runs if run.text.strip())
                if is_bold:
                    # Ignore this as a new speaker; if we already have one, append to their text
                    if current_norm:
                        buffer.append(ln)
                    continue

            # if we reach here its really anew speaker
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





if __name__ == "__main__":
    main()
