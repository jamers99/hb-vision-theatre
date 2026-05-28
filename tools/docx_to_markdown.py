"""
Theatre Script — DOCX to Markdown Converter
============================================
Converts Vision Community Theatre .docx scripts to Markdown.

USAGE
-----
    python3 docx_to_markdown.py

Edit the CONVERSIONS list at the bottom to point at your input/output files.

REQUIREMENTS
------------
    pip install python-docx

FORMATTING STANDARD PRODUCED
-----------------------------
Element                    Markdown
-----------------------    ------------------------------------
Document title             # Title
Act header                 # ACT ONE / TWO / THREE
Scene / part title         ## Scene Title
Song title                 ### 🎵 Song Title
LIGHTS cue                 💡 **LIGHTS:** description
MUSIC cue                  🎵 **MUSIC:** song name
Music time marker          ⏱ **TIME:** 0:42
SOUND cue                  🔊 **SOUND:** description
END OF SONG / MUSIC        *— END OF SONG —*
Ensemble label             *(Male Ensemble)*
Stage direction (parens)   *(stage direction)*
Stage action (sentence)    *Stage action here.*
Character speaker label    **CHARACTER NAME**
Inline dialogue            **CHAR** — "dialogue"
Lyrics / speech            plain text line
Intermission               ## ❖ INTERMISSION  between  --- dividers
Bible verse block          > "verse text"

HOW IT WORKS
------------
The converter reads run-level bold/plain formatting from each Word paragraph:

  - All-bold paragraph  → classified as a cue, direction, song title, or character name
  - Mixed bold+plain    → inline dialogue (bold = character name, plain = their line)
  - All-plain paragraph → lyrics, speech, section header, or narrative text

CUSTOMISATION
-------------
- KNOWN_CHARS  : add character names for your show so the classifier doesn't
                 mistake a short song title for a speaker label (or vice versa).
- start_active : set to True for files that jump straight into content without
                 an explicit "First Act / Second Act" plain-text header.
"""

import docx
import re


# ---------------------------------------------------------------------------
# CONFIGURATION — edit these for each new show
# ---------------------------------------------------------------------------

# Character names in your show.  Add every speaking role.
# This prevents short song titles (e.g. "Shalom") being misread as char names.
KNOWN_CHARS = {
    'Narrator',
    'Naomi', 'Ruth', 'Elimalek',
    'Mahlon', 'Chilion',
    'Hannah', 'Abigail', 'Elizabeth', 'Rahab',
    'Granny', 'Josh', 'Israelite Leader', 'Zechariah',
    'Lewis', 'Jonathan', 'Daniel',
    'Boaz', 'Eliza', 'Orpah', 'Moab Thug',
}

# Files to convert: list of (input_path, output_path, title, start_active)
# start_active=True  → file has no "First/Second/Third Act" plain-text header;
#                      content processing begins from the very first paragraph.
# start_active=False → file has an act header that triggers content processing.
CONVERSIONS = [
    (
        '../script/RUTH FIRST DRAFT (FIRST & SECOND ACT).docx',
        '../script/Ruth - Act I & II.md',
        'Ruth the Musical — Act I & II',
        False,
    ),
    (
        '../script/RUTH FIRST DRAFT (THIRD ACT).docx',
        '../script/Ruth - Act III.md',
        'Ruth the Musical — Act III',
        True,   # This file has no "Third Act" plain-text marker
        {
            # Maps raw docx paragraph text → desired ## heading (case-insensitive lookup)
            'RUTH ACT TWO—RETURNING HOME':  'Returning Home—PART ONE',
            'BETHLEHEMS GOSSIP':            "Bethlehem's Gossip—PART TWO",
            'RUTH & BOAZ':                  'Ruth & Boaz—PART THREE',
            "NAOMI's plan":                 "Naomi's Plan—PART FIVE",
            'ACT SIX—THE THRESHING FLOOR':  'The Threshing Floor—PART SIX',
            'ACT SEVEN—THE ACCORD':         'The Accord—PART SEVEN',
            'Rumors':                       'Rumors—PART EIGHT',
            'THE WEDDING':                  'The Wedding—PART NINE',
            'THE FINAL SCENE':              'THE END—PART TEN',
        },
    ),
]


# ---------------------------------------------------------------------------
# Paragraph helpers
# ---------------------------------------------------------------------------

def is_all_bold(para):
    runs = [r for r in para.runs if r.text.strip()]
    return bool(runs) and all(r.bold for r in runs)


def has_mixed_bold(para):
    """True when the paragraph starts with bold run(s) followed by plain run(s)."""
    runs = [r for r in para.runs if r.text.strip()]
    if not runs:
        return False
    bolds = [r.bold for r in runs]
    return bolds[0] and not all(bolds)


def strip_outer_parens(text):
    """Remove only the outermost ( … ) wrapper, preserving any inner parens."""
    t = text.strip()
    return t[1:-1] if t.startswith('(') and t.endswith(')') else t


# ---------------------------------------------------------------------------
# Classification — all-bold paragraphs
# ---------------------------------------------------------------------------

# Verbs / words that signal a stage-action sentence rather than a character name.
STAGE_VERBS = [
    'grabs', 'takes', 'walks', 'opens', 'enters', 'exits', 'leans', 'breaks',
    'rushes', 'whispers', 'motions', 'kisses', 'collapses', 'standing',
    'everyone', 'scene', 'shifts', 'switches', 'disperses', 'leaves', 'leads',
    'steps', 'moves', 'returns', 'emerges', 'dances', 'claps', 'cheers',
    'witnesses', 'tries', 'attempts', 'purposely', 'startled', 'turns', 'bows',
    'focus', 'looks', 'hugs', 'weep', 'watch', 'joins', 'pulls', 'stabs',
    'lays', 'falls', 'hits', 'dies', 'instructs', 'drops', 'picks', 'snaps',
    'crowd', 'gleaners', 'people', 'thug',
]


def starts_with_known_char(text):
    t = text.strip()
    for ch in KNOWN_CHARS:
        if t == ch:
            return True
        if t.startswith(ch + ' ') or t.startswith(ch + '&') or t.startswith(ch + '('):
            return True
        if t.startswith(ch + ','):
            return True
    return False


def is_stage_action(text):
    tl = text.lower()
    if any(tl.startswith(v) or (' ' + v) in tl for v in STAGE_VERBS):
        return True
    if text.endswith('.') and len(text) > 30:
        return True
    return False


def is_song_title_candidate(text):
    """
    Heuristics for identifying song titles among ambiguous bold paragraphs.
    Override or extend this function for a show with unusual song naming.
    """
    t = text.strip()
    if starts_with_known_char(t):
        return False
    if re.match(r'^(The |A |An )', t):          # starts with article
        return True
    if '?' in t:                                  # question-title
        return True
    if t == t.upper() and len(t) > 4 and ' ' not in t:   # all-caps single word
        return True
    words = t.split()
    if len(words) >= 2:
        # Common song-starting words — extend for your show
        song_starters = [
            'god', 'is', 'know', 'love', 'stay', 'wonder', 'learning',
            'together', 'falling', 'gossip', 'rumors', 'shalom',
            "boaz's", "rehab's",
        ]
        if words[0].lower() in song_starters:
            return True
    return False


def lookahead_has_music_or_lights(next_texts):
    """Scan upcoming paragraphs for a MUSIC or LIGHTS cue → suggests song title."""
    for nt in next_texts[:10]:
        if not nt:
            continue
        if re.match(r'^\((LIGHTS|Lights|MUSIC|Music)', nt):
            return True
        # Stop at first plain-text lyric line
        if not nt.startswith('(') and len(nt) > 0 and not nt[0].isupper():
            break
    return False


def classify_bold_para(text, next_texts):
    """
    Classify an all-bold paragraph into one of:
      lights_cue | music_cue | time_marker | sound_cue | end_marker |
      intermission | ensemble | stage_direction | stage_action |
      song_title | character_name | bible_verse
    """
    t = text.strip()
    tl = t.lower()

    if tl in ('end of song', 'end of music'):
        return 'end_marker'
    if t == '(END OF MUSIC)':
        return 'end_marker'
    if 'intermission' in tl:
        return 'intermission'

    if t.startswith('('):
        inner = strip_outer_parens(t)
        if re.match(r'LIGHTS', inner, re.IGNORECASE):
            return 'lights_cue'
        if re.match(r'MUSIC\s+TIME', inner, re.IGNORECASE):
            return 'time_marker'
        if re.match(r'Music\s+Time', inner):
            return 'time_marker'
        if re.match(r'Time\s*:', inner, re.IGNORECASE):
            return 'time_marker'
        if re.match(r'Fade out', inner, re.IGNORECASE):
            return 'time_marker'
        if re.match(r'MUSIC\s*[:\-–—]', inner, re.IGNORECASE):
            return 'music_cue'
        if re.match(r'MUSIC\s*$', inner, re.IGNORECASE):
            return 'music_cue'
        if re.match(r'Music\s*:', inner):
            return 'music_cue'
        if re.match(r'SOUND', inner, re.IGNORECASE):
            return 'sound_cue'
        if re.match(r'^\d+:\d+$', inner):
            return 'time_marker'
        if any(w in tl for w in ['ensemble', 'chorus', 'speak at', 'everyone']):
            return 'ensemble'
        return 'stage_direction'

    if is_stage_action(t):
        return 'stage_action'
    if t.startswith('"') and len(t) > 100:
        return 'bible_verse'
    if starts_with_known_char(t):
        return 'character_name'
    if is_song_title_candidate(t):
        return 'song_title'
    if lookahead_has_music_or_lights(next_texts):
        return 'song_title'
    if len(t) <= 40 and not t.endswith('.'):
        return 'character_name'
    return 'stage_action'


# ---------------------------------------------------------------------------
# Inline dialogue parser
# ---------------------------------------------------------------------------

def parse_inline_dialogue(para):
    """
    Handles lines like  **Josh** — "dialogue text"
    Returns (character_name, rest_of_line) or None.
    """
    bold_parts, plain_parts, in_bold = [], [], True
    for r in para.runs:
        if not r.text:
            continue
        if r.bold and in_bold:
            bold_parts.append(r.text)
        else:
            in_bold = False
            plain_parts.append(r.text)
    char = ''.join(bold_parts).strip().rstrip('—').rstrip('-').strip()
    rest = ''.join(plain_parts).strip().lstrip('—').lstrip('-').strip()
    return (char, rest) if char else None


# ---------------------------------------------------------------------------
# Plain-text section header detection
# ---------------------------------------------------------------------------

ACT_MAP = {
    'first act':  '# ACT ONE',
    'second act': '# ACT TWO',
    'third act':  '# ACT THREE',
}


def is_plain_section_header(text):
    t = text.strip()
    if re.search(r'PART (ONE|TWO|THREE|FOUR)', t, re.IGNORECASE):
        return True
    if re.search(r'ACT (ONE|TWO|THREE|FOUR)', t, re.IGNORECASE):
        return True
    if '—' in t and t == t.upper() and len(t) > 5:
        return True
    return False


# ---------------------------------------------------------------------------
# Markdown formatters
# ---------------------------------------------------------------------------

def fmt_lights(t):
    inner = re.sub(r'^LIGHTS\s*[:\-–—]\s*', '', strip_outer_parens(t), flags=re.IGNORECASE)
    return '💡 **LIGHTS:** ' + inner.strip()

def fmt_music(t):
    inner = strip_outer_parens(t)
    inner = re.sub(r'^MUSIC\s+TIME\s*[:\-–—]\s*', '', inner, flags=re.IGNORECASE)
    inner = re.sub(r'^MUSIC\s*[:\-–—]\s*', '', inner, flags=re.IGNORECASE)
    inner = re.sub(r'^Music\s*[:\-–—]\s*', '', inner)
    return '🎵 **MUSIC:** ' + inner.strip()

def fmt_time(t):
    inner = strip_outer_parens(t)
    inner = re.sub(r'^(MUSIC\s+)?TIME\s*[:\-–—]\s*', '', inner, flags=re.IGNORECASE)
    inner = re.sub(r'^Time\s*[:\-–—]\s*', '', inner, flags=re.IGNORECASE)
    inner = re.sub(r'^Fade out at\s*', '', inner, flags=re.IGNORECASE)
    return '⏱ **TIME:** ' + inner.strip()

def fmt_sound(t):
    inner = re.sub(r'^SOUND\s+AFFECT\s*[:\-–—]\s*', '', strip_outer_parens(t), flags=re.IGNORECASE)
    return '🔊 **SOUND:** ' + inner.strip()

def fmt_ensemble(t):  return f'*({strip_outer_parens(t).strip()})*'
def fmt_stage_dir(t): return f'*({strip_outer_parens(t).strip()})*'

def fmt_stage_action(t):
    t = t.strip()
    if not t.endswith(('.', '!', '?', '"')):
        t += '.'
    return f'*{t}*'

def fmt_char(t):   return f'**{t.strip().upper()}**'
def fmt_song(t):   return f'### 🎵 {t.strip()}'
def fmt_inline(char, rest): return f'**{char.strip().upper()}** — {rest.strip()}'


def last_content_line(lines):
    """Return the last non-empty line in the output list."""
    for ln in reversed(lines):
        if ln != '':
            return ln
    return ''


def append_cue(lines, cue_line):
    """
    Append a blockquote cue line (MUSIC, LIGHTS, TIME, SOUND).
    If the previous content line was also a blockquote, insert a bare '>'
    separator so consecutive cues render as visually distinct items.
    """
    prev = last_content_line(lines)
    if prev.startswith('>'):
        if lines[-1] != '':
            lines.append('>')
        else:
            lines[-1] = '>'   # replace trailing blank with separator
    lines.append(cue_line)


# ---------------------------------------------------------------------------
# Main converter
# ---------------------------------------------------------------------------

def convert(input_path, output_path, doc_title, start_active=False, section_map=None):
    """
    Convert a single .docx script to Markdown.

    input_path   : path to the source .docx file
    output_path  : path for the output .md file
    doc_title    : shown as the # H1 at the top of the Markdown file
    start_active : True if the file has no plain-text act header and content
                   should be processed from the first paragraph onwards
    section_map  : optional dict mapping raw docx paragraph text (case-insensitive)
                   to the desired ## heading text, overriding auto-detection and
                   reformatting for that paragraph.
    """
    section_lookup = {k.strip().lower(): v for k, v in (section_map or {}).items()}

    doc = docx.Document(input_path)
    paras = doc.paragraphs
    texts = [p.text.strip() for p in paras]
    bold_flags = [is_all_bold(p) for p in paras]

    lines = [f'# {doc_title}', '']
    subtitle_done = start_active
    prev_blank = True
    i = 0

    def ensure_blank():
        if lines and lines[-1] != '':
            lines.append('')

    def ensure_blank_after_cue():
        """Add a blank line when transitioning from a blockquote cue to plain text."""
        if last_content_line(lines).startswith('>'):
            ensure_blank()

    while i < len(paras):
        para = paras[i]
        text = texts[i]

        if not text:
            if not prev_blank:
                lines.append('')
            prev_blank = True
            i += 1
            continue

        prev_blank = False
        all_bold = bold_flags[i]
        mixed = has_mixed_bold(para)

        # section_map overrides normal classification, but only for plain-text
        # paragraphs — bold paragraphs with the same text (e.g. a song title
        # "Rumors") must still go through normal bold classification.
        if not all_bold and text.strip().lower() in section_lookup:
            ensure_blank()
            lines += [f'## {section_lookup[text.strip().lower()]}', '']
            subtitle_done = True
            i += 1
            continue

        # ── All-bold paragraph ───────────────────────────────────────────────
        if all_bold:
            act = ACT_MAP.get(text.strip().lower())
            if act:
                ensure_blank()
                lines += ['---', '', act, '']
                subtitle_done = True
                i += 1
                continue

            next_texts = [texts[j] for j in range(i + 1, min(i + 12, len(texts)))]
            kind = classify_bold_para(text, next_texts)

            if kind == 'lights_cue':
                append_cue(lines, fmt_lights(text))
            elif kind == 'music_cue':
                append_cue(lines, fmt_music(text))
            elif kind == 'time_marker':
                append_cue(lines, fmt_time(text))
            elif kind == 'sound_cue':
                append_cue(lines, fmt_sound(text))
            elif kind == 'end_marker':
                ensure_blank()
                lines.append('*— END OF SONG —*')
                lines.append('')
            elif kind == 'intermission':
                ensure_blank()
                lines += ['---', '', '# ❖ INTERMISSION', '', '---', '']
            elif kind == 'ensemble':
                ensure_blank_after_cue()
                lines.append(fmt_ensemble(text) + '  ')  # hard line break before lyrics
            elif kind == 'stage_direction':
                ensure_blank_after_cue()
                lines.append(fmt_stage_dir(text))
            elif kind == 'stage_action':
                ensure_blank_after_cue()
                lines.append(fmt_stage_action(text))
            elif kind == 'song_title':
                ensure_blank()
                lines.append(fmt_song(text))
                lines.append('')
            elif kind == 'bible_verse':
                lines.append(f'> {text}')
            else:  # character_name
                ensure_blank()
                lines.append(fmt_char(text))
                lines.append('')

        # ── Mixed bold/plain — inline dialogue ───────────────────────────────
        elif mixed:
            result = parse_inline_dialogue(para)
            if result:
                ensure_blank_after_cue()
                lines.append(fmt_inline(*result))
                lines.append('')   # blank line so consecutive exchanges don't merge
            else:
                lines.append(text)

        # ── Plain text ───────────────────────────────────────────────────────
        else:
            act = ACT_MAP.get(text.strip().lower())
            if act:
                ensure_blank()
                lines += ['---', '', act, '']
                subtitle_done = True
                i += 1
                continue

            if not subtitle_done:
                # Files without an explicit act header: first section marker
                # activates content processing.
                if is_plain_section_header(text):
                    ensure_blank()
                    lines += [f'## {text}', '']
                    subtitle_done = True
                # Everything before that (title, subtitle) is skipped silently.
                i += 1
                continue

            if is_plain_section_header(text):
                ensure_blank()
                lines += [f'## {text}', '']
                i += 1
                continue

            if text.startswith('"') and len(text) > 120:
                lines.append(f'> {text}')
                i += 1
                continue

            ensure_blank_after_cue()
            lines.append(text + '  ')  # two trailing spaces = Markdown hard line break

        i += 1

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines) + '\n')
    print(f'✓  {output_path}  ({len(lines)} lines)')


# ---------------------------------------------------------------------------
# Run
# ---------------------------------------------------------------------------

if __name__ == '__main__':
    for args in CONVERSIONS:
        input_path, output_path, title, start_active = args[:4]
        section_map = args[4] if len(args) > 4 else None
        convert(input_path, output_path, title, start_active, section_map)
