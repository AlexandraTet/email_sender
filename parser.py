"""Parse emails.docx into data/queue.csv (HTML body with the Word formatting + plain-text fallback).

Expected layout of emails.docx (what the current file uses):

    Table of contents:  "1. Sigma Software Labs  [ENG]", "2. Intellias  [ENG]", ...
    Then, per letter:
        letterhead            "NATIONAL TECHNICAL UNIVERSITY OF UKRAINE" (EN) or
                              "НАЦІОНАЛЬНИЙ ТЕХНІЧНИЙ УНІВЕРСИТЕТ УКРАЇНИ" (UA), then address lines
        date line             "«____» ____________ 2026"
        addressee (+ phone)   "To the Management of Sigma Software Labs"
        e-mail line           "E-mail: info@sigma.software"     -> recipient_email
        subject line          "Request for sponsorship ..."     -> subject
        salutation ... end    "Dear Sigma Software Labs Team," ... contacts

The whole letter, from the letterhead to the contacts, becomes the e-mail body, laid out as in
Word. The blank date "«____» ____________ 2026" becomes {{date}}, filled in with the sending date.
It is converted to e-mail-safe HTML with inline styles only: bold/italic/underline/
strike/super-/subscript, colours, highlights, hyperlinks, paragraph alignment, indents,
spacing, borders, line breaks, bulleted/numbered lists and simple tables. The plain-text
`body` is derived from that HTML, so both always say the same thing.

An explicit attachment can be named anywhere in a letter with a line such as
"Attachment: brochure_en.pdf" (or "Вкладення:", "Додаток:", "Файл:"). Otherwise the
attachment defaults to config.DEFAULT_ATTACHMENTS[lang].

Run standalone:  python parser.py            (re-import, keep 'sent' history)
                 python parser.py --fresh    (re-import, everything back to pending)
"""

from __future__ import annotations

import argparse
import html
import itertools
import re
import shutil
from collections import Counter, defaultdict, deque
from dataclasses import dataclass
from pathlib import Path

import docx
import pandas as pd
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn

import config
from config import (
    ATTACHMENTS_DIR,
    DEFAULT_ATTACHMENTS,
    QUEUE_BACKUP_CSV,
    QUEUE_COLUMNS,
    STATUS_ERROR,
    STATUS_PENDING,
    STATUS_SENT,
)
from sender import DATE_TOKEN, RECIPIENT_TOKEN, html_to_text, linkify, queue_store

LETTERHEADS = {
    "NATIONAL TECHNICAL UNIVERSITY OF UKRAINE": "en",
    "НАЦІОНАЛЬНИЙ ТЕХНІЧНИЙ УНІВЕРСИТЕТ УКРАЇНИ": "ua",
}
# "E-mail:", "Email:", "e-mail :" - also with a Cyrillic "Е".
EMAIL_LINE_RE = re.compile(r"^\s*[EЕeе]\s*[-‐‑–]?\s*mail\s*:\s*(.*)$", re.IGNORECASE)
ADDRESS_RE = re.compile(r"[A-Za-z0-9._%+'-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+")
ATTACHMENT_LINE_RE = re.compile(r"^\s*(?:attachments?|attached|вкладення|додаток|додатки|файл)\s*:\s*(.+)$", re.IGNORECASE)
TOC_RE = re.compile(r"^\s*(\d+)\.\s+(.+?)\s*(?:\[(ENG|EN|UKR|UA)\])?\s*$", re.IGNORECASE)
DATE_LINE_RE = re.compile(r"(?:«.*»|\d{1,2}\.\d{1,2}\.)\s*.*20\d\d")
# The unfilled date "«____» ____________ 2026 р." -> DATE_TOKEN, replaced by the sending date.
BLANK_DATE_RE = re.compile(r"«\s*_+\s*»\s*_+\s*20\d\d(?:\s*р\.)?")
# The address(es) or blank in the recipient block's "E-mail:" paragraph (as rendered HTML).
RECIPIENT_VALUE_RE = re.compile(r'<a href="mailto:[^"]*">[^<]*</a>(?:\s*[,;]\s*<a href="mailto:[^"]*">[^<]*</a>)*|_{3,}')
PHONE_LINE_RE = re.compile(r"^\s*(?:тел|tel|phone)", re.IGNORECASE)
CYRILLIC_RE = re.compile(r"[а-яіїєґ]", re.IGNORECASE)

MISSING_EMAIL_MESSAGE = "No e-mail address in emails.docx: fill in recipient_email and set status to pending"
ADDRESSEE_PREFIXES = re.compile(
    r"^(to the management of|to the|керівництву компанії|керівництву|директору)\s+", re.IGNORECASE
)


# --------------------------------------------------------------------------- DOCX -> HTML

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
P_TAG, TBL_TAG = W + "p", W + "tbl"
SAFE_URL_RE = re.compile(r"^(https?:|mailto:|tel:)", re.IGNORECASE)
FIELD_URL_RE = re.compile(r'HYPERLINK\s+"([^"]+)"', re.IGNORECASE)
# Containers whose runs are shown as normal text (tracked insertions, content controls, ...).
TRANSPARENT = {W + t for t in ("ins", "moveTo", "smartTag", "customXml", "sdt", "sdtContent")}
ALIGNMENTS = {"center": "center", "right": "right", "end": "right", "both": "justify", "distribute": "justify"}
ORDERED_TYPES = {"lowerLetter": "a", "upperLetter": "A", "lowerRoman": "i", "upperRoman": "I"}
HIGHLIGHTS = {
    "yellow": "#ffff00", "green": "#00ff00", "cyan": "#00ffff", "magenta": "#ff00ff", "blue": "#0000ff",
    "red": "#ff0000", "darkBlue": "#000080", "darkCyan": "#008080", "darkGreen": "#008000",
    "darkMagenta": "#800080", "darkRed": "#800000", "darkYellow": "#808000", "darkGray": "#808080",
    "lightGray": "#c0c0c0", "black": "#000000", "white": "#ffffff",
}
SERIF_FONTS = {"times new roman", "times", "georgia", "cambria", "garamond", "book antiqua", "palatino linotype"}
# CSS line-height for Word's "single" spacing. 1.5 matches the letters as they were actually sent
# and read in Gmail (measured against a screenshot: within ±3 px over the letterhead and body).
SINGLE_LINE_HEIGHT = 1.5


def _on(el) -> bool:
    """Word on/off property: <w:b/> and <w:b w:val="1"/> are on, <w:b w:val="0"/> is off."""
    return el is not None and (el.get(W + "val") or "true").lower() not in ("0", "false", "off", "none")


def _props(el) -> dict:
    """Index a w:pPr / w:rPr element by child tag (plain dict lookups are much faster than find())."""
    return {child.tag: child for child in el} if el is not None else {}


def _find(chain: list[dict], tag: str):
    """First `tag` element in a property chain (direct formatting first, defaults last)."""
    for props in chain:
        el = props.get(W + tag)
        if el is not None:
            return el
    return None


def _attr(chain: list[dict], tag: str, *names: str) -> str | None:
    """Attribute-wise lookup, so e.g. spacing 'before' from a style and 'after' from the paragraph combine."""
    for props in chain:
        el = props.get(W + tag)
        if el is not None:
            for name in names:
                value = el.get(W + name)
                if value is not None:
                    return value
    return None


def _twips(value: str | None) -> int:
    try:
        return int(float(value)) if value else 0
    except ValueError:
        return 0


def _css_len(twips: int) -> str:
    px = round(twips / 15)  # 1 px = 0.75 pt = 15 twips
    return f"{px}px" if px else "0"


def _pt_to_px(points: float) -> str:
    return f"{round(points * 4 / 3, 1):g}"


def _font_css(name: str) -> str:
    fallbacks = ["Times", "serif"] if name.lower() in SERIF_FONTS else ["Arial", "Helvetica", "sans-serif"]
    quoted = f"'{name}'" if " " in name else name
    return ",".join([quoted] + [f for f in fallbacks if f.lower() != name.lower()])


@dataclass(frozen=True)
class _RunFormat:
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strike: bool = False
    vert: str = ""  # "sup" or "sub"
    caps: bool = False
    small_caps: bool = False
    color: str = ""
    background: str = ""
    size: float = 0.0  # points
    font: str = ""
    hidden: bool = False

    @property
    def decorated(self) -> bool:
        """Formatting that is visible even on whitespace."""
        return self.underline or self.strike or bool(self.background)


@dataclass(frozen=True)
class _Base:
    """The letter's dominant font/size/line height; only deviations are written per element."""

    font: str
    size: float
    line_height: str


class _Styles:
    def __init__(self, document):
        root = document.styles.element
        self.by_id = {s.get(W + "styleId"): s for s in root.iter(W + "style")}
        self.default_paragraph = next(
            (sid for sid, s in self.by_id.items() if s.get(W + "type") == "paragraph" and _on_value(s.get(W + "default"))),
            None,
        )
        defaults = root.find(W + "docDefaults")
        rpr_default = defaults.find(f"{W}rPrDefault/{W}rPr") if defaults is not None else None
        ppr_default = defaults.find(f"{W}pPrDefault/{W}pPr") if defaults is not None else None
        self.rpr_default = [_props(rpr_default)] if rpr_default is not None else []
        self.ppr_default = [_props(ppr_default)] if ppr_default is not None else []
        self._chains: dict[tuple, list[dict]] = {}

    def _chain(self, style_id: str | None, tag: str) -> list[dict]:
        """Property dicts of a style and the styles it is based on."""
        key = (style_id, tag)
        if key not in self._chains:
            chain, seen = [], set()
            while style_id and style_id in self.by_id and style_id not in seen:
                seen.add(style_id)
                style = self.by_id[style_id]
                props = style.find(W + tag)
                if props is not None:
                    chain.append(_props(props))
                based_on = style.find(W + "basedOn")
                style_id = based_on.get(W + "val") if based_on is not None else None
            self._chains[key] = chain
        return self._chains[key]

    def _paragraph_style(self, own: dict) -> str | None:
        ref = own.get(W + "pStyle")
        style_id = ref.get(W + "val") if ref is not None else None
        return style_id if style_id in self.by_id else self.default_paragraph

    def paragraph_chain(self, p) -> list[dict]:
        own = _props(p.find(W + "pPr"))
        return [own, *self._chain(self._paragraph_style(own), "pPr"), *self.ppr_default]

    def run_chain(self, r, p) -> list[dict]:
        own = _props(r.find(W + "rPr"))
        char_style = own.get(W + "rStyle")
        chain = [own]
        if char_style is not None:
            chain += self._chain(char_style.get(W + "val"), "rPr")
        paragraph_style = self._paragraph_style(_props(p.find(W + "pPr")))
        return chain + self._chain(paragraph_style, "rPr") + self.rpr_default


def _on_value(value: str | None) -> bool:
    return (value or "").lower() in ("1", "true", "on")


class _Numbering:
    """Resolves w:numPr to list kind/number using the document's numbering definitions."""

    def __init__(self, document):
        try:
            root = document.part.part_related_by(RT.NUMBERING).element
        except KeyError:
            root = None
        self.nums = {n.get(W + "numId"): n for n in root.iter(W + "num")} if root is not None else {}
        self.abstract = {a.get(W + "abstractNumId"): a for a in root.iter(W + "abstractNum")} if root is not None else {}
        self.counters: dict[tuple[str, int], int] = {}

    def _level(self, num_id: str, ilvl: int) -> tuple[str, int, int]:
        """(numFmt, start, left indent in twips) for one list level."""
        num = self.nums[num_id]
        ref = num.find(W + "abstractNumId")
        abstract = self.abstract.get(ref.get(W + "val")) if ref is not None else None
        level = start_override = None
        for override in num.iter(W + "lvlOverride"):
            if override.get(W + "ilvl") == str(ilvl):
                start = override.find(W + "startOverride")
                start_override = int(start.get(W + "val")) if start is not None else None
                level = override.find(W + "lvl")
        if level is None and abstract is not None:
            level = next((lvl for lvl in abstract.iter(W + "lvl") if lvl.get(W + "ilvl") == str(ilvl)), None)
        if level is None:
            return "bullet", 1, 720 * (ilvl + 1)
        fmt = level.find(W + "numFmt")
        start = level.find(W + "start")
        ind = level.find(f"{W}pPr/{W}ind")
        left = _twips((ind.get(W + "left") or ind.get(W + "start")) if ind is not None else None)
        return (
            fmt.get(W + "val") if fmt is not None else "decimal",
            start_override or (int(start.get(W + "val")) if start is not None else 1),
            left or 720 * (ilvl + 1),
        )

    def item(self, chain: list) -> dict | None:
        """List info for a paragraph, or None if it is not a list item. Advances the counters."""
        num_pr = _find(chain, "numPr")
        if num_pr is None:
            return None
        num_el, lvl_el = num_pr.find(W + "numId"), num_pr.find(W + "ilvl")
        num_id = num_el.get(W + "val") if num_el is not None else None
        if not num_id or num_id == "0" or num_id not in self.nums:
            return None
        ilvl = int(lvl_el.get(W + "val")) if lvl_el is not None else 0
        fmt, start, left = self._level(num_id, ilvl)
        if fmt == "none":
            return None
        number = self.counters.get((num_id, ilvl), start - 1) + 1
        self.counters[(num_id, ilvl)] = number
        for key in [k for k in self.counters if k[0] == num_id and k[1] > ilvl]:
            del self.counters[key]  # a higher-level item restarts its sub-levels
        parent_left = self._level(num_id, ilvl - 1)[2] if ilvl else 0
        ordered = fmt != "bullet"
        return {
            "key": (num_id, ordered),
            "ilvl": ilvl,
            "tag": "ol" if ordered else "ul",
            "type": ORDERED_TYPES.get(fmt, ""),
            "number": number,
            "indent": max(18, round((left - parent_left) / 15)),
        }


class DocxHtmlConverter:
    """Converts runs of body elements (w:p / w:tbl) into lightweight, e-mail-safe HTML."""

    def __init__(self, document):
        self.part = document.part
        self.styles = _Styles(document)
        self.numbering = _Numbering(document)
        self._chains: dict = {}  # per-letter memo: paragraph -> property chain
        self._formats: dict = {}  # per-letter memo: run -> _RunFormat

    # ---- public

    def letter(self, blocks: list) -> str:
        self._chains.clear()
        self._formats.clear()
        blocks = _trim_empty(blocks)
        base = self._base(blocks)
        style = f"font-family:{_font_css(base.font)};font-size:{_pt_to_px(base.size)}px;line-height:{base.line_height};"
        return f'<div style="{style}">\n{self._blocks(blocks, base)}\n</div>'

    # ---- blocks

    def _blocks(self, blocks: list, base: _Base) -> str:
        out: list[str] = []
        lists: list[dict] = []  # open lists, innermost last; each has an open <li>

        def close_lists(to_level: int = -1) -> None:
            while lists and lists[-1]["ilvl"] > to_level:
                out.append(f"</li></{lists.pop()['tag']}>")

        for el in blocks:
            if el.tag == TBL_TAG:
                close_lists()
                out.append(self._table(el, base))
                continue
            chain = self._paragraph_chain(el)
            item = self.numbering.item(chain)
            inner = self._inline(el, base)
            if item is None:
                close_lists()
                out.append(f'<p style="{self._paragraph_css(chain, base)}">{inner or "&nbsp;"}</p>')
                continue

            close_lists(item["ilvl"])
            if lists and lists[-1]["ilvl"] == item["ilvl"] and lists[-1]["key"] != item["key"]:
                close_lists(item["ilvl"] - 1)  # same level, different list: start a new one
            if lists and lists[-1]["ilvl"] == item["ilvl"]:
                out.append("</li>")
            else:
                attrs = f' type="{item["type"]}"' if item["type"] else ""
                attrs += f' start="{item["number"]}"' if item["tag"] == "ol" and item["number"] != 1 else ""
                out.append(f'<{item["tag"]}{attrs} style="margin:0;padding-left:{item["indent"]}px;">')
                lists.append(item)
            out.append(f'<li style="{self._paragraph_css(chain, base, in_list=True)}">{inner}')
        close_lists()
        return "\n".join(out)

    def _table(self, tbl, base: _Base) -> str:
        tbl_pr = tbl.find(W + "tblPr")
        borders = tbl_pr.find(W + "tblBorders") if tbl_pr is not None else None
        bordered = (
            any((b.get(W + "val") or "none") not in ("none", "nil") for b in borders)
            if borders is not None
            else tbl_pr is not None and tbl_pr.find(W + "tblStyle") is not None
        )
        cell_css = "vertical-align:top;" + ("border:1px solid #999999;" if bordered else "")
        rows = []
        for tr in tbl.iterchildren(W + "tr"):
            cells = []
            for tc in tr.iterchildren(W + "tc"):
                tc_pr = tc.find(W + "tcPr")
                span = tc_pr.find(W + "gridSpan") if tc_pr is not None else None
                shade = tc_pr.find(W + "shd") if tc_pr is not None else None
                fill = shade.get(W + "fill") if shade is not None else None
                css = cell_css + (f"background-color:#{fill};" if fill and fill.lower() not in ("auto", "ffffff") else "")
                colspan = f' colspan="{span.get(W + "val")}"' if span is not None else ""
                content = self._blocks(_trim_empty(list(tc.iterchildren(P_TAG, TBL_TAG))), base)
                cells.append(f'<td{colspan} style="{css}">{content}</td>')
            rows.append("<tr>" + "".join(cells) + "</tr>")
        return (
            '<table cellpadding="6" cellspacing="0" border="0" style="border-collapse:collapse;margin:0 0 12px 0;">'
            + "".join(rows)
            + "</table>"
        )

    # ---- paragraphs

    def _paragraph_css(self, chain: list, base: _Base, in_list: bool = False) -> str:
        before, after = _twips(_attr(chain, "spacing", "before")), _twips(_attr(chain, "spacing", "after"))
        if _on_value(_attr(chain, "spacing", "beforeAutospacing")):
            before = 280  # Word's automatic spacing is 14 pt
        if _on_value(_attr(chain, "spacing", "afterAutospacing")):
            after = 280
        left = 0 if in_list else _twips(_attr(chain, "ind", "left", "start"))
        right = _twips(_attr(chain, "ind", "right", "end"))
        css = [f"margin:{_css_len(before)} {_css_len(right)} {_css_len(after)} {_css_len(left)}"]

        align = ALIGNMENTS.get(_attr(chain, "jc", "val") or "")
        if align:
            css.append(f"text-align:{align}")
        if not in_list:
            hanging = _twips(_attr(chain, "ind", "hanging"))
            indent = -hanging if hanging else _twips(_attr(chain, "ind", "firstLine"))
            if indent:
                css.append(f"text-indent:{_css_len(indent)}")
        line_height = self._line_height(chain)
        if line_height != base.line_height:
            css.append(f"line-height:{line_height}")

        borders = _find(chain, "pBdr")
        for side in ("top", "right", "bottom", "left"):
            border = borders.find(W + side) if borders is not None else None
            if border is not None and (border.get(W + "val") or "none") not in ("none", "nil"):
                width = max(1, round(_twips(border.get(W + "sz") or "4") / 6))  # eighths of a point -> px
                color = border.get(W + "color") or "auto"
                css.append(f"border-{side}:{width}px solid #{'000000' if color == 'auto' else color}")
                padding = round(_twips(border.get(W + "space")) * 4 / 3)
                if padding:
                    css.append(f"padding-{side}:{padding}px")
        fill = _attr(chain, "shd", "fill")
        if fill and fill.lower() not in ("auto", "ffffff"):
            css.append(f"background-color:#{fill}")
        return ";".join(css) + ";"

    @staticmethod
    def _line_height(chain: list) -> str:
        line = _twips(_attr(chain, "spacing", "line"))
        if not line:
            return f"{SINGLE_LINE_HEIGHT:g}"
        if (_attr(chain, "spacing", "lineRule") or "auto") == "auto":
            return f"{round(SINGLE_LINE_HEIGHT * line / 240, 2):g}"
        return f"{round(line / 15)}px"  # exact / at-least, in twips

    def _base(self, blocks: list) -> _Base:
        fonts, sizes, heights = Counter(), Counter(), Counter()
        for block in blocks:
            for p in block.iter(P_TAG):
                heights[self._line_height(self._paragraph_chain(p))] += 1
                for r in p.iter(W + "r"):
                    weight = sum(len(t.text or "") for t in r.iter(W + "t"))
                    if weight:
                        fmt = self._format(r, p)
                        fonts[fmt.font] += weight
                        sizes[fmt.size] += weight
        font = next((f for f, _ in fonts.most_common() if f), "Arial")
        size = next((s for s, _ in sizes.most_common() if s), 11.0)
        height = heights.most_common(1)[0][0] if heights else f"{SINGLE_LINE_HEIGHT:g}"
        return _Base(font=font, size=size, line_height=height)

    # ---- runs

    def _paragraph_chain(self, p) -> list[dict]:
        if p not in self._chains:
            self._chains[p] = self.styles.paragraph_chain(p)
        return self._chains[p]

    def _format(self, r, p) -> _RunFormat:
        if r not in self._formats:
            self._formats[r] = self._resolve_format(r, p)
        return self._formats[r]

    def _resolve_format(self, r, p) -> _RunFormat:
        chain = self.styles.run_chain(r, p)
        underline = _find(chain, "u")
        color = (_attr(chain, "color", "val") or "").lower()
        highlight = _attr(chain, "highlight", "val") or ""
        shade = (_attr(chain, "shd", "fill") or "").lower()
        size = _attr(chain, "sz", "val")
        return _RunFormat(
            bold=_on(_find(chain, "b")),
            italic=_on(_find(chain, "i")),
            underline=underline is not None and (underline.get(W + "val") or "single") != "none",
            strike=_on(_find(chain, "strike")) or _on(_find(chain, "dstrike")),
            vert={"superscript": "sup", "subscript": "sub"}.get(_attr(chain, "vertAlign", "val") or "", ""),
            caps=_on(_find(chain, "caps")),
            small_caps=_on(_find(chain, "smallCaps")),
            color="" if color in ("", "auto", "000000") else color,
            background=HIGHLIGHTS.get(highlight, "") or ("" if shade in ("", "auto", "ffffff") else f"#{shade}"),
            size=int(size) / 2 if size and size.isdigit() else 0.0,
            font=_attr(chain, "rFonts", "ascii", "hAnsi", "cs") or "",
            hidden=_on(_find(chain, "vanish")),
        )

    def _inline(self, p, base: _Base) -> str:
        segments: list[tuple] = []  # (kind, text, format, href)
        self._collect(p, p, None, segments, [])

        merged: list[list] = []
        for kind, text, fmt, href in segments:
            prev = merged[-1] if merged else None
            if prev and kind == "text" == prev[0] and prev[3] == href and (
                prev[2] == fmt or (not text.strip() and not fmt.decorated)
            ):
                prev[1] += text  # same formatting (or plain whitespace): one element instead of two
            else:
                merged.append([kind, text, fmt, href])

        parts, line_start = [], True
        for href, group in itertools.groupby(merged, key=lambda seg: seg[3]):
            inner = []
            for kind, text, fmt, _ in group:
                if kind == "br":
                    inner.append("<br>")
                    line_start = True
                    continue
                inner.append(self._text_html(text, fmt, base, autolink=href is None, line_start=line_start))
                line_start = False
            joined = "".join(inner)
            parts.append(f'<a href="{html.escape(href)}">{joined}</a>' if href else joined)
        return "".join(parts).strip()

    def _collect(self, el, p, href: str | None, out: list, fields: list) -> None:
        for child in el:
            tag = child.tag
            if tag == W + "r":
                self._run(child, p, href, out, fields)
            elif tag == W + "hyperlink":
                self._collect(child, p, self._hyperlink_url(child) or href, out, fields)
            elif tag == W + "fldSimple":
                self._collect(child, p, _field_url(child.get(W + "instr")) or href, out, fields)
            elif tag in TRANSPARENT:
                self._collect(child, p, href, out, fields)
            # w:del, w:moveFrom, bookmarks, proofing marks, pPr: not visible text

    def _run(self, r, p, href: str | None, out: list, fields: list) -> None:
        fmt = self._format(r, p)
        for child in r:
            tag = child.tag
            if tag == W + "fldChar":  # complex fields, e.g. HYPERLINK inserted as a field
                kind = child.get(W + "fldCharType")
                if kind == "begin":
                    fields.append({"instr": "", "url": None, "result": False})
                elif kind == "separate" and fields:
                    fields[-1].update(result=True, url=_field_url(fields[-1]["instr"]))
                elif kind == "end" and fields:
                    fields.pop()
                continue
            if tag == W + "instrText":
                if fields:
                    fields[-1]["instr"] += child.text or ""
                continue
            if fmt.hidden or (fields and not fields[-1]["result"]):
                continue
            link = href or next((f["url"] for f in reversed(fields) if f["url"]), None)
            if tag == W + "t":
                out.append(("text", child.text or "", fmt, link))
            elif tag == W + "tab":
                out.append(("text", "\t", fmt, link))
            elif tag in (W + "br", W + "cr"):
                if child.get(W + "type") not in ("page", "column"):
                    out.append(("br", "", fmt, link))
            elif tag == W + "noBreakHyphen":
                out.append(("text", "‑", fmt, link))
            elif tag == W + "sym":
                code = int(child.get(W + "char") or "0", 16)
                out.append(("text", "•" if code in (0xF0B7, 0xF06C, 0xF0A7) else chr(code % 0xF000 or 0x20), fmt, link))

    def _hyperlink_url(self, link) -> str | None:
        rel_id = link.get(qn("r:id"))
        rel = self.part.rels.get(rel_id) if rel_id else None
        if rel is None or not rel.is_external:
            return None  # links to bookmarks inside the .docx mean nothing in an e-mail
        url = rel.target_ref + (f"#{link.get(W + 'anchor')}" if link.get(W + "anchor") else "")
        return url if SAFE_URL_RE.match(url) else None

    @staticmethod
    def _text_html(text: str, fmt: _RunFormat, base: _Base, autolink: bool, line_start: bool) -> str:
        s = html.escape(text, quote=False)
        if autolink:
            s = linkify(s)
        s = s.replace("\t", "&emsp;")
        s = re.sub(r" {2,}", lambda m: " " + "&nbsp;" * (len(m.group()) - 1), s)  # keep Word's extra spaces
        if line_start and s.startswith(" "):
            s = "&nbsp;" + s[1:]
        if not text.strip() and not fmt.decorated:
            return s

        css = []
        if fmt.color:
            css.append(f"color:#{fmt.color}")
        if fmt.background:
            css.append(f"background-color:{fmt.background}")
        if fmt.size and abs(fmt.size - base.size) > 0.01:
            css.append(f"font-size:{_pt_to_px(fmt.size)}px")
        if fmt.font and fmt.font != base.font:
            css.append(f"font-family:{_font_css(fmt.font)}")
        if fmt.caps:
            css.append("text-transform:uppercase")
        if fmt.small_caps:
            css.append("font-variant:small-caps")
        if css:
            s = f'<span style="{";".join(css)};">{s}</span>'
        for on, tag in ((fmt.vert, fmt.vert), (fmt.strike, "s"), (fmt.underline, "u"), (fmt.italic, "em"), (fmt.bold, "strong")):
            if on:
                s = f"<{tag}>{s}</{tag}>"
        return s


def _field_url(instruction: str | None) -> str | None:
    match = FIELD_URL_RE.search(instruction or "")
    return match.group(1) if match and SAFE_URL_RE.match(match.group(1)) else None


def _recipient_placeholder(body_html: str) -> str:
    """Swap the address in the recipient block ("E-mail: info@x.com" / "E-mail: ____") for
    RECIPIENT_TOKEN, so the letter always shows the queue's current recipient_email.
    The converter writes one paragraph per line; the first "E-mail:" paragraph is the recipient's
    (the letterhead's "... e-mail: pl.kpi@ukr.net" line does not start with it)."""
    lines = body_html.split("\n")
    for i, line in enumerate(lines):
        if line.startswith("<p") and EMAIL_LINE_RE.match(html_to_text(line)):
            lines[i] = RECIPIENT_VALUE_RE.sub(RECIPIENT_TOKEN, line, count=1)
            break
    return "\n".join(lines)


def _paragraph_text(p) -> str:
    """Paragraph text for layout detection (tabs/breaks as whitespace, deleted text excluded)."""
    parts = []
    for el in p.iter(W + "t", W + "tab", W + "br", W + "cr"):
        parts.append(el.text or "" if el.tag == W + "t" else " ")
    return " ".join("".join(parts).split())


def _visible_text(el) -> str:
    return "".join(t.text or "" for t in el.iter(W + "t"))


def _trim_empty(blocks: list) -> list:
    """Drop empty paragraphs (e.g. the page break between letters) at both ends."""
    keep = [i for i, el in enumerate(blocks) if el.tag == TBL_TAG or _visible_text(el).strip()]
    return blocks[keep[0] : keep[-1] + 1] if keep else []


# --------------------------------------------------------------------------- letters


def _detect_lang(letterhead_lang: str | None, toc_marker: str | None, sample: str) -> str:
    if letterhead_lang:
        return letterhead_lang
    if toc_marker:
        return "en" if toc_marker.lower().startswith("en") else "ua"
    letters = [c for c in sample if c.isalpha()]
    cyrillic = sum(1 for c in letters if CYRILLIC_RE.match(c))
    return "ua" if letters and cyrillic / len(letters) > 0.3 else "en"


def _find_attachment(lines: dict[int, str]) -> tuple[str, set[int]]:
    """Explicitly named attachment(s) (';'-joined) and the block indexes of those lines."""
    found: list[str] = []
    used: set[int] = set()
    existing = {p.name.lower(): p.name for p in ATTACHMENTS_DIR.glob("*") if p.is_file()}
    for idx, text in lines.items():
        match = ATTACHMENT_LINE_RE.match(text)
        if not match:
            continue
        used.add(idx)
        for name in re.split(r"[;,]", match.group(1)):
            name = name.strip().strip("«»\"'")
            if name:
                found.append(existing.get(name.lower(), name))
    return "; ".join(dict.fromkeys(found)), used


def parse_docx(path: Path = config.DOCX_PATH) -> pd.DataFrame:
    document = docx.Document(str(path))
    converter = DocxHtmlConverter(document)
    blocks = list(document.element.body.iterchildren(P_TAG, TBL_TAG))
    texts = [_paragraph_text(el) if el.tag == P_TAG else "" for el in blocks]

    starts = [i for i, t in enumerate(texts) if t in LETTERHEADS]
    if not starts:
        raise ValueError(
            "No letters found: expected each letter to start with the letterhead line "
            f"{' or '.join(repr(k) for k in LETTERHEADS)}."
        )

    toc = [m for m in (TOC_RE.match(t) for t in texts[: starts[0]]) if m]
    rows = []
    bounds = starts + [len(blocks)]
    for n, start in enumerate(starts):
        end = bounds[n + 1]
        email_idx = next((i for i in range(start, min(start + 25, end)) if EMAIL_LINE_RE.match(texts[i])), None)
        if email_idx is None:
            raise ValueError(f"Letter {n + 1} (block {start}) has no 'E-mail:' line near its top.")

        addresses = ADDRESS_RE.findall(EMAIL_LINE_RE.match(texts[email_idx]).group(1))
        date_idx = max((i for i in range(start, email_idx) if DATE_LINE_RE.search(texts[i])), default=start)
        addressee = [texts[i] for i in range(date_idx + 1, email_idx) if texts[i] and not PHONE_LINE_RE.match(texts[i])]

        subject_idx = next((i for i in range(email_idx + 1, end) if texts[i]), None)
        if subject_idx is None:
            raise ValueError(f"Letter {n + 1} (block {start}) has no subject/body after the e-mail line.")
        # The e-mail is the whole letter as laid out in Word: letterhead, date, addressee block,
        # subject line, salutation, body, signatures and contacts.
        attachment, attachment_lines = _find_attachment({i: texts[i] for i in range(subject_idx + 1, end) if texts[i]})
        body_html = converter.letter([blocks[i] for i in range(start, end) if i not in attachment_lines])
        body_html = _recipient_placeholder(BLANK_DATE_RE.sub(DATE_TOKEN, body_html, count=1))
        body = html_to_text(body_html)

        toc_match = toc[n] if len(toc) == len(starts) else None
        lang = _detect_lang(LETTERHEADS.get(texts[start]), toc_match and toc_match.group(3), texts[subject_idx] + body[:500])
        if toc_match:
            organization = toc_match.group(2).strip()
        else:
            organization = ADDRESSEE_PREFIXES.sub("", addressee[0]).strip("«» ") if addressee else ""

        rows.append(
            {
                "id": n + 1,
                "organization": organization,
                "recipient_email": ", ".join(dict.fromkeys(addresses)),
                "subject": " ".join(texts[subject_idx].split()),
                "body": body,
                "body_html": body_html,
                "lang": lang,
                "attachment_filename": attachment or DEFAULT_ATTACHMENTS.get(lang, ""),
                "status": STATUS_PENDING if addresses else STATUS_ERROR,
                "error_message": "" if addresses else MISSING_EMAIL_MESSAGE,
                "sent_at": "",
            }
        )
    return pd.DataFrame(rows, columns=QUEUE_COLUMNS)


# --------------------------------------------------------------------------- queue


def _history_key(row) -> tuple[str, str]:
    return (
        " ".join(str(row["recipient_email"]).lower().split()),
        " ".join(str(row["subject"]).lower().split()),
    )


def carry_over_sent(new: pd.DataFrame, old: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """Keep 'sent' status for letters that were already sent (matched by recipient + subject).

    Matching on content rather than id means reordering letters in the .docx never
    causes a second e-mail to someone who already got one.
    """
    history: dict[tuple[str, str], deque] = defaultdict(deque)
    for _, row in old[old["status"] == STATUS_SENT].iterrows():
        history[_history_key(row)].append(row)

    new = new.copy()
    kept = 0
    for idx, row in new.iterrows():
        queue = history.get(_history_key(row))
        if row["recipient_email"] and queue:
            previous = queue.popleft()
            new.loc[idx, ["status", "sent_at", "error_message"]] = [STATUS_SENT, previous["sent_at"], ""]
            kept += 1
    return new, kept


def import_docx(docx_path: Path = config.DOCX_PATH, keep_sent: bool = True) -> dict:
    """Parse the .docx and (re)write the queue. Returns a summary for the UI."""
    new = parse_docx(docx_path)
    kept = 0
    if queue_store.exists():
        shutil.copyfile(queue_store.path, QUEUE_BACKUP_CSV)
        if keep_sent:
            new, kept = carry_over_sent(new, queue_store.load())
    queue_store.save(new)
    return {
        "total": len(new),
        "ua": int((new["lang"] == "ua").sum()),
        "en": int((new["lang"] == "en").sum()),
        "missing_email": int((new["recipient_email"] == "").sum()),
        "kept_sent": kept,
        "pending": int((new["status"] == STATUS_PENDING).sum()),
    }


def main() -> None:
    cli = argparse.ArgumentParser(description="Parse emails.docx into data/queue.csv")
    cli.add_argument("--docx", type=Path, default=config.DOCX_PATH)
    cli.add_argument("--fresh", action="store_true", help="do not keep 'sent' status from the existing queue")
    args = cli.parse_args()
    summary = import_docx(args.docx, keep_sent=not args.fresh)
    print(
        f"{summary['total']} letters (UA {summary['ua']}, EN {summary['en']}) -> {config.QUEUE_CSV}\n"
        f"pending: {summary['pending']}, missing e-mail: {summary['missing_email']}, "
        f"kept as sent: {summary['kept_sent']}"
    )


if __name__ == "__main__":
    main()
