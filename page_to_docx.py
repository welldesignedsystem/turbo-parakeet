#!/usr/bin/env python3
"""
page_to_docx.py
---------------
Downloads web pages listed in pages.txt and combines them into a single Word
document (output.docx).

Features
--------
- Headings h1–h6
- Paragraphs, divs, blockquotes, pre / code blocks
- Bold, italic, underline, strikethrough, superscript, subscript, <mark>
- Inline style="" parsing: font-weight, font-style, font-size, color,
  background-color, text-decoration, font-family, vertical-align
- Hyperlinks (real clickable links in the docx, with Hyperlink style)
- Nested inline formatting (e.g. bold inside a link inside italic)
- Ordered and unordered lists (including nested lists)
- Tables (with header-row bolding and colspan support)
- Images (remote URLs + data URIs + SVG via cairosvg)
- YouTube / video iframes → embedded thumbnail image + caption link
- Mermaid diagrams rendered to PNG via mmdc CLI

Dependencies:
    pip install requests beautifulsoup4 python-docx pillow lxml
    (optional SVG)   pip install cairosvg
    (Mermaid render) npm install -g @mermaid-js/mermaid-cli
"""

import os
import re
import sys
import time
import base64
import shutil
import subprocess
import tempfile
import requests
from io import BytesIO
from urllib.parse import urljoin, urlparse, parse_qs

from bs4 import BeautifulSoup, NavigableString, Tag
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from PIL import Image


# ---------------------------------------------------------------------------
# HTTP helpers
# ---------------------------------------------------------------------------

HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; PageToDocx/1.0)"}


def fetch(url: str, timeout: int = 30) -> requests.Response:
    r = requests.get(url, headers=HEADERS, timeout=timeout)
    r.raise_for_status()
    return r


def svg_to_png_bytes(svg_bytes: bytes) -> bytes | None:
    try:
        import cairosvg
        return cairosvg.svg2png(bytestring=svg_bytes)
    except Exception:
        return None


def fetch_image_bytes(url: str) -> tuple[bytes, str] | tuple[None, None]:
    try:
        r = fetch(url)
        content_type = r.headers.get("Content-Type", "")
        data = r.content
        is_svg = (
            "svg" in content_type
            or url.lower().endswith(".svg")
            or data[:100].lstrip().startswith(b"<svg")
            or data[:100].lstrip().startswith(b"<?xml")
        )
        if is_svg:
            png = svg_to_png_bytes(data)
            return (png, "PNG") if png else (None, None)
        img = Image.open(BytesIO(data))
        fmt = img.format or "PNG"
        if img.mode not in ("RGB", "RGBA", "L"):
            img = img.convert("RGBA")
        buf = BytesIO()
        img.save(buf, format=fmt if fmt in ("JPEG", "PNG", "GIF", "BMP") else "PNG")
        return buf.getvalue(), fmt
    except Exception:
        return None, None


# ---------------------------------------------------------------------------
# YouTube / video iframe helpers
# ---------------------------------------------------------------------------

def _extract_youtube_id(url: str) -> str | None:
    """Extract YouTube video ID from various URL formats."""
    if not url:
        return None
    # youtu.be/ID
    m = re.search(r"youtu\.be/([A-Za-z0-9_-]{11})", url)
    if m:
        return m.group(1)
    # youtube.com/watch?v=ID  or  /embed/ID  or  /v/ID
    m = re.search(r"youtube\.com/(?:watch\?v=|embed/|v/)([A-Za-z0-9_-]{11})", url)
    if m:
        return m.group(1)
    # youtube-nocookie.com/embed/ID
    m = re.search(r"youtube-nocookie\.com/embed/([A-Za-z0-9_-]{11})", url)
    if m:
        return m.group(1)
    return None


def _youtube_thumbnail_bytes(video_id: str) -> bytes | None:
    """Try to fetch the best-quality YouTube thumbnail available."""
    for quality in ("maxresdefault", "hqdefault", "mqdefault", "default"):
        url = f"https://img.youtube.com/vi/{video_id}/{quality}.jpg"
        try:
            img_bytes, _ = fetch_image_bytes(url)
            if img_bytes:
                return img_bytes
        except Exception:
            continue
    return None


def _extract_vimeo_id(url: str) -> str | None:
    """Extract Vimeo video ID."""
    m = re.search(r"vimeo\.com/(?:video/)?(\d+)", url)
    return m.group(1) if m else None


def _vimeo_thumbnail_bytes(video_id: str) -> tuple[bytes | None, str]:
    """Fetch Vimeo thumbnail via oEmbed API."""
    try:
        api_url = f"https://vimeo.com/api/v2/video/{video_id}.json"
        r = fetch(api_url)
        data = r.json()
        thumb_url = data[0].get("thumbnail_large") or data[0].get("thumbnail_medium")
        if thumb_url:
            img_bytes, _ = fetch_image_bytes(thumb_url)
            title = data[0].get("title", "")
            return img_bytes, title
    except Exception:
        pass
    return None, ""


def _get_iframe_src(node: Tag) -> str:
    """Return the src (or data-src) of an iframe."""
    return (node.get("src") or node.get("data-src") or "").strip()


# ---------------------------------------------------------------------------
# Mermaid rendering
# ---------------------------------------------------------------------------

_MMDC_PATH = shutil.which("mmdc")


def render_mermaid_to_png(mermaid_code: str) -> bytes | None:
    if not _MMDC_PATH:
        print("  [mermaid] mmdc not found. Install: npm install -g @mermaid-js/mermaid-cli")
        return None
    mermaid_code = mermaid_code.strip()
    if not mermaid_code:
        return None
    with tempfile.TemporaryDirectory() as tmpdir:
        inp = os.path.join(tmpdir, "diagram.mmd")
        out = os.path.join(tmpdir, "diagram.png")
        with open(inp, "w", encoding="utf-8") as f:
            f.write(mermaid_code)
        try:
            result = subprocess.run(
                [_MMDC_PATH, "-i", inp, "-o", out, "-b", "white",
                 "--width", "1200", "--height", "800", "--quiet"],
                capture_output=True, timeout=30,
            )
            if result.returncode != 0:
                print(f"  [mermaid] error: {result.stderr.decode(errors='replace')[:200]}")
                return None
            if not os.path.exists(out):
                return None
            with open(out, "rb") as f:
                return f.read()
        except Exception as e:
            print(f"  [mermaid] {e}")
            return None


def _is_mermaid_tag(node: Tag) -> bool:
    if node.name not in ("pre", "div", "code"):
        return False
    classes = node.get("class") or []
    if isinstance(classes, str):
        classes = classes.split()
    if "mermaid" in classes:
        return True
    if node.name == "pre":
        code = node.find("code")
        if code:
            cc = code.get("class") or []
            if isinstance(cc, str):
                cc = cc.split()
            if any("mermaid" in c for c in cc):
                return True
    if node.name == "code" and any("mermaid" in c for c in classes):
        return True
    return False


def _extract_mermaid_code(node: Tag) -> str:
    code_child = node.find("code") if node.name == "pre" else None
    return code_child.get_text() if code_child else node.get_text()


# ---------------------------------------------------------------------------
# CSS inline style parsing
# ---------------------------------------------------------------------------

def _parse_css_color(value: str) -> RGBColor | None:
    value = value.strip().lower()
    m = re.match(r"#([0-9a-f]{6})$", value)
    if m:
        h = m.group(1)
        return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    m = re.match(r"#([0-9a-f]{3})$", value)
    if m:
        h = m.group(1)
        return RGBColor(int(h[0] * 2, 16), int(h[1] * 2, 16), int(h[2] * 2, 16))
    m = re.match(r"rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)", value)
    if m:
        return RGBColor(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    NAMED = {
        "red": (255, 0, 0), "green": (0, 128, 0), "blue": (0, 0, 255),
        "black": (0, 0, 0), "white": (255, 255, 255), "gray": (128, 128, 128),
        "grey": (128, 128, 128), "orange": (255, 165, 0), "yellow": (255, 255, 0),
        "purple": (128, 0, 128), "pink": (255, 192, 203), "brown": (165, 42, 42),
        "navy": (0, 0, 128), "teal": (0, 128, 128), "silver": (192, 192, 192),
        "maroon": (128, 0, 0), "lime": (0, 255, 0), "aqua": (0, 255, 255),
        "cyan": (0, 255, 255), "fuchsia": (255, 0, 255), "magenta": (255, 0, 255),
    }
    return RGBColor(*NAMED[value]) if value in NAMED else None


def _parse_css_pt(value: str) -> Pt | None:
    value = value.strip().lower()
    m = re.match(r"([\d.]+)(px|pt|em|rem)?$", value)
    if not m:
        return None
    num = float(m.group(1))
    unit = m.group(2) or "px"
    if unit == "pt":
        return Pt(num)
    if unit in ("em", "rem"):
        return Pt(num * 12)
    return Pt(num * 0.75)  # px → pt at 96dpi


# ---------------------------------------------------------------------------
# Run formatting accumulator
# ---------------------------------------------------------------------------

class _RunFmt:
    """
    Immutable-ish bag of formatting that is passed down through the
    inline-content tree and applied to each Run.
    """
    __slots__ = (
        "bold", "italic", "underline", "strike",
        "superscript", "subscript",
        "color", "font_name", "font_size",
        "is_code", "is_link", "link_url",
    )

    def __init__(self):
        self.bold: bool | None = None
        self.italic: bool | None = None
        self.underline: bool | None = None
        self.strike: bool | None = None
        self.superscript: bool | None = None
        self.subscript: bool | None = None
        self.color: RGBColor | None = None
        self.font_name: str | None = None
        self.font_size: Pt | None = None
        self.is_code: bool = False
        self.is_link: bool = False
        self.link_url: str = ""

    def copy(self) -> "_RunFmt":
        n = _RunFmt()
        for s in self.__slots__:
            setattr(n, s, getattr(self, s))
        return n

    def apply_tag(self, tag: str, node: Tag) -> "_RunFmt":
        """Return a new _RunFmt with this HTML tag's formatting layered on."""
        f = self.copy()

        # Semantic tags
        if tag in ("strong", "b"):
            f.bold = True
        elif tag in ("em", "i"):
            f.italic = True
        elif tag == "u":
            f.underline = True
        elif tag in ("s", "del", "strike"):
            f.strike = True
        elif tag == "sup":
            f.superscript = True
            f.subscript = None
        elif tag == "sub":
            f.subscript = True
            f.superscript = None
        elif tag in ("code", "tt", "kbd", "samp", "var"):
            f.is_code = True
            f.font_name = "Courier New"
            f.font_size = Pt(9)
        elif tag == "mark":
            f.color = RGBColor(0xFF, 0xC0, 0x00)   # gold highlight approximation
        elif tag == "small":
            f.font_size = Pt(8)
        elif tag == "a":
            href = node.get("href", "").strip()
            if href and not href.startswith("#") and not href.startswith("javascript:"):
                f.is_link = True
                f.link_url = href
                f.color = RGBColor(0x1F, 0x69, 0xC0)
                f.underline = True

        # <font> tag legacy support
        if tag == "font":
            color_attr = node.get("color", "")
            if color_attr:
                c = _parse_css_color(color_attr)
                if c:
                    f.color = c
            face = node.get("face", "")
            if face:
                f.font_name = face.split(",")[0].strip()
            size_attr = node.get("size", "")
            # HTML font size 1–7 mapped to pt
            if size_attr.isdigit():
                sizes = {1: 8, 2: 10, 3: 12, 4: 14, 5: 18, 6: 24, 7: 36}
                f.font_size = Pt(sizes.get(int(size_attr), 12))

        # Inline style="" attribute (highest priority)
        style_attr = node.get("style", "")
        if style_attr:
            f._apply_inline_style(style_attr)

        return f

    def _apply_inline_style(self, style_attr: str):
        for decl in style_attr.split(";"):
            decl = decl.strip()
            if ":" not in decl:
                continue
            prop, _, val = decl.partition(":")
            prop = prop.strip().lower()
            val = val.strip()
            if not val:
                continue

            if prop == "font-weight":
                if val in ("bold", "bolder") or (val.isdigit() and int(val) >= 600):
                    self.bold = True
                elif val in ("normal", "lighter") or (val.isdigit() and int(val) < 600):
                    self.bold = False

            elif prop == "font-style":
                self.italic = val in ("italic", "oblique")

            elif prop == "text-decoration":
                if "underline" in val:
                    self.underline = True
                if "line-through" in val:
                    self.strike = True
                if "none" in val:
                    self.underline = False
                    self.strike = False

            elif prop == "color":
                c = _parse_css_color(val)
                if c:
                    self.color = c

            elif prop == "font-size":
                pt = _parse_css_pt(val)
                if pt:
                    self.font_size = pt

            elif prop == "font-family":
                first = val.split(",")[0].strip().strip("'\"")
                if first:
                    self.font_name = first

            elif prop == "vertical-align":
                if val == "super":
                    self.superscript = True
                    self.subscript = None
                elif val == "sub":
                    self.subscript = True
                    self.superscript = None

    def apply_to_run(self, run):
        if self.bold is not None:
            run.bold = self.bold
        if self.italic is not None:
            run.italic = self.italic
        if self.underline is not None:
            run.underline = self.underline
        if self.strike:
            run.font.strike = True
        if self.superscript:
            run.font.superscript = True
        if self.subscript:
            run.font.subscript = True
        if self.color:
            run.font.color.rgb = self.color
        if self.font_name:
            run.font.name = self.font_name
        if self.font_size:
            run.font.size = self.font_size


# ---------------------------------------------------------------------------
# Hyperlink helper
# ---------------------------------------------------------------------------

def _add_hyperlink(paragraph, url: str, text: str, fmt: "_RunFmt"):
    """
    Insert a real clickable hyperlink into *paragraph*.
    """
    try:
        part = paragraph.part
        r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)
    except Exception:
        run = paragraph.add_run(text)
        fmt.apply_to_run(run)
        return

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    hyperlink.set(qn("w:history"), "1")

    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")

    rStyle = OxmlElement("w:rStyle")
    rStyle.set(qn("w:val"), "Hyperlink")
    rPr.append(rStyle)

    if fmt.bold:
        rPr.append(OxmlElement("w:b"))
    if fmt.italic:
        rPr.append(OxmlElement("w:i"))
    if fmt.strike:
        rPr.append(OxmlElement("w:strike"))
    if fmt.font_size:
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), str(int(fmt.font_size.pt * 2)))
        rPr.append(sz)
        szCs = OxmlElement("w:szCs")
        szCs.set(qn("w:val"), str(int(fmt.font_size.pt * 2)))
        rPr.append(szCs)
    if fmt.font_name:
        fonts = OxmlElement("w:rFonts")
        fonts.set(qn("w:ascii"), fmt.font_name)
        fonts.set(qn("w:hAnsi"), fmt.font_name)
        rPr.append(fonts)
    if fmt.color:
        col = OxmlElement("w:color")
        col.set(qn("w:val"), f"{fmt.color[0]:02X}{fmt.color[1]:02X}{fmt.color[2]:02X}")
        rPr.append(col)

    new_run.append(rPr)

    t = OxmlElement("w:t")
    t.text = text
    if text != text.strip():
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


# ---------------------------------------------------------------------------
# Tag classification
# ---------------------------------------------------------------------------

_INLINE_TAGS = {
    "a", "abbr", "acronym", "b", "bdi", "bdo", "cite", "code", "data",
    "del", "dfn", "em", "font", "i", "ins", "kbd", "mark", "q", "s",
    "samp", "small", "span", "strike", "strong", "sub", "sup", "time",
    "tt", "u", "var",
}

_BLOCK_TAGS = {
    "address", "article", "aside", "blockquote", "dd", "details",
    "dialog", "div", "dl", "dt", "fieldset", "figcaption", "figure",
    "footer", "form", "h1", "h2", "h3", "h4", "h5", "h6", "header",
    "hgroup", "hr", "li", "main", "nav", "ol", "p", "pre", "section",
    "summary", "table", "ul",
}

_SKIP_TAGS = {
    "script", "style", "noscript", "head",
    "button", "input", "select", "textarea", "template", "canvas",
}


# ---------------------------------------------------------------------------
# HTML → python-docx renderer
# ---------------------------------------------------------------------------

class HtmlToDocx:
    IMG_MAX_WIDTH_IN = 5.5

    def __init__(self, doc: Document, base_url: str = ""):
        self.doc = doc
        self.base_url = base_url
        self._para = None   # currently open paragraph

    # ------------------------------------------------------------------
    # Public
    # ------------------------------------------------------------------

    def feed(self, node):
        self._walk_block(node)
        self._close_para()

    # ------------------------------------------------------------------
    # Block-level walk
    # ------------------------------------------------------------------

    def _walk_block(self, node):
        if isinstance(node, NavigableString):
            text = str(node)
            if text.strip():
                self._open_para()
                self._emit(text, _RunFmt())
            return

        if not isinstance(node, Tag):
            return

        tag = node.name.lower() if node.name else ""

        if tag in _SKIP_TAGS:
            return

        # ── iframe: video embeds ──────────────────────────────────────
        if tag == "iframe":
            self._handle_iframe(node)
            return

        # ── SVG: skip (no useful text content) ───────────────────────
        if tag == "svg":
            return

        # Mermaid diagrams
        if _is_mermaid_tag(node):
            self._handle_mermaid(node)
            return

        # Headings
        if tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
            self._close_para()
            level = int(tag[1])
            self._para = self.doc.add_heading("", level=min(level, 9))
            base = _RunFmt()
            base.bold = True
            for child in node.children:
                self._walk_inline(child, base)
            self._close_para()
            return

        # Generic block containers
        if tag in ("p", "div", "section", "article", "main", "header",
                   "footer", "nav", "aside", "address", "details",
                   "summary", "figure", "figcaption", "hgroup", "dialog", "form",
                   "fieldset"):
            self._close_para()
            for child in node.children:
                self._walk_block(child)
            self._close_para()
            return

        if tag == "blockquote":
            self._close_para()
            for child in node.children:
                self._walk_block(child)
            self._close_para()
            return

        # Preformatted text (non-mermaid)
        if tag == "pre":
            self._close_para()
            text = node.get_text()
            for line in text.splitlines():
                para = self.doc.add_paragraph()
                run = para.add_run(line)
                run.font.name = "Courier New"
                run.font.size = Pt(9)
            self._para = None
            return

        if tag in ("ul", "ol"):
            self._close_para()
            self._handle_list(node, ordered=(tag == "ol"), depth=0)
            return

        if tag == "table":
            self._close_para()
            self._handle_table(node)
            return

        if tag == "img":
            self._handle_img(node)
            return

        if tag == "br":
            self._open_para()
            self._para.add_run().add_break()
            return

        if tag == "hr":
            self._close_para()
            self._add_hr()
            return

        if tag == "dl":
            self._close_para()
            for child in node.children:
                self._walk_block(child)
            self._close_para()
            return

        if tag in ("dt", "dd"):
            self._close_para()
            self._open_para()
            fmt = _RunFmt()
            if tag == "dt":
                fmt.bold = True
            for child in node.children:
                self._walk_inline(child, fmt)
            self._close_para()
            return

        # Inline tags at block level → open a para and go inline
        if tag in _INLINE_TAGS:
            self._open_para()
            self._walk_inline(node, _RunFmt())
            return

        # Default: recurse
        for child in node.children:
            self._walk_block(child)

    # ------------------------------------------------------------------
    # Inline walk
    # ------------------------------------------------------------------

    def _walk_inline(self, node, fmt: "_RunFmt"):
        if isinstance(node, NavigableString):
            text = str(node)
            text = re.sub(r"[ \t\r\n]+", " ", text)
            if text:
                self._open_para()
                self._emit(text, fmt)
            return

        if not isinstance(node, Tag):
            return

        tag = node.name.lower() if node.name else ""

        if tag in _SKIP_TAGS:
            return

        if tag == "iframe":
            self._close_para()
            self._handle_iframe(node)
            return

        if tag == "svg":
            return

        if _is_mermaid_tag(node):
            self._close_para()
            self._handle_mermaid(node)
            return

        if tag in _BLOCK_TAGS and tag not in _INLINE_TAGS:
            self._close_para()
            self._walk_block(node)
            return

        if tag == "img":
            self._handle_img(node)
            return

        if tag == "br":
            self._open_para()
            self._para.add_run().add_break()
            return

        child_fmt = fmt.apply_tag(tag, node)

        if child_fmt.is_link and child_fmt.link_url:
            link_text = re.sub(r"[ \t\r\n]+", " ", node.get_text())
            if link_text.strip():
                self._open_para()
                url = urljoin(self.base_url, child_fmt.link_url)
                try:
                    _add_hyperlink(self._para, url, link_text, child_fmt)
                except Exception:
                    run = self._para.add_run(link_text)
                    child_fmt.apply_to_run(run)
            return

        for child in node.children:
            self._walk_inline(child, child_fmt)

    # ------------------------------------------------------------------
    # Emit a text run
    # ------------------------------------------------------------------

    def _emit(self, text: str, fmt: "_RunFmt"):
        run = self._para.add_run(text)
        fmt.apply_to_run(run)

    # ------------------------------------------------------------------
    # Paragraph lifecycle
    # ------------------------------------------------------------------

    def _open_para(self):
        if self._para is None:
            self._para = self.doc.add_paragraph()

    def _close_para(self):
        self._para = None

    # ------------------------------------------------------------------
    # iframe / video embed handler
    # ------------------------------------------------------------------

    def _handle_iframe(self, node: Tag):
        """
        Detect YouTube / Vimeo iframes and embed the video thumbnail
        plus a clickable caption. Other iframes are noted as plain text.
        """
        src = _get_iframe_src(node)
        if not src:
            return

        self._close_para()

        # ── YouTube ──────────────────────────────────────────────────
        yt_id = _extract_youtube_id(src)
        if yt_id:
            watch_url = f"https://www.youtube.com/watch?v={yt_id}"
            print(f"  [iframe] YouTube video {yt_id} – fetching thumbnail…")
            thumb_bytes = _youtube_thumbnail_bytes(yt_id)
            if thumb_bytes:
                self._embed_video_thumbnail(thumb_bytes, watch_url, label="▶ Watch on YouTube")
            else:
                # Fallback: plain link
                para = self.doc.add_paragraph()
                fmt = _RunFmt()
                fmt.is_link = True
                fmt.link_url = watch_url
                fmt.color = RGBColor(0x1F, 0x69, 0xC0)
                fmt.underline = True
                _add_hyperlink(para, watch_url, f"▶ YouTube video: {yt_id}", fmt)
            return

        # ── Vimeo ────────────────────────────────────────────────────
        vm_id = _extract_vimeo_id(src)
        if vm_id:
            watch_url = f"https://vimeo.com/{vm_id}"
            print(f"  [iframe] Vimeo video {vm_id} – fetching thumbnail…")
            thumb_bytes, title = _vimeo_thumbnail_bytes(vm_id)
            label = f"▶ Watch on Vimeo" + (f": {title}" if title else "")
            if thumb_bytes:
                self._embed_video_thumbnail(thumb_bytes, watch_url, label=label)
            else:
                para = self.doc.add_paragraph()
                fmt = _RunFmt()
                fmt.is_link = True
                fmt.link_url = watch_url
                fmt.color = RGBColor(0x1F, 0x69, 0xC0)
                fmt.underline = True
                _add_hyperlink(para, watch_url, label, fmt)
            return

        # ── Generic iframe: emit as a note ───────────────────────────
        title = node.get("title", "") or node.get("aria-label", "")
        label = f"[Embedded content: {title or src}]"
        # Attempt to make the src a clickable link if it looks like a URL
        if src.startswith("http"):
            para = self.doc.add_paragraph()
            fmt = _RunFmt()
            fmt.italic = True
            fmt.color = RGBColor(0x88, 0x88, 0x88)
            run = para.add_run(label)
            fmt.apply_to_run(run)
        else:
            para = self.doc.add_paragraph()
            run = para.add_run(label)
            run.italic = True
            run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

    def _embed_video_thumbnail(self, img_bytes: bytes, url: str, label: str):
        """Embed a video thumbnail image with a clickable caption below it."""
        try:
            pil = Image.open(BytesIO(img_bytes))
            w_in = min(pil.size[0] / 96, self.IMG_MAX_WIDTH_IN)
            para = self.doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.add_run().add_picture(BytesIO(img_bytes), width=Inches(w_in))
        except Exception as e:
            print(f"  [iframe] Could not embed thumbnail: {e}")
            return

        # Caption paragraph with clickable link
        cap_para = self.doc.add_paragraph()
        cap_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fmt = _RunFmt()
        fmt.is_link = True
        fmt.link_url = url
        fmt.color = RGBColor(0x1F, 0x69, 0xC0)
        fmt.underline = True
        fmt.font_size = Pt(9)
        try:
            _add_hyperlink(cap_para, url, label, fmt)
        except Exception:
            run = cap_para.add_run(label)
            fmt.apply_to_run(run)

    # ------------------------------------------------------------------
    # Element handlers
    # ------------------------------------------------------------------

    def _handle_mermaid(self, node: Tag):
        code = _extract_mermaid_code(node)
        print(f"  [mermaid] Rendering diagram ({len(code)} chars)…")
        png = render_mermaid_to_png(code)
        if not png:
            print("  [mermaid] Falling back to code block.")
            para = self.doc.add_paragraph()
            run = para.add_run(code.strip())
            run.font.name = "Courier New"
            run.font.size = Pt(8)
            self._para = None
            return
        self._close_para()
        try:
            pil = Image.open(BytesIO(png))
            w_in = min(pil.size[0] / 96, self.IMG_MAX_WIDTH_IN)
            para = self.doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.add_run().add_picture(BytesIO(png), width=Inches(w_in))
            print(f"  [mermaid] Embedded {pil.size[0]}×{pil.size[1]}px → {w_in:.1f}in.")
        except Exception as e:
            print(f"  [mermaid] Could not embed PNG: {e}")

    def _handle_img(self, node: Tag):
        # Try src, data-src (lazy-load), srcset (first URL)
        src = (
            node.get("src")
            or node.get("data-src")
            or node.get("data-lazy-src")
            or ""
        ).strip()

        # If src is a tiny placeholder, try srcset
        if not src or src.startswith("data:image/gif") or src.startswith("data:image/png"):
            srcset = node.get("srcset") or node.get("data-srcset") or ""
            if srcset:
                # Take the last (largest) URL from srcset
                candidates = [s.strip().split()[0] for s in srcset.split(",") if s.strip()]
                if candidates:
                    src = candidates[-1]

        if not src:
            return

        if src.startswith("data:image"):
            self._embed_data_uri(src)
            return
        if src.startswith("data:"):
            return

        img_url = urljoin(self.base_url, src)
        img_bytes, _ = fetch_image_bytes(img_url)
        if not img_bytes:
            return
        self._close_para()
        try:
            pil = Image.open(BytesIO(img_bytes))
            # Skip tiny tracking pixels / icons (< 10px in either dimension)
            if pil.size[0] < 10 or pil.size[1] < 10:
                return
            w_in = min(pil.size[0] / 96, self.IMG_MAX_WIDTH_IN)
            para = self.doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.add_run().add_picture(BytesIO(img_bytes), width=Inches(w_in))
        except Exception as e:
            print(f"  Could not embed image {img_url}: {e}")

    def _embed_data_uri(self, data_uri: str):
        try:
            _, encoded = data_uri.split(",", 1)
            img_bytes = base64.b64decode(encoded)
            pil = Image.open(BytesIO(img_bytes))
            if pil.size[0] < 10 or pil.size[1] < 10:
                return
            w_in = min(pil.size[0] / 96, self.IMG_MAX_WIDTH_IN)
            self._close_para()
            para = self.doc.add_paragraph()
            para.add_run().add_picture(BytesIO(img_bytes), width=Inches(w_in))
        except Exception:
            pass

    def _handle_list(self, node: Tag, ordered: bool, depth: int):
        style = "List Number" if ordered else "List Bullet"
        for child in node.children:
            if not isinstance(child, Tag):
                continue
            if child.name == "li":
                para = self.doc.add_paragraph(style=style)
                for grandchild in child.children:
                    if isinstance(grandchild, Tag) and grandchild.name in ("ul", "ol"):
                        self._para = None
                        self._handle_list(grandchild,
                                          ordered=(grandchild.name == "ol"),
                                          depth=depth + 1)
                    else:
                        self._para = para
                        self._walk_inline(grandchild, _RunFmt())
                self._para = None
            elif child.name in ("ul", "ol"):
                self._handle_list(child, ordered=(child.name == "ol"), depth=depth + 1)

    def _handle_table(self, node: Tag):
        rows_html = node.find_all("tr")
        if not rows_html:
            return
        col_count = max(
            sum(int(c.get("colspan", 1)) for c in row.find_all(["td", "th"]))
            for row in rows_html
        )
        if col_count == 0:
            return
        table = self.doc.add_table(rows=0, cols=col_count)
        table.style = "Table Grid"
        for row_html in rows_html:
            cells_html = row_html.find_all(["td", "th"])
            if not cells_html:
                continue
            row = table.add_row()
            col_idx = 0
            for cell_html in cells_html:
                if col_idx >= col_count:
                    break
                cell = row.cells[col_idx]
                self._para = cell.paragraphs[0]
                for child in cell_html.children:
                    self._walk_inline(child, _RunFmt())
                self._para = None
                if cell_html.name == "th":
                    for para in cell.paragraphs:
                        for run in para.runs:
                            run.bold = True
                col_idx += int(cell_html.get("colspan", 1))

    def _add_hr(self):
        para = self.doc.add_paragraph()
        pPr = para._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "6")
        bottom.set(qn("w:space"), "1")
        bottom.set(qn("w:color"), "AAAAAA")
        pBdr.append(bottom)
        pPr.append(pBdr)


# ---------------------------------------------------------------------------
# Main converter
# ---------------------------------------------------------------------------

class WebPageToDocx:
    def __init__(self, input_file: str = "pages.txt", output_file: str = "output.docx"):
        self.input_file = input_file
        self.output_file = output_file
        self.urls: list[str] = []

    def read_urls(self):
        try:
            with open(self.input_file) as f:
                self.urls = [line.strip() for line in f if line.strip()]
            if not self.urls:
                print(f"Error: {self.input_file} is empty or has no valid URLs")
                sys.exit(1)
            print(f"Read {len(self.urls)} URLs from {self.input_file}")
        except FileNotFoundError:
            print(f"Error: {self.input_file} not found")
            sys.exit(1)

    def _download_page(self, url: str) -> BeautifulSoup | None:
        print(f"Downloading: {url}")
        try:
            r = fetch(url)
            return BeautifulSoup(r.content, "html.parser")
        except Exception as e:
            print(f"  Error: {e}")
            return None

    def convert(self):
        doc = Document()
        style = doc.styles["Normal"]
        style.font.name = "Arial"
        style.font.size = Pt(11)

        first_page = True
        for url in self.urls:
            soup = self._download_page(url)
            if not soup:
                continue

            if not first_page:
                doc.add_page_break()
            first_page = False

            # Source URL banner
            url_para = doc.add_paragraph()
            run = url_para.add_run(f"Source: {url}")
            run.italic = True
            run.font.size = Pt(8)
            run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

            # Strip non-content tags — iframes and mermaid containers are preserved
            for tag in soup(["script", "style", "noscript", "nav",
                              "footer", "svg"]):
                tag.decompose()

            content_root = (
                soup.find("main")
                or soup.find("article")
                or soup.find(id=re.compile(r"content|main", re.I))
                or soup.find("body")
                or soup
            )

            renderer = HtmlToDocx(doc, base_url=url)
            renderer.feed(content_root)

            time.sleep(1)

        try:
            doc.save(self.output_file)
            print(f"\nWord document saved: {self.output_file}")
            return True
        except Exception as e:
            print(f"Error saving document: {e}")
            return False


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    try:
        import docx   # noqa: F401
        import bs4    # noqa: F401
        from PIL import Image  # noqa: F401
    except ImportError as e:
        print(f"Missing dependency: {e}")
        print("Install: pip install requests beautifulsoup4 python-docx pillow")
        sys.exit(1)

    if not _MMDC_PATH:
        print("Warning: mmdc not found – Mermaid diagrams will be plain text.\n"
              "Install with: npm install -g @mermaid-js/mermaid-cli")

    input_file  = sys.argv[1] if len(sys.argv) > 1 else "pages.txt"
    output_file = sys.argv[2] if len(sys.argv) > 2 else "output.docx"

    converter = WebPageToDocx(input_file=input_file, output_file=output_file)
    converter.read_urls()
    success = converter.convert()

    print("Process completed successfully!" if success else "Process failed.")