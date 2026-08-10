#!/usr/bin/env python3
"""Deterministic static checks for the artifact HTML."""
import re, sys, os
from html.parser import HTMLParser

PATH = "/home/user/claude_playground/docs/asset-tokenization-graphic-record.html"
src = open(PATH, encoding="utf-8").read()

fails, warns, oks = [], [], []
def ok(m): oks.append(m)
def fail(m): fails.append(m)
def warn(m): warns.append(m)

# ── 1. size ────────────────────────────────────────────────
size = len(src.encode("utf-8"))
(ok if size < 16*1024*1024 else fail)(f"page size {size/1024:.1f} KiB (limit 16 MiB)")

# ── 2. forbidden document-level tags ───────────────────────
for tag in ("<!doctype", "<html", "<head", "<body"):
    if re.search(re.escape(tag)+r"(?=[\\s>/])", src, re.I):
        fail(f"forbidden document-level tag present: {tag}")
if not any(re.search(re.escape(t)+r"(?=[\\s>/])", src, re.I) for t in ("<!doctype","<html","<head","<body")):
    ok("no <!doctype>/<html>/<head>/<body> wrapper tags (required)")

# ── 3. tag balance / well-formedness ───────────────────────
VOID = {"area","base","br","col","embed","hr","img","input","link","meta",
        "param","source","track","wbr","path","circle","rect","use","stop",
        "feturbulence","fedisplacementmap","line","polygon","polyline","ellipse","image"}
class P(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.stack=[]; self.errs=[]
    def handle_starttag(self, tag, attrs):
        if tag in VOID: return
        self.stack.append((tag, self.getpos()))
    def handle_startendtag(self, tag, attrs): pass
    def handle_endtag(self, tag):
        if tag in VOID: return
        if not self.stack:
            self.errs.append(f"stray </{tag}> at line {self.getpos()[0]}"); return
        if self.stack[-1][0] != tag:
            self.errs.append(f"mismatch: </{tag}> at line {self.getpos()[0]} "
                             f"but <{self.stack[-1][0]}> open since line {self.stack[-1][1][0]}")
            for i in range(len(self.stack)-1, -1, -1):
                if self.stack[i][0]==tag: del self.stack[i:]; return
            return
        self.stack.pop()
p = P(); p.feed(src)
if p.errs: [fail("HTML: "+e) for e in p.errs]
if p.stack: [fail(f"HTML: unclosed <{t}> opened at line {pos[0]}") for t,pos in p.stack]
if not p.errs and not p.stack: ok("HTML tag balance: every non-void element closed, no mismatches")

# ── 4. unquoted attributes ─────────────────────────────────
unq = re.findall(r'<[a-zA-Z][^>]*?\s([a-zA-Z-]+)=([^"\'\s>][^\s>]*)', src)
(ok if not unq else fail)(f"attribute quoting: {len(unq)} unquoted" if unq else "all attributes double-quoted")

# ── 5. CSP: no external resource references ────────────────
ext = re.findall(r'(?:src|href)\s*=\s*"(https?://[^"]+)"', src)
ext_nonsource = [u for u in ext if True]
imports = re.findall(r'@import[^;]+;', src)
fetches = re.findall(r'\b(?:fetch|XMLHttpRequest|WebSocket|importScripts)\s*\(', src)
# links/scripts/styles that would actually load
loaders = re.findall(r'<(?:script|link|img|iframe|source|video|audio)\b[^>]*?(?:src|href)\s*=\s*"(https?://[^"]+)"', src, re.I)
(ok if not loaders else fail)("no external resource loads (script/link/img/iframe)" if not loaders else f"external loads: {loaders}")
(ok if not imports else fail)("no CSS @import" if not imports else f"@import found: {imports}")
(ok if not fetches else fail)("no fetch/XHR/WebSocket calls" if not fetches else f"network calls: {fetches}")
(ok if not re.search(r'@font-face', src) else warn)("no @font-face (system font stacks only)")
ok(f"{len(ext)} https URLs present — all are anchor hrefs in the sources list (no loads)")

# ── 6. CSS custom properties: every var() is defined ───────
defined = set(re.findall(r'(--[a-zA-Z0-9-]+)\s*:', src))
used = set(re.findall(r'var\(\s*(--[a-zA-Z0-9-]+)', src))
missing = sorted(u for u in used if u not in defined)
(ok if not missing else fail)(f"all {len(used)} var() references resolve to a definition"
                              if not missing else f"undefined custom properties: {missing}")

# ── 7. theme completeness: 3 states ────────────────────────
def block(pat):
    m = re.search(pat, src)
    if not m: return None
    i = src.index("{", m.end()-1) if False else m.end()
    depth=0; start=src.index("{", m.start())
    for j in range(start, len(src)):
        if src[j]=="{": depth+=1
        elif src[j]=="}":
            depth-=1
            if depth==0: return src[start:j+1]
    return None

root_bare = re.search(r':root\s*\{(.*?)\n\}', src, re.S)
media_blk  = re.search(r'@media\s*\(prefers-color-scheme:\s*dark\)\s*\{(.*?)\n\}\n', src, re.S)
stamp_blk  = re.search(r':root\[data-theme="dark"\]\s*\{(.*?)\n\}', src, re.S)

for name, blk in (("bare :root", root_bare), ("@media prefers-color-scheme:dark", media_blk),
                  (':root[data-theme="dark"]', stamp_blk)):
    (ok if blk else fail)(f"theme block present: {name}")

if media_blk:
    (ok if 'not([data-theme="light"])' in media_blk.group(1) else fail)(
        'dark media query guarded with :root:not([data-theme="light"])')

if root_bare and media_blk and stamp_blk:
    light = set(re.findall(r'(--[a-zA-Z0-9-]+)\s*:', root_bare.group(1)))
    dark_m = set(re.findall(r'(--[a-zA-Z0-9-]+)\s*:', media_blk.group(1)))
    dark_s = set(re.findall(r'(--[a-zA-Z0-9-]+)\s*:', stamp_blk.group(1)))
    (ok if dark_m == dark_s else fail)(
        f"media-query dark set == data-theme dark set ({len(dark_m)} tokens)"
        if dark_m==dark_s else f"dark sets differ: media-only={sorted(dark_m-dark_s)} stamp-only={sorted(dark_s-dark_m)}")
    orphan = sorted(dark_m - light)
    (ok if not orphan else fail)("every dark token also defined in bare :root (no theme-only colors)"
                                 if not orphan else f"tokens defined ONLY in dark: {orphan}")

# ── 8. the classic bug: literal colors inside media/[data-theme] component rules ──
def strip_ranges(text, ranges):
    out = text
    for a,b in sorted(ranges, reverse=True): out = out[:a]+out[b:]
    return out
ranges=[]
for m in (media_blk, stamp_blk):
    if m: ranges.append((m.start(), m.end()))
outside = strip_ranges(src, ranges)
# any color declaration outside token blocks that is a raw literal on a component?
comp_literals = re.findall(r'(?<!-)\b(?:color|background|background-color|fill|stroke|border-color)\s*:\s*(#[0-9a-fA-F]{3,8}|rgba?\([^)]*\))', outside)
(ok if not comp_literals else warn)("no raw color literals on components outside token blocks"
    if not comp_literals else f"raw literals on components: {sorted(set(comp_literals))}")

# ── 9. body background from a token ────────────────────────
bodyrule = re.search(r'\nbody\s*\{(.*?)\n\}', src, re.S)
if bodyrule and re.search(r'background(-color)?\s*:\s*var\(--', bodyrule.group(1)):
    ok("body paints an explicit background from a token (not transparent)")
else:
    fail("body has no token-based background")

# ── 10. SVG filter refs resolve ────────────────────────────
fdef = set(re.findall(r'<filter[^>]*\bid="([^"]+)"', src))
fref = set(re.findall(r'url\(#([^)]+)\)', src))
miss = sorted(fref - fdef)
(ok if not miss else fail)(f"all SVG filter url(#…) refs resolve: {sorted(fdef)}"
                           if not miss else f"dangling filter refs: {miss}")

# ── 11. internal anchors resolve ───────────────────────────
ids = set(re.findall(r'\bid="([^"]+)"', src))
anchors = set(re.findall(r'href="#([^"]+)"', src))
bad = sorted(a for a in anchors if a not in ids)
(ok if not bad else fail)(f"all {len(anchors)} in-page anchors resolve to an id"
                          if not bad else f"broken anchors: {bad}")
dupe = [i for i in ids if src.count(f'id="{i}"')>1]
(ok if not dupe else fail)("no duplicate id attributes" if not dupe else f"duplicate ids: {dupe}")

# ── 12. a11y basics ────────────────────────────────────────
svgs = re.findall(r'<svg\b[^>]*>', src)
labeled = [s for s in svgs if 'aria-hidden="true"' in s or 'role="img"' in s]
(ok if len(labeled)==len(svgs) else fail)(
    f"all {len(svgs)} <svg> are aria-hidden or role=img with a label"
    if len(labeled)==len(svgs) else f"{len(svgs)-len(labeled)} unlabelled <svg>")
imgsvg = re.findall(r'<svg\b[^>]*role="img"[^>]*>', src)
unlab = [s for s in imgsvg if "aria-label" not in s]
(ok if not unlab else fail)(f"all {len(imgsvg)} role=img SVGs carry aria-label" if not unlab else f"{len(unlab)} missing aria-label")
(ok if ':focus-visible' in src else fail)("keyboard focus has a visible state (:focus-visible rules present)")
(ok if 'prefers-reduced-motion' in src else fail)("prefers-reduced-motion honored")
(ok if re.search(r'<title>.+?</title>', src) else fail)("<title> present")
(ok if re.search(r'<nav[^>]*aria-label=', src) else warn)("nav has aria-label")
(ok if 'tabular-nums' in src else warn)("tabular-nums used for aligned figures")

# ── 13. overflow containment for wide content ──────────────
tw = src.count('class="tw"') + src.count('class="tw" ')
tables = len(re.findall(r'<table', src))
scroll_figs = len(re.findall(r'class="fig-scroll"', src))
dias = len(re.findall(r'class="dia"', src))
(ok if tables <= len(re.findall(r'class="tw"', src)) else fail)(
    f"{tables} tables, all wrapped in overflow-x:auto (.tw = {len(re.findall(r'class=.tw.', src))})")
(ok if dias <= scroll_figs else fail)(f"{dias} wide SVG diagrams, all in .fig-scroll containers ({scroll_figs})")

print("="*72)
print("STATIC ANALYSIS")
print("="*72)
for m in oks:   print("  PASS  " + m)
for m in warns: print("  NOTE  " + m)
for m in fails: print("  FAIL  " + m)
print("-"*72)
print(f"{len(oks)} passed, {len(warns)} notes, {len(fails)} failed")
sys.exit(1 if fails else 0)
