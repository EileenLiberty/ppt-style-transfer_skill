# -*- coding: utf-8 -*-
"""Source slide image geometry extraction (EMU)."""
# -*- coding: utf-8 -*-
from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

# fallback if previous output is locked by PowerPoint/WPS

FOOTER_RE = re.compile(r"^(中国民航大学\s*)?(CAUC)?$", re.I)
MASTER_JUNK = ("单击此处编辑",)
MIN_EMU = 50000  # ignore broken nested OLE preview frames
PT_TO_EMU = 12700

def xml_text(s: str) -> str:
    return escape(s, {"'": "&apos;", '"': "&quot;"})

def parse_vml_length(s: str) -> int:
    s = s.strip().lower()
    if s.endswith("pt"):
        return int(round(float(s[:-2]) * PT_TO_EMU))
    if s.endswith("in"):
        return int(round(float(s[:-2]) * 914400))
    if s.endswith("mm"):
        return int(round(float(s[:-2]) * 36000))
    if s.endswith("cm"):
        return int(round(float(s[:-2]) * 360000))
    return int(round(float(re.sub(r"[^\d.]+$", "", s)) * PT_TO_EMU))

def _parse_rels_map(rels_xml: str) -> dict[str, str]:
    rid_to_target = {}
    for m in re.finditer(r'<Relationship\s[^>]*/>', rels_xml):
        tag = m.group(0)
        rid = re.search(r'Id="(rId\d+)"', tag)
        tgt = re.search(r'Target="([^"]+)"', tag)
        if rid and tgt:
            rid_to_target[rid.group(1)] = tgt.group(1)
    return rid_to_target

def extract_image_placements(slide_idx: int, src_unpacked: Path) -> list[dict]:
    """Exact source image geometries: [{media,x,y,cx,cy}, ...] in EMU."""
    slide_xml_path = src_unpacked / "ppt" / "slides" / f"slide{slide_idx + 1}.xml"
    rels_path = src_unpacked / "ppt" / "slides" / "_rels" / f"slide{slide_idx + 1}.xml.rels"
    if not slide_xml_path.exists() or not rels_path.exists():
        return []

    slide_xml = slide_xml_path.read_text(encoding="utf-8")
    rels_xml = rels_path.read_text(encoding="utf-8")
    rid_map = _parse_rels_map(rels_xml)

    placements: list[dict] = []

    def add(media: str, x: int, y: int, cx: int, cy: int, source: str):
        if cx < MIN_EMU or cy < MIN_EMU:
            return
        # dedupe same media at nearly same origin
        for p in placements:
            if p["media"] == media and abs(p["x"] - x) < 20000 and abs(p["y"] - y) < 20000:
                # keep larger box
                if cx * cy > p["cx"] * p["cy"]:
                    p.update(x=x, y=y, cx=cx, cy=cy, source=source)
                return
        placements.append(
            {"media": media, "x": x, "y": y, "cx": cx, "cy": cy, "source": source}
        )

    # 1) VML drawings (Visio/OLE true on-slide position in pt)
    for rid, tgt in rid_map.items():
        if "vmlDrawing" not in tgt:
            continue
        vml_name = Path(tgt).name
        vml_path = src_unpacked / "ppt" / "drawings" / vml_name
        vml_rels = src_unpacked / "ppt" / "drawings" / "_rels" / f"{vml_name}.rels"
        if not vml_path.exists():
            continue
        vml = vml_path.read_text(encoding="utf-8", errors="ignore")
        vml_rid_map = _parse_rels_map(vml_rels.read_text(encoding="utf-8")) if vml_rels.exists() else {}
        for sm in re.finditer(r"<v:shape\b([^>]*)>(.*?)</v:shape>", vml, re.S | re.I):
            attrs, body = sm.group(1), sm.group(2)
            style = re.search(r"style\s*=\s*['\"]([^'\"]+)['\"]", attrs, re.I)
            img = re.search(r'o:relid\s*=\s*["\'](rId\d+)["\']', body, re.I)
            if not style or not img:
                continue
            st = style.group(1)
            vals = {}
            for key in ("left", "top", "width", "height", "margin-left", "margin-top"):
                m = re.search(rf"{key}\s*:\s*([^;]+)", st, re.I)
                if m:
                    vals[key] = m.group(1).strip()
            left = vals.get("left") or vals.get("margin-left")
            top = vals.get("top") or vals.get("margin-top")
            width = vals.get("width")
            height = vals.get("height")
            if not all([left, top, width, height]):
                continue
            media_tgt = vml_rid_map.get(img.group(1), "")
            media = Path(media_tgt).name
            if not media:
                continue
            add(
                media,
                parse_vml_length(left),
                parse_vml_length(top),
                parse_vml_length(width),
                parse_vml_length(height),
                "vml",
            )

    def _xfrm(block: str):
        m = re.search(
            r"<a:xfrm[^>]*>\s*<a:off x=\"(-?\d+)\" y=\"(-?\d+)\"/>\s*<a:ext cx=\"(\d+)\" cy=\"(\d+)\"/>",
            block,
            re.S,
        )
        if not m:
            m = re.search(
                r'<a:off x="(-?\d+)" y="(-?\d+)"/>\s*<a:ext cx="(\d+)" cy="(\d+)"/>',
                block,
            )
        return tuple(map(int, m.groups())) if m else None

    def _grp_xfrm(grp_block: str):
        """Return (ox,oy,ecx,ecy,chx,chy,chcx,chcy) or None."""
        m = re.search(
            r"<p:grpSpPr[^>]*>.*?<a:xfrm[^>]*>\s*"
            r"<a:off x=\"(-?\d+)\" y=\"(-?\d+)\"/>\s*<a:ext cx=\"(\d+)\" cy=\"(\d+)\"/>\s*"
            r"<a:chOff x=\"(-?\d+)\" y=\"(-?\d+)\"/>\s*<a:chExt cx=\"(\d+)\" cy=\"(\d+)\"/>",
            grp_block,
            re.S,
        )
        return tuple(map(int, m.groups())) if m else None

    def map_child(gx, child):
        ox, oy, ecx, ecy, chx, chy, chcx, chcy = gx
        cx, cy, ccx, ccy = child
        if chcx == 0 or chcy == 0:
            return cx, cy, ccx, ccy
        x = ox + (cx - chx) * ecx / chcx
        y = oy + (cy - chy) * ecy / chcy
        w = ccx * ecx / chcx
        h = ccy * ecy / chcy
        return int(round(x)), int(round(y)), int(round(w)), int(round(h))

    # 2) Pictures (including inside groups — apply group transform)
    # Walk grpSp recursively via simple stack of group frames
    def walk_groups(xml: str, parent_gx=None):
        # Find top-level groups and pics in this fragment (non-greedy nesting is hard;
        # use iterative approach on grpSp blocks then leftover pics)
        pos = 0
        while True:
            gstart = xml.find("<p:grpSp>", pos)
            pstart = xml.find("<p:pic>", pos)
            if gstart < 0 and pstart < 0:
                break
            if gstart >= 0 and (pstart < 0 or gstart < pstart):
                # find matching close with depth
                depth = 1
                i = gstart + len("<p:grpSp>")
                while depth and i < len(xml):
                    if xml.startswith("<p:grpSp>", i):
                        depth += 1
                        i += len("<p:grpSp>")
                    elif xml.startswith("</p:grpSp>", i):
                        depth -= 1
                        i += len("</p:grpSp>")
                    else:
                        i += 1
                grp_block = xml[gstart:i]
                gx = _grp_xfrm(grp_block)
                # compose with parent
                if gx and parent_gx:
                    # map group origin/size through parent
                    x, y, w, h = map_child(parent_gx, (gx[0], gx[1], gx[2], gx[3]))
                    # child coords still in this group's ch space
                    gx = (x, y, w, h, gx[4], gx[5], gx[6], gx[7])
                elif gx is None:
                    gx = parent_gx
                inner = grp_block[len("<p:grpSp>") : -len("</p:grpSp>")]
                walk_groups(inner, gx)
                pos = i
            else:
                pend = xml.find("</p:pic>", pstart)
                if pend < 0:
                    break
                block = xml[pstart : pend + len("</p:pic>")]
                embed = re.search(r'r:embed="(rId\d+)"', block)
                child = _xfrm(block)
                pos = pend + len("</p:pic>")
                if not embed or not child:
                    continue
                tgt = rid_map.get(embed.group(1), "")
                if "/media/" not in tgt.replace("\\", "/"):
                    continue
                media = Path(tgt).name
                if parent_gx:
                    x, y, cx, cy = map_child(parent_gx, child)
                else:
                    x, y, cx, cy = child
                add(media, x, y, cx, cy, "pic")

    walk_groups(slide_xml, None)

    # 3) OLE graphicFrames with valid OOXML xfrm
    for m in re.finditer(r"<p:graphicFrame>(.*?)</p:graphicFrame>", slide_xml, re.S):
        block = m.group(1)
        off = re.search(
            r"<p:xfrm>\s*<a:off x=\"(-?\d+)\" y=\"(-?\d+)\"/>\s*<a:ext cx=\"(\d+)\" cy=\"(\d+)\"/>",
            block,
            re.S,
        )
        if not off:
            continue
        embed = re.search(r'r:embed="(rId\d+)"', block)
        if not embed:
            continue
        tgt = rid_map.get(embed.group(1), "")
        if "/media/" not in tgt.replace("\\", "/"):
            continue
        media = Path(tgt).name
        x, y, cx, cy = map(int, off.groups())
        if cx < MIN_EMU or cy < MIN_EMU:
            # broken nested preview; prefer imgW/imgH only if VML didn't place this media
            img = re.search(r'imgW="(\d+)" imgH="(\d+)"', block)
            already = any(p["media"] == media for p in placements)
            if img and not already:
                cx, cy = int(img.group(1)), int(img.group(2))
                # center in content area as last resort (should be rare if VML exists)
                x = max(0, (9144000 - cx) // 2)
                y = max(0, (6858000 - cy) // 2)
                add(media, x, y, cx, cy, "ole-imgWH")
            continue
        add(media, x, y, cx, cy, "ole")

    placements.sort(key=lambda p: (p["y"], p["x"]))
    return placements
