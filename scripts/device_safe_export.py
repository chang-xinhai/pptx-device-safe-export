#!/usr/bin/env python3
"""Create a cross-device-safe PPTX by flattening slide visuals to images.

The script keeps the original source deck unchanged. It creates an output PPTX
where each slide contains a full-slide PNG background plus selected preserved
media objects, typically videos.
"""

from __future__ import annotations

import argparse
import json
import os
import posixpath
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple
from zipfile import ZIP_DEFLATED, ZipFile
from xml.etree import ElementTree as ET


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "p15": "http://schemas.microsoft.com/office/powerpoint/2012/main",
}

for prefix, uri in NS.items():
    if prefix == "rel":
        ET.register_namespace("", uri)
    elif prefix == "ct":
        ET.register_namespace("", uri)
    else:
        ET.register_namespace(prefix, uri)

P = f"{{{NS['p']}}}"
A = f"{{{NS['a']}}}"
R = f"{{{NS['r']}}}"
REL = f"{{{NS['rel']}}}"
CT = f"{{{NS['ct']}}}"

VIDEO_EXTS = {
    ".mp4",
    ".mov",
    ".m4v",
    ".avi",
    ".wmv",
    ".mpeg",
    ".mpg",
    ".webm",
}
AUDIO_EXTS = {".mp3", ".m4a", ".wav", ".aac", ".aiff", ".aif"}
DEFAULT_PRESERVE_EXTS = {".gif"}
MEDIA_REL_MARKERS = ("/video", "/audio", "/media")
BACKGROUND_EXTS = {".png", ".jpg", ".jpeg"}
IMAGE_CONTENT_TYPES = {
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
}


class ExportError(RuntimeError):
    pass


def run(cmd: Sequence[str], *, cwd: Optional[Path] = None) -> subprocess.CompletedProcess[str]:
    proc = subprocess.run(
        list(cmd),
        cwd=str(cwd) if cwd else None,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
    )
    if proc.returncode != 0:
        raise ExportError(f"Command failed ({proc.returncode}): {' '.join(cmd)}\n{proc.stdout}")
    return proc


def natural_key(path: Path) -> Tuple[int, str]:
    match = re.search(r"-(\d+)\.png$", path.name)
    if match:
        return int(match.group(1)), path.name
    return 10**9, path.name


def which(name: str, explicit: Optional[str] = None) -> Optional[str]:
    if explicit:
        return explicit
    return shutil.which(name)


def export_pdf_with_soffice(pptx: Path, out_dir: Path, soffice: Optional[str]) -> Path:
    exe = which("soffice", soffice)
    if not exe:
        raise ExportError("Missing soffice. Provide --pdf or install LibreOffice/OpenOffice.")
    run([exe, "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(pptx)])
    pdfs = sorted(out_dir.glob("*.pdf"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not pdfs:
        raise ExportError("soffice completed but no PDF was created.")
    return pdfs[0]


def rasterize_pdf(pdf: Path, out_dir: Path, dpi: int, pdftoppm: Optional[str]) -> List[Path]:
    exe = which("pdftoppm", pdftoppm)
    if not exe:
        raise ExportError("Missing pdftoppm from Poppler. Install Poppler or provide --pdftoppm.")
    prefix = out_dir / "slide"
    run([exe, "-png", "-r", str(dpi), str(pdf), str(prefix)])
    pages = sorted(out_dir.glob("slide-*.png"), key=natural_key)
    if not pages:
        raise ExportError("pdftoppm completed but no PNG pages were created.")
    return pages


def collect_background_images(background_dir: Path) -> List[Path]:
    if not background_dir.exists() or not background_dir.is_dir():
        raise ExportError(f"Background image directory does not exist: {background_dir}")
    pages = sorted(
        [
            p
            for p in background_dir.iterdir()
            if p.is_file() and p.suffix.lower() in BACKGROUND_EXTS
        ],
        key=natural_key,
    )
    if not pages:
        raise ExportError(
            f"No background images found in {background_dir}. Use PNG/JPG files, one per slide."
        )
    return pages


def rels_path_for(part_name: str) -> str:
    directory, filename = posixpath.split(part_name)
    return posixpath.join(directory, "_rels", filename + ".rels")


def load_rels(zf: ZipFile, rels_path: str) -> Tuple[ET.Element, Dict[str, ET.Element], int]:
    if rels_path in zf.namelist():
        root = ET.fromstring(zf.read(rels_path))
    else:
        root = ET.Element(REL + "Relationships")
    rels: Dict[str, ET.Element] = {}
    max_id = 0
    for rel in list(root):
        rid = rel.attrib.get("Id", "")
        rels[rid] = rel
        if rid.startswith("rId"):
            try:
                max_id = max(max_id, int(rid[3:]))
            except ValueError:
                pass
    return root, rels, max_id


def next_rid(existing: Dict[str, ET.Element], max_id: int) -> str:
    candidate = max_id + 1
    while f"rId{candidate}" in existing:
        candidate += 1
    return f"rId{candidate}"


def normalize_target(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    base = posixpath.dirname(source_part)
    return posixpath.normpath(posixpath.join(base, target))


def relative_media_target(media_name: str) -> str:
    return f"../media/{media_name}"


def slide_order(zf: ZipFile) -> List[str]:
    pres = ET.fromstring(zf.read("ppt/presentation.xml"))
    pres_rels_root, pres_rels, _ = load_rels(zf, "ppt/_rels/presentation.xml.rels")
    _ = pres_rels_root
    ordered: List[str] = []
    for sld_id in pres.findall(".//p:sldId", NS):
        rid = sld_id.attrib.get(R + "id")
        if not rid or rid not in pres_rels:
            continue
        target = pres_rels[rid].attrib.get("Target", "")
        part = normalize_target("ppt/presentation.xml", target)
        ordered.append(part)
    if not ordered:
        ordered = sorted(
            [
                n
                for n in zf.namelist()
                if n.startswith("ppt/slides/slide")
                and n.endswith(".xml")
                and "/_rels/" not in n
            ]
        )
    return ordered


def slide_size(zf: ZipFile) -> Tuple[int, int]:
    pres = ET.fromstring(zf.read("ppt/presentation.xml"))
    sld_sz = pres.find("p:sldSz", NS)
    if sld_sz is None:
        return 12192000, 6858000
    return int(sld_sz.attrib["cx"]), int(sld_sz.attrib["cy"])


def max_cnvpr_id(root: ET.Element) -> int:
    max_id = 0
    for elem in root.iter(P + "cNvPr"):
        try:
            max_id = max(max_id, int(elem.attrib.get("id", "0")))
        except ValueError:
            continue
    return max_id


def elem_rids(elem: ET.Element) -> Iterable[str]:
    for node in elem.iter():
        for key, value in node.attrib.items():
            if key in {R + "embed", R + "link", R + "id"} and value.startswith("rId"):
                yield value


def is_preserved_shape(
    shape: ET.Element,
    slide_part: str,
    rels: Dict[str, ET.Element],
    preserve_exts: set[str],
    preserve_audio: bool,
) -> bool:
    if shape.tag != P + "pic":
        return False
    has_media_marker = False
    for node in shape.iter():
        local = node.tag.split("}")[-1]
        if local in {"videoFile", "audioFile"}:
            has_media_marker = True
            break
    for rid in elem_rids(shape):
        rel = rels.get(rid)
        if rel is None:
            continue
        target = rel.attrib.get("Target", "")
        rel_type = rel.attrib.get("Type", "")
        ext = Path(target).suffix.lower()
        is_media_rel = any(marker in rel_type.lower() for marker in MEDIA_REL_MARKERS)
        if ext in VIDEO_EXTS:
            return True
        if preserve_audio and ext in AUDIO_EXTS:
            return True
        if ext in preserve_exts:
            return True
        if has_media_marker and is_media_rel:
            return True
    return False


def make_background_pic(rid: str, pic_id: int, slide_w: int, slide_h: int, slide_no: int) -> ET.Element:
    pic = ET.Element(P + "pic")
    nv_pic_pr = ET.SubElement(pic, P + "nvPicPr")
    ET.SubElement(
        nv_pic_pr,
        P + "cNvPr",
        {"id": str(pic_id), "name": f"Flattened slide {slide_no}"},
    )
    c_nv_pic_pr = ET.SubElement(nv_pic_pr, P + "cNvPicPr")
    ET.SubElement(
        c_nv_pic_pr,
        A + "picLocks",
        {"noChangeAspect": "1", "noMove": "1", "noResize": "1"},
    )
    ET.SubElement(nv_pic_pr, P + "nvPr")
    blip_fill = ET.SubElement(pic, P + "blipFill")
    ET.SubElement(blip_fill, A + "blip", {R + "embed": rid})
    stretch = ET.SubElement(blip_fill, A + "stretch")
    ET.SubElement(stretch, A + "fillRect")
    sp_pr = ET.SubElement(pic, P + "spPr")
    xfrm = ET.SubElement(sp_pr, A + "xfrm")
    ET.SubElement(xfrm, A + "off", {"x": "0", "y": "0"})
    ET.SubElement(xfrm, A + "ext", {"cx": str(slide_w), "cy": str(slide_h)})
    prst_geom = ET.SubElement(sp_pr, A + "prstGeom", {"prst": "rect"})
    ET.SubElement(prst_geom, A + "avLst")
    return pic


def add_content_type_defaults(root: ET.Element, defaults: Dict[str, str]) -> None:
    existing = {
        node.attrib.get("Extension", "").lower()
        for node in root.findall("ct:Default", NS)
    }
    for ext, content_type in defaults.items():
        if ext.lower() not in existing:
            ET.SubElement(root, CT + "Default", {"Extension": ext, "ContentType": content_type})


def available_converter() -> Optional[str]:
    if shutil.which("ffmpeg"):
        return "ffmpeg"
    if shutil.which("avconvert"):
        return "avconvert"
    return None


def convert_mov_to_mp4(zf: ZipFile, media_part: str, tmp_dir: Path) -> Tuple[Optional[Path], Optional[str]]:
    converter = available_converter()
    if not converter:
        return None, "No ffmpeg or avconvert found; kept MOV media."
    src = tmp_dir / Path(media_part).name
    dst = tmp_dir / (Path(media_part).stem + ".mp4")
    src.write_bytes(zf.read(media_part))
    try:
        if converter == "ffmpeg":
            run(["ffmpeg", "-y", "-i", str(src), "-c", "copy", str(dst)])
        else:
            run([
                "avconvert",
                "--source",
                str(src),
                "--preset",
                "PresetPassthrough",
                "--output",
                str(dst),
                "--replace",
            ])
    except ExportError as exc:
        return None, f"MOV to MP4 conversion failed for {media_part}: {exc}"
    if not dst.exists() or dst.stat().st_size == 0:
        return None, f"MOV to MP4 conversion produced no file for {media_part}."
    return dst, None


def normalize_font_metadata(xml_bytes: bytes, font: str) -> bytes:
    text = xml_bytes.decode("utf-8", errors="ignore")
    text = re.sub(r'typeface="[^"]*"', f'typeface="{font}"', text)
    return text.encode("utf-8")


def validate_pptx(path: Path) -> None:
    with ZipFile(path) as zf:
        bad = zf.testzip()
    if bad:
        raise ExportError(f"Output PPTX is corrupt at entry: {bad}")


def build_device_safe_deck(args: argparse.Namespace) -> Dict[str, object]:
    input_pptx = Path(args.input).expanduser().resolve()
    if not input_pptx.exists():
        raise ExportError(f"Input PPTX does not exist: {input_pptx}")
    output = Path(args.output).expanduser().resolve()
    output.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="pptx-device-safe-") as temp_name:
        temp_dir = Path(temp_name)
        render_dir = temp_dir / "render"
        render_dir.mkdir()
        pdf_dir = temp_dir / "pdf"
        pdf_dir.mkdir()
        media_tmp = temp_dir / "media"
        media_tmp.mkdir()

        render_warnings: List[str] = []
        if args.pdf and args.background_dir:
            raise ExportError("Use either --pdf or --background-dir, not both.")
        if args.pdf:
            pdf = Path(args.pdf).expanduser().resolve()
            if not pdf.exists():
                raise ExportError(f"PDF render source does not exist: {pdf}")
            pages = rasterize_pdf(pdf, render_dir, args.dpi, args.pdftoppm)
            render_source_type = "trusted_pdf"
            render_source = str(pdf)
        elif args.background_dir:
            background_dir = Path(args.background_dir).expanduser().resolve()
            pages = collect_background_images(background_dir)
            render_source_type = "trusted_images"
            render_source = str(background_dir)
            pdf = None
        elif args.allow_soffice_renderer:
            pdf = export_pdf_with_soffice(input_pptx, pdf_dir, args.soffice)
            pages = rasterize_pdf(pdf, render_dir, args.dpi, args.pdftoppm)
            render_source_type = "untrusted_soffice"
            render_source = str(pdf)
            render_warnings.append(
                "Used LibreOffice/OpenOffice soffice to render the PPTX. "
                "This may not match WPS, PowerPoint, or Keynote rendering."
            )
        else:
            raise ExportError(
                "No trusted visual source was provided. Export the deck to PDF from the "
                "presentation app whose appearance you trust, then pass --pdf; or pass "
                "--background-dir with one PNG/JPG per slide. Use --allow-soffice-renderer "
                "only for a draft when you accept possible layout drift."
            )

        preserve_exts = {
            ext.lower() if ext.startswith(".") else "." + ext.lower()
            for ext in args.preserve_ext
        }

        modified: Dict[str, bytes] = {}
        added_files: Dict[str, Path] = {}
        converted_media: Dict[str, str] = {}
        conversion_warnings: List[str] = list(render_warnings)
        preserved_report: List[Dict[str, object]] = []
        slide_reports: List[Dict[str, object]] = []

        with ZipFile(input_pptx, "r") as zin:
            slides = slide_order(zin)
            if len(pages) != len(slides):
                raise ExportError(
                    f"Rendered page count ({len(pages)}) does not match PPTX slide count ({len(slides)})."
                )
            slide_w, slide_h = slide_size(zin)

            for slide_index, slide_part in enumerate(slides, start=1):
                rels_part = rels_path_for(slide_part)
                rels_root, rels, max_rid = load_rels(zin, rels_part)
                slide_root = ET.fromstring(zin.read(slide_part))
                c_sld = slide_root.find("p:cSld", NS)
                if c_sld is None:
                    raise ExportError(f"Slide has no cSld: {slide_part}")
                sp_tree = c_sld.find("p:spTree", NS)
                if sp_tree is None:
                    raise ExportError(f"Slide has no spTree: {slide_part}")

                required_group = list(sp_tree)[:2]
                preserved_shapes: List[ET.Element] = []
                for child in list(sp_tree)[2:]:
                    if is_preserved_shape(child, slide_part, rels, preserve_exts, args.preserve_audio):
                        preserved_shapes.append(child)

                if not preserved_shapes and args.drop_timing_without_media:
                    timing = slide_root.find("p:timing", NS)
                    if timing is not None:
                        slide_root.remove(timing)

                bg_ext = pages[slide_index - 1].suffix.lower().lstrip(".")
                if bg_ext == "jpg":
                    bg_ext = "jpeg"
                bg_name = f"device_safe_slide_{slide_index:02d}.{bg_ext}"
                bg_arc = f"ppt/media/{bg_name}"
                added_files[bg_arc] = pages[slide_index - 1]
                bg_rid = next_rid(rels, max_rid)
                ET.SubElement(
                    rels_root,
                    REL + "Relationship",
                    {
                        "Id": bg_rid,
                        "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                        "Target": relative_media_target(bg_name),
                    },
                )
                rels[bg_rid] = list(rels_root)[-1]
                bg_pic = make_background_pic(
                    bg_rid,
                    max_cnvpr_id(slide_root) + 1000,
                    slide_w,
                    slide_h,
                    slide_index,
                )

                if args.convert_mov_to_mp4:
                    for shape in preserved_shapes:
                        for rid in elem_rids(shape):
                            rel = rels.get(rid)
                            if rel is None:
                                continue
                            target = rel.attrib.get("Target", "")
                            if Path(target).suffix.lower() != ".mov":
                                continue
                            old_part = normalize_target(slide_part, target)
                            if old_part not in converted_media:
                                converted, warning = convert_mov_to_mp4(zin, old_part, media_tmp)
                                if warning:
                                    conversion_warnings.append(warning)
                                    continue
                                new_name = Path(old_part).with_suffix(".mp4").name
                                new_part = f"ppt/media/{new_name}"
                                suffix = 2
                                while new_part in zin.namelist() or new_part in added_files:
                                    new_name = f"{Path(old_part).stem}_{suffix}.mp4"
                                    new_part = f"ppt/media/{new_name}"
                                    suffix += 1
                                added_files[new_part] = converted
                                converted_media[old_part] = new_part
                            if old_part in converted_media:
                                rel.set("Target", relative_media_target(Path(converted_media[old_part]).name))

                new_sp_tree = ET.Element(P + "spTree")
                for node in required_group:
                    new_sp_tree.append(node)
                new_sp_tree.append(bg_pic)
                for node in preserved_shapes:
                    new_sp_tree.append(node)
                c_sld.remove(sp_tree)
                c_sld.append(new_sp_tree)

                modified[slide_part] = ET.tostring(slide_root, encoding="utf-8", xml_declaration=True)
                modified[rels_part] = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)

                media_targets: List[str] = []
                for shape in preserved_shapes:
                    for rid in elem_rids(shape):
                        rel = rels.get(rid)
                        if rel is not None:
                            target = rel.attrib.get("Target", "")
                            if Path(target).suffix.lower() in VIDEO_EXTS | AUDIO_EXTS | preserve_exts:
                                media_targets.append(target)
                if preserved_shapes:
                    preserved_report.append(
                        {
                            "slide": slide_index,
                            "shape_count": len(preserved_shapes),
                            "targets": sorted(set(media_targets)),
                        }
                    )
                slide_reports.append(
                    {
                        "slide": slide_index,
                        "background": bg_name,
                        "preserved_shapes": len(preserved_shapes),
                    }
                )

            if "[Content_Types].xml" in zin.namelist():
                ct_root = ET.fromstring(zin.read("[Content_Types].xml"))
                add_content_type_defaults(
                    ct_root,
                    {
                        **{
                            ext: IMAGE_CONTENT_TYPES[ext]
                            for ext in sorted(
                                {
                                    ("jpeg" if page.suffix.lower() == ".jpg" else page.suffix.lower().lstrip("."))
                                    for page in pages
                                }
                            )
                        },
                        "mp4": "video/mp4",
                    },
                )
                modified["[Content_Types].xml"] = ET.tostring(
                    ct_root, encoding="utf-8", xml_declaration=True
                )

            converted_sources = set(converted_media.keys())
            if output.exists():
                output.unlink()
            with ZipFile(output, "w", ZIP_DEFLATED) as zout:
                seen = set()
                for item in zin.infolist():
                    if item.filename in seen:
                        continue
                    seen.add(item.filename)
                    if args.remove_converted_sources and item.filename in converted_sources:
                        continue
                    if item.filename in added_files:
                        continue
                    data = modified.get(item.filename)
                    if data is None:
                        data = zin.read(item.filename)
                        if args.normalize_fonts and item.filename.endswith((".xml", ".rels")):
                            data = normalize_font_metadata(data, args.font)
                    zout.writestr(item, data)
                for arc, source in added_files.items():
                    zout.write(source, arc)

        validate_pptx(output)
        report = {
            "input": str(input_pptx),
            "output": str(output),
            "render_source_type": render_source_type,
            "render_source": render_source,
            "pdf_source": str(pdf) if pdf else None,
            "slide_count": len(pages),
            "dpi": args.dpi,
            "preserved_media": preserved_report,
            "converted_media": converted_media,
            "warnings": conversion_warnings,
            "slides": slide_reports,
            "size_bytes": output.stat().st_size,
        }
        if args.report:
            report_path = Path(args.report).expanduser().resolve()
            report_path.parent.mkdir(parents=True, exist_ok=True)
            report_path.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")
        return report


def parse_args(argv: Sequence[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Flatten PPTX slides into full-slide images while preserving playable media."
    )
    parser.add_argument("input", help="Source .pptx file")
    parser.add_argument("--output", required=True, help="Output .pptx file")
    parser.add_argument(
        "--pdf",
        help="Trusted PDF render of the deck, exported from the presentation app whose appearance should be preserved",
    )
    parser.add_argument(
        "--background-dir",
        help="Trusted directory containing one PNG/JPG background image per slide, in slide order",
    )
    parser.add_argument("--report", help="Optional JSON report path")
    parser.add_argument("--dpi", type=int, default=144, help="Rasterization DPI for slide backgrounds")
    parser.add_argument(
        "--allow-soffice-renderer",
        action="store_true",
        help="Allow LibreOffice/OpenOffice to render PPTX to PDF. This is a fallback and may not match WPS/PowerPoint.",
    )
    parser.add_argument("--soffice", help="Path to soffice when --allow-soffice-renderer is used")
    parser.add_argument("--pdftoppm", help="Path to pdftoppm")
    parser.add_argument(
        "--preserve-ext",
        action="append",
        default=sorted(DEFAULT_PRESERVE_EXTS),
        help="Additional picture/media extension to preserve above the background. Repeatable. Default: .gif",
    )
    parser.add_argument("--preserve-audio", action="store_true", help="Preserve audio media shapes")
    parser.add_argument(
        "--no-convert-mov-to-mp4",
        dest="convert_mov_to_mp4",
        action="store_false",
        help="Keep .mov media as-is instead of attempting MP4 passthrough conversion",
    )
    parser.set_defaults(convert_mov_to_mp4=True)
    parser.add_argument(
        "--keep-converted-sources",
        dest="remove_converted_sources",
        action="store_false",
        help="Keep original MOV files after successful MP4 conversion",
    )
    parser.set_defaults(remove_converted_sources=True)
    parser.add_argument(
        "--no-normalize-fonts",
        dest="normalize_fonts",
        action="store_false",
        help="Do not normalize residual XML font metadata",
    )
    parser.set_defaults(normalize_fonts=True)
    parser.add_argument("--font", default="Arial", help="Font name for residual metadata normalization")
    parser.add_argument(
        "--keep-timing-without-media",
        dest="drop_timing_without_media",
        action="store_false",
        help="Keep slide timing data even on slides with no preserved media",
    )
    parser.set_defaults(drop_timing_without_media=True)
    return parser.parse_args(argv)


def main(argv: Sequence[str]) -> int:
    args = parse_args(argv)
    try:
        report = build_device_safe_deck(args)
    except ExportError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    print(json.dumps(report, indent=2, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
