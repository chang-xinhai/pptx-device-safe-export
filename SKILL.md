---
name: pptx-device-safe-export
description: Create cross-device safe PowerPoint files by flattening each PPTX slide into a full-slide raster image from a trusted visual source while preserving playable media such as videos and animated GIFs. Use when a user needs to present or send a .pptx on another computer and wants to avoid missing fonts, layout drift, broken SVG rendering, WPS/PowerPoint compatibility issues, or other device-dependent formatting problems while keeping videos playable. Requires a trusted PDF export or per-slide images when exact WPS/PowerPoint/Keynote appearance matters.
---

# PPTX Device-Safe Export

## Overview

Use this skill to produce a presentation-ready `.pptx` that behaves like a PDF for layout stability but still keeps selected dynamic objects playable. The key rule is: **flatten from the same renderer the user trusts**. If the user says the deck looks correct in WPS, use a WPS-exported PDF or WPS-exported slide images as the visual source; do not silently use LibreOffice as the layout source.

## Quick Start

Use a trusted PDF export that already looks correct:

```bash
python3 /path/to/pptx-device-safe-export/scripts/device_safe_export.py \
  input.pptx \
  --pdf input.pdf \
  --output input_pre.pptx
```

Or use a folder of trusted per-slide PNG/JPG backgrounds:

```bash
python3 /path/to/pptx-device-safe-export/scripts/device_safe_export.py \
  input.pptx \
  --background-dir rendered_slides/ \
  --output input_pre.pptx
```

## Workflow

1. Inspect the source deck for slide count, size, and media objects.
2. Obtain a trusted visual source.
   - Best: ask the user for a PDF exported from WPS/PowerPoint/Keynote after they confirm it looks correct.
   - Also acceptable: use `--background-dir` with one PNG/JPG per slide exported from the trusted app.
   - Fallback only: use `--allow-soffice-renderer` when the user accepts that LibreOffice/OpenOffice may render differently from WPS/PowerPoint.
3. Replace each slide's visible contents with one full-slide PNG.
4. Preserve playable media shapes above the background.
   - Videos are preserved by default.
   - Animated GIFs are preserved by default.
   - Static SVGs usually should not be preserved separately because they are already baked into the background. Preserve them only when the user explicitly needs live SVG objects.
5. Normalize residual font metadata to `Arial` unless the user asks not to.
6. Validate by opening or rendering the output and checking media-heavy slides.

## Tool Requirements

The script is mostly standard-library Python. It needs external tools for rendering:

- `pdftoppm` from Poppler to rasterize PDF pages.
- `soffice` only when `--allow-soffice-renderer` is explicitly used.
- `ffmpeg` or macOS `avconvert` only when converting `.mov` videos to `.mp4`.

If no trusted visual source is available, stop and ask for a PDF or per-slide images. Do not silently produce an unverified deck with a different renderer.

## Verification

After creating the output:

- Run `unzip -t output.pptx`.
- Render or open the output deck and inspect a contact sheet if possible.
- Check the slides containing preserved media in PowerPoint/WPS/Keynote before submitting or presenting.
- Report the output path, slide count, `render_source_type`, preserved media count, conversion warnings, and any residual risk.

## Safety Notes

- Keep the original `.pptx` unchanged.
- Expect the output deck to be less editable because each slide is a screenshot-like background.
- Use `.mp4` for maximum cross-device video compatibility.
- If videos must autoplay, test the final deck in the actual presentation app because autoplay semantics vary across PowerPoint, WPS, and Keynote.
- LibreOffice/OpenOffice is not a reliable proxy for WPS/PowerPoint visual fidelity. Treat it as a draft renderer only.
