---
name: pptx-device-safe-export
description: Create cross-device safe PowerPoint files by flattening each PPTX slide into a full-slide raster image while preserving playable media such as videos and animated GIFs. Use when a user needs to present or send a .pptx on another computer and wants to avoid missing fonts, layout drift, broken SVG rendering, WPS/PowerPoint compatibility issues, or other device-dependent formatting problems while keeping videos playable.
---

# PPTX Device-Safe Export

## Overview

Use this skill to produce a presentation-ready `.pptx` that behaves like a PDF for layout stability but still keeps selected dynamic objects playable. The default workflow renders every slide to a PNG background, replaces the slide contents with that background, preserves video-like media shapes, normalizes residual font metadata, and validates the package.

## Quick Start

Run the bundled script:

```bash
python3 /path/to/pptx-device-safe-export/scripts/device_safe_export.py \
  input.pptx \
  --output input_pre.pptx
```

When a trusted PDF export already exists and matches the PPTX slide order, use it as the rendering source:

```bash
python3 /path/to/pptx-device-safe-export/scripts/device_safe_export.py \
  input.pptx \
  --pdf input.pdf \
  --output input_pre.pptx
```

## Workflow

1. Inspect the source deck for slide count, size, and media objects.
2. Render slide visuals to PNG backgrounds.
   - Prefer a user-provided `--pdf` when the user has already checked that the PDF looks correct.
   - Otherwise let the script call LibreOffice/OpenOffice `soffice` to export a temporary PDF.
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
- `soffice` only when no PDF is provided.
- `ffmpeg` or macOS `avconvert` only when converting `.mov` videos to `.mp4`.

If a required renderer is missing, stop and tell the user exactly which tool is missing and why. Do not silently produce an unverified deck.

## Verification

After creating the output:

- Run `unzip -t output.pptx`.
- Render or open the output deck and inspect a contact sheet if possible.
- Check the slides containing preserved media in PowerPoint/WPS/Keynote before submitting or presenting.
- Report the output path, slide count, preserved media count, conversion warnings, and any residual risk.

## Safety Notes

- Keep the original `.pptx` unchanged.
- Expect the output deck to be less editable because each slide is a screenshot-like background.
- Use `.mp4` for maximum cross-device video compatibility.
- If videos must autoplay, test the final deck in the actual presentation app because autoplay semantics vary across PowerPoint, WPS, and Keynote.
