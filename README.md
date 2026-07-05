# PPTX Device-Safe Export

Turn a fragile presentation into a cross-device-safe `.pptx`.

This Codex skill flattens every slide into a full-slide PNG background while
preserving playable media such as videos and animated GIFs. It is designed for
the very real moment when you need to present on another computer and do not
want missing fonts, SVG quirks, WPS/PowerPoint differences, or layout drift to
ruin the deck.

## What It Does

- Renders each slide to a high-resolution image.
- Rebuilds the deck with one full-slide image per slide.
- Keeps video-like media objects above the background so they can still play.
- Attempts to convert `.mov` media to `.mp4` for better compatibility.
- Normalizes residual font metadata to reduce missing-font warnings.
- Leaves the original `.pptx` untouched.

## Quick Use

```bash
python3 ~/.codex/skills/pptx-device-safe-export/scripts/device_safe_export.py \
  input.pptx \
  --output input_pre.pptx
```

If you already have a PDF export that looks correct, use it as the visual
source. This is the most reliable path because the PDF becomes the layout truth:

```bash
python3 ~/.codex/skills/pptx-device-safe-export/scripts/device_safe_export.py \
  input.pptx \
  --pdf input.pdf \
  --output input_pre.pptx
```

## Install As A Codex Skill

```bash
git clone https://github.com/chang-xinhai/pptx-device-safe-export.git \
  ~/.codex/skills/pptx-device-safe-export
```

Restart Codex, then invoke:

```text
Use $pptx-device-safe-export to create a cross-device safe PPTX.
```

## Requirements

- Python 3.9+
- Poppler `pdftoppm` for PDF-to-PNG rendering
- LibreOffice/OpenOffice `soffice` when no `--pdf` is provided
- `ffmpeg` or macOS `avconvert` for optional `.mov` to `.mp4` conversion

On macOS:

```bash
brew install poppler
```

## Good Workflow

1. Export the original deck to PDF and visually check it.
2. Run this tool with `--pdf`.
3. Open the output `.pptx` on the target presentation machine.
4. Test the slides containing videos.

## Tradeoffs

The output deck is intentionally less editable: most slide content becomes a
screenshot-like background. That is the point. It trades editability for
presentation stability while keeping media playback where possible.

## Example

```bash
python3 scripts/device_safe_export.py talk.pptx \
  --pdf talk.pdf \
  --output talk_pre.pptx \
  --report talk_pre_report.json
```

