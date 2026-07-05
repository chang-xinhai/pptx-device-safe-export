# PPTX Device-Safe Export

Turn a fragile presentation into a cross-device-safe `.pptx`.

This Codex skill flattens every slide into a full-slide image background while
preserving playable media such as videos and animated GIFs. It is designed for
the very real moment when you need to present on another computer and do not
want missing fonts, SVG quirks, WPS/PowerPoint differences, or layout drift to
ruin the deck.

The important design choice: **the visual source must come from the app whose
appearance you trust**. If your deck looks right in WPS, export a PDF from WPS
and use that PDF. LibreOffice is only a draft fallback because it can render
PPTX differently.

## What It Does

- Uses a trusted PDF or per-slide image export as the layout truth.
- Rebuilds the deck with one full-slide image per slide.
- Keeps video-like media objects above the background so they can still play.
- Attempts to convert `.mov` media to `.mp4` for better compatibility.
- Normalizes residual font metadata to reduce missing-font warnings.
- Leaves the original `.pptx` untouched.

## Quick Use

Use a PDF export that already looks correct:

```bash
python3 ~/.codex/skills/pptx-device-safe-export/scripts/device_safe_export.py \
  input.pptx \
  --pdf input.pdf \
  --output input_pre.pptx
```

Or use one trusted slide image per page:

```bash
python3 ~/.codex/skills/pptx-device-safe-export/scripts/device_safe_export.py \
  input.pptx \
  --background-dir rendered_slides/ \
  --output input_pre.pptx
```

For draft-only conversion with LibreOffice/OpenOffice:

```bash
python3 ~/.codex/skills/pptx-device-safe-export/scripts/device_safe_export.py \
  input.pptx \
  --allow-soffice-renderer \
  --output input_pre_draft.pptx
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
- LibreOffice/OpenOffice `soffice` only for explicit draft fallback
- `ffmpeg` or macOS `avconvert` for optional `.mov` to `.mp4` conversion

On macOS:

```bash
brew install poppler
```

## Good Workflow

1. Open the original deck in the presentation app whose appearance you trust.
2. Export it to PDF and visually check the PDF.
3. Run this tool with `--pdf`.
4. Open the output `.pptx` on the target presentation machine.
5. Test the slides containing videos.

## Tradeoffs

The output deck is intentionally less editable: most slide content becomes a
screenshot-like background. That is the point. It trades editability for
presentation stability while keeping media playback where possible.

This tool does not try to make LibreOffice render like WPS or PowerPoint. It
uses a trusted render as input and then rebuilds a safer PPTX around it.

## Example

```bash
python3 scripts/device_safe_export.py talk.pptx \
  --pdf talk.pdf \
  --output talk_pre.pptx \
  --report talk_pre_report.json
```
