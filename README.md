B71 template

# Block71 Startup Profile – PowerPoint Template

A Canva-compatible 1920×1080 PowerPoint template for the Block71/Program-style Startup Profile presentation.

## Deliverables

| File | Description |
|------|-------------|
| `templates/Startup_Profile_Template_1920x1080_CanvaReady.pptx` | Ready-to-use PPTX (13 slides) |
| `templates/generate.js` | Node.js generation script (PptxGenJS) |
| `templates/package.json` | Node dependencies |

## Slides included

1. **Cover** – Title, subtitle (Block71 Singapore), location/date/time placeholders
2. **Startup Index** – Two-column list of Startup 1–10
3–12. **Startup Profile** × 10 – Per-startup slide with website, tagline, Problem/Solution bullets, and quick-facts chips (Industry, Stage, HQ, Team Size, Ask)
13. **Closing** – "Thank You" + contact placeholders

## Design tokens

| Token | Value |
|-------|-------|
| Orange accent | `#F57C00` |
| Navy | `#0B2E5E` |
| Dark text | `#1E1E1E` |
| Light gray placeholder | `#EEF1F5` |
| Fonts | Montserrat (titles) / Arial (body) |
| Slide size | 1920 × 1080 px (16:9) |

## How to regenerate

1. Install Node.js (≥ 18) if not already available.
2. Install dependencies:
   ```bash
   cd templates
   npm install
   ```
3. Run the generator:
   ```bash
   npm run generate
   # or: node generate.js
   ```
4. The file `templates/Startup_Profile_Template_1920x1080_CanvaReady.pptx` will be overwritten.

## Using in Canva

1. Go to [canva.com](https://canva.com) → **Create a design → Import file**.
2. Upload `Startup_Profile_Template_1920x1080_CanvaReady.pptx`.
3. All text boxes and shapes remain editable.
4. Replace "MESH IMAGE" rectangles with your chosen tech-mesh image or illustration.
5. Update the orange "DEMO 71" badge text as needed.
