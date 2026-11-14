# Restyle Next Office Slides to the new design system

You are restyling **Next Office Slides**: a PowerPoint-style editor built with React + TypeScript, shipped in the browser and as a Tauri Windows 11 desktop app. Styling is **plain hand-written CSS with custom properties**. Do not add Tailwind, CSS-in-JS, UI kits or other styling dependencies.

The goal is a full visual restyle of the app's chrome plus Windows branding. **Behaviour must not change.** Slide rendering (PPTX themes, the 8 built-in themes, gradients, transitions) must not change either.

The brand pack sits next to this file in `next-office-brand/`:

- `tokens.json` — the source of truth for every token (light + dark values)
- `fonts/PlusJakartaSans-Variable.woff2` — UI and brand font (weights 200–800)
- `logos/*.svg` — the wordmarks and the app icon (already outlined; never retype the logo in a font)
- `reference-components.css` — reference CSS for Button, Ribbon and SlideThumbnail. Copy the patterns, not the class names.

Copy the pack into the repo first, e.g. `src/assets/brand/` (or wherever the project keeps static assets).

---

## 0. Ground rules

1. **Explore first, then plan.** Before editing anything:
   - Map every CSS file, every place a colour, font, radius, shadow or spacing is hard-coded, and every existing custom property.
   - Map the UI surfaces listed in §4.
   - Write a short plan in `docs/restyle-plan.md`: the files you'll touch, the old→new token mapping, and the risks. Then implement.
2. **Separate the two worlds.**
   - **App chrome** (ribbon, panels, dialogs, rail, status bar, start screen) uses the design tokens.
   - **Slide content** (anything rendered inside a slide, thumbnails' content, slide-show output, PPTX/PDF export) keeps using the presentation theme.
   - Never let UI tokens leak into slide rendering or export. Never let the dark theme change slide colours.
3. **Tokens only.** After the restyle, chrome CSS has no raw hex/rgb values, font stacks or magic radii/shadows. Everything goes through `var(--…)`. The only exceptions are the token file itself and slide-content code.
4. **Work in small commits**, one per step below. Run the type-check, lint, tests and both builds (`web` + `tauri build`, or the project's equivalents) after each step. Fix what you break.
5. Ask me before deleting files, renaming public components, changing the file format or adding dependencies.

---

## 1. Tokens → `src/styles/tokens.css`

Generate `tokens.css` from `tokens.json` (a tiny Node script in `scripts/build-tokens.mjs` is fine; commit both the script and the output). Shape:

```css
:root, [data-theme="light"] { /* colour + shadow tokens, light values */ }
[data-theme="dark"] { /* colour + shadow tokens, dark values */ }
:root { /* spacing, radius, size, font families */ }
@font-face { font-family: "Plus Jakarta Sans"; src: url(...woff2) format("woff2"); font-weight: 200 800; font-display: swap; }
```

`{brand}`-style values in tokens.json are aliases, so emit them as `var(--brand)`. For reference, the values are:

**Colour (light / dark)**

| token | light | dark | use |
|---|---|---|---|
| `--surface` | #ffffff | #1b1b26 | panels, dialogs, ribbon body, inputs |
| `--surface-100` | #f6f6fb | #14141d | app background, title bar, tab row, slide rail |
| `--surface-200` | #ececf4 | #242432 | hover, quiet fills, table headers |
| `--surface-300` | #e0e0ec | #2e2e3e | pressed; stage backdrop behind the slide |
| `--border` | #d3d3e0 | #383849 | decorative hairlines |
| `--border-strong` | #8a8aa3 | #6e6e88 | input/control borders (3:1) |
| `--ink` | #16152b | #ededf6 | primary text/icons |
| `--ink-muted` | #55556e | #a6a6bd | secondary text, group labels, slide numbers |
| `--ink-disabled` | #9a9ab0 | #5e5e75 | disabled only |
| `--brand` | #4338ca | #8f8cf7 | primary buttons, File button, active tab underline, selected thumbnail |
| `--brand-hover` | #3730a3 | #a9a6ff | hover/pressed of brand fills |
| `--brand-subtle` | #ecebfd | #26244d | pressed toggles, selected rows |
| `--on-brand` | #ffffff | #12103a | text on brand fills |
| `--brand-ink` | #3730a3 | #b9b7ff | brand-coloured text, links, active tab label |
| `--focus-ring` | var(--brand) | var(--brand) | 2px outline, 2px offset |
| `--guide` | #d9480f | #d9480f | smart guides + distance labels (same in both themes) |
| `--selection` | #4338ca | #4338ca | selection box + handles (same in both themes) |
| `--success` | #1d7a4b | #4cc38a | saved / export done |
| `--warning` | #a15500 | #f0a54a | unsaved-changes dot, compatibility notices |
| `--danger` | #c42b2b | #ff7070 | destructive actions, errors |
| `--slide-bg` | #ffffff | #ffffff | default slide paper (never darkens) |
| `--overlay` | rgba(12,11,30,.5) | rgba(0,0,0,.6) | modal scrim |

**Type** — `--font-sans: "Plus Jakarta Sans", "Segoe UI Variable", "Segoe UI", system-ui, sans-serif`; `--font-mono: "Cascadia Code", Consolas, ui-monospace, monospace`.

| style | size/line | weight | use |
|---|---|---|---|
| display | 40/44, -0.02em | 800 | start screen, empty states |
| title | 24/30, -0.01em | 700 | dialog & backstage titles |
| heading | 16/22 | 700 | pane headings |
| body | 14/20 | 400 | dialog copy, notes pane |
| ui | 13/18 | 500 | buttons, tabs, menus, inputs (default chrome text) |
| caption | 12/16 | 500 | status bar, slide numbers, tooltips, small ribbon buttons |
| label | 11/14, 0.02em | 600 | ribbon group captions |
| mono | 12/16 | 400 | file names, size/position readouts |

Expose these as utility classes (`.t-display`, `.t-ui`, …) or as matching declarations in the component CSS.

**Spacing** `--space-1..8` = 4, 8, 12, 16, 24, 32px · **Radius** `--radius-sm` 4px, `--radius-md` 6px, `--radius-lg` 10px, `--radius-full` 9999px · **Shadow** `--shadow-slide`, `--shadow-popover` (values in tokens.json; dark values are heavier) · **Size** `--size-ribbon` 96px, `--size-tabrow` 36px, `--size-control` 28px, `--size-thumb` 168px.

Import `tokens.css` first in the global stylesheet. Then replace existing custom properties and hard-coded values with these tokens. Keep a temporary alias layer (`--old-name: var(--new-name)`) only if it shrinks the diff, and delete it by the end.

---

## 2. Theming: system + manual toggle

- Set `data-theme` on `<html>`. The default follows `prefers-color-scheme`. A user choice of **System / Light / Dark** overrides it.
- Persist the choice:
  - **web:** `localStorage`, wrapped in try/catch
  - **Tauri:** the app's existing settings store, or `localStorage` if there isn't one
- Apply the theme before first paint (an inline script in `index.html`) so there's no flash.
- Put the toggle in **View** tab → "Theme" group, and in File/backstage → Options if that exists.
- **Tauri / Windows 11:** match the window title bar and background to the theme. Use Tauri's window `theme` API and set the window `backgroundColor`, so there's no white flash on dark startup. If the app uses a custom title bar, style it with `--surface-100`.
- **The slide canvas, thumbnails' contents, slide show and exports are theme-independent.** Only the stage backdrop (`--surface-300`) and the chrome change.

---

## 3. Global foundations

- `body`: `--surface-100` background, `--ink` text, `--font-sans`, 13px/18px weight 500 (the `ui` style). Antialiasing on.
- Focus: `:focus-visible { outline: 2px solid var(--focus-ring); outline-offset: 2px; }` on every interactive element. Remove any `outline: none` that has no replacement.
- Scrollbars: thin and neutral (`--border-strong` thumb on transparent) in the rail, notes pane and dialogs.
- Text selection inside chrome inputs: `--brand-subtle` background.
- Motion: chrome transitions are 120–180ms ease-out. Wrap them in `@media (prefers-reduced-motion: reduce)` to disable. Don't touch slide-show transitions.
- Icons: line icons, 1.8px stroke on a 24px grid (1.5px at 28px), round caps/joins, `currentColor`. If the app uses an icon set, keep it and normalise size and colour. Don't swap libraries without asking.

---

## 4. Surface-by-surface spec

Restyle each of these. Tick each one off in `docs/restyle-plan.md`.

1. **Title bar / window chrome** (Tauri): `--surface-100`, file name in `ui`. Add an unsaved-changes dot (`--warning`, 8px) before the file name when the file is dirty. Use the existing dirty state; don't add new logic.
2. **Ribbon**
   - Tab row: `--size-tabrow` high, on `--surface-100`.
   - **File** button first: `--brand` fill, `--on-brand` text, `--radius-sm`, weight 600.
   - Tabs: `ink-muted` → hover `ink` on `--surface-200`. The active tab is `--brand-ink`, weight 700, with a 3px `--brand` underline (rounded top).
   - Body: `--size-ribbon`, `--surface` background, `--border` top and bottom.
   - Groups are columns divided by `--border` hairlines, with a `label` caption in `--ink-muted` centred at the bottom.
   - Buttons:
     - small: 25px high, `caption` text, 16px icon, `--radius-sm`, transparent → hover `--surface-200`
     - large: ~74px high, 28px icon over the label; at most one per group, placed first
     - toggles (Bold/Italic/alignment): pressed = `--brand-subtle` fill + 1px `--brand` border + `--brand-ink` icon
   - Galleries and dropdowns (layouts, themes, shapes) use `--shadow-popover`, `--radius-lg` and `--surface`.
   - Narrow window: the body scrolls horizontally and never wraps.
   - See `reference-components.css` → `.nx-ribbon*`, `.nx-rbtn*`.
3. **Slide rail**
   - `--surface-100`, gap `--space-3`, thumbnails `--size-thumb` wide (16:9) with `--radius-sm`, `--border` and `--shadow-slide`.
   - Slide number to the left in `caption` `--ink-muted`.
   - **Selected:** 2px `--brand` ring + number in `--brand-ink`, bold.
   - Hidden slides: 45% opacity and a struck-through number.
   - Drag-reorder insertion line: 2px `--brand`.
4. **Stage / canvas**
   - Backdrop `--surface-300`. The slide sits on `--slide-bg` (or its theme) with `--shadow-slide` and a `--space-6` minimum margin.
   - Selection box: 1px `--selection`. Handles: 8px white squares with a 1px `--selection` border. Rotate handle: a round `--radius-full` 10px dot above the box on a 1px stem.
   - **Smart guides:** 1px `--guide` lines. Distance/size labels go on a `--guide` pill with white `caption` text.
   - Text-edit caret and selection keep their current behaviour.
5. **Notes pane:** `--surface`, `--border` top, `body` text, placeholder "Click to add notes" in `--ink-muted`.
6. **Status bar:** `--surface-100`, `caption`, `--ink-muted`. Shows "Slide 3 of 12", the zoom control and the view buttons as small ghost buttons.
7. **Buttons everywhere:** `primary | secondary | ghost | danger`, `md` 32px / `sm` 28px, `--radius-md`. One primary per view. Disabled = `--surface-200` + `--ink-disabled`. See `.nx-btn*`.
8. **Inputs, selects, spinners** (font size, position/size fields): 28px high, `--surface` background, 1px `--border-strong` border, `--radius-sm`. Focus = `--focus-ring`. Numeric readouts use `mono`.
9. **Menus, context menus, tooltips:** `--surface`, `--radius-md` (menus `--radius-lg`), `--shadow-popover`, 28px items, hover `--surface-200`. Tooltips are `caption` text on `--ink` with `--surface` text.
10. **Dialogs** (open, export to PDF/PPTX, layout picker, unsaved changes, about)
    - `--overlay` scrim, `--surface` panel, `--radius-lg`, `--shadow-popover`, `--space-6` padding, `title` heading, `body` copy.
    - **Unsaved changes:** "Save changes to {filename}?" with buttons **Save** (primary) · **Don't save** (danger) · **Cancel** (ghost), in that order. Copy changes only; keep the logic.
11. **File / backstage + start screen**
    - Horizontal wordmark top-left (`wordmark-horizontal.svg`, or `-reversed.svg` in dark).
    - "New" and recent files as cards on `--surface` with `--radius-lg` and `--border`. Hover `--border-strong`.
    - Empty state in `display`/`body`.
12. **Slide show controls** (overlay toolbar during presentation): dark translucent pill (`--overlay` background, white icons, `--radius-full`). Keep it subtle, and don't change slide output.
13. **Toasts / notices** (font substitution, export done, errors): `--surface` with `--shadow-popover`, a leading status icon in `--success` / `--warning` / `--danger`, `body` text.

Contrast must hold in **both** themes: 4.5:1 for text, 3:1 for control borders, icons and focus rings. Never use colour alone for state; pressed/selected also get a border or ring.

---

## 5. Branding & Windows assets

- **Wordmark usage:**
  - horizontal (`wordmark-horizontal.svg`) in backstage/start and the About dialog
  - `-reversed` on dark grounds
  - stacked in square spaces (splash, installer banner)
  - clear space = the height of the "N"
  - minimum width 120px horizontal, 64px stacked
  - never recolour, re-space or retype it
- **App icon** (`app-icon.svg`): generate the full Tauri icon set (`npx tauri icon logos/app-icon.svg` or the project's equivalent). That covers `icons/32x32.png`, `128x128.png`, `128x128@2x.png`, `icon.ico` (16/24/32/48/64/256) and `icon.png`. Check that the 16px and 24px sizes stay legible.
- **`.nxslides` file association icon:** use the same icon (or a document variant, if the installer config supports a separate one) in `tauri.conf.json` → `bundle.fileAssociations`.
- **Installer (NSIS/WiX):** use the brand header/sidebar images if the installer supports them. Generate BMP/PNG at the sizes Tauri's bundler expects from the stacked wordmark on `#ffffff`, with indigo `#4338ca` accents.
- **Web:** favicon (`app-icon.svg` + 32px PNG fallback), `apple-touch-icon` 180px, `<meta name="theme-color">` light `#f6f6fb` / dark `#14141d`, page title "Next Office Slides".
- **Name rules in UI copy:** "Next Office Slides" in full on first mention and in the window title. "Next Slides" is acceptable afterwards. Never write "NextSlides" or "NOS".

---

## 6. Voice & copy pass

While restyling, fix chrome copy to match:

- Plain, short, task-first. Mirror PowerPoint command names (New Slide, Layout, Arrange, Slide Show).
- Use Title Case only for ribbon commands and sentence case everywhere else. No exclamation marks or emoji.
- Errors say what happened and what to do, e.g. "Some fonts in this file aren't installed. Text uses Segoe UI instead."

Don't change copy that comes from the PPTX itself.

---

## 7. Verification (required before you say you're done)

1. Every build passes: type-check, lint, unit tests, the web build and `tauri build`.
2. `grep` the chrome CSS for raw `#hex`, `rgb(` and font-family literals outside `tokens.css` and slide-rendering code. Report anything left and why.
3. Run the app and check each surface in §4 in **light and dark**:
   - Open a real `.pptx` with theme colours, tables, pictures and speaker notes.
   - Confirm the slides render identically to before (compare with screenshots from before the restyle).
   - Export to PPTX and PDF and confirm the output is unchanged.
4. Keyboard pass: tab through the ribbon, rail and dialogs. Focus is always visible.
5. Contrast spot-check of the key pairs in both themes (ink/muted on each surface, on-brand on brand, guide on slide).
6. Windows: run the installer and check the app icon (taskbar, Start, Alt-Tab), the `.nxslides` file icon, the dark title bar, and that there's no white flash on dark startup.
7. Update `docs/restyle-plan.md` with before/after screenshots of the main editor, rail, a dialog and backstage in both themes. List anything you deferred.

Finish with a short summary: what changed, the files touched, anything you couldn't match, and follow-ups (e.g. an icon set, more components).
