# Restyle plan: Next Office design system

Brief: `next-office-brand/next-office-brand/RESTYLE_PROMPT.md`. Brand pack copied to `src/assets/brand/`.

## What exists today

- Styling is plain CSS: `src/styles.css` (Slides + shared chrome), `src/sheet/sheet.css` (Next Spreadsheet + launcher), `src/doc/doc.css` (Next Writer), `src/motion.css` (transitions), `src/fonts/fonts.css` (bundled Arabic slide fonts). The splash screen is inline CSS in `index.html`.
- Custom properties in `:root`: `--brand` (orange #c2410c), `--brand-dark`, `--brand-soft`, `--bg`, `--panel`, `--border`, `--text`, `--muted`, `--hover`, `--active`, `--sel`. Sheet and Writer override `--brand*` and `--active` with green/blue on `.sheet-app` / `.doc-app`.
- About 190 raw colour literals across the three CSS files. The font stack is `'Segoe UI'`.
- The ribbon classes (`.ribbon`, `.rb-btn`, `.rb-group`, `.menu`, `.modal`, `.titlebar`, `.statusbar`) and components (`Btn`, `Group`, `Popover`) are **shared by all three apps**, so a restyle of the Slides chrome also restyles Sheet and Writer.
- Desktop-only Tauri 2 app with a frameless window and a custom title bar (`decorations: false`). The unsaved-changes prompt is a **native** Windows message box (`askSaveChanges` in `src/desktop.ts`), so it can't be styled. Only its copy can change.
- There's no git repository, so the brief's "one commit per step" isn't possible. A backup of the pre-restyle sources is kept outside the repo.

## Decisions

1. **Scope = the whole suite chrome.** Tokens are global. Slides and the launcher use the indigo brand. Sheet and Writer keep their own accents (green / blue) as overrides of `--brand`, `--brand-hover`, `--brand-subtle`, `--brand-ink` and `--on-brand`, with light and dark values. Everything else comes from the shared tokens.
2. **Content never themes:** slide content, thumbnails' contents, the slide show, exports, Writer's pages and the spreadsheet cells keep fixed colours. Only the chrome around them changes.
3. **Theme setting** (`System | Light | Dark`) lives in `src/theme.ts` and is stored in the app's state store (`next-office:theme`). It is mirrored to `localStorage` so an inline script in `index.html` can set `data-theme` before first paint. Rust reads the same state file at start-up and sets the window background and theme, so there's no white flash on a dark start.
4. Toggle placement: a **View** tab gets a "Theme" group in all three apps (Slides gains a View tab). The File menu also gets theme items under the Language items, and the launcher and start screen get a compact theme select next to the language select.

## Old → new token mapping

| old | new |
|---|---|
| `--bg` | `--surface-100` |
| `--panel` | `--surface` |
| `--hover` | `--surface-200` |
| `--text` | `--ink` |
| `--muted` | `--ink-muted` |
| `--active` | `--brand-subtle` |
| `--brand-soft` | `--brand-subtle` |
| `--brand-dark` (text) | `--brand-ink` |
| `--brand-dark` (fill hover) | `--brand-hover` |
| `--sel` | `--selection` |
| `#fff` on brand | `--on-brand` |
| `#fdba8c` (orange hover rings) | `--border-strong` / `--brand` |
| `#ec4899` guides | `--guide` |
| `rgba(0,0,0,.4)` scrims | `--overlay` |
| ad-hoc shadows | `--shadow-popover` / `--shadow-slide` |
| `#e8e6e4` stage | `--surface-300` |

## Files

- new: `scripts/build-tokens.mjs`, `src/styles/tokens.css`, `src/theme.ts`, `src/components/ThemePicker.tsx`, `src/assets/brand/**`
- edited: `src/main.tsx`, `index.html`, `src/styles.css`, `src/sheet/sheet.css`, `src/doc/doc.css`, `src/motion.css`, `src/components/{Chrome,Ribbon,LanguagePicker,StartScreen}.tsx`, `src/shell/{Launcher,HomeButton}.tsx`, `src/sheet/components/Ribbon.tsx`, `src/doc/components/Ribbon.tsx`, `src/desktop.ts` (copy), `src/i18n*.ar.ts` (Arabic for new strings), `src-tauri/src/lib.rs`, `src-tauri/icons/*`, `package.json` (a `tokens` script)

## Risks

- Shared class names: Sheet and Writer change look too. Mitigation: they keep their accent colour, and both get checked.
- Slide rendering must not change. Mitigation: no edits to `SlideView`, `ElementView`, `ShapeSvg`, `ChartView`, the theme code or export code. `.slide-view`/`.slide-surface` rules are left alone. Thumbnails keep the `SlideView` renderer and only the frame around them changes.
- Arabic UI: `[dir=rtl] body` sets a font stack. Plus Jakarta Sans has no Arabic glyphs, so the Arabic stack keeps Segoe UI / Noto Kufi first.
- The i18n test requires Arabic for every new `t()` string.

## Surface checklist (§4)

- [x] 1 Title bar: `--surface-100`, unsaved dot from the existing `doc !== savedDoc` (Slides only; Sheet/Writer title bars restyled, no dot)
- [x] 2 Ribbon: tab row, File button, tabs, body, groups, small/large buttons, toggles, galleries
- [x] 3 Slide rail: ring, numbers, drop line. *Hidden slides don't exist in the app, so not added.*
- [x] 4 Stage: backdrop, shadow, selection, square handles, round rotate dot, `--guide` lines. *The app has no distance/size labels, so none were added.*
- [x] 5 Notes pane: placeholder copy is now "Click to add notes"
- [x] 6 Status bar
- [x] 7 Buttons: `.primary-btn`, `.secondary-btn`, `.danger-btn`, `.icon-btn`
- [x] 8 Inputs, selects, number boxes
- [x] 9 Menus, context menus. *There are no custom tooltips (native `title`), so nothing to restyle there.*
- [x] 10 Dialogs: in-app modals restyled. The unsaved-changes box is native, so only the copy changed ("Save changes to {name}?", Save / Don't save / Cancel).
- [x] 11 Start screen / backstage with the wordmark. The launcher also uses it.
- [x] 12 Slide show controls: pill on `--overlay`
- [ ] 13 Toasts: the app has none (status-bar messages + native message boxes), so nothing was added

## Verification log

### 1. Builds

- `tsc --noEmit`: passes.
- `vitest run`: 20 files, 232 tests pass. The only test change: the splash test now selects `.splash-by img`, because the splash has a second image (the app icon).
- `vite build`: passes.
- `cargo check`: passes.
- `tauri build`: passes. It produced `Next Office_0.4.1_x64-setup.exe` (NSIS, with the brand header and sidebar) and `Next Office_0.4.1_x64_en-US.msi` (WiX banner and dialog).
- There is no lint script in the project.

### 2. Raw values left in chrome CSS

Every remaining `#hex`, `rgb(` or font-family literal outside `tokens.css` and `app-accents.css` is one of these:

- **Slide content:** the slide font pin (`.slide-view`, `.stage`, `.present-slide`, `.print-page`), video/audio placeholders, the dashed placeholder outline over slides, the slide show's black backdrop and end text, and print `#fff`.
- **Document content:** Writer pages, text, tables, links, formatting marks and find highlights; spreadsheet grid lines, headers, cells, selection and the cell editor; style and font samples that show the document's own colours.
- **Windows conventions:** the caption-glyph font (`Segoe Fluent Icons`) and the red Close-button hover, `#c42b1c`. Every Windows 11 title bar uses these.
- **Slide-show pill:** white icons on `--overlay`, because the show is always on black.
- **Splash (`index.html`):** inline light/dark brand values, because tokens.css hasn't loaded yet when it paints.

### 3. Rendering

Shots of the built app were taken in headless Edge, with a stubbed Tauri bridge, before and after the restyle (`docs/restyle/`). The *Project Status Report* template renders with the same fonts, colours and layout. It looks slightly larger after because the stage margin is now `--space-6`.

The PPTX/PDF export code and the slide renderers weren't edited. A round-trip of a real `.pptx` (theme colours, tables, pictures, notes) in the desktop app is **still to do by hand**.

### 4. Keyboard

A scripted Tab pass covered the title bar, File, every ribbon tab, the Home tab controls, the rail and the status bar. Every stop shows a 2px `--focus-ring` outline, with two designed exceptions: the notes pane shows focus as a top line, and the title field shows a ring.

### 5. Contrast

All key pairs pass in both themes, including the Sheet and Writer accents (`scratchpad contrast.mjs`). The lowest are `--border-strong` on `--surface-100` at 3.12 (light) and `--guide` on the slide at 4.30. **One miss:** white text on `--guide` is 4.30 against the 4.5 target. It doesn't matter today because the app has no distance labels, but the guide colour needs darkening before labels are added.

### 6. Windows: still to do by hand

These need a real install:

- [ ] Run the installer and check its header and sidebar art.
- [ ] Check the app icon in the taskbar, Start and Alt-Tab.
- [ ] Check the `.nxslides` file icon.
- [ ] Check the dark title bar.
- [ ] Check there's no white flash on a dark start.

### 7. Screenshots (`docs/restyle/`)

| surface | before | after, light | after, dark |
|---|---|---|---|
| launcher | before-light-1-launcher | after-light-1-launcher | after-dark-1-launcher |
| start screen | before-light-2-start | after-light-2-start | after-dark-2-start |
| editor + rail + selection | before-light-4-editor-selected | after-light-4-editor-selected | after-dark-4-editor-selected |
| File menu | before-light-6-file-menu | after-light-6-file-menu | after-dark-6-file-menu |
| dialog | before-light-7-dialog | after-light-7-dialog | after-dark-7-dialog |
| View › Theme | — | after-light-5-view-tab | after-dark-5-view-tab |
| Spreadsheet / Writer | — | — | after-dark-8-spreadsheet, after-dark-9-writer |

### Deferred

- Hidden slides, distance labels and toasts: the app doesn't have these features, so they weren't added.
- The brand pack's wordmark reads "Next Office Slides". The suite launcher, splash and installer say "Next Office", so they use the app icon with text, except the installer art, which uses the Slides wordmark. A suite wordmark would fix this.
- The `.nxslides` file icon is the app icon. Windows takes it from the exe's icon, and the bundler has no separate document icon.
- The spreadsheet grid and Writer pages stay light in dark mode on purpose, as document content.
