# Next Office

A Windows desktop office suite (React + TypeScript + Vite + Tauri). It opens on a
launcher where you choose an app:

- **Next Office Slides**: a PowerPoint-style presentation editor (opens and exports .pptx).
- **Next Spreadsheet**: an Excel-style spreadsheet (`src/sheet/`): formulas with about 100 Excel functions, number
  formats, fill handle, sort, freeze panes, several sheets, and opening and saving .xlsx and .csv files (native format: .nxsheet).
  Not supported yet: charts, merged cells, borders, conditional formatting (they are left out when an .xlsx is opened,
  and the status bar says so).
- **Next Writer**: a Word-style word processor (`src/doc/`), described below. Opens and saves .docx (native format: .nxdoc).

The launcher and app switching live in `src/shell/`; the logo button in each app's title bar goes back to the launcher.

### File types and default apps

`src/fileTypes.json` is the one list of file types Next Office opens (.xlsx, .csv, .pptx…). From it:

- each type opens in the right app (Spreadsheet or Slides);
- at every start, the Windows app registers itself for these types (per user, no admin rights), so it appears in
  **Open with** and in **Settings › Default apps**. This never takes a type over: Windows only lets the user change a default;
- the first time, Next Office asks whether it should open these files. **Set as default** opens its page in Windows
  Settings, where one click makes it the default. The question is asked once per file type, so a type added to the
  list later is asked about once more. **File › Default apps…** (and the link on the launcher) shows it again any time;
- the installer claims Next Office's own formats (.nxslides, .nxsheet), and the uninstaller removes the registration.

To support a new type, add it to `src/fileTypes.json` and give the app a reader for it; `npm run app:build` updates the
installer (`npm run sync:file-types` does it on its own), and `npm test` fails if something is out of step.

## Run

Next Office is a native Windows app, built with [Tauri 2](https://tauri.app) on the WebView2 runtime that comes with Windows 11. The installer is only a few MB. There is no browser version: opening the Vite dev server in a browser tab only shows a notice.

```bash
npm install
npm run app:dev     # run the app with hot reload
npm run app:build   # build the installers
npm test            # unit tests, including a PowerPoint export/import round trip
```

`npm run dev` and `npm run build` are the frontend steps Tauri runs for these.

The installers are written to `src-tauri/target/release/bundle/`:

- `nsis/Next Office Slides_0.1.0_x64-setup.exe`: a per-user installer that doesn't need admin rights.
- `msi/Next Office Slides_0.1.0_x64_en-US.msi`: for IT or Group Policy deployment.

Building needs Rust (MSVC toolchain) and the Visual Studio C++ Build Tools.

What the app does as a Windows app:

- A **start screen**, like PowerPoint's. It greets you with ready-made templates (Pitch Deck, Project Status Report, Quarterly Business Review, Product Launch, Classroom Lesson, Training Workshop, Personal Portfolio and Tech Talk, plus four Arabic and bilingual ones, see below), a blank deck for each theme, a card to continue where you left off, and recent and pinned files with search. Clicking a template shows a preview of all its slides before you create it. You can come back to it any time with **File → Home**.
- A **Windows 11 window**: frameless, with rounded corners, a drag-to-move title bar and Fluent minimise, maximise and close buttons.
- **Real files**: native Open and Save dialogs. **Ctrl+S** saves in place and **F12** (or Ctrl+Shift+S) is Save As, to `.nxslides` or `.pptx`. An imported PowerPoint file is never overwritten silently; the first Save always asks where to save it.
- **An unsaved-changes prompt** (Save / Don't Save / Cancel) when you close, open or create a presentation.
- **A `.nxslides` file association**, so double-clicking a presentation in Explorer opens it.
- **Slide shows fill the whole screen.**
- **Native behaviour**: no browser context menu, reload or page zoom.

The desktop code lives in `src/desktop.ts` (frontend bridge) and `src-tauri/` (Rust shell: window setup and file read/write commands).

## Features

- **Ribbon UI**: File menu plus Home, Insert, Design, Transitions and Slide Show tabs.
- **Slides**: six layouts (Title, Title and Content, Section Header, Two Content, Title Only, Blank). You can add, duplicate, delete and drag-reorder slides in the thumbnail panel.
- **Charts**: column, bar, line, area, pie and doughnut. Double-click a chart to edit its data in a grid, where you can paste a table from Excel or Google Sheets. The **Chart Design** tab changes the type, title, legend, data labels and stacking. Charts export to PowerPoint as native, editable charts.
- **Elements**: text boxes, 10 shapes (with text inside), lines, arrows, pictures, **video** and **audio**, added from disk, by paste, or by drag-and-drop. Each video gets a poster frame. The **Playback** tab sets start automatically, loop and mute. In the slide show, media plays in place.
- **Editing**: move, resize (8 handles, Shift locks aspect) and rotate (Shift snaps to 15°). You can marquee-select and multi-select, and smart guides snap to the slide and to other objects (hold Alt to disable).
- **Formatting**: font, size, bold, italic, underline, color, bullets, text direction, horizontal and vertical alignment, fill, outline, weight, opacity and z-order. Text formatting can be mixed within a box: select words while editing and the ribbon applies to just those. With the box selected, it applies to all its text. Tab and Shift+Tab indent and outdent bullets, and Shift+Enter inserts a line break.
- **Design**: 12 themes, 4 of them with Arabic fonts (they restyle title and body placeholders), solid or gradient backgrounds, and **slide sizes**: Widescreen 16:9, Standard 4:3, Square 1:1, Portrait 4:5 and Vertical 9:16. Changing the size scales content to fit, like PowerPoint's "Ensure Fit". PowerPoint files keep their own aspect ratio when opened and exported.
- **Animations tab**: entrance effects (Appear, Fade, Fly In, Wipe, Zoom, Float In), emphasis effects (Pulse, Spin, Teeter) and exit effects. Each effect starts On Click, With Previous or After Previous, with its own duration, delay and direction. Bulleted text can build one paragraph per click. The **Animation Pane** lists and reorders effects, and numbered badges mark animated objects. Animations export to PowerPoint and play there.
- **Slide show**: full screen, with Fade, Push and Zoom transitions and keyboard or mouse navigation. Animations play click by click; going back undoes the last step.
- **Speaker notes**, **undo/redo**, and **autosave**. The document is a [Yjs](https://yjs.dev) CRDT, so edits made offline or by several people merge without conflicts. This is the base for cloud sync and co-editing.
- **File**: open **PowerPoint (.pptx / .ppsx / .potx)** files, open and save `.nxslides` (a zip holding `presentation.json` and a `media/` folder; older `.nxslides.json` files still open), export to **PowerPoint (.pptx)**, and export to **PDF** via print. You can also drag a .pptx onto the slide to open it.

## Arabic and right-to-left text

- **Text direction follows the text.** A paragraph that starts with an Arabic letter is right-to-left: it is right-aligned, its bullet sits on the right, and indent levels grow from the right. English paragraphs in the same box stay left-to-right, so Arabic and English mix freely. Paragraphs without letters (empty or numbers only) follow the paragraph before them, so pressing Enter after Arabic keeps the caret on the right.
- **Home › Paragraph** has **Right-to-Left** and **Left-to-Right** buttons, like PowerPoint's. They set the direction of the selected paragraphs (or the whole box) and mirror left and right alignment. An alignment you pick yourself always wins: "Align Left" on Arabic text left-aligns it.
- **Arabic fonts.** The font list has three groups: the usual fonts, the Arabic fonts that come with Windows (Sakkal Majalla, Traditional Arabic, Simplified Arabic, Arabic Typesetting, Aldhabi, Andalus), and five open-source Arabic fonts included with the app (Cairo, Tajawal, Noto Kufi Arabic, Noto Naskh Arabic and Amiri, under the SIL Open Font License in `src/fonts/`). Arabic typed in a font without Arabic letters falls back to a matching Arabic font, never to a random one.
- **PowerPoint**: paragraph direction (`rtl="1"`), alignment and bullets go both ways. Arabic runs are exported with the Arabic language tag and complex-script font (`a:cs`), and on import Arabic runs use their complex-script font or the theme's Arabic font.
- **Templates**: Company Profile (Arabic/English), Pitch Deck (Arabic), Classroom Lesson (Arabic) and Event / Conference (Arabic/English). Bilingual slides put Arabic in the right-hand column and English in the left. They use four new themes with Arabic fonts: Sahara (Cairo/Tajawal), Nile (Noto Kufi/Tajawal), Andalus (Amiri/Noto Naskh) and Diwan (Segoe UI, which every Windows PC has).

### Arabic interface

The whole interface is translated into Arabic and mirrored: the ribbon reads from the right, the thumbnails are on the right, the Animation Pane on the left, and menus open towards the left. It uses Arabic-Indic digits (٠١٢٣) and Arabic plural forms ("شريحتان", "٣ شرائح"). Slides themselves are never mirrored or translated.

The language follows the Windows display language. Change it under **File › Language** or with the picker at the bottom of the start screen; the choice is remembered. Translations live in `src/i18n.ar.ts`, keyed by the English text; `npm test` fails if any interface string is missing an Arabic translation.

## Opening PowerPoint files

The importer (`src/importPptx.ts`) comes close to PowerPoint on most decks. It brings in:

- **Placeholders**, with position, fonts, sizes, colors and bullets inherited from the slide layout and master.
- **Theme colors** (including tints and shades) and theme fonts.
- **Master and layout artwork**, such as logos and bars.
- **Shapes**: preset shapes, freeforms, groups (including rotated and flipped ones), and lines and connectors with arrowheads.
- **Pictures**, with cropping and transparency, and **embedded video and audio**.
- **Tables**, including header and banded-row styling.
- **SmartArt**, from PowerPoint's pre-rendered drawing.
- **Charts** (column, bar, line, area, pie and doughnut) come in as editable charts, with their data, colours, title, legend and labels.
- **Backgrounds**: solid, gradient and picture.
- **Speaker notes and transitions.**
- **Animations**: entrance, emphasis and exit effects, with their order, triggers, timing and by-paragraph builds.

Known limitations:

- **Scatter, radar, bubble and stock charts** show as placeholders, and combination charts show only their first chart type.
- **Pictures in EMF, WMF or TIFF format** are skipped, because WebView2 can't display them.
- **Animations** outside the supported effects are simplified (to Fade, or Pulse for emphasis). Motion paths and triggers on other shapes are not supported.
- **Linked (not embedded) video and audio** are skipped, and WMV, WMA and AVI media is imported but can't play. Convert it to MP4 or MP3.
- **Old binary .ppt files** must be re-saved as .pptx first.
- **Arabic**: the fonts included with the app (Cairo, Tajawal, Noto, Amiri) can't be embedded in a .pptx, so PowerPoint shows them only where they are installed. Use the Diwan theme or the Windows Arabic fonts for decks you send. Chart axes are not mirrored for Arabic charts. Windows message boxes (Save / Don't Save / Cancel) show Arabic text but keep their left-to-right layout.

## Next Writer (documents)

A word processor laid out on real pages, like Word's Print Layout (`src/doc/`).

- **Pages**: A4, Letter, Legal, A5 and Executive, portrait or landscape, with preset or custom margins (Layout tab).
  Text flows from page to page: a paragraph that doesn't fit moves to the next page, or splits between two lines when
  it is longer than the space left. **Ctrl+Enter** inserts a page break.
- **Text**: font, size (Ctrl+] / Ctrl+[), bold, italic, underline, strike-through, subscript and superscript, colour,
  highlight, Clear All Formatting and the **Format Painter**.
- **Paragraphs**: alignment (Ctrl+L/E/R/J), bulleted and numbered lists (1. a. i. by level; typing "* " or "1. " starts
  one), indents (Tab / Shift+Tab), line spacing, space before and after, and the styles gallery (Normal, Title, Subtitle,
  Heading 1–4, Quote, Code; Ctrl+Alt+1/2/3).
- **Arabic**: each paragraph's direction follows its first letter (or the Right-to-Left button), as in Slides.
- **Insert**: tables (with a contextual **Table** tab to add or delete rows and columns, merge and split cells, shade
  cells and make a header row), pictures (from disk, paste or drag-and-drop; drag the corner to resize), links (Ctrl+K),
  horizontal lines, the date, symbols, a header and footer, and page numbers.
- **Find and Replace** (Ctrl+F / Ctrl+H), word count, formatting marks (¶), spell check, zoom (One Page, Page Width).
- **Files**: open .docx (and .docm, .dotx), .html, .txt and .nxdoc; save .docx or .nxdoc; export web page or plain text;
  **Print or save as PDF** (Ctrl+P) prints with the page size, margins, header, footer and page numbers. Autosaves between
  sessions.

The .docx reader and writer (`src/doc/docx.ts`) are written on JSZip. Known limitations: text boxes and shapes,
comments, footnotes, equations, columns and multiple sections are left out when a .docx is opened (the status bar
lists what was dropped); floating pictures become in-line pictures; headers and footers are plain text, the same on
every page; tables move to the next page whole rather than splitting between rows.

## Keyboard shortcuts

| Keys | Action |
| --- | --- |
| Ctrl+Z / Ctrl+Y | Undo / Redo |
| Ctrl+C / X / V | Copy / Cut / Paste (images and text from other apps too) |
| Ctrl+D | Duplicate selection (or slide) |
| Ctrl+A | Select all on slide |
| Ctrl+M | New slide |
| Ctrl+T | Insert text box |
| Ctrl+B / I / U | Bold / Italic / Underline |
| Enter / F2 | Edit text of selected object |
| Del | Delete selection (or slide when thumbnails are focused) |
| Arrows | Nudge 1px (Shift: 10px); with nothing selected, change slide |
| F5 / Shift+F5 | Present from beginning / current slide |
| Ctrl+S / Ctrl+O / Ctrl+P | Save / Open (.pptx or saved file) / PDF |
| Ctrl+N | New blank presentation |
| F12 / Ctrl+Shift+S | Save As (desktop app) |

## Project layout

Next Writer lives in `src/doc/`: `schema.ts` (ProseMirror schema), `model.ts` (styles, page setup, .nxdoc),
`pagination.ts` (pages), `commands.ts`, `editor.ts` (plugins, shortcuts, paste), `plugins.ts` (list numbers, direction,
find), `docx.ts` (Word files), `convert.ts` (HTML and text), `store.ts`, `io.ts` (files, printing) and `components/`.
Its Arabic text is in `src/i18n.doc.ar.ts`.

```
src/
  types.ts          data model (Presentation → Slide → SlideElement)
  store.ts          zustand store: document snapshot (backed by ydoc.ts), selection, undo, autosave
  factory.ts        element/slide/layout constructors
  themes.ts         themes and font helpers
  geometry.ts       bounds, hit-testing, snapping
  io.ts             open/save, images, clipboard, PDF
  desktop.ts        Windows app bridge (Tauri): dialogs, file I/O, window controls
  ydoc.ts           Yjs document layer: snapshot ↔ CRDT diffing, undo manager, persistence
  textSchema.ts     ProseMirror schema + paragraph model conversion (shared by editor and Yjs)
  richtext.ts       rich text model helpers (paragraphs → runs), shared rendering
  migrate.ts        upgrades files saved by older versions
  i18n.ts           interface language: t()/tp(), Arabic-Indic digits, page direction
  i18n.ar.ts        Arabic translations (keyed by the English text)
  fonts/            bundled Arabic fonts (woff2) and their licences
  anim.ts           animation model, click steps, keyframes
  pptxAnim.ts       PowerPoint <p:timing> export/import
  chart.ts          chart model, palette, scales, table paste parsing
  media.ts          media store: files named by content hash (files in the app data folder)
  nxfile.ts         .nxslides package format (zip: presentation.json + media/)
  storage.ts        app state storage (files in the app data folder)
  templates.ts      start-screen templates and theme presets
  recent.ts         recent and pinned files
  importPptx.ts     PowerPoint import (JSZip + DOMParser, lazy-loaded)
  exportPptx.ts     PowerPoint export (pptxgenjs, lazy-loaded)
  color.ts          colour conversion helpers
  components/
    Editor.tsx      canvas: selection, drag, resize, rotate, marquee, context menu
    ElementView.tsx renders one element (shared by editor, thumbnails, slide show)
    RichTextEditor.tsx in-place text editor (ProseMirror) and selection formatting
    Ribbon.tsx      ribbon tabs
    SlidePanel.tsx  thumbnails + reordering
    Presenter.tsx   slide show + print view
    AnimationPane.tsx animation list; slidePlayer.ts plays a slide's animations in the slide show
    ChartView.tsx   chart renderer (SVG); ChartDataEditor.tsx the data grid
    MediaView.tsx   pictures, video and audio (stills in the editor, players in the slide show)
    StartScreen.tsx start screen: templates, recent files, open
    LanguagePicker.tsx interface language menu items and start-screen picker
```
