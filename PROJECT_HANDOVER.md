# Gemini Copilot for PowerPoint — Complete Architecture, Development & Handover Guide

This document provides a comprehensive technical reference for the **Gemini Copilot for PowerPoint** project. It is specifically structured so that another AI agent or developer can immediately pick up development with full context, understand every design decision, run tests, and continue building new features.

---

## 1. Project Overview & System Architecture

### What is Gemini Copilot for PowerPoint?
A Microsoft Office Add-in (Task Pane) embedded directly inside **Microsoft PowerPoint** (Desktop for Windows/Mac and Office on the Web). It provides an autonomous, multi-turn AI presentation copilot powered by the **Google Gemini API** (`gemini-3.8-flash` with automatic fallback to `gemini-2.5-flash` and `gemini-2.5-pro`).

### Key Capabilities
1. **Autonomous Tool Calling Loop**: Directly reads, formats, edits, and creates PowerPoint slides using Gemini Function Calling across up to 6 iterative rounds per turn.
2. **Deck-Aware & Multi-Slide Editing**: Not restricted to the active slide. Can inspect the entire deck outline (`read_deck_outline`) and target any specific slide by 1-based `slide_number` (e.g. `slide_number: 6`, `slide_number: 7`), allowing multi-slide batch edits in a single turn.
3. **Table & Shape Native Manipulation**: Reads and writes text, font color, size, bold, italic, and table cell background fills using Office.js PowerPointApi 1.8+.
4. **Safety & Prompt Injection Defense**: Slide document text is segregated into an untrusted content block (`[SLIDE DATA CONTEXT - UNTRUSTED DOCUMENT CONTENT]`). Destructive actions (like `delete_shape`) strictly require `user_confirmed: true`.
5. **In-Memory Undo History Stack**: Tracks all mutations (`set_font`, `set_shape_text`, `table_cells`) with previous values, accessible via an interactive `↩ Undo` button in the UI and via the `undo_last_edit` agent tool.
6. **Multi-Tab Task Pane UI**:
   - **Chat**: Autonomous conversation agent with full multi-turn memory.
   - **Create**: 16:9 widescreen slide layout generator (cards, bullet lists, problem-solution, timeline, stats).
   - **Refine**: Selection-based or whole-slide text rewriting and polish.
   - **Critique & Notes**: Slide design scoring and speaker note generator.

---

## 2. Workspace & File Hierarchy

### Local Workspace Path
- **Project Root**: `C:\Users\shiva\.gemini\antigravity\scratch\gemini-powerpoint-addin\`
- **Sample Presentation**: `C:\Users\shiva\Documents\antigravity\Presentation2.pptx`

### Repository & Live Hosting
- **GitHub Repository**: `svsrr/gemini-powerpoint-addin` (Branch: `main`)
- **Live Production URL**: `https://ppt.distillatelabs.com`
- **Cloudflare Pages URL**: `https://ppt-distillatelabs.pages.dev` (Project: `ppt-distillatelabs`)

### File Structure Map
```text
C:\Users\shiva\.gemini\antigravity\scratch\gemini-powerpoint-addin\
├── .github/
│   └── workflows/
│       └── deploy.yml              # GitHub Actions CI/CD to Cloudflare Pages
├── assets/                         # Add-in ribbon and store icons
│   ├── icon-16.png
│   ├── icon-32.png
│   ├── icon-64.png
│   ├── icon-80.png
│   └── icon-128.png
├── dist/                           # Compiled production build (Vite output)
│   ├── _headers                    # Cloudflare security & CORS headers
│   ├── CNAME                       # Custom domain binding (ppt.distillatelabs.com)
│   ├── index.html
│   ├── manifest.xml
│   ├── taskpane.html
│   └── assets/                     # Bundled JS and CSS chunks
├── public/                         # Static assets copied verbatim to dist/
│   ├── _headers
│   ├── CNAME
│   ├── manifest.xml
│   └── assets/
├── scripts/
│   └── generate-icons.js           # Icon generation utility
├── src/
│   ├── taskpane.css                # Modern responsive UI stylesheet
│   └── taskpane.js                 # CORE ADD-IN ENGINE: UI, Office.js, Gemini API, Agent loop
├── tests/
│   ├── audit_suite.test.js         # Schema, syntax, role validation, and table selection tests
│   ├── mock_office_tools.test.js   # Mocked Office.js integration & runtime lifecycle test suite
│   └── pptx_fidelity.test.js       # Real slide structure fidelity tests (against Presentation2.pptx)
├── index.html                      # Standalone entry/preview page
├── manifest.xml                    # Production Office Add-in XML manifest (with PowerPointApi 1.8)
├── manifest.prod.xml               # Duplicate production manifest
├── package.json                    # Scripts and dependencies
├── privacy.html                    # Store compliance Privacy Policy
├── support.html                    # Store compliance Support page
├── terms.html                      # Store compliance Terms of Service
├── vite.config.js                  # Vite bundler configuration
└── PROJECT_HANDOVER.md             # This comprehensive handover document
```

---

## 3. Technology Stack & Key Dependencies

| Component | Technology | Version | Purpose |
| :--- | :--- | :--- | :--- |
| **Bundler** | Vite | `^6.1.0` | Fast ESM build & local development |
| **Typings** | `@types/office-js` | `^1.0.616` | TypeScript definitions for Office.js |
| **Deployment** | `wrangler` | `^4.158.0` | Cloudflare Pages deployment CLI |
| **LLM Provider** | Google AI Studio | `v1beta` | Direct client-side API via `x-goog-api-key` header |
| **Default Model** | `gemini-3.8-flash` | Latest | Primary fast agent model |
| **Fallback Models**| `gemini-2.5-flash`, `gemini-2.5-pro` | Latest | Automatic fallback on 404/429/503 |
| **Host Runtime** | Microsoft PowerPoint | API 1.8+ | Desktop Windows/Mac and Office Online |

---

## 4. Complete Tool Registry (`GEMINI_AGENT_TOOLS`)

The agent loop in `src/taskpane.js` declares and executes the following 9 function tools:

| Tool Name | Parameter Schema | Description & Functionality |
| :--- | :--- | :--- |
| `read_deck_outline` | `{}` | Scans the entire presentation and returns `{ totalSlides, slides: [ { slideNumber, slideId, title, shapeCount, hasTable } ] }`. |
| `read_slide` | `slide_number?: number` | Reads all shapes, text boxes, and table cells with IDs, types, positions, texts, and cell font colors on the target or active slide. |
| `set_font` | `slide_number?: number`, `shape_id: string`, `color?: string`, `size?: number`, `bold?: boolean`, `italic?: boolean` | Applies font styling to a shape, title, or heading (`shape_id: "heading"` targets title). Records previous state to undo stack. |
| `set_shape_text` | `slide_number?: number`, `shape_id: string`, `text: string` | Updates or rewrites text content of a shape or text box on the target or active slide. Records undo state. |
| `set_table_cell` | `slide_number?: number`, `shape_id?: string`, `row: number`, `col: number`, `text?: string`, `font_color?: string`, `fill_color?: string` | Updates text, font color, and/or fill background of a single table cell (0-based indices). |
| `batch_update_table`| `slide_number?: number`, `shape_id?: string`, `updates: Array<{row, col, text?, font_color?, fill_color?}>` | Batch updates multiple cells in a table in a single atomic Office.js call. |
| `add_slide` | `title: string`, `subtitle?: string`, `layout?: string`, `items?: Array`, `bullets?: Array`, `takeaway?: string`, `theme?: string` | Inserts a new 16:9 widescreen slide formatted with the selected theme into the presentation. |
| `delete_shape` | `slide_number?: number`, `shape_id: string`, `user_confirmed: boolean` | Deletes a shape. Blocks execution if `user_confirmed` is not `true`. |
| `undo_last_edit` | `{}` | Pops the most recent mutation from `state.undoStack` and restores previous values on the target slide. |

---

## 5. Chronological Bug Fixes & Architectural Evolution

During development, several critical challenges and edge cases were discovered and resolved:

### 1. Gemini Function Response Role Standard
- **Bug**: Gemini API rejected requests with `Error calling Gemini API: Role 'function' is not supported. Please use a valid role: USER, MODEL...`.
- **Root Cause**: Earlier OpenAI-style code used `{ role: 'function', ... }`. Google AI Studio Gemini API (`v1beta`) mandates `role: 'user'` containing `{ functionResponse: { name, response: { result } } }`.
- **Fix**: Standardized all tool response turns to `role: 'user'`.

### 2. Typo Fallthrough on Shape Lookups
- **Bug**: If Gemini passed an unknown shape ID (e.g. `sh-nonexistent`), the code fell through to `withText[0] || textual[0]`, accidentally editing the slide title!
- **Fix**: Implemented strict lookup in `executeSetFont` and `executeSetShapeText`:
  ```javascript
  let targetShape;
  if (!shape_id || ['heading', 'title', 'slidetitle', 'main'].includes(lowerId)) {
    targetShape = titleShape || withText[0];
  } else {
    targetShape = shapes.items.find(s => s.id === shape_id || s.name.toLowerCase() === lowerId || s.id.toLowerCase() === lowerId);
  }
  if (!targetShape) {
    return { success: false, error: `Shape not found: "${shape_id}". Available shapes: ${available}. Call read_slide to verify.` };
  }
  ```

### 3. Native Table Support (Office.js API Set 1.8)
- **Bug**: Reading/writing PowerPoint tables threw runtime errors because early code attempted illegal assignments to `cell.textRuns = ...`.
- **Fix**: Switched to native `shape.getTable()`, `tbl.load('rowCount, columnCount, values')`, and `cell.font.color = font_color`, `cell.fill.setSolidColor(fill_color)`. Added `<Requirements><Sets DefaultMinVersion="1.8"><Set Name="PowerPointApi" MinVersion="1.8"/></Sets></Requirements>` to `manifest.xml`.

### 4. Cell Font Color Extraction in `read_slide`
- **Bug**: The model could see table text values but not which cells were green or black, preventing commands like *"change green remarks to black"*.
- **Fix**: In `executeReadSlide`, batched proxy loading loads `cell.font.color` for tables with up to 50 cells into `cellColors[row][col]`, giving Gemini full color telemetry.

### 5. In-Memory Undo Stack & UI Reversion
- **Feature**: Added `state.undoStack` storing `{ type, slideId, slideNumber, shapeId, description, previous, current }`. Added interactive `↩ Undo` toolbar button and `undo_last_edit` tool.

### 6. Prompt Injection Defense
- **Security**: Prevented untrusted document content from overriding system prompts by segregating slide telemetry inside `[SLIDE DATA CONTEXT - UNTRUSTED DOCUMENT CONTENT]` and requiring `user_confirmed: true` for destructive tools.

### 7. Target Slide Locking & Multi-Slide Targeting (`slide_number`)
- **Bug**: User asked: *"change color of headings in slide 6,7 to green"*. Copilot only changed Slide 6 and asked the user to manually click Slide 7 in the thumbnail pane!
- **Root Cause**: Tools were exclusively locked to `getSelectedSlides().getItemAt(0)`. There was no `slide_number` parameter in schemas, and the system prompt told Gemini it could only touch the active slide.
- **Fix**:
  1. Updated `getTargetSlide(context, targetSlideId, slide_number)` to access `context.presentation.slides.items[slide_number - 1]`.
  2. Added `slide_number` to all tool parameter schemas in `GEMINI_AGENT_TOOLS`.
  3. Added `read_deck_outline` tool.
  4. Updated `baseInstruction` with explicit directives to execute multi-slide edits in parallel in the same turn without asking the user to manually switch slides.

---

## 6. Automated Verification & Test Suites

The project has **16 passing automated tests** across three specialized test runners:
```bash
npm test
```

### 1. `tests/audit_suite.test.js` (5 tests)
- **Test 1**: Verifies syntax integrity, no `role: 'function'` regressions, clean cell API assignments, and core function exports.
- **Test 2**: Validates all 9 tools declared in `GEMINI_AGENT_TOOLS` with required `slide_number` parameter support.
- **Test 3**: Tests multi-table dimension & name matching (e.g. Slide 4 with Table 7 [2x3] and Table 8 [7x3]).
- **Test 4**: Validates title resolution across geometric shapes (banners), placeholders, and standard text boxes.
- **Test 5**: Tests parallel multi-tool execution payloads and 4-turn conversation chaining.

### 2. `tests/pptx_fidelity.test.js` (5 tests)
Tests real structures extracted from `Presentation2.pptx`:
- **Test A**: Slide 4 title resolution (`Rectangle 2`, GeometricShape).
- **Test B**: Remarks column resolution (6 data rows in Table 8).
- **Test C**: Slide 6 title resolution (`TextBox 3`, "TRAINING ON SMP").
- **Test D**: Slide 7 title resolution (`TextBox 6` at `top: 197319` identified as title over `TextBox 11` at `top: 1052736`).
- **Test E**: Multi-slide simultaneous tool call targets Slide 6 & 7 heading colors.

### 3. `tests/mock_office_tools.test.js` (6 tests)
- **Case 1**: Title placeholder priority and `set_font` styling.
- **Case 2**: Strict shape lookup (typo protection rejects without mutating title).
- **Case 3a & 3b**: Table batch update of cell font colors and full `Undo` restoration.
- **Case 4**: Prompt injection defense blocks unconfirmed shape deletions.
- **Case 5**: Multi-slide direct targeting simultaneously recolors Slide 6 & Slide 7 headings.
- **Case 6**: Out-of-bounds slide number safely caught and rejected.

---

## 7. Developer Workflow & Commands

### Running Locally
```powershell
# Navigate to project root
cd C:\Users\shiva\.gemini\antigravity\scratch\gemini-powerpoint-addin

# Install dependencies (if needed)
npm install

# Start local dev server with HTTPS
npm run dev
```

### Running Tests
```powershell
npm test
```

### Building for Production
```powershell
npm run build
```
Build output is generated in `dist/`.

### Deploying to Cloudflare Pages
```powershell
npx wrangler pages deploy dist --project-name=ppt-distillatelabs --branch=main --commit-dirty=true
```

### Validating Office Add-in Manifest
```powershell
npm run validate
```

---

## 8. Known Nuances, Edge Cases & Roadmap for Next AI Agent

Here are the specific areas and potential future enhancements for the next developer/agent:

### 1. Mid-Turn Slide Navigation Drift
- **Nuance**: If a user submits a prompt without specifying `slide_number` (e.g. *"make title blue"*), `targetSlideId` is locked to the selected slide at the moment the user clicks send. If the user clicks another slide in PowerPoint while Gemini is processing, subsequent tool iterations still target `targetSlideId` via `getItemOrNullObject(targetSlideId)`. If `targetSlideId` ever becomes invalid (e.g. slide deleted), it falls back to active selection.
- **Recommendation**: Keep prompt encouraging explicit `slide_number` or continue locking `targetSlideId`.

### 2. Large Deck Outline Scanning (>30 slides)
- **Nuance**: `read_deck_outline` iterates through all slides synchronously. In decks with 50+ slides, loading textFrames for every slide can take 3-5 seconds.
- **Improvement**: Add pagination parameters to `read_deck_outline({ start_slide: 1, count: 20 })`.

### 3. Table Creation via Agent Tools
- **Nuance**: `add_slide` currently generates styled shapes (title, cards, bullet boxes, takeaways), but does not create native PowerPoint tables (`slide.shapes.addTable(rows, cols, left, top, width, height)`).
- **Improvement**: Implement an `add_table` or `insert_table_slide` tool so Gemini can generate native editable tables when requested.

### 4. Undo Persistence Across Task Pane Reloads
- **Nuance**: `state.undoStack` is held in memory in JavaScript heap. If the user closes and re-opens the taskpane, the in-memory undo stack is reset.
- **Improvement**: Serialize `state.undoStack` to `sessionStorage` or Office document settings (`Office.context.document.settings.set`).

### 5. Native PowerPoint Dark Mode Theme
- **Nuance**: The task pane UI currently uses an executive light theme.
- **Improvement**: Add CSS `@media (prefers-color-scheme: dark)` or Office theme detection (`Office.context.officeTheme`) to automatically sync background and font colors with PowerPoint's Dark Mode.

---

## 9. Quick Verification Checklist for Any Incoming Agent

When continuing development:
1. `npm test` must output `ALL ... TESTS PASSED (16/16)`.
2. `npm run build` must compile `dist/` cleanly with no module errors.
3. Every new tool must be declared in `GEMINI_AGENT_TOOLS` with types, descriptions, and added to `TOOL_HANDLERS` in `runGeminiAgentLoop`.
4. All function response turns must keep `role: 'user'` (never `role: 'function'`).
5. Ensure `manifest.xml` retains `<Requirements><Sets DefaultMinVersion="1.8"><Set Name="PowerPointApi" MinVersion="1.8"/></Sets></Requirements>`.

---
*Document prepared on 2026-11-07 for autonomous handover.*
