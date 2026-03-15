# Assignment — Spreadsheet Application

A browser-based spreadsheet built with React + Vite. No external spreadsheet libraries — the formula engine, dependency tracker, and all spreadsheet logic are written from scratch.

## How to Run

```bash
npm install
npm run dev
```

## What's Built

- **Formula engine** — supports `=A1+B1*2`, `=SUM(A1:A5)`, `=AVG(...)`, `=MIN(...)`, `=MAX(...)`, nested parentheses, operator precedence
- **Dependency graph** — tracks which cells depend on which, recalculates in topological order, detects circular references (`#CYCLE!`)
- **Sort & filter** — click column headers to sort (asc/desc/none), filter dropdown with checkbox selection per column
- **Copy/paste** — supports multi-cell selection, tab-separated format (compatible with Excel/Google Sheets clipboard)
- **Undo/redo** — Ctrl+Z / Ctrl+Y, including single-step undo for batch paste operations
- **Insert/delete rows & columns** — right-click context menu, auto-shifts all formula references
- **Persistence** — auto-saves to LocalStorage with debouncing, restores on reload
- **Cell styling** — bold, italic, text color, background color per cell

## Key Design Decisions

### 1. Custom parser instead of `eval()`

Formulas go through three stages: **tokenize → parse (shunting-yard) → evaluate AST**. This avoids any use of `eval()` or `Function()`, which would be a security risk and makes error handling unpredictable. The shunting-yard algorithm handles operator precedence naturally without needing recursive descent.

*Relevant code:* `src/engine/core.js` — `tokenize()` (line 163), `parseTokensToAST()` (line 232), `evaluateAST()` (line 353)

### 2. Dependency graph with cycle detection

Every formula's cell references are tracked in a directed graph (forward + reverse edges). When a cell changes, we find all transitive dependents and recalculate them in **topological order** — so a cell is never evaluated before its dependencies. Circular references are caught via DFS before they cause infinite loops.

*Relevant code:* `src/engine/core.js` — `createDependencyGraph()` (line 41)

### 3. Sort and filter as a view layer, not data mutation

Sorting and filtering **do not move or hide actual cell data**. Instead, the engine computes a `viewRowOrder` — an array of row indices in display order. The React grid maps over this array instead of iterating `0..N`. This means formulas like `=A3+A5` always refer to the real row 3 and row 5, regardless of how the view is sorted.

*Relevant code:* `src/engine/core.js` — `getViewRowOrder()` (line 1005)

### 4. Batch operations with single undo entry

Pasting 50 cells creates one undo entry, not 50. The engine's `executeBatchSet()` records all old values upfront, applies all changes, then pushes a single `batch` entry to the undo stack. This keeps undo/redo behavior intuitive.

*Relevant code:* `src/engine/core.js` — `executeBatchSet()` (line 919)

### 5. Formula references shift on structural changes

When you insert or delete a row/column, every formula in the sheet is scanned and its cell references are shifted accordingly. For example, inserting a row above row 3 turns `=A3` into `=A4` in all formulas below the insertion point.

*Relevant code:* `src/engine/core.js` — `shiftCellReferences()` (line 431)

### 6. Engine/UI separation

The engine (`core.js`) is a pure data layer — no React, no DOM. It exposes a clean API (`setCell`, `getCell`, `undo`, `insertRow`, etc.) and the React component (`App.jsx`) is purely responsible for rendering and dispatching user events. This makes the engine independently testable and the UI easy to swap.

## Project Structure

```
src/
  engine/
    core.js      — Formula engine, dependency graph, sort/filter, serialization
  App.jsx        — React UI component (grid, toolbar, context menus, keyboard handling)
  App.css        — All styling
```
