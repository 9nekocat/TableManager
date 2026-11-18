## Why

index.html >1600 lines. CSS/HTML/JS single file. Global state scattered. Change one feature — edit three places. Interaction complexity (drag/dropdown/panel/responsive/autohide) outgrew vanilla JS. Plain tool UI lacks game-project visual identity.

## What Changes

1. **React refactor**: Vite + React components. Single file → 8-10 components. React Context state.
2. **Responsive**: <700px auto-hide sidebar, flat table list, min_size → 400
3. **Always-on-top**: toolbar pin toggle `alwaysOnTop`
4. **Edge autohide**: snap edge → hide on mouse-out, expand on mouse-in (4px grab)
5. **Game UI theme**: character art (`art/charater.png`) + title font (`art/Title_Font.ttf`); Canvas particle BG; dark sci-fi palette

## Capabilities

### New

- `react-refactor`: Vite + React components, Context state
- `responsive-sidebar`: narrow auto-hide, flat list
- `always-on-top`: pin window topmost
- `edge-autohide`: edge snap hide/expand
- `game-ui-theme`: charater art + title font + particles + sci-fi

### Modified

(none — all existing features kept)

## Impact

| Layer | Impact |
|-------|--------|
| New `gui/` | Vite + React: `src/`, `public/`, `vite.config.ts` |
| `templates/index.html` | **Remove** — replaced by `gui/dist/` |
| `table_server.py` | Add window API (`always-on-top`, `autohide`); static → `gui/dist/` |
| `TableManager.spec` | datas: `('gui/dist', 'templates')`; add `('art', 'art')` |
| `art/` | Charater art + title font → bundled in EXE |
