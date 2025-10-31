# Explorer Pane Hooking Design

## Goals
- Subclass the navigation tree (`SysTreeView32`) and folder view (`SysListView32`) panes that Explorer hosts so ShellTabs can participate in drawing events.
- Provide a reusable routing layer that normalizes the custom-draw, selection-change, and other notifications that originate from those panes.
- Expose a registry-style API that tab modules can use to assign foreground/background colours for specific filesystem paths or tags.

## Hook attachment model
1. **Discovery** – `CExplorerBHO` already locates the view window, list view, and tree view handles when a shell view activates. Once the handles are found the BHO calls `PaneHookRouter::SetListView` / `SetTreeView` to associate them with the router.
2. **Subclass participation** – the BHO continues to own the `SetWindowSubclass` lifecycle for Explorer windows. Inside the shared subclass procedure we forward every `WM_NOTIFY` payload to the router before processing other ShellTabs features.
3. **Notification coverage**
   - `SysListView32`: handle `NM_CUSTOMDRAW` (prepaint + per-item/sub-item) so we can set `clrText`/`clrTextBk` for highlighted entries. The router returns `CDRF_NOTIFYITEMDRAW | CDRF_NOTIFYSUBITEMDRAW` on `CDDS_PREPAINT` and emits `CDRF_NEWFONT` when we actually adjust colours.
   - `SysTreeView32`: handle `NM_CUSTOMDRAW` (prepaint + item-prepaint) to push per-node colours, and `TVN_SELCHANGED` to trigger invalidation so custom backgrounds repaint when selection moves.
4. **Detachment** – when the subclass receives `WM_NCDESTROY` for a pane, the router is cleared for that handle. `RemoveExplorerViewSubclass` also calls `PaneHookRouter::Reset` to guarantee a clean state during navigation or band teardown.

## Highlight data flow
- A dedicated `PaneHighlightRegistry` (implemented in `PaneHooks.cpp`) stores highlight definitions keyed by normalized filesystem paths. The registry exposes `RegisterPaneHighlight`, `UnregisterPaneHighlight`, `ClearPaneHighlights`, and `TryGetPaneHighlight` helpers.
- `PaneHookRouter` does not perform PIDL or filesystem translation. Instead it requests highlight information via the `PaneHighlightProvider` interface:
  - `TryGetListViewHighlight(HWND listView, int itemIndex, PaneHighlight* out)`
  - `TryGetTreeViewHighlight(HWND treeView, HTREEITEM item, PaneHighlight* out)`
- `CExplorerBHO` implements the provider contract. It resolves PIDLs from list/tree items, maps them to filesystem paths, and queries the registry. This keeps hooking logic focused on message flow while allowing tab/colour subsystems to populate the registry without Explorer-specific knowledge.

## Integration points for tabs and tagging
- Tab management or tagging modules can register highlight entries whenever a tab activates or tag assignments change. The registry methods live in `PaneHooks.h` so they are available across ShellTabs components.
- When Explorer navigates to a folder the router automatically updates the colours on the next paint pass. No explicit repaint messages are required unless the selection changes, which the router handles by invalidating the tree view.

## Testing strategy
- Unit tests under `tests/PaneHooksTests.cpp` exercise the router without needing real Explorer controls. We construct synthetic `NMLVCUSTOMDRAW` / `NMTVCUSTOMDRAW` structures and a mock provider to verify that:
  - `CDDS_PREPAINT` stages request item callbacks.
  - Highlight colours propagate into the draw structures for the correct indices.
  - Non-highlighted items fall back to `CDRF_DODEFAULT` behaviour.
- These smoke tests compile under the existing `SHELLTABS_BUILD_TESTS` option and validate the routing logic in isolation.

## Future work
- Extend the registry to support font overrides and selection-aware theming.
- Introduce asynchronous invalidation (e.g., via `PostMessage`) when tag definitions change while a view is visible.
- Add COM-based hooks (`INameSpaceTreeControlCustomDraw`) for scenarios where Explorer hosts DirectUI wrappers instead of classic common controls.
