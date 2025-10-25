# Explorer Behavior, Architecture, and Reliability Reference

## Explorer Behavior Reference

### View Columns
- **Column Sets per Namespace:**
  - Default filesystem columns (Name, Date Modified, Type, Size) plus optional ones (Attributes, Owner) to mirror historical Explorer layouts from Windows Vista/7 era.
  - Allow shell namespace extensions to expose custom columns via `IColumnProvider` or the column handler registration; persist user column choices in `HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Streams`.
- **Auto-Sizing & Sorting:**
  - Support double-click auto-size, persisted sort column/direction, and grouping metadata (e.g., "Date Modified") consistent with `FOLDERFLAGS` defaults.
- **Per-Folder Customization:**
  - Honor folder templates (General Items, Documents, Pictures, Music, Videos) resolved through `GetDetailsEx` / `SHGetDetailsOf` to map to the correct column set and view mode.

### Verbs & Commands
- **Context Menu Verbs:**
  - Populate via `IContextMenu` and canonical verbs (`open`, `explore`, `properties`, custom extension verbs).
  - Integrate the standard command routing (accelerators, toolbar buttons) by using `IExplorerCommand` implementations where possible for Win7+ parity.
- **Keyboard & Ribbon Integration:**
  - Preserve Explorer hotkeys (`F2`, `Ctrl+C/V`, `Alt+Enter`) by mapping to commands on `IShellView`.
  - When hosted in Explorer frame, ensure `IShellBrowser::SetToolbarItems` participates in the command target chain.

### Drag & Drop
- **Source/Target Contracts:**
  - Implement `IDropTarget` on views and `IDropSource` for originating drags, negotiating `CFSTR_SHELLIDLIST`/`CF_HDROP`.
  - Support async file operations (IFileOperation/ITransferSource) for long transfers to keep UI responsive.
- **Visual Feedback:**
  - Provide standard `DROPEFFECT` cursor semantics and overlay insertion point for list views; honor `DragDropHelper` for ghosted icons.
- **Shell Integration:**
  - Call `SHDoDragDrop` to inherit system behaviors (right-drag menu, copy/move/link default logic).

### Preview Pane & Ancillary UI
- **Preview Handlers:**
  - Instantiate registered preview handlers via `IPreviewHandler` and the Preview Handler Framework based on file type association.
- **Details Pane / Info Tips:**
  - Query properties with `IPropertyStore` to populate the details pane and infotips (hover).
- **Thumbnail & Icon Caching:**
  - Use `ISharedBitmap`/`IThumbnailCache` to request cached thumbnails; fallback to `IShellItemImageFactory`.

---

## COM Architecture & Threading

### Threading Model
- **Apartment-Threaded (STA):**
  - Explorer historically hosts shell extensions in STA; continue this to align with UI components and COM reentrancy expectations.
  - Use a dedicated UI thread that pumps messages for `IShellView` and associated window handles.
  - Offload blocking operations to background MTA threads while marshaling results back to STA via `CoMarshalInterThreadInterfaceInStream` or simple `SendMessage/PostMessage`.

### Core Interfaces & Relationships
```
+---------------------+          +---------------------+
| Shell Frame (Host)  |<>--------| IShellBrowser       |
+---------------------+          +---------------------+
           ^
           |
           v
+---------------------+   implements   +---------------------+
| Folder Instance     |--------------->| IFolder (custom)    |
|  - owns IShellView  |                | IShellFolder        |
|  - caches PIDLs     |                | IPersistFolder      |
+---------------------+                | IEnumIDList         |
           |                           +---------------------+
           | creates/owns
           v
+---------------------+   uses         +---------------------+
| View (list/tree)    |--------------->| IDataObject         |
|  - HWND             |   (selection)  | IContextMenu        |
|  - Drop target      |                | IBindStatusCallback |
+---------------------+                +---------------------+
```

- **`IPersistFolder`:** Initialize folder with root PIDL; maintain canonical namespace identity.
- **`IShellFolder`:** Enumerate children (`EnumObjects`), bind to sub-folders (`BindToObject`), resolve display names (`GetDisplayNameOf`), attribute flags.
- **`IFolder` (custom helper):** Encapsulate domain logic (e.g., caching, background worker coordination) to keep COM interface implementations slim.
- **`IShellView`:** Manage view window lifecycle, selection changes, command routing; hosts list/tree controls.
- **`IEnumIDList`:** Backed by enumerator objects fed from cache/worker thread; supports `SHCONTF` filters and cancellation.
- **`IBindStatusCallback`:** For asynchronous bindings (e.g., network folders) and progress UI integration.

---

## Caching & Background Work

### Directory Listing Cache
- **Shared Cache Service:**
  - Centralize enumeration results keyed by absolute PIDL/`IShellItem`. Maintain metadata (timestamp, item attributes, error state).
  - Use reader-writer locks or version tokens to allow concurrent reads from UI while background updates refresh data.

### Background Worker Model
- **Producer/Consumer Queue:**
  - UI thread posts enumeration jobs (folder PIDL + filters) to a worker pool (MTA).
  - Workers perform filesystem/namespace enumeration using `IShellFolder::EnumObjects`, coalescing duplicate requests and honoring cancellation via `IEnumIDList::Reset/Skip`.
- **Async Updates:**
  - Workers marshal results back to UI via posted messages or `IDataObject` notifications, allowing incremental updates (e.g., chunked batches).
  - Support change notifications through `SHChangeNotifyRegister` to invalidate caches reactively.

### Synchronization & Freshness
- **Staleness Policy:**
  - Define TTL for cached results (e.g., 5 seconds for local disk, longer for network).
  - Provide explicit refresh command (`F5`) bypassing cache.
- **Error Propagation:**
  - Cache error states (access denied, offline) with last attempt timestamp to avoid spamming prompts; surface to UI with actionable messages.

---

## Error Handling & Credentials

### Authentication & Prompts
- **SSO First:**
  - Attempt implicit authentication using current user’s token (Windows Integrated Auth, cached creds).
  - For SMB/SharePoint/webdav scenarios, leverage `CredUIPromptForWindowsCredentials` when access is denied and no valid token exists.
- **Credential Persistence:**
  - Offer “Remember my credentials” via `CredWrite` to Windows Credential Manager; respect enterprise policy (Group Policy).
  - Support per-target credential caching keyed by UNC or URL.

### UI/UX for Errors
- **Inline Notifications:**
  - Display toast/banner in view for transient errors; double-click to retry or open Credential UI.
- **Standard Dialogs:**
  - Use `SHOutOfMemoryMessageBox`, `SHRestrictedMessageBox`, and `IUserNotification2` where appropriate for consistency with Explorer.

### Retry & Offline Handling
- **Graceful Degradation:**
  - If network resource unavailable, show placeholder items with offline glyph; allow explicit retry.
- **Background Retry Logic:**
  - Implement exponential backoff for repeated failures; avoid blocking UI thread while credential prompts are open.

### Logging & Telemetry
- **Diagnostics:**
  - Instrument key error paths (failed bindings, enumeration timeouts) to ETW or Windows Event Log for supportability.
- **Privacy/Security:**
  - Ensure logs redact credentials and sensitive paths; comply with enterprise logging standards.

---

## Recommendations Summary
1. Mirror historical Explorer behaviors to ensure user familiarity—persist view state per folder, respect shell verbs, and use system drag/drop helpers.
2. Adopt an STA COM threading model with background MTA workers to keep UI responsive while aligning with Explorer extension expectations.
3. Build a layered caching system with asynchronous workers and change notifications to deliver fast, incremental updates.
4. Integrate credential handling that prefers SSO but gracefully falls back to Windows credential UI, with robust error messaging and logging.
