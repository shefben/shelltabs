# Tab Seeding Audit

This note captures the current behavior of `TabBand` when Explorer spawns or reuses
windows and how pending window seeds are consumed. It also enumerates the failure
scenarios we observed prior to introducing explicit window identifiers in the
`TabManager`.

## Pending Window Seed Flow

1. When a tab or group is detached (`OnDetachTabRequested` /
   `OnDetachGroupRequested`), the source band enqueues a `PendingWindowSeed`
   entry and calls `BrowseObject` to let Explorer materialize a new frame.
2. The next `TabBand` instance created within the process dequeues the oldest
   pending seed during `InitializeTabs()` and uses it to hydrate the new window's
   `TabManager`.
3. If no seeds exist, the newly created band falls back to querying the current
   folder and seeds a default tab for the active view.

## Failure Cases Identified

- **Late `NavigateComplete` on recycled frames.** Explorer sometimes recycles a
  top-level window rather than creating a fresh instance. When this happens, the
  old band's `NavigateComplete` can arrive after the frame has been reused. The
  recycled band would dequeue a seed that belonged to the previous window and
  incorrectly apply it to the new navigation.
- **FIFO queue mismatch across concurrent launches.** The global deque assumes
  windows will materialize strictly in the order they were requested. Explorer
  may reorder or cancel window creation, leaving seeds stranded in the queue or
  assigned to the wrong frame on the next attach.
- **Session teardown leaking seeds.** When a recycled frame detaches without
  clearing its tab set, the next band to consume the queue would inherit stale
  tabs that no longer belong to any live Explorer surface.

## Mitigation

The updated implementation tracks a concrete window identifier (frame `HWND`
plus the connection cookie derived from the `IWebBrowser2` automation pointer)
inside `TabManager`. `TabBand::DisconnectSite` now clears the registration for
that identifier so recycled frames start with a clean slate. Future work will
match pending seeds to these identifiers so the deque can be drained safely even
under heavy reuse.
