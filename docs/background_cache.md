# Background Cache Maintenance

ShellTabs caches folder background images on disk so Explorer can reuse them without re-reading the original files. The cache
maintenance sweep removes stale entries by comparing the files on disk against the currently referenced image list and by
expiring files that have not been touched recently.

## Throttled purge scheduling

`UpdateCachedImageUsage` now throttles cache purges so repeated calls within five minutes reuse the most recent sweep. The
function records the last successful purge timestamp and skips new work unless the throttle window has expired or the caller
explicitly requests a forced run. This prevents UI-heavy code paths, such as option saves, from blocking the Explorer thread on
redundant directory scans.

When a purge is eligible, ShellTabs queues the work on a background thread so UI calls return immediately. The helper guards the
worker with a mutex so only one purge runs at a time and updates the last-run timestamp after completion.

## Manual diagnostics hook

For diagnostics or maintenance tooling, call `ForceBackgroundCacheMaintenance(const ShellTabsOptions&)`. The helper bypasses the
throttle, runs synchronously on the calling thread, and still serializes execution with in-flight background sweeps. This is
useful when a test or support scenario needs to validate cache cleanup deterministically.
