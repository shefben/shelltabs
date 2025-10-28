# ShellTabs FTP Namespace Manual Test Plan

This test plan validates ShellTabs FTP navigation, namespace registration, and transfer workflows. Execute the following suites
on a clean Windows Explorer session with the updated ShellTabs build installed via `regsvr32 /i`.

## 1. Credential Management
- Launch Explorer and open the **Shell Tabs FTP Sites** namespace node in the navigation pane.
- Connect to an FTP host that requires anonymous access and confirm no credential prompt appears.
- Connect to an FTP host that requires basic credentials:
  - Verify the credentials prompt appears.
  - Test "Remember my credentials" toggle; reconnect to confirm persistence.
  - Use incorrect credentials and ensure retry prompt surfaces with the correct error message.
- Clear stored credentials through Windows Credential Manager and confirm ShellTabs reprompts on the next connection.

## 2. Permissions & Directory Operations
- Navigate through nested directories on a host where the account has read/write permissions; confirm folder expansion in the
  navigation tree and item enumeration in the view.
- Attempt to enter a directory without permissions; verify access-denied errors surface and no phantom tabs remain.
- Rename a writable directory and confirm the tab display name updates; repeat on a read-only directory and ensure the operation
  fails gracefully.

## 3. Upload & Download Scenarios
- Download a small file and confirm the transfer result shows the correct byte count and that the file opens locally.
- Upload a new file to a writable directory and refresh to confirm it appears with accurate metadata.
- Attempt to overwrite an existing file and validate the overwrite prompt occurs; cancel to ensure the original file remains
  unchanged.
- Queue sequential uploads and downloads in different tabs to confirm session reuse and isolation across tabs.

## 4. Offline & Recovery Behavior
- While connected to an FTP site, disconnect the network adapter:
  - Validate ShellTabs surfaces a connection-lost message and tabs remain responsive.
  - Reconnect the adapter and retry navigation; ensure automatic reconnection succeeds.
- Close and reopen Explorer; verify pinned FTP namespace nodes are restored and that reconnecting to recent hosts succeeds
  without stale session errors.

## 5. Error Handling & Edge Cases
- Navigate to an invalid FTP host name and confirm Explorer reports name resolution failure while leaving existing tabs intact.
- Force an SSL/TLS negotiation error on an FTPS endpoint and verify the resulting message indicates the secure transport issue.
- Attempt uploads to a full quota account and ensure the reported error reflects insufficient storage.
- Test cancellation of long-running transfers from the ShellTabs UI and confirm temporary files are cleaned up.
- Remove ShellTabs via `regsvr32 /u /i` and ensure all FTP namespace nodes disappear from Explorer.

Document observed results, screenshots of failures, and any unexpected log entries for triage.
