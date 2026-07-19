# Install the Word add-in

Word Order's hosted add-in is an alpha developer build. It runs the same local,
deterministic OOXML engine as the web app; Word Order has no server that accepts
the open document or its text.

## Install the current hosted version on Mac

1. Download the [current Word Order manifest](https://tamirgoldd.github.io/word-order/addin/manifest.xml).
2. Quit Word.
3. In Finder, press **Command+Shift+G** and open
   `~/Library/Containers/com.microsoft.Word/Data/Documents/wef`.
4. Copy the downloaded `manifest.xml` into that folder.
5. Reopen Word, open a document, and choose **Home → Add-ins → Word Order**.

These are the sideloading steps Microsoft documents for add-in-only manifests
on Word for Mac. Microsoft 365 administrators can deploy the same manifest
through Integrated Apps after choosing a stable hosted origin. Microsoft's
[Mac sideloading guide](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac)
contains the platform-specific details.

## Build against another HTTPS origin

Build the repository, serve `packages/addin/dist` over HTTPS, and change the
icon and `SourceLocation` URLs in `packages/addin/manifest.xml` to that origin.

The task pane scans before applying and requires confirmation that a backup was
saved. Word's OOXML insertion APIs differ by host version, so the add-in reports
unsupported hosts rather than attempting a partial repair.
