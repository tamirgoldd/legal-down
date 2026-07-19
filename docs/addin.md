# Install the Word add-in

Word Order's hosted add-in is an alpha developer build. It runs the same local,
deterministic OOXML engine as the web app; Word Order has no server that accepts
the open document or its text.

## Install the current hosted version on Mac

1. Download the [current Word Order manifest](https://tamirgoldd.github.io/word-order/addin/manifest.xml).
2. Quit Word.
3. In Finder, press **Command+Shift+G** and open
   `~/Library/Containers/com.microsoft.Word/Data/Documents/wef`.
4. Remove the old `Legal Down` manifest if it is present, then copy the new
   `manifest.xml` into that folder.
5. Reopen Word, open a document, and choose **Home → Add-ins → Word Order**.

These are the sideloading steps Microsoft documents for add-in-only manifests
on Word for Mac. Microsoft 365 administrators can deploy the same manifest
through Integrated Apps after choosing a stable hosted origin.

## Fix a “Legal Down” 404 after the rename

The old manifest points to `https://tamirgoldd.github.io/legal-down/addin/`.
That GitHub Pages path stopped existing when the repository became Word Order.
The current manifest is version `0.2.0.0` and points to
`https://tamirgoldd.github.io/word-order/addin/`.

Replacing the file in the `wef` folder and restarting Word is normally enough.
If Word still shows the old name or page, clear the Office add-in cache and then
copy the current manifest into the `wef` folder again. Microsoft's supported
command is:

```bash
npx office-addin-cache clear
```

This clears sideloaded Office add-ins, not just Word Order. Review Microsoft's
[Office cache guide](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/clear-cache)
before using it. You can also follow Microsoft's
[Mac sideloading guide](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac).

## Build against another HTTPS origin

Build the repository, serve `packages/addin/dist` over HTTPS, and change the
icon and `SourceLocation` URLs in `packages/addin/manifest.xml` to that origin.

The task pane scans before applying and requires confirmation that a backup was
saved. Word's OOXML insertion APIs differ by host version, so the add-in reports
unsupported hosts rather than attempting a partial repair.
