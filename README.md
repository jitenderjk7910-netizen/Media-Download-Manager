# Media Download Manager

Single GUI tool for mixed brand links: direct image URLs, public image folders, Dropbox, Google Drive, SharePoint, OneDrive, ImgBB, and generic HTTP links.

## Input Format

Paste rows or load Excel/CSV.

```text
BarcodeOrFolder    Link1    Link2    Link3
ABC001             https://...jpg    https://drive.google.com/...
ABC002             https://dropbox.com/...    https://ibb.co/...
```

The first non-link column becomes the output folder name. Every URL in the row is downloaded into that folder.

## Working

- Direct image/file link: downloads the file.
- Zip/folder download: extracts the zip into the folder.
- Dropbox: converts public shared folder links to direct zip download, then extracts them.
- Google Drive: uses `gdown` for folders where possible, and direct Google download fallback for `open?id=...` files.
- SharePoint/OneDrive: reuses the existing SharePoint downloader core in this workspace.
- ImgBB/generic image pages: opens the page HTML and downloads image links found in `og:image`, `src`, or `href`.
- Every run creates `download_report.xlsx` in the output folder.

## Limitations

No public downloader can bypass private permissions. If Google Drive, SharePoint, OneDrive, Dropbox, or any website requires login/access approval, the tool will fail or report a login/HTML page. For such cases, make the link public/downloadable, or use a browser-session fallback for that platform.

Dropbox note: folder links like `https://www.dropbox.com/scl/fo/...&dl=0` are supported. The tool switches them to `dl=1` and uses a downloader-style request so Dropbox returns a zip instead of the web page.

## Recommended Run

Use the browser dashboard:

```bat
run_web_ui.bat
```

It opens `http://127.0.0.1:8080/` and uses the same downloader backend.
Running `python universal_link_downloader.py` also opens the modern web dashboard.

Modes:

- Excel / CSV Batch: select a file and preview detected links before downloading.
- Manual Link Test: test one folder name with one or more links.
- Search & Download: paste Koton style codes/search terms; the app searches product images and downloads them by style folder.

Main controls:

- Preview Queue: checks link count, folder count, and platform split.
- Pause / Resume / Stop: controls the active queue.
- Retry Failed Only: reads `download_report.xlsx` and retries failed links only.
- Skip Existing: skips folders that already contain downloaded files.
- File Mode: choose all files, images + videos, or images only.
- Folder Naming: choose folder-only, row+folder, or platform+folder naming.
- Result Filters: filter table by all, OK, failed, skipped, or search text.
- Failure Groups: groups failures by login/access, timeout, not found, page/not direct, no media, and other.
- Browser Fallback: after direct/public download fails, tries a logged-in Chrome/Edge browser download button pass.
- Output Audit: shows how many expected folders contain files and which are empty/missing.
- Preview Thumbnails: shows the first image found in each completed folder.
- Resume From Report: skips links already marked OK in `download_report.xlsx`.
- Image Quality: choose best available image or all discovered sizes for search modes.
- Auto Fixes: keeps platform link normalization enabled by default.
- Create Template: creates a starter Excel template in the output folder.
- Open Output / Open Report: opens the save folder or `download_report.xlsx`.

## Classic Tkinter Run

1. Double-click `install_requirements.bat` once.
2. Double-click `run_gui.bat`, or run `python universal_link_downloader.py --classic-gui`.

CLI:

```bat
python universal_link_downloader.py --cli links.xlsx downloads
```
