# downloads/

Place the Windows installer here and it will be served by the landing page
`index.html`'s **Download for Windows** button.

## Expected filename

The button currently links to:

    downloads/PDF2Excel-win64.exe

Build it with:

    scripts/build_windows.bat

which produces `dist\PDF2Excel.exe`. Copy or rename that file to
`web/downloads/PDF2Excel-win64.exe` (or adjust the href in `web/index.html`).

## Notes

- Everything in this folder is committed/uploaded when you deploy the site to
  GitHub Pages, Netlify, or any static host — so keep the installer small if
  you can (the PyInstaller onefile build is a few tens of MB).
- For a proper `Setup.exe` with shortcuts and an icon, wrap the `.exe` with
  NSIS or Inno Setup after building (see `README.md`).
