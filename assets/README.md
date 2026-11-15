# 🎨 Assets

`assets/` holds the repository's images: the screenshots and artwork used by
`README.md` and the GitHub Pages landing page (`index.html`), and the social
preview.

## Social preview

[`social-preview.png`](social-preview.png) is the image GitHub shows when the
repository is shared, and the `og:image` of the landing page. It is rendered at
GitHub's recommended 1280×640 from the editable
[`social-preview.svg`](social-preview.svg). Change the SVG, then regenerate the
PNG; never edit the PNG alone.

The SVG uses two fonts, both under the SIL Open Font License and available from
Google Fonts: **Oswald Bold** for the title and **Inter** for the other text.
Install both before rendering; without them the browser substitutes other fonts
and the PNG changes.

Render with headless Chrome, Edge or Chromium from the repository root:

```text
chrome --headless --hide-scrollbars --force-device-scale-factor=1 --window-size=1280,640 --screenshot=assets/social-preview.png assets/social-preview.svg
```

On Windows, call the browser by its full path, for example
`"C:\Program Files\Google\Chrome\Application\chrome.exe"`, and pass absolute
paths for both files. Check that the result is 1280×640 and under GitHub's
1 MB limit, then upload it in **Settings → General → Social preview**. GitHub
offers no API for that step.

Keep assets free of personal or confidential information, and update every
reference when an asset moves.
