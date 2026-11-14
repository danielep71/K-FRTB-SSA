# 🌐 Project website

The public address is <https://danielep71.github.io/K-FRTB-SSA/>. The page
says what the workbook calculates, how to build it and what has and has not
been checked. It runs no calculation.

## 📦 Files

| File | Holds |
| --- | --- |
| [`index.html`](../index.html) | The page, at the repository root, where GitHub Pages looks for it |
| [`docs/assets/site.css`](assets/site.css) | Its stylesheet |
| [`docs/assets/social-preview.png`](assets/social-preview.png) | The image of the page's first screen, also the image of shared links ([images](assets/README.md)) |
| `.nojekyll` | An empty file at the repository root that turns Jekyll off |

The page needs no JavaScript, no third-party font, no package installation
and no build step.

## 🚀 Publication

GitHub Pages serves the site from the `main` branch, folder `/ (root)`:
**Settings → Pages → Build and deployment → Deploy from a branch**. A change
goes live only after it is merged to `main`. Pages is outside the
[settings baseline](setup/REPOSITORY_SETUP.md), so
**Detect settings drift** does not read it: check the setting and the public
address yourself.

Jekyll must stay off. The repository's Markdown holds text that Jekyll would
read as template code, such as the `${{ … }}` expressions of the workflow
guides, and a Jekyll build of it can fail. With `.nojekyll`, Pages publishes
the files as they are.

Pages publishes every file of the branch, not only the page: each tracked
file is reachable under the public address. The repository is public, so
this exposes nothing new; keep it in mind before the visibility changes.

## 🧹 Maintenance

- **Status.** The status box and the *Evidence and boundaries* section must
  say what the [README](../README.md#-status) says. Change them in the same
  pull request that changes the README's status.
- **Figures.** Every figure on the page is a MAR40 parameter that the `Params`
  sheet of `src/workbook/Template.xlsx` holds. The page shows no result.
- **Scope.** The list of what is not implemented follows the `README` sheet of
  the template.
- **Links.** `python tools/check_links.py` reads Markdown only. After a
  change, check by hand that each relative link of the page names a file of
  the repository and that each link into the repository names a path on
  `main`.
- **Rendering.** Open `index.html` in a browser, at a desktop and at a phone
  width: no horizontal scrolling, a visible focus ring on every link with
  the keyboard, and a text alternative on the image.
