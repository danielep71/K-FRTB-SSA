# 🖼️ Documentation images

The images in this directory illustrate K-FRTB-SSA. They do not certify the
workbook or replace the [methodology](../methodology/README.md).

## 🪧 Social preview

`social-preview.png` (1280 x 640) is rendered from `social-preview.html`. The
README shows it. To use it as the social preview too, upload the same file in
**Settings → General → Social preview**: GitHub offers no API for it.

It shows the stage, *Prototype*, and that the workbook has not been run in
Excel yet, and no result of a calculation. Its only figures are the four scaling
factors of the simplified standardised approach, 1.30 for interest rate risk,
3.50 for equity risk, 1.90 for commodity risk and 1.20 for foreign exchange
risk. They are those of MAR40.2 of `SRC-02` in the
[sources](../methodology/SOURCES.md), with which I compared the image on
2026-11-16.

To render it again (Chromium, then crop to the 1280 x 640 frame):

```sh
chromium --headless=new --hide-scrollbars --window-size=1280,900 \
  --virtual-time-budget=8000 --screenshot=preview.png \
  file://$PWD/docs/assets/social-preview.html
convert preview.png -crop 1280x640+0+0 +repage docs/assets/social-preview.png
```

The page loads its fonts from Google Fonts, so rendering needs network access.
After rendering, update the image's SHA-256 in
[`.github/binary-allowlist.json`](../../.github/binary-allowlist.json):
**Repository integrity** fails while they differ.

Update the image when the status it states changes, such as *Prototype* or
the statement that the workbook has not been run in Excel. Show a result figure
only if it is an expected value that a [test case](../methodology/TEST_CASES.md)
registers, with its class.
