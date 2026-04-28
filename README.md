# Line Check Counter - Zebra QR Labels

A simple GitHub Pages app for printing Zebra/sticker paper QR labels and counting checked receiving lines.

## What it does

- Creates printable 4x2 Zebra-style QR labels
- QR includes PO/Transfer number, item number, quantity, description, and line ID
- Scanner marks each line checked once
- Tracks daily total checked lines
- Tracks selected worker's checked lines
- Exports checked line log to CSV

## GitHub Pages Setup

1. Create a new GitHub repo named `line-check-counter`
2. Upload these files:
   - `index.html`
   - `style.css`
   - `app.js`
3. Go to repo Settings
4. Click Pages
5. Source: Deploy from branch
6. Branch: `main`
7. Folder: `/root`
8. Save

Your app link will look like:

`https://YOURUSERNAME.github.io/line-check-counter/`

## Zebra Printing

This version uses browser printing.

Recommended print settings:

- Paper/label size: 4 in x 2 in
- Scale: 100%
- Margins: None or Minimum
- Headers/Footers: Off

If your Zebra printer is installed on Windows as a normal printer, this should print like a standard label.

## Bulk Paste Format

One line per label:

PO, Item, Qty, Description

Example:

463090, SXFR0202592, 4, 36 x 80 PVC Louver Bifold Door
