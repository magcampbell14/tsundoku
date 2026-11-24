# Night Shelf

Maggie's reading tracker: an owned-books shelf that lights up as books get read, a yearly goal, streaks, and past-year shelves.

- `index.html` – the whole app
- `sw.js`, `manifest.json`, `icon-*.png` – home-screen app and offline support
- `sheet-sync.gs` – Google Apps Script that saves the app's data to a Google Sheet (paste into the Sheet's Extensions > Apps Script and deploy as a web app)

Data lives on the phone and is saved to the Google Sheet after every change.
