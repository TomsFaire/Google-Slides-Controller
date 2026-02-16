# Speaker notes encoding (U+FFFD / broken line breaks)

## Where the text comes from

1. **Notes window (Electron)**  
   The speaker notes **window** is a BrowserWindow opened by Google Slides when you start presenter view. We don’t create its content. We only:
   - Call `getSpeakerNotesWindowOptions()` and pass that to `setWindowOpenHandler` so the popup uses our size/position.
   - The window loads a Google URL; Google’s HTML/JS build the DOM, including the notes div (`div.punch-viewer-speakernotes-text-body-scrollable`).

2. **Our only use of that content**  
   We read the notes in one place: **GET /api/get-speaker-notes**. There we:
   - Run `executeJavaScript` in the notes window to get `el.innerText` (or `textContent`) from that div.
   - The string we get **already contains U+FFFD** (replacement character) where line breaks should be—that’s what’s in the DOM when we read it.
   - We then call `normalizeSpeakerNotes(rawNotes)` (replace U+FFFD, etc., with newlines) and send the result in the API response.

3. **Web UI**  
   The Web UI requests `/api/get-speaker-notes` and displays the **normalized** response, so it shows correct line breaks.

4. **Electron notes window**  
   What the user sees and copies in the **notes window** is the raw DOM content that **Google Slides** put there. We never write to that div; we only read from it for the API. So the “corruption” (U+FFFD instead of newlines) happens **upstream** of our code: when Google’s page is loaded and their JS fills the notes div.

## Conclusion

- There is **no earlier step in our process** where we create or modify the notes text before it’s shown in the notes window. We don’t inject or set that content.
- The only place we touch it is **reading** it for the API and normalizing for the Web UI.
- So the broken characters in the **Electron notes window** come from how the Google Slides page (and possibly its response encoding) produces that DOM.

## What we can try upstream

- **Session / response encoding**  
  We could try to ensure the Google session treats responses as UTF-8 (e.g. via `session.webRequest.onHeadersReceived` for the Google partition and forcing or adding `charset=utf-8` on text responses). That might help if the issue is a missing or wrong charset on Google’s responses.  
  This is implemented in `main.js` by calling `setupGoogleSessionEncoding()` from `app.whenReady()`.

- **Fixing only in the window**  
  Any fix that changes what the user sees in the notes window (e.g. a script that rewrites text nodes in that div) would run **in** that window and was rolled back because it caused other issues (notes not advancing with slides, or characters deleted instead of replaced). Re-enabling such a fix would require making it robust (e.g. correct newlines, no full-div replacement so slide updates still work).
