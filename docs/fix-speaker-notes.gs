/**
 * fix-speaker-notes.gs
 *
 * Google Apps Script to batch-clean speaker notes in a Google Slides presentation.
 *
 * WHY:
 * Google Slides sometimes stores line breaks in speaker notes using characters
 * that its own Presenter View cannot render correctly. These appear as U+FFFD
 * (replacement character) instead of line breaks. This script reads every
 * slide's speaker notes, replaces those broken characters with real newlines,
 * and writes the cleaned text back.
 *
 * HOW TO USE:
 * 1. Open the Google Slides presentation you want to fix.
 * 2. Go to Extensions > Apps Script.
 * 3. Delete any existing code in the editor and paste this entire file.
 * 4. Click the Save icon (or Ctrl+S / Cmd+S).
 * 5. Select "cleanSpeakerNotes" from the function dropdown at the top.
 * 6. Click Run. The first time, Google will ask you to authorize the script.
 * 7. After it finishes, check View > Logs (or Execution log) for a summary
 *    of which slides were fixed.
 *
 * SAFE TO RE-RUN:
 * The script only modifies slides that contain problematic characters.
 * Running it on an already-clean presentation does nothing.
 */

function cleanSpeakerNotes() {
  var presentation = SlidesApp.getActivePresentation();
  var slides = presentation.getSlides();
  var fixed = 0;
  var scanned = 0;

  slides.forEach(function (slide, i) {
    scanned++;
    var notesShape = slide.getNotesPage().getSpeakerNotesShape();
    var textRange = notesShape.getText();
    var raw = textRange.asString();

    // Check for problematic characters:
    //   U+FFFD  (replacement character - most common)
    //   U+FFFC  (object replacement character)
    //   U+0000  (null)
    if (/[\uFFFD\uFFFC\u0000]/.test(raw)) {
      var cleaned = raw
        .replace(/\uFFFD+/g, '\n')
        .replace(/\uFFFC/g, '\n')
        .replace(/\u0000/g, '');

      // Remove trailing newline that Google Slides always appends
      if (cleaned.endsWith('\n') && !raw.endsWith('\n')) {
        cleaned = cleaned.slice(0, -1);
      }

      textRange.setText(cleaned);
      fixed++;
      Logger.log('Fixed slide ' + (i + 1) + ' (' + raw.length + ' chars -> ' + cleaned.length + ' chars)');
    }
  });

  var summary = 'Done. Scanned ' + scanned + ' slides, fixed ' + fixed + '.';
  Logger.log(summary);

  // Show a toast so you don't have to open the log
  if (typeof SpreadsheetApp === 'undefined') {
    // In Slides there's no toast, use a simple UI alert instead
    try {
      SlidesApp.getUi().alert(summary);
    } catch (e) {
      // If UI is not available (e.g. running as a trigger), just log
      Logger.log('(Could not show UI alert: ' + e.message + ')');
    }
  }
}
