/**
 * URL checks for window-open and navigation handlers.
 * Slides popups (speaker notes) open via window.open() to http(s) or about:blank;
 * anything else (file:, javascript:, data:, custom schemes) is refused.
 */
function isSafeWindowOpenUrl(url) {
  if (url === 'about:blank') return true;
  try {
    const { protocol } = new URL(url);
    return protocol === 'https:' || protocol === 'http:';
  } catch (_) {
    return false;
  }
}

/** True for URLs that may be handed to the OS browser. */
function isExternalHttpUrl(url) {
  try {
    const { protocol } = new URL(url);
    return protocol === 'https:' || protocol === 'http:';
  } catch (_) {
    return false;
  }
}

module.exports = { isSafeWindowOpenUrl, isExternalHttpUrl };
