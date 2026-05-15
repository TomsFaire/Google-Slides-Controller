const { normalizeAdapterId } = require('./perfectcue-adapter-presets');

/**
 * Normalize perfectCuePorts to PortConfig[] regardless of the stored format.
 * Handles:
 *   - Legacy single number `perfectCuePort`
 *   - Array of plain numbers (old multi-port format)
 *   - Array of PortConfig objects (current format)
 * Always returns a non-empty PortConfig[].
 * @param {object} prefs
 * @returns {{ port: number, name: string, enabled: boolean, adapter: 'dsan'|'waveshare' }[]}
 */
function normalizePerfectCuePorts(prefs) {
  const raw = Array.isArray(prefs.perfectCuePorts) ? prefs.perfectCuePorts : [];
  const configs = raw.map(entry => {
    if (typeof entry === 'number') {
      return {
        port: entry,
        name: '',
        enabled: true,
        adapter: normalizeAdapterId(undefined)
      };
    }
    return {
      port: Number(entry.port),
      name: typeof entry.name === 'string' ? entry.name : '',
      enabled: entry.enabled !== false,
      adapter: normalizeAdapterId(entry.adapter)
    };
  }).filter(c => c.port > 0);

  if (configs.length === 0) {
    const legacyPort = prefs.perfectCuePort ? Number(prefs.perfectCuePort) : 0;
    return [{
      port: legacyPort > 0 ? legacyPort : 8899,
      name: '',
      enabled: true,
      adapter: normalizeAdapterId(undefined)
    }];
  }
  return configs;
}

module.exports = { normalizePerfectCuePorts };
