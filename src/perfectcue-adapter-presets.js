/**
 * TCP keep-alive presets per serial-TCP converter family.
 * DSAN (USR-TCP232-style) needs frequent application-level 0xFF traffic.
 * WaveShare firmware typically tolerates longer intervals (field-tunable).
 */

const PRESETS = {
  dsan: { pingIntervalMs: 15_000, idleTimeoutMs: 50_000 },
  waveshare: { pingIntervalMs: 45_000, idleTimeoutMs: 120_000 }
};

/** @typedef {'dsan'|'waveshare'} PerfectCueAdapterId */

/**
 * @param {unknown} value
 * @returns {PerfectCueAdapterId}
 */
function normalizeAdapterId(value) {
  if (value === 'waveshare') return 'waveshare';
  return 'dsan';
}

/**
 * @param {unknown} adapter
 * @returns {{ pingIntervalMs: number, idleTimeoutMs: number }}
 */
function getPerfectCueAdapterPreset(adapter) {
  const id = normalizeAdapterId(adapter);
  return PRESETS[id];
}

module.exports = {
  normalizeAdapterId,
  getPerfectCueAdapterPreset,
  PERFECTCUE_ADAPTER_IDS: /** @type {const} */ (['dsan', 'waveshare'])
};
