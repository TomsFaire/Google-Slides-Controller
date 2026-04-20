// ─── Web Remote refinements — artboards ───
// All artboards render the Remote tab at phone width (390x844)

const wrStyles = {
  font: "'Inter', -apple-system, BlinkMacSystemFont, sans-serif",
  serif: "'Lora', Georgia, serif",
  text: '#333',
  sub: '#757575',
  border: '#dfe0e1',
  surface: '#ffffff',
  warm: '#fbf8f6',
  page: '#f5f5f5',
};

// Tiny icon set used across artboards
const wrIcon = {
  chevL: <path d="M15 6 9 12l6 6" stroke="currentColor" strokeWidth="2" fill="none" strokeLinecap="round" strokeLinejoin="round"/>,
  chevR: <path d="m9 6 6 6-6 6" stroke="currentColor" strokeWidth="2" fill="none" strokeLinecap="round" strokeLinejoin="round"/>,
  notes: <><path d="M5 4h10l4 4v12H5z" stroke="currentColor" fill="none" strokeWidth="1.5" strokeLinejoin="round"/><path d="M15 4v4h4M8 13h8M8 16h6" stroke="currentColor" fill="none" strokeWidth="1.5" strokeLinecap="round"/></>,
  previews: <><rect x="3" y="4" width="8" height="8" rx="1" stroke="currentColor" fill="none" strokeWidth="1.5"/><rect x="13" y="4" width="8" height="8" rx="1" stroke="currentColor" fill="none" strokeWidth="1.5"/><rect x="3" y="14" width="18" height="6" rx="1" stroke="currentColor" fill="none" strokeWidth="1.5"/></>,
  settings: <><circle cx="12" cy="12" r="3" stroke="currentColor" fill="none" strokeWidth="1.5"/><path d="M12 2v3M12 19v3M4.22 4.22l2.12 2.12M17.66 17.66l2.12 2.12M2 12h3M19 12h3M4.22 19.78l2.12-2.12M17.66 6.34l2.12-2.12" stroke="currentColor" fill="none" strokeWidth="1.5" strokeLinecap="round"/></>,
  expand: <path d="M5 8V5h3M19 8V5h-3M5 16v3h3M19 16v3h-3" stroke="currentColor" fill="none" strokeWidth="1.5" strokeLinecap="round"/>,
  more: <><circle cx="12" cy="6" r="1.2" fill="currentColor"/><circle cx="12" cy="12" r="1.2" fill="currentColor"/><circle cx="12" cy="18" r="1.2" fill="currentColor"/></>,
};

// Fake slide thumbnail – used in all variants
function SlideThumb({ w = 160, h, title, body, accent = '#5a7b9a', label }) {
  const height = h ?? Math.round(w * 9 / 16);
  return (
    <div style={{
      width: w, height, background: '#e6e9ec', borderRadius: 2,
      position: 'relative', overflow: 'hidden', border: '1px solid ' + wrStyles.border,
    }}>
      <div style={{ position: 'absolute', inset: 0, background: 'linear-gradient(180deg,#fff,#eef1f4)' }}/>
      <div style={{ position: 'absolute', top: '18%', left: '10%', width: '45%', height: 4, background: accent, borderRadius: 2 }}/>
      <div style={{ position: 'absolute', top: '30%', left: '10%', right: '12%', height: 10, background: '#333', borderRadius: 2, opacity: 0.85 }}/>
      <div style={{ position: 'absolute', top: '46%', left: '10%', width: '60%', height: 5, background: '#aaa', borderRadius: 2 }}/>
      <div style={{ position: 'absolute', top: '54%', left: '10%', width: '70%', height: 5, background: '#bbb', borderRadius: 2 }}/>
      <div style={{ position: 'absolute', top: '62%', left: '10%', width: '40%', height: 5, background: '#c9c9c9', borderRadius: 2 }}/>
      {label && (
        <div style={{
          position: 'absolute', bottom: 6, right: 6, fontFamily: 'ui-monospace, Menlo, monospace',
          fontSize: 9, color: '#757575', letterSpacing: 0.5,
        }}>{label}</div>
      )}
    </div>
  );
}

function StatusDotWR({ tone = 'idle', size = 6 }) {
  const c = { ok: '#49694c', warn: '#907c3a', bad: '#921100', idle: '#b5a998' };
  return <span style={{ width: size, height: size, borderRadius: '50%', background: c[tone], display: 'inline-block' }}/>;
}

// ════════════════════════════════════════════════════════════
// V0 — current "Light (minimalist)" theme reproduced for reference
// ════════════════════════════════════════════════════════════
function V0Current() {
  return (
    <div style={{ width: 390, height: 844, background: '#f5f5f5', fontFamily: wrStyles.font, padding: '16px 12px', overflow: 'hidden' }}>
      <div style={{ background: '#fff', border: '1px solid #e0e0e0', borderRadius: 0, boxShadow: '0 1px 3px rgba(0,0,0,0.08)', padding: 20 }}>
        <h1 style={{ fontSize: 22, fontWeight: 700, color: '#212121', marginBottom: 16, display: 'flex', alignItems: 'center', gap: 8 }}>
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="2" y="4" width="20" height="12" rx="2"/><line x1="6" y1="20" x2="18" y2="20"/><line x1="8" y1="16" x2="8" y2="20"/><line x1="16" y1="16" x2="16" y2="20"/></svg>
          Stage Left Mac
        </h1>
        <div style={{ display: 'flex', gap: 6, marginBottom: 16 }}>
          {['Remote', 'Controls', 'Settings'].map((t, i) => (
            <button key={t} style={{
              padding: '8px 16px', fontSize: 13, fontWeight: 600,
              background: i === 0 ? '#4285f4' : '#f0f0f0',
              color: i === 0 ? '#fff' : '#555',
              border: 'none', borderRadius: 0, cursor: 'pointer',
            }}>{t}</button>
          ))}
        </div>
        <h2 style={{ fontSize: 18, fontWeight: 700, color: '#212121', marginBottom: 12 }}>Remote Control</h2>
        <div style={{ display: 'flex', gap: 6, marginBottom: 14 }}>
          <button style={{ flex: 1, padding: '8px', fontSize: 12, background: '#f8f9fa', border: '2px solid #e0e0e0', borderRadius: 0 }}>◐ Notes</button>
          <button style={{ flex: 1, padding: '8px', fontSize: 12, background: '#f8f9fa', border: '2px solid #e0e0e0', borderRadius: 0 }}>▦ Previews</button>
        </div>
        {/* Stagetimer card – gradient, not themed */}
        <div style={{
          background: 'linear-gradient(135deg, #4caf50 0%, #388e3c 100%)', borderRadius: 0,
          padding: 16, color: '#fff', textAlign: 'center', marginBottom: 14,
          boxShadow: '0 4px 12px rgba(0,0,0,0.15)',
        }}>
          <div style={{ fontSize: 14, fontWeight: 600, opacity: 0.95 }}>Stage Timer</div>
          <div style={{ fontSize: 38, fontWeight: 700, letterSpacing: 2, fontFamily: 'ui-monospace, Menlo, monospace' }}>12:34</div>
          <div style={{ fontSize: 13, opacity: 0.9 }}>Running</div>
        </div>
        {/* Controls */}
        <div style={{ display: 'flex', gap: 8, marginBottom: 14 }}>
          <button style={{ flex: 1, padding: '18px 10px', background: '#f0f0f0', border: '2px solid #e0e0e0', borderRadius: 0, fontSize: 14, fontWeight: 600, color: '#333' }}>◀ Previous Slide</button>
          <button style={{ flex: 1, padding: '18px 10px', background: '#4285f4', color: '#fff', border: 'none', borderRadius: 0, fontSize: 14, fontWeight: 600 }}>Next Slide ▶</button>
        </div>
        {/* Previews */}
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8, marginBottom: 14 }}>
          <div style={{ background: '#f8f9fa', border: '1px solid #e0e0e0', padding: 8 }}>
            <div style={{ fontSize: 11, color: '#666', marginBottom: 4, fontWeight: 600 }}>Current Slide</div>
            <SlideThumb w={154}/>
          </div>
          <div style={{ background: '#f8f9fa', border: '1px solid #e0e0e0', padding: 8 }}>
            <div style={{ fontSize: 11, color: '#666', marginBottom: 4, fontWeight: 600 }}>Next Slide</div>
            <SlideThumb w={154} accent="#a35a7b"/>
          </div>
        </div>
        {/* Notes */}
        <div style={{ background: '#f8f9fa', border: '1px solid #e0e0e0', padding: 12, fontSize: 13, lineHeight: 1.5, color: '#333' }}>
          <div style={{ display: 'flex', gap: 4, marginBottom: 8 }}>
            <button style={{ fontSize: 11, padding: '4px 8px', background: '#fff', border: '2px solid #e0e0e0', fontWeight: 600 }}>↑ Scroll Up</button>
            <button style={{ fontSize: 11, padding: '4px 8px', background: '#fff', border: '2px solid #e0e0e0', fontWeight: 600 }}>– Zoom Out</button>
            <button style={{ fontSize: 11, padding: '4px 8px', background: '#fff', border: '2px solid #e0e0e0', fontWeight: 600 }}>+ Zoom In</button>
          </div>
          Welcome everyone — thanks for joining us today. Three things we want to cover: quarterly results, the new venue rollout…
        </div>
      </div>
    </div>
  );
}

// ════════════════════════════════════════════════════════════
// V1 — Editorial minimal (Faire-literal)
// Direct application of the desktop's Faire language. Serif title,
// monochrome, no tabs (bottom bar), previews and notes in one rail.
// ════════════════════════════════════════════════════════════
function V1Editorial() {
  return (
    <div style={{ width: 390, height: 844, background: wrStyles.page, fontFamily: wrStyles.font, display: 'flex', flexDirection: 'column' }}>
      {/* Header */}
      <header style={{ padding: '18px 20px 12px', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
        <div>
          <div style={{ fontSize: 10.5, fontWeight: 500, color: wrStyles.sub, letterSpacing: 0.8, textTransform: 'uppercase' }}>Live · slide 3 of 24</div>
          <div style={{ fontFamily: wrStyles.serif, fontSize: 22, color: wrStyles.text, marginTop: 2 }}>Stage Left Mac</div>
        </div>
        <button style={{ width: 32, height: 32, background: 'transparent', border: 'none', color: wrStyles.sub, cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
          <svg width="18" height="18" viewBox="0 0 24 24">{wrIcon.more}</svg>
        </button>
      </header>

      {/* Stagetimer — warm card, no gradient */}
      <div style={{ margin: '4px 20px 16px', background: wrStyles.warm, border: '1px solid ' + wrStyles.border, borderRadius: 4, padding: '14px 16px', display: 'flex', alignItems: 'center', gap: 12 }}>
        <StatusDotWR tone="ok" size={7}/>
        <div style={{ fontSize: 11, color: wrStyles.sub, letterSpacing: 0.5, textTransform: 'uppercase', flex: 1 }}>Stage timer</div>
        <div style={{ fontFamily: 'ui-monospace, Menlo, monospace', fontSize: 22, color: wrStyles.text, letterSpacing: 1, fontWeight: 500 }}>12:34</div>
      </div>

      {/* Slide rail: current + next side-by-side */}
      <div style={{ padding: '0 20px', display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10, marginBottom: 16 }}>
        <div>
          <div style={{ fontSize: 10.5, color: wrStyles.sub, letterSpacing: 0.6, textTransform: 'uppercase', marginBottom: 6 }}>Now · 3/24</div>
          <SlideThumb w={160}/>
        </div>
        <div>
          <div style={{ fontSize: 10.5, color: wrStyles.sub, letterSpacing: 0.6, textTransform: 'uppercase', marginBottom: 6 }}>Next · 4/24</div>
          <SlideThumb w={160} accent="#a35a7b"/>
        </div>
      </div>

      {/* Notes — editorial serif card */}
      <div style={{ flex: 1, margin: '0 20px', background: wrStyles.surface, border: '1px solid ' + wrStyles.border, borderRadius: 4, padding: '16px 18px', display: 'flex', flexDirection: 'column', minHeight: 0 }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 10 }}>
          <div style={{ fontSize: 10.5, color: wrStyles.sub, letterSpacing: 0.6, textTransform: 'uppercase' }}>Speaker notes</div>
          <div style={{ display: 'flex', gap: 4 }}>
            <button style={{ width: 26, height: 26, border: '1px solid ' + wrStyles.border, background: '#fff', borderRadius: 4, fontSize: 13, cursor: 'pointer', color: wrStyles.sub }}>−</button>
            <button style={{ width: 26, height: 26, border: '1px solid ' + wrStyles.border, background: '#fff', borderRadius: 4, fontSize: 13, cursor: 'pointer', color: wrStyles.sub }}>+</button>
          </div>
        </div>
        <div style={{ fontFamily: wrStyles.serif, fontSize: 16, lineHeight: '26px', color: wrStyles.text, overflow: 'hidden', flex: 1 }}>
          Welcome everyone — thanks for joining us today.<br/><br/>
          Three things to cover:<br/>
          • Quarterly results<br/>
          • New venue rollout<br/>
          • How you can get involved
        </div>
      </div>

      {/* Bottom bar: tabs + primary controls in one */}
      <div style={{ marginTop: 14, background: wrStyles.surface, borderTop: '1px solid ' + wrStyles.border, padding: '12px 16px 20px' }}>
        {/* Prev/Next */}
        <div style={{ display: 'flex', gap: 10, marginBottom: 10 }}>
          <button style={{
            flex: 1, height: 56, background: wrStyles.surface, border: '1px solid ' + wrStyles.border,
            borderRadius: 4, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 8,
            fontSize: 14, fontWeight: 500, color: wrStyles.text, cursor: 'pointer',
          }}>
            <svg width="16" height="16" viewBox="0 0 24 24">{wrIcon.chevL}</svg>
            Previous
          </button>
          <button style={{
            flex: 1.5, height: 56, background: '#333', border: '1px solid #333',
            borderRadius: 4, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 8,
            fontSize: 14, fontWeight: 500, color: '#fff', cursor: 'pointer',
          }}>
            Next slide
            <svg width="16" height="16" viewBox="0 0 24 24">{wrIcon.chevR}</svg>
          </button>
        </div>
        {/* Tiny tab strip */}
        <div style={{ display: 'flex', gap: 24, justifyContent: 'center' }}>
          {[
            { t: 'Remote', active: true },
            { t: 'Controls', active: false },
            { t: 'Settings', active: false },
          ].map(t => (
            <button key={t.t} style={{
              background: 'transparent', border: 'none', cursor: 'pointer',
              fontSize: 11.5, color: t.active ? wrStyles.text : wrStyles.sub,
              fontWeight: t.active ? 500 : 400, padding: '4px 0',
              borderBottom: t.active ? '1px solid ' + wrStyles.text : '1px solid transparent',
              letterSpacing: 0.3,
            }}>{t.t}</button>
          ))}
        </div>
      </div>
    </div>
  );
}

// ════════════════════════════════════════════════════════════
// V2 — Stage-ready (operator-first)
// Bigger hit targets, monochrome info-dense header, notes dominant.
// Tuned for live operators in low light: high contrast, giant tap zones.
// ════════════════════════════════════════════════════════════
function V2Stage() {
  return (
    <div style={{ width: 390, height: 844, background: '#fafaf8', fontFamily: wrStyles.font, display: 'flex', flexDirection: 'column' }}>
      {/* Dense status strip */}
      <div style={{ padding: '14px 18px 10px', display: 'flex', alignItems: 'center', gap: 10, fontSize: 11.5, color: wrStyles.sub, letterSpacing: 0.3 }}>
        <StatusDotWR tone="ok"/>
        <span style={{ color: wrStyles.text, fontWeight: 500 }}>Stage Left Mac</span>
        <span>·</span>
        <span>3 / 24</span>
        <span style={{ flex: 1 }}/>
        <span style={{ fontFamily: 'ui-monospace, Menlo, monospace', color: wrStyles.text, fontWeight: 500 }}>12:34</span>
        <StatusDotWR tone="ok"/>
      </div>

      {/* Current slide big */}
      <div style={{ padding: '0 18px 10px' }}>
        <SlideThumb w={354} h={199} label="3 / 24"/>
      </div>

      {/* Next slide mini row */}
      <div style={{ padding: '0 18px 14px', display: 'flex', alignItems: 'center', gap: 10 }}>
        <div style={{ fontSize: 10.5, color: wrStyles.sub, letterSpacing: 0.7, textTransform: 'uppercase', minWidth: 44 }}>Next</div>
        <SlideThumb w={80} h={45} accent="#a35a7b"/>
        <div style={{ flex: 1, minWidth: 0 }}>
          <div style={{ fontSize: 13, color: wrStyles.text, fontWeight: 500, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>The new venue rollout</div>
          <div style={{ fontSize: 11.5, color: wrStyles.sub, marginTop: 1 }}>Slide 4 of 24</div>
        </div>
      </div>

      {/* Notes — dominant area, plain-text mono-hint */}
      <div style={{ flex: 1, margin: '0 18px', minHeight: 0, display: 'flex', flexDirection: 'column', background: '#fff', border: '1px solid ' + wrStyles.border, borderRadius: 4 }}>
        <div style={{ padding: '10px 14px', borderBottom: '1px solid ' + wrStyles.border, display: 'flex', alignItems: 'center', gap: 6, fontSize: 10.5, color: wrStyles.sub, letterSpacing: 0.7, textTransform: 'uppercase' }}>
          <span style={{ flex: 1 }}>Speaker notes</span>
          <button style={{ width: 26, height: 26, border: 'none', background: 'transparent', color: wrStyles.sub, cursor: 'pointer', borderRadius: 4, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
            <svg width="14" height="14" viewBox="0 0 24 24">{wrIcon.expand}</svg>
          </button>
        </div>
        <div style={{ padding: '14px 16px', fontFamily: wrStyles.serif, fontSize: 17, lineHeight: '28px', color: wrStyles.text, overflow: 'hidden', flex: 1 }}>
          Welcome everyone — thanks for joining us today.<br/><br/>
          Three things to cover: quarterly results, the new venue rollout, and how you can get involved.
        </div>
      </div>

      {/* Giant tap controls — full width, equal weight, 80pt */}
      <div style={{ padding: '14px 18px 14px', display: 'flex', gap: 10 }}>
        <button style={{
          flex: 1, height: 80, background: '#fff', border: '1px solid ' + wrStyles.border,
          borderRadius: 4, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', gap: 2,
          color: wrStyles.text, cursor: 'pointer',
        }}>
          <svg width="22" height="22" viewBox="0 0 24 24">{wrIcon.chevL}</svg>
          <span style={{ fontSize: 11.5, color: wrStyles.sub, letterSpacing: 0.5, textTransform: 'uppercase' }}>Previous</span>
        </button>
        <button style={{
          flex: 1, height: 80, background: '#333', border: '1px solid #333',
          borderRadius: 4, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', gap: 2,
          color: '#fff', cursor: 'pointer',
        }}>
          <svg width="22" height="22" viewBox="0 0 24 24">{wrIcon.chevR}</svg>
          <span style={{ fontSize: 11.5, color: 'rgba(255,255,255,0.7)', letterSpacing: 0.5, textTransform: 'uppercase' }}>Next slide</span>
        </button>
      </div>

      {/* Footer tab bar — iOS-style, minimal */}
      <nav style={{ borderTop: '1px solid ' + wrStyles.border, background: '#fff', display: 'flex', padding: '8px 0 16px' }}>
        {[
          { t: 'Remote', i: 'previews', active: true },
          { t: 'Controls', i: 'settings', active: false },
          { t: 'Settings', i: 'settings', active: false },
        ].map((t, i) => (
          <button key={i} style={{
            flex: 1, background: 'transparent', border: 'none', cursor: 'pointer',
            display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 3, padding: '6px 0',
            color: t.active ? wrStyles.text : wrStyles.sub,
          }}>
            <svg width="18" height="18" viewBox="0 0 24 24">{wrIcon[t.i]}</svg>
            <span style={{ fontSize: 10, fontWeight: t.active ? 500 : 400, letterSpacing: 0.2 }}>{t.t}</span>
          </button>
        ))}
      </nav>
    </div>
  );
}

// ════════════════════════════════════════════════════════════
// V3 — Focus mode (notes-first landscape of reading)
// A cinema-like card: giant serif notes with a slim control shelf.
// Best for speakers themselves — reading at distance.
// ════════════════════════════════════════════════════════════
function V3Focus() {
  return (
    <div style={{ width: 390, height: 844, background: '#1d1d1b', fontFamily: wrStyles.font, color: '#e8e4dd', display: 'flex', flexDirection: 'column' }}>
      {/* Minimal bar */}
      <div style={{ padding: '14px 18px 10px', display: 'flex', alignItems: 'center', gap: 10, fontSize: 11.5, color: 'rgba(232,228,221,0.55)', letterSpacing: 0.4 }}>
        <StatusDotWR tone="ok"/>
        <span>Stage Left</span>
        <span style={{ flex: 1 }}/>
        <span style={{ fontFamily: 'ui-monospace, Menlo, monospace', color: '#e8e4dd', fontWeight: 500 }}>12:34</span>
        <span>· 3/24</span>
      </div>

      {/* Giant serif notes */}
      <div style={{ flex: 1, padding: '28px 24px 18px', overflow: 'hidden', display: 'flex', flexDirection: 'column' }}>
        <div style={{ fontSize: 10.5, color: 'rgba(232,228,221,0.4)', letterSpacing: 0.8, textTransform: 'uppercase', marginBottom: 12 }}>
          Slide 3 · Opening
        </div>
        <div style={{
          fontFamily: wrStyles.serif, fontSize: 22, lineHeight: '36px',
          color: '#f4f0e8', flex: 1, overflow: 'hidden',
        }}>
          Welcome everyone — thanks for joining us today.<br/><br/>
          Three things to cover:<br/>
          • Quarterly results<br/>
          • New venue rollout<br/>
          • How you can get involved<br/><br/>
          <span style={{ color: 'rgba(232,228,221,0.5)' }}>Let's dive in.</span>
        </div>
      </div>

      {/* Next slide peek */}
      <div style={{ padding: '0 24px 14px', display: 'flex', alignItems: 'center', gap: 12 }}>
        <div style={{ width: 72, height: 40, background: '#2e2e2b', borderRadius: 2, border: '1px solid rgba(255,255,255,0.1)', overflow: 'hidden' }}>
          <div style={{ width: '100%', height: '100%', background: 'linear-gradient(135deg,#3a3a37,#2a2a28)', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 10, color: 'rgba(232,228,221,0.4)' }}>4</div>
        </div>
        <div style={{ flex: 1, minWidth: 0 }}>
          <div style={{ fontSize: 10, color: 'rgba(232,228,221,0.4)', letterSpacing: 0.7, textTransform: 'uppercase' }}>Up next</div>
          <div style={{ fontSize: 13, color: '#e8e4dd', marginTop: 2, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>The new venue rollout</div>
        </div>
      </div>

      {/* Control shelf */}
      <div style={{ padding: '12px 18px 22px', borderTop: '1px solid rgba(255,255,255,0.08)', display: 'flex', gap: 10, alignItems: 'center' }}>
        <button style={{
          width: 56, height: 56, background: 'rgba(255,255,255,0.06)', border: '1px solid rgba(255,255,255,0.1)',
          borderRadius: 4, color: '#e8e4dd', cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center',
        }}>
          <svg width="20" height="20" viewBox="0 0 24 24">{wrIcon.chevL}</svg>
        </button>
        <button style={{
          flex: 1, height: 56, background: '#e8e4dd', border: 'none',
          borderRadius: 4, color: '#1d1d1b', fontSize: 14, fontWeight: 500, letterSpacing: 0.3,
          cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 8,
        }}>
          Next slide
          <svg width="18" height="18" viewBox="0 0 24 24">{wrIcon.chevR}</svg>
        </button>
        <button style={{
          width: 56, height: 56, background: 'rgba(255,255,255,0.06)', border: '1px solid rgba(255,255,255,0.1)',
          borderRadius: 4, color: '#e8e4dd', cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center',
        }}>
          <svg width="20" height="20" viewBox="0 0 24 24">{wrIcon.more}</svg>
        </button>
      </div>
    </div>
  );
}

Object.assign(window, { V0Current, V1Editorial, V2Stage, V3Focus });
