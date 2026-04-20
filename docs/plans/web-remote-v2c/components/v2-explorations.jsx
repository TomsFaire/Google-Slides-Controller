// ─── V2 Stage-ready — deep explorations ───
// Focus: Stagetimer states (idle/green/yellow/red/overtime + messages),
// speaker notes zoom+scroll controls, and how they coexist without clutter.

const v2s = {
  font: "'Inter', -apple-system, BlinkMacSystemFont, sans-serif",
  serif: "'Lora', Georgia, serif",
  mono: "ui-monospace, 'SF Mono', Menlo, monospace",
  text: '#333',
  sub: '#757575',
  muted: '#b5a998',
  border: '#dfe0e1',
  surface: '#ffffff',
  warm: '#fbf8f6',
  page: '#fafaf8',
  // Stagetimer tone palette — calm, editorial, not neon
  timerTones: {
    idle:     { bg: '#fbf8f6', border: '#dfe0e1', text: '#757575', clock: '#333', label: 'Idle' },
    running:  { bg: '#eef2ed', border: '#c8d4c8', text: '#49694c', clock: '#2d4a30', label: 'Running' },
    warning:  { bg: '#f6efdb', border: '#d1b985', text: '#907c3a', clock: '#5c4e1e', label: 'Warning' },
    critical: { bg: '#f5dcd6', border: '#d9a79a', text: '#921100', clock: '#6e1100', label: 'Critical' },
    overtime: { bg: '#3a1510', border: '#6e1100', text: '#ffd3c9', clock: '#ffffff', label: 'Overtime' },
  },
};

const v2Icon = {
  chevL: <path d="M15 6 9 12l6 6" stroke="currentColor" strokeWidth="2" fill="none" strokeLinecap="round" strokeLinejoin="round"/>,
  chevR: <path d="m9 6 6 6-6 6" stroke="currentColor" strokeWidth="2" fill="none" strokeLinecap="round" strokeLinejoin="round"/>,
  expand: <path d="M5 8V5h3M19 8V5h-3M5 16v3h3M19 16v3h-3" stroke="currentColor" fill="none" strokeWidth="1.5" strokeLinecap="round"/>,
  arrowUp: <path d="M12 19V5M5 12l7-7 7 7" stroke="currentColor" fill="none" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>,
  arrowDn: <path d="M12 5v14M5 12l7 7 7-7" stroke="currentColor" fill="none" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>,
  plus: <path d="M12 5v14M5 12h14" stroke="currentColor" fill="none" strokeWidth="1.6" strokeLinecap="round"/>,
  minus: <path d="M5 12h14" stroke="currentColor" fill="none" strokeWidth="1.6" strokeLinecap="round"/>,
  home: <path d="M3 11 12 3l9 8M5 9.5V20h14V9.5" stroke="currentColor" fill="none" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>,
  control: <><rect x="3" y="4" width="18" height="4" rx="1" stroke="currentColor" fill="none" strokeWidth="1.5"/><rect x="3" y="10" width="18" height="4" rx="1" stroke="currentColor" fill="none" strokeWidth="1.5"/><rect x="3" y="16" width="18" height="4" rx="1" stroke="currentColor" fill="none" strokeWidth="1.5"/></>,
  settings: <><circle cx="12" cy="12" r="3" stroke="currentColor" fill="none" strokeWidth="1.5"/><path d="M12 2v3M12 19v3M4.22 4.22l2.12 2.12M17.66 17.66l2.12 2.12M2 12h3M19 12h3M4.22 19.78l2.12-2.12M17.66 6.34l2.12-2.12" stroke="currentColor" fill="none" strokeWidth="1.5" strokeLinecap="round"/></>,
};

function V2SlideThumb({ w = 160, h, title, accent = '#5a7b9a', label, placeholder }) {
  const height = h ?? Math.round(w * 9 / 16);
  if (placeholder) {
    return (
      <div style={{
        width: w, height, background: v2s.warm, borderRadius: 2,
        border: '1px dashed ' + v2s.border, display: 'flex',
        alignItems: 'center', justifyContent: 'center',
        color: v2s.muted, fontSize: 10, letterSpacing: 0.5, textTransform: 'uppercase',
      }}>{placeholder}</div>
    );
  }
  return (
    <div style={{
      width: w, height, background: '#e6e9ec', borderRadius: 2,
      position: 'relative', overflow: 'hidden', border: '1px solid ' + v2s.border,
    }}>
      <div style={{ position: 'absolute', inset: 0, background: 'linear-gradient(180deg,#fff,#eef1f4)' }}/>
      <div style={{ position: 'absolute', top: '18%', left: '10%', width: '45%', height: 4, background: accent, borderRadius: 2 }}/>
      <div style={{ position: 'absolute', top: '30%', left: '10%', right: '12%', height: 10, background: '#333', borderRadius: 2, opacity: 0.85 }}/>
      <div style={{ position: 'absolute', top: '46%', left: '10%', width: '60%', height: 5, background: '#aaa', borderRadius: 2 }}/>
      <div style={{ position: 'absolute', top: '54%', left: '10%', width: '70%', height: 5, background: '#bbb', borderRadius: 2 }}/>
      <div style={{ position: 'absolute', top: '62%', left: '10%', width: '40%', height: 5, background: '#c9c9c9', borderRadius: 2 }}/>
      {label && (
        <div style={{
          position: 'absolute', bottom: 6, right: 6, fontFamily: v2s.mono,
          fontSize: 9, color: v2s.sub, letterSpacing: 0.5,
        }}>{label}</div>
      )}
    </div>
  );
}

function V2Dot({ tone = 'idle', size = 6 }) {
  const c = { ok: '#49694c', warn: '#907c3a', bad: '#921100', idle: '#b5a998' };
  return <span style={{ width: size, height: size, borderRadius: '50%', background: c[tone], display: 'inline-block' }}/>;
}

// ════════════════════════════════════════════════════════════
// Stagetimer card — single component, state-driven
// ════════════════════════════════════════════════════════════
function StagetimerCard({ state = 'running', time = '12:34', name = 'Opening remarks', message }) {
  const t = v2s.timerTones[state];
  const isOvertime = state === 'overtime';
  const timeDisplay = isOvertime ? '-' + time : time;
  return (
    <div style={{
      background: t.bg, border: '1px solid ' + t.border,
      borderRadius: 4, padding: '12px 14px', position: 'relative', overflow: 'hidden',
    }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
        <div style={{ width: 6, height: 6, borderRadius: '50%', background: t.text, flexShrink: 0 }}/>
        <div style={{
          fontSize: 10.5, color: t.text, letterSpacing: 0.7, textTransform: 'uppercase', fontWeight: 500,
          flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap',
        }}>
          {name} · {t.label}
        </div>
        <div style={{
          fontFamily: v2s.mono, fontSize: 28, color: t.clock,
          fontWeight: 500, letterSpacing: 1, lineHeight: 1, fontVariantNumeric: 'tabular-nums',
        }}>{timeDisplay}</div>
      </div>
      {message && (
        <div style={{
          marginTop: 10, paddingTop: 10, borderTop: '1px solid ' + t.border,
          fontSize: 12.5, color: t.text, lineHeight: '17px',
        }}>
          {message}
        </div>
      )}
    </div>
  );
}

// ════════════════════════════════════════════════════════════
// V2-A: Notes-dominant with inline zoom + scroll rail
// Zoom and scroll sit on the LEFT edge of the notes card — always visible,
// out of the reading flow, fat hit targets.
// ════════════════════════════════════════════════════════════
function V2A_NotesRail({ timerState = 'running' }) {
  return (
    <div style={{ width: 390, height: 844, background: v2s.page, fontFamily: v2s.font, display: 'flex', flexDirection: 'column' }}>
      <div style={{ padding: '14px 18px 10px', display: 'flex', alignItems: 'center', gap: 10, fontSize: 11.5, color: v2s.sub, letterSpacing: 0.3 }}>
        <V2Dot tone="ok"/>
        <span style={{ color: v2s.text, fontWeight: 500 }}>Stage Left Mac</span>
        <span style={{ flex: 1 }}/>
        <span>3 / 24</span>
      </div>

      {/* Stagetimer */}
      <div style={{ padding: '0 18px 12px' }}>
        <StagetimerCard state={timerState}/>
      </div>

      {/* Current slide */}
      <div style={{ padding: '0 18px 10px' }}>
        <V2SlideThumb w={354} h={199} label="3 / 24"/>
      </div>

      {/* Next row */}
      <div style={{ padding: '0 18px 14px', display: 'flex', alignItems: 'center', gap: 10 }}>
        <div style={{ fontSize: 10.5, color: v2s.sub, letterSpacing: 0.7, textTransform: 'uppercase', minWidth: 44 }}>Next</div>
        <V2SlideThumb w={72} h={40} accent="#a35a7b"/>
        <div style={{ flex: 1, minWidth: 0 }}>
          <div style={{ fontSize: 13, color: v2s.text, fontWeight: 500, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>The new venue rollout</div>
          <div style={{ fontSize: 11.5, color: v2s.sub, marginTop: 1 }}>Slide 4 of 24</div>
        </div>
      </div>

      {/* Notes card with LEFT-edge control rail */}
      <div style={{ flex: 1, margin: '0 18px', minHeight: 0, background: v2s.surface, border: '1px solid ' + v2s.border, borderRadius: 4, display: 'flex', overflow: 'hidden' }}>
        {/* Rail */}
        <div style={{
          width: 44, borderRight: '1px solid ' + v2s.border, background: v2s.warm,
          display: 'flex', flexDirection: 'column', alignItems: 'center', padding: '8px 0',
        }}>
          <button title="Scroll notes up (on presenter screen)" style={{ width: 36, height: 44, border: 'none', background: 'transparent', color: v2s.text, cursor: 'pointer', borderRadius: 2, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
            <svg width="16" height="16" viewBox="0 0 24 24">{v2Icon.arrowUp}</svg>
          </button>
          <div style={{ width: 24, height: 1, background: v2s.border, margin: '4px 0' }}/>
          <button title="Zoom in" style={{ width: 36, height: 44, border: 'none', background: 'transparent', color: v2s.text, cursor: 'pointer', borderRadius: 2, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
            <svg width="14" height="14" viewBox="0 0 24 24">{v2Icon.plus}</svg>
          </button>
          <div style={{ fontFamily: v2s.mono, fontSize: 10, color: v2s.sub }}>18</div>
          <button title="Zoom out" style={{ width: 36, height: 44, border: 'none', background: 'transparent', color: v2s.text, cursor: 'pointer', borderRadius: 2, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
            <svg width="14" height="14" viewBox="0 0 24 24">{v2Icon.minus}</svg>
          </button>
          <div style={{ flex: 1 }}/>
          <button title="Scroll notes down (on presenter screen)" style={{ width: 36, height: 44, border: 'none', background: 'transparent', color: v2s.text, cursor: 'pointer', borderRadius: 2, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
            <svg width="16" height="16" viewBox="0 0 24 24">{v2Icon.arrowDn}</svg>
          </button>
        </div>
        {/* Notes content */}
        <div style={{ flex: 1, padding: '12px 16px', overflow: 'hidden', display: 'flex', flexDirection: 'column' }}>
          <div style={{ fontSize: 10.5, color: v2s.sub, letterSpacing: 0.7, textTransform: 'uppercase', marginBottom: 8, display: 'flex', alignItems: 'center', gap: 6 }}>
            <span style={{ flex: 1 }}>Speaker notes</span>
            <span style={{ fontFamily: v2s.mono, fontSize: 10, color: v2s.muted }}>slide 3</span>
          </div>
          <div style={{ fontFamily: v2s.serif, fontSize: 18, lineHeight: '28px', color: v2s.text, overflow: 'hidden', flex: 1 }}>
            Welcome everyone — thanks for joining us today.<br/><br/>
            Three things to cover:<br/>
            • Quarterly results<br/>
            • New venue rollout<br/>
            • How you can get involved
          </div>
        </div>
      </div>

      {/* Controls */}
      <div style={{ padding: '14px 18px 14px', display: 'flex', gap: 10 }}>
        <button style={{
          flex: 1, height: 72, background: v2s.surface, border: '1px solid ' + v2s.border,
          borderRadius: 4, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', gap: 2,
          color: v2s.text, cursor: 'pointer',
        }}>
          <svg width="20" height="20" viewBox="0 0 24 24">{v2Icon.chevL}</svg>
          <span style={{ fontSize: 11, color: v2s.sub, letterSpacing: 0.5, textTransform: 'uppercase' }}>Previous</span>
        </button>
        <button style={{
          flex: 1, height: 72, background: '#333', border: '1px solid #333',
          borderRadius: 4, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', gap: 2,
          color: '#fff', cursor: 'pointer',
        }}>
          <svg width="20" height="20" viewBox="0 0 24 24">{v2Icon.chevR}</svg>
          <span style={{ fontSize: 11, color: 'rgba(255,255,255,0.7)', letterSpacing: 0.5, textTransform: 'uppercase' }}>Next slide</span>
        </button>
      </div>

      <nav style={{ borderTop: '1px solid ' + v2s.border, background: v2s.surface, display: 'flex', padding: '6px 0 12px' }}>
        {[
          { t: 'Remote', i: 'home', active: true },
          { t: 'Controls', i: 'control', active: false },
          { t: 'Settings', i: 'settings', active: false },
        ].map((x, i) => (
          <button key={i} style={{
            flex: 1, background: 'transparent', border: 'none', cursor: 'pointer',
            display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 2, padding: '6px 0',
            color: x.active ? v2s.text : v2s.sub,
          }}>
            <svg width="18" height="18" viewBox="0 0 24 24">{v2Icon[x.i]}</svg>
            <span style={{ fontSize: 10, fontWeight: x.active ? 500 : 400, letterSpacing: 0.2 }}>{x.t}</span>
          </button>
        ))}
      </nav>
    </div>
  );
}

// ════════════════════════════════════════════════════════════
// V2-B: Notes-dominant with toolbar on top
// Zoom+scroll in a thin toolbar above the notes body. More discoverable
// than V2-A's side rail, but costs a little vertical real estate.
// ════════════════════════════════════════════════════════════
function V2B_NotesToolbar({ timerState = 'warning' }) {
  return (
    <div style={{ width: 390, height: 844, background: v2s.page, fontFamily: v2s.font, display: 'flex', flexDirection: 'column' }}>
      <div style={{ padding: '14px 18px 10px', display: 'flex', alignItems: 'center', gap: 10, fontSize: 11.5, color: v2s.sub, letterSpacing: 0.3 }}>
        <V2Dot tone="ok"/>
        <span style={{ color: v2s.text, fontWeight: 500 }}>Stage Left Mac</span>
        <span style={{ flex: 1 }}/>
        <span>3 / 24</span>
      </div>

      <div style={{ padding: '0 18px 12px' }}>
        <StagetimerCard state={timerState} message="Wrap it up — Q&A coming"/>
      </div>

      <div style={{ padding: '0 18px 10px' }}>
        <V2SlideThumb w={354} h={199} label="3 / 24"/>
      </div>

      <div style={{ padding: '0 18px 14px', display: 'flex', alignItems: 'center', gap: 10 }}>
        <div style={{ fontSize: 10.5, color: v2s.sub, letterSpacing: 0.7, textTransform: 'uppercase', minWidth: 44 }}>Next</div>
        <V2SlideThumb w={72} h={40} accent="#a35a7b"/>
        <div style={{ flex: 1, minWidth: 0 }}>
          <div style={{ fontSize: 13, color: v2s.text, fontWeight: 500, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>The new venue rollout</div>
          <div style={{ fontSize: 11.5, color: v2s.sub, marginTop: 1 }}>Slide 4 of 24</div>
        </div>
      </div>

      <div style={{ flex: 1, margin: '0 18px', minHeight: 0, background: v2s.surface, border: '1px solid ' + v2s.border, borderRadius: 4, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
        {/* Toolbar */}
        <div style={{
          display: 'flex', alignItems: 'center', padding: '6px 8px 6px 14px', borderBottom: '1px solid ' + v2s.border,
          background: v2s.warm, gap: 6,
        }}>
          <div style={{ fontSize: 10.5, color: v2s.sub, letterSpacing: 0.7, textTransform: 'uppercase', flex: 1 }}>
            Speaker notes
          </div>
          {/* Scroll cluster */}
          <div style={{ display: 'flex', border: '1px solid ' + v2s.border, borderRadius: 4, background: v2s.surface }}>
            <button title="Scroll up on presenter screen" style={{ width: 32, height: 30, border: 'none', background: 'transparent', color: v2s.text, cursor: 'pointer', borderRight: '1px solid ' + v2s.border, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              <svg width="14" height="14" viewBox="0 0 24 24">{v2Icon.arrowUp}</svg>
            </button>
            <button title="Scroll down on presenter screen" style={{ width: 32, height: 30, border: 'none', background: 'transparent', color: v2s.text, cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              <svg width="14" height="14" viewBox="0 0 24 24">{v2Icon.arrowDn}</svg>
            </button>
          </div>
          {/* Zoom cluster */}
          <div style={{ display: 'flex', border: '1px solid ' + v2s.border, borderRadius: 4, background: v2s.surface, alignItems: 'center' }}>
            <button title="Zoom out" style={{ width: 30, height: 30, border: 'none', background: 'transparent', color: v2s.text, cursor: 'pointer', borderRight: '1px solid ' + v2s.border, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              <svg width="14" height="14" viewBox="0 0 24 24">{v2Icon.minus}</svg>
            </button>
            <div style={{ fontFamily: v2s.mono, fontSize: 11, color: v2s.sub, padding: '0 8px', minWidth: 32, textAlign: 'center' }}>18px</div>
            <button title="Zoom in" style={{ width: 30, height: 30, border: 'none', background: 'transparent', color: v2s.text, cursor: 'pointer', borderLeft: '1px solid ' + v2s.border, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              <svg width="14" height="14" viewBox="0 0 24 24">{v2Icon.plus}</svg>
            </button>
          </div>
        </div>
        {/* Notes body */}
        <div style={{ padding: '14px 18px', fontFamily: v2s.serif, fontSize: 18, lineHeight: '28px', color: v2s.text, flex: 1, overflow: 'hidden' }}>
          Welcome everyone — thanks for joining us today.<br/><br/>
          Three things to cover:<br/>
          • Quarterly results<br/>
          • New venue rollout<br/>
          • How you can get involved
        </div>
      </div>

      <div style={{ padding: '14px 18px 14px', display: 'flex', gap: 10 }}>
        <button style={{
          flex: 1, height: 72, background: v2s.surface, border: '1px solid ' + v2s.border,
          borderRadius: 4, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', gap: 2,
          color: v2s.text, cursor: 'pointer',
        }}>
          <svg width="20" height="20" viewBox="0 0 24 24">{v2Icon.chevL}</svg>
          <span style={{ fontSize: 11, color: v2s.sub, letterSpacing: 0.5, textTransform: 'uppercase' }}>Previous</span>
        </button>
        <button style={{
          flex: 1, height: 72, background: '#333', border: '1px solid #333',
          borderRadius: 4, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', gap: 2,
          color: '#fff', cursor: 'pointer',
        }}>
          <svg width="20" height="20" viewBox="0 0 24 24">{v2Icon.chevR}</svg>
          <span style={{ fontSize: 11, color: 'rgba(255,255,255,0.7)', letterSpacing: 0.5, textTransform: 'uppercase' }}>Next slide</span>
        </button>
      </div>
      <nav style={{ borderTop: '1px solid ' + v2s.border, background: v2s.surface, display: 'flex', padding: '6px 0 12px' }}>
        {[
          { t: 'Remote', i: 'home', active: true },
          { t: 'Controls', i: 'control', active: false },
          { t: 'Settings', i: 'settings', active: false },
        ].map((x, i) => (
          <button key={i} style={{
            flex: 1, background: 'transparent', border: 'none', cursor: 'pointer',
            display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 2, padding: '6px 0',
            color: x.active ? v2s.text : v2s.sub,
          }}>
            <svg width="18" height="18" viewBox="0 0 24 24">{v2Icon[x.i]}</svg>
            <span style={{ fontSize: 10, fontWeight: x.active ? 500 : 400, letterSpacing: 0.2 }}>{x.t}</span>
          </button>
        ))}
      </nav>
    </div>
  );
}

// ════════════════════════════════════════════════════════════
// V2-C: Swap next-slide thumbnail ⇄ notes (notes-as-primary mode)
// The preview section collapses to a small strip, notes take the full
// middle of the screen. Good for speakers; previews stay visible.
// ════════════════════════════════════════════════════════════
function V2C_NotesExpanded({ timerState = 'critical' }) {
  return (
    <div style={{ width: 390, height: 844, background: v2s.page, fontFamily: v2s.font, display: 'flex', flexDirection: 'column' }}>
      <div style={{ padding: '14px 18px 10px', display: 'flex', alignItems: 'center', gap: 10, fontSize: 11.5, color: v2s.sub, letterSpacing: 0.3 }}>
        <V2Dot tone="ok"/>
        <span style={{ color: v2s.text, fontWeight: 500 }}>Stage Left Mac</span>
        <span style={{ flex: 1 }}/>
        <span>3 / 24</span>
      </div>

      <div style={{ padding: '0 18px 12px' }}>
        <StagetimerCard state={timerState} name="Closing" message="30 seconds to hard out"/>
      </div>

      {/* Slide strip — small, shows both current and next */}
      <div style={{ padding: '0 18px 12px', display: 'flex', gap: 10, alignItems: 'center' }}>
        <div style={{ flex: 1, display: 'flex', gap: 8, alignItems: 'center' }}>
          <V2SlideThumb w={96} h={54} label="3"/>
          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ fontSize: 10.5, color: v2s.sub, letterSpacing: 0.7, textTransform: 'uppercase' }}>Now</div>
            <div style={{ fontSize: 12.5, color: v2s.text, fontWeight: 500, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', marginTop: 1 }}>Three things to cover</div>
          </div>
        </div>
        <div style={{ width: 1, height: 36, background: v2s.border }}/>
        <div style={{ flex: 1, display: 'flex', gap: 8, alignItems: 'center' }}>
          <V2SlideThumb w={72} h={40} accent="#a35a7b"/>
          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ fontSize: 10.5, color: v2s.sub, letterSpacing: 0.7, textTransform: 'uppercase' }}>Next</div>
            <div style={{ fontSize: 12.5, color: v2s.text, fontWeight: 500, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', marginTop: 1 }}>The new venue rollout</div>
          </div>
        </div>
      </div>

      {/* Notes — dominant, serif, toolbar on top */}
      <div style={{ flex: 1, margin: '0 18px', minHeight: 0, background: v2s.surface, border: '1px solid ' + v2s.border, borderRadius: 4, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
        <div style={{
          display: 'flex', alignItems: 'center', padding: '6px 8px 6px 14px', borderBottom: '1px solid ' + v2s.border,
          background: v2s.warm, gap: 6,
        }}>
          <div style={{ fontSize: 10.5, color: v2s.sub, letterSpacing: 0.7, textTransform: 'uppercase', flex: 1 }}>
            Speaker notes · slide 3
          </div>
          <div style={{ display: 'flex', border: '1px solid ' + v2s.border, borderRadius: 4, background: v2s.surface }}>
            <button title="Scroll notes up" style={{ width: 32, height: 30, border: 'none', background: 'transparent', color: v2s.text, cursor: 'pointer', borderRight: '1px solid ' + v2s.border, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              <svg width="14" height="14" viewBox="0 0 24 24">{v2Icon.arrowUp}</svg>
            </button>
            <button title="Scroll notes down" style={{ width: 32, height: 30, border: 'none', background: 'transparent', color: v2s.text, cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              <svg width="14" height="14" viewBox="0 0 24 24">{v2Icon.arrowDn}</svg>
            </button>
          </div>
          <div style={{ display: 'flex', border: '1px solid ' + v2s.border, borderRadius: 4, background: v2s.surface, alignItems: 'center' }}>
            <button title="Zoom out" style={{ width: 30, height: 30, border: 'none', background: 'transparent', color: v2s.text, cursor: 'pointer', borderRight: '1px solid ' + v2s.border, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              <svg width="14" height="14" viewBox="0 0 24 24">{v2Icon.minus}</svg>
            </button>
            <button title="Zoom in" style={{ width: 30, height: 30, border: 'none', background: 'transparent', color: v2s.text, cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              <svg width="14" height="14" viewBox="0 0 24 24">{v2Icon.plus}</svg>
            </button>
          </div>
        </div>
        <div style={{ padding: '16px 18px', fontFamily: v2s.serif, fontSize: 19, lineHeight: '30px', color: v2s.text, flex: 1, overflow: 'hidden' }}>
          Welcome everyone — thanks for joining us today.<br/><br/>
          Three things to cover:<br/>
          • Quarterly results<br/>
          • The new venue rollout<br/>
          • How you can get involved<br/><br/>
          Let's dive in.
        </div>
      </div>

      <div style={{ padding: '14px 18px 14px', display: 'flex', gap: 10 }}>
        <button style={{
          flex: 1, height: 72, background: v2s.surface, border: '1px solid ' + v2s.border,
          borderRadius: 4, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', gap: 2,
          color: v2s.text, cursor: 'pointer',
        }}>
          <svg width="20" height="20" viewBox="0 0 24 24">{v2Icon.chevL}</svg>
          <span style={{ fontSize: 11, color: v2s.sub, letterSpacing: 0.5, textTransform: 'uppercase' }}>Previous</span>
        </button>
        <button style={{
          flex: 1, height: 72, background: '#333', border: '1px solid #333',
          borderRadius: 4, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', gap: 2,
          color: '#fff', cursor: 'pointer',
        }}>
          <svg width="20" height="20" viewBox="0 0 24 24">{v2Icon.chevR}</svg>
          <span style={{ fontSize: 11, color: 'rgba(255,255,255,0.7)', letterSpacing: 0.5, textTransform: 'uppercase' }}>Next slide</span>
        </button>
      </div>
      <nav style={{ borderTop: '1px solid ' + v2s.border, background: v2s.surface, display: 'flex', padding: '6px 0 12px' }}>
        {[
          { t: 'Remote', i: 'home', active: true },
          { t: 'Controls', i: 'control', active: false },
          { t: 'Settings', i: 'settings', active: false },
        ].map((x, i) => (
          <button key={i} style={{
            flex: 1, background: 'transparent', border: 'none', cursor: 'pointer',
            display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 2, padding: '6px 0',
            color: x.active ? v2s.text : v2s.sub,
          }}>
            <svg width="18" height="18" viewBox="0 0 24 24">{v2Icon[x.i]}</svg>
            <span style={{ fontSize: 10, fontWeight: x.active ? 500 : 400, letterSpacing: 0.2 }}>{x.t}</span>
          </button>
        ))}
      </nav>
    </div>
  );
}

// Standalone stagetimer state strip (for the study section)
function TimerStateStrip({ state, name = 'Opening remarks', message }) {
  return (
    <div style={{ width: 320 }}>
      <div style={{ fontSize: 10.5, color: v2s.sub, letterSpacing: 0.7, textTransform: 'uppercase', marginBottom: 6 }}>
        {state}
      </div>
      <StagetimerCard state={state} time={state === 'overtime' ? '00:42' : state === 'critical' ? '00:28' : state === 'warning' ? '01:15' : state === 'running' ? '12:34' : '—:—'} name={name} message={message}/>
    </div>
  );
}

Object.assign(window, { V2A_NotesRail, V2B_NotesToolbar, V2C_NotesExpanded, TimerStateStrip, StagetimerCard });
