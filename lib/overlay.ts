// Drop target and status card, isolated from X's styles in a shadow root.

const CSS = `
:host { all: initial; }
.drop {
  position: fixed; inset: 0; z-index: 2147483646;
  display: none; align-items: center; justify-content: center;
  background: rgba(15, 20, 25, 0.55); backdrop-filter: blur(2px);
  font: 600 22px/1.3 -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
  color: #fff; pointer-events: none;
}
.drop.on { display: flex; }
.drop > div {
  border: 3px dashed rgba(255, 255, 255, 0.85); border-radius: 20px;
  padding: 48px 64px; text-align: center; background: rgba(29, 155, 240, 0.25);
}
.drop small { display: block; margin-top: 8px; font-weight: 400; font-size: 15px; opacity: 0.85; }
.card {
  position: fixed; right: 20px; bottom: 20px; z-index: 2147483647;
  display: none; gap: 12px; align-items: flex-start; max-width: 360px;
  padding: 14px 16px; border-radius: 14px;
  background: #0f1419; color: #e7e9ea; box-shadow: 0 8px 28px rgba(0, 0, 0, 0.35);
  font: 400 14px/1.4 -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
}
.card.on { display: flex; }
.card.error { background: #5c1a1a; }
.card ul { margin: 6px 0 0; padding-left: 18px; color: #c4c9cc; }
.card button {
  margin-left: auto; border: 0; background: none; color: inherit; cursor: pointer;
  font-size: 18px; line-height: 1; padding: 0 0 0 8px; opacity: 0.7;
}
.spin {
  flex: none; width: 16px; height: 16px; margin-top: 1px; border-radius: 50%;
  border: 2px solid rgba(255, 255, 255, 0.25); border-top-color: #1d9bf0;
  animation: spin 0.8s linear infinite;
}
@keyframes spin { to { transform: rotate(360deg); } }
`;

export interface Overlay {
  showDropZone(on: boolean): void;
  busy(text: string): void;
  result(kind: 'ok' | 'error', title: string, details?: string[]): void;
}

export function createOverlay(): Overlay {
  const host = document.createElement('docx2x-overlay');
  const root = host.attachShadow({ mode: 'closed' });
  const el = (tag: string, props: Partial<HTMLElement> = {}, ...children: (Node | string)[]) => {
    const node = Object.assign(document.createElement(tag), props);
    node.append(...children);
    return node;
  };
  const drop = el(
    'div',
    { className: 'drop' },
    el('div', {}, 'Drop your .docx to import it', el('small', {}, 'The title, formatting and images go into this article')),
  );
  const card = el('div', { className: 'card', role: 'status', ariaLive: 'polite' });
  root.append(el('style', { textContent: CSS }), drop, card);
  document.documentElement.append(host);
  let hideTimer: number | undefined;

  const render = (spinner: boolean, title: string, details: string[] = [], error = false) => {
    window.clearTimeout(hideTimer);
    card.replaceChildren();
    card.className = `card on${error ? ' error' : ''}`;
    if (spinner) card.append(el('div', { className: 'spin' }));
    const body = el('div', {}, el('strong', { textContent: title }));
    if (details.length) {
      const list = el('ul', {}, ...details.map((d) => el('li', { textContent: d })));
      body.append(list);
    }
    card.append(body);
    if (!spinner) {
      const close = el('button', { textContent: '×', title: 'Dismiss' });
      close.addEventListener('click', () => card.classList.remove('on'));
      card.append(close);
      hideTimer = window.setTimeout(() => card.classList.remove('on'), error ? 20_000 : 10_000);
    }
  };

  return {
    showDropZone: (on) => drop.classList.toggle('on', on),
    busy: (text) => render(true, text),
    result: (kind, title, details) => render(false, title, details, kind === 'error'),
  };
}
