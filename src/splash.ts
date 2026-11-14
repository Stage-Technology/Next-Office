/**
 * The "Next Office — By Stage Tech" splash screen (index.html). It covers the window while the app loads,
 * then fades out. It stays up at least briefly, so a fast start doesn't flash, but never holds the app
 * back: a click dismisses it, and it goes on its own if start-up takes unusually long.
 */

/** Shortest time on screen, counted from when the window started loading. */
export const SPLASH_MIN_MS = 900;
/** Give up waiting for the app after this long (something went wrong; don't hide the app behind it). */
export const SPLASH_MAX_MS = 8000;

let done = false;
let pending: ReturnType<typeof setTimeout> | undefined;

function remove(el: HTMLElement) {
  el.classList.add('splash-hide');
  const gone = () => el.remove();
  el.addEventListener('transitionend', gone, { once: true });
  setTimeout(gone, 400); // reduced motion or no transition
}

/** Fade the splash out (after the minimum time). Safe to call more than once. */
export function hideSplash({ now = false } = {}) {
  const el = document.getElementById('splash');
  if (!el) return;
  // Already fading, or waiting out the minimum time: only "now" (a click) changes anything.
  if (done && (!now || pending === undefined)) return;
  done = true;
  clearTimeout(pending);
  const wait = now ? 0 : Math.max(0, SPLASH_MIN_MS - performance.now());
  pending = setTimeout(() => {
    pending = undefined;
    remove(el);
  }, wait);
}

/** Click to skip, and the safety timeout. */
export function installSplash() {
  const el = document.getElementById('splash');
  if (!el) return;
  el.addEventListener('pointerdown', () => hideSplash({ now: true }));
  setTimeout(() => hideSplash({ now: true }), Math.max(0, SPLASH_MAX_MS - performance.now()));
}
