// @vitest-environment jsdom
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';

/** The splash markup from index.html, as the page starts with it. */
function mountSplash() {
  const html = readFileSync(join(__dirname, '../../index.html'), 'utf8');
  const markup = /<div id="splash"[\s\S]*?<div id="root">/.exec(html)![0].replace('<div id="root">', '');
  document.body.innerHTML = markup;
}

const splash = () => document.getElementById('splash');

async function load(elapsedMs: number) {
  vi.resetModules();
  vi.spyOn(performance, 'now').mockReturnValue(elapsedMs);
  return import('../splash');
}

beforeEach(() => {
  vi.useFakeTimers();
  vi.restoreAllMocks();
  mountSplash();
});

describe('splash screen', () => {
  it('is in the page from the first frame, with the Stage Tech logo and credit', () => {
    expect(splash()).not.toBeNull();
    expect(splash()!.querySelector('.splash-by img')!.getAttribute('src')).toBe('/stage-tech-logo.png');
    expect(splash()!.textContent).toContain('By Stage Tech');
    expect(splash()!.textContent).toContain('Next Office');
    expect(splash()!.getAttribute('role')).toBe('status');
  });

  it('stays up at least briefly when the app is ready very fast', async () => {
    const { hideSplash, SPLASH_MIN_MS } = await load(300);
    hideSplash();
    vi.advanceTimersByTime(SPLASH_MIN_MS - 300 - 1);
    expect(splash()!.classList.contains('splash-hide')).toBe(false);
    vi.advanceTimersByTime(1);
    expect(splash()!.classList.contains('splash-hide')).toBe(true);
    vi.advanceTimersByTime(400);
    expect(splash()).toBeNull();
  });

  it('adds no wait when loading already took longer than the minimum', async () => {
    const { hideSplash } = await load(2500);
    hideSplash();
    vi.advanceTimersByTime(0);
    expect(splash()!.classList.contains('splash-hide')).toBe(true);
  });

  it('a click skips it right away', async () => {
    const { installSplash } = await load(100);
    installSplash();
    splash()!.dispatchEvent(new Event('pointerdown'));
    vi.advanceTimersByTime(0);
    expect(splash()!.classList.contains('splash-hide')).toBe(true);
  });

  it('a click still skips it while it waits out the minimum time after the app is ready', async () => {
    const { hideSplash, installSplash } = await load(200);
    installSplash();
    hideSplash(); // app ready at 200 ms: would fade at 900 ms
    splash()!.dispatchEvent(new Event('pointerdown'));
    vi.advanceTimersByTime(0);
    expect(splash()!.classList.contains('splash-hide')).toBe(true);
  });

  it('never stays forever if start-up fails', async () => {
    const { installSplash, SPLASH_MAX_MS } = await load(100);
    installSplash();
    vi.advanceTimersByTime(SPLASH_MAX_MS - 100 - 1);
    expect(splash()!.classList.contains('splash-hide')).toBe(false);
    vi.advanceTimersByTime(2); // the fade starts on the next tick
    expect(splash()!.classList.contains('splash-hide')).toBe(true);
  });

  it('can be told to hide more than once without errors', async () => {
    const { hideSplash } = await load(5000);
    hideSplash();
    hideSplash({ now: true });
    vi.advanceTimersByTime(1000);
    expect(splash()).toBeNull();
    expect(() => hideSplash()).not.toThrow();
  });
});
