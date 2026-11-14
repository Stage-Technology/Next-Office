import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import './styles/tokens.css';
import './styles/app-accents.css';
import './fonts/fonts.css';
import './styles.css';
import './sheet/sheet.css';
import './doc/doc.css';
import './motion.css';
import { inTauri } from './desktop';
import { preloadStorage } from './storage';
import { initLang, useLang } from './i18n';
import { initTheme } from './theme';
import { Shell } from './shell/Launcher';
import { initShell, refreshTitle } from './shell/shell';
import { hideSplash, installSplash } from './splash';

// Next Office is a Windows app: the page only works inside its Tauri window.
if (!inTauri) {
  document.getElementById('splash')?.remove();
  document.getElementById('root')!.textContent = 'Next Office runs as a desktop app. Start it with: npm run app:dev';
  throw new Error('Not running in the Next Office window');
}

installSplash();

// Saved state must be loaded before the app stores read it (they are imported when an app is opened).
preloadStorage().then(() => {
  initLang();
  initTheme();
  // Re-mount on a language change so every label (including memoised ones) is translated.
  const Root = () => <Shell key={useLang((s) => s.lang)} />;
  createRoot(document.getElementById('root')!).render(
    <StrictMode>
      <Root />
    </StrictMode>,
  );
  useLang.subscribe(refreshTitle);
  // Hide the splash once the first screen has been drawn.
  initShell({ onStarted: () => requestAnimationFrame(() => requestAnimationFrame(() => hideSplash())) });
});
