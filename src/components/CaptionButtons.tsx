import { useEffect, useState } from 'react';
import { windowControls } from '../desktop';
import { t } from '../i18n';

/** Minimize / maximize / close, drawn like Windows 11 since the window is frameless. */
export function CaptionButtons() {
  const [maximized, setMaximized] = useState(false);
  useEffect(() => {
    const off = windowControls.onMaximizedChange(setMaximized);
    return () => void off.then((f) => f());
  }, []);
  return (
    <div className="caption-buttons">
      <button aria-label={t('Minimize')} title={t('Minimize')} onClick={windowControls.minimize}>
        {'\uE921'}
      </button>
      <button
        aria-label={t(maximized ? 'Restore' : 'Maximize')}
        title={t(maximized ? 'Restore Down' : 'Maximize')}
        onClick={windowControls.toggleMaximize}
      >
        {maximized ? '\uE923' : '\uE922'}
      </button>
      <button className="caption-close" aria-label={t('Close')} title={t('Close')} onClick={windowControls.close}>
        {'\uE8BB'}
      </button>
    </div>
  );
}
