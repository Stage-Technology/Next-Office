import { LayoutGrid } from 'lucide-react';
import { t } from '../i18n';
import { goHome } from './shell';

/** The app logo in the title bar; clicking it goes back to the Next Office launcher. */
export function HomeButton({ letter = 'N' }: { letter?: string }) {
  return (
    <button className="brand-logo home-btn" onClick={goHome} title={t('Next Office: all apps')} aria-label={t('Next Office: all apps')}>
      <span className="home-btn-letter">{letter}</span>
      <LayoutGrid size={15} className="home-btn-icon" />
    </button>
  );
}
