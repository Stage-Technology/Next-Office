import { Check, Languages } from 'lucide-react';
import { setLang, t, useLang, type LangSetting } from '../i18n';

/** Language names are written in their own language, as in Office. */
const CHOICES: { id: LangSetting; label: () => string }[] = [
  { id: 'auto', label: () => t('Same as Windows') },
  { id: 'en', label: () => 'English' },
  { id: 'ar', label: () => 'العربية' },
];

/** The language choices as menu items (File menu). */
export function LanguageMenuItems({ onPick }: { onPick?: () => void }) {
  const setting = useLang((s) => s.setting);
  return (
    <>
      <div className="menu-label">
        <Languages size={14} /> {t('Language')}
      </div>
      {CHOICES.map((c) => (
        <button
          key={c.id}
          className="menu-item"
          role="menuitemradio"
          aria-checked={setting === c.id}
          onClick={() => {
            onPick?.();
            setLang(c.id);
          }}
        >
          <span className="menu-icon">{setting === c.id && <Check size={15} />}</span>
          <span>{c.label()}</span>
        </button>
      ))}
    </>
  );
}

/** Compact picker for the start screen. */
export function LanguageSelect() {
  const setting = useLang((s) => s.setting);
  return (
    <label className="lang-select" title={t('Language')}>
      <Languages size={16} />
      <select value={setting} aria-label={t('Language')} onChange={(e) => setLang(e.target.value as LangSetting)}>
        {CHOICES.map((c) => (
          <option key={c.id} value={c.id}>
            {c.label()}
          </option>
        ))}
      </select>
    </label>
  );
}
