import { Check, Monitor, Moon, Sun, SunMoon } from 'lucide-react';
import { t } from '../i18n';
import { setThemeSetting, useTheme, type ThemeSetting } from '../theme';
import { Btn, Group } from './controls';

const CHOICES: { id: ThemeSetting; label: string; icon: JSX.Element }[] = [
  { id: 'system', label: 'System', icon: <Monitor size={16} /> },
  { id: 'light', label: 'Light', icon: <Sun size={16} /> },
  { id: 'dark', label: 'Dark', icon: <Moon size={16} /> },
];

/** The theme choices as menu items (File menu), under the language choices. */
export function ThemeMenuItems({ onPick }: { onPick?: () => void }) {
  const setting = useTheme((s) => s.setting);
  return (
    <>
      <div className="menu-label">
        <SunMoon size={14} /> {t('Theme')}
      </div>
      {CHOICES.map((c) => (
        <button
          key={c.id}
          className="menu-item"
          role="menuitemradio"
          aria-checked={setting === c.id}
          onClick={() => {
            onPick?.();
            setThemeSetting(c.id);
          }}
        >
          <span className="menu-icon">{setting === c.id && <Check size={15} />}</span>
          <span>{t(c.label)}</span>
        </button>
      ))}
    </>
  );
}

/** Compact picker for the launcher and the start screen. */
export function ThemeSelect() {
  const setting = useTheme((s) => s.setting);
  return (
    <label className="lang-select" title={t('Theme')}>
      <SunMoon size={16} />
      <select value={setting} aria-label={t('Theme')} onChange={(e) => setThemeSetting(e.target.value as ThemeSetting)}>
        {CHOICES.map((c) => (
          <option key={c.id} value={c.id}>
            {t(c.label)}
          </option>
        ))}
      </select>
    </label>
  );
}

/** View tab › Theme: System / Light / Dark as toggles. */
export function ThemeGroup() {
  const setting = useTheme((s) => s.setting);
  return (
    <Group label={t('Theme')}>
      <div className="rb-col">
        {CHOICES.map((c) => (
          <Btn key={c.id} icon={c.icon} label={t(c.label)} active={setting === c.id} onClick={() => setThemeSetting(c.id)} />
        ))}
      </div>
    </Group>
  );
}
