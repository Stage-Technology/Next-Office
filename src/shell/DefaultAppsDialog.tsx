import { CheckCircle2, Circle, FileSpreadsheet, FileText, Loader2, Presentation, X } from 'lucide-react';
import { t } from '../i18n';
import { FILE_TYPES } from './fileTypes';
import { closeDefaultApps, dontAskAgain, notNow, setAsDefault, useDefaultApps } from './defaultApps';

const GROUPS = [
  { app: 'sheet' as const, label: 'Spreadsheets', icon: <FileSpreadsheet size={16} /> },
  { app: 'slides' as const, label: 'Presentations', icon: <Presentation size={16} /> },
  { app: 'doc' as const, label: 'Documents', icon: <FileText size={16} /> },
];

/** "Open your files with Next Office?" — asked at start-up, and from File › Default apps. */
export function DefaultAppsDialog() {
  const { open, prompted, status, waiting, error } = useDefaultApps();
  if (!open) return null;
  const allSet = !!status && FILE_TYPES.every((f) => status[f.ext]);

  return (
    <div className="modal-backdrop" onMouseDown={(e) => e.target === e.currentTarget && !prompted && closeDefaultApps()}>
      <div className="modal default-apps" role="dialog" aria-modal="true" aria-labelledby="default-apps-title">
        <div className="modal-head">
          <h2 id="default-apps-title">{t('Open your files with Next Office?')}</h2>
          <button className="icon-btn" aria-label={t('Close')} title={t('Close')} onClick={prompted ? notNow : closeDefaultApps}>
            <X size={16} />
          </button>
        </div>

        <div className="default-apps-body">
          <p>
            {t(
              allSet
                ? 'Next Office opens all of these file types when you double-click them.'
                : 'Make Next Office the app that opens these files when you double-click them. You can change this any time in Windows Settings.',
            )}
          </p>
          <div className="default-apps-groups">
            {GROUPS.map((g) => (
              <section key={g.app}>
                <h3>
                  {g.icon} {t(g.label)}
                </h3>
                <ul>
                  {FILE_TYPES.filter((f) => f.app === g.app).map((f) => {
                    const ours = status?.[f.ext];
                    return (
                      <li key={f.ext} className={ours ? 'ours' : ''}>
                        {status === null ? (
                          <Loader2 size={15} className="spin" />
                        ) : ours ? (
                          <CheckCircle2 size={15} />
                        ) : (
                          <Circle size={15} />
                        )}
                        <span className="default-apps-ext" dir="ltr">
                          .{f.ext}
                        </span>
                        <span className="default-apps-name">{t(f.name)}</span>
                        <span className="default-apps-state">{status === null ? '' : t(ours ? 'Next Office' : 'Another app')}</span>
                      </li>
                    );
                  })}
                </ul>
              </section>
            ))}
          </div>
          {waiting && (
            <p className="default-apps-hint">
              {t(
                'In Settings, select "Set default" at the top of the Next Office page, then come back here. On Windows 10, choose Next Office for each file type.',
              )}
            </p>
          )}
          {error && <p className="default-apps-error">{t('Could not open Windows Settings.')}</p>}
        </div>

        <div className="modal-foot">
          {prompted && !allSet && (
            <>
              <button className="link-btn default-apps-never" onClick={dontAskAgain}>
                {t("Don't ask again")}
              </button>
              <button className="secondary-btn" onClick={notNow}>
                {t('Not now')}
              </button>
            </>
          )}
          {(!prompted || allSet) && (
            <button className="secondary-btn" onClick={closeDefaultApps}>
              {t(allSet ? 'Done' : 'Close')}
            </button>
          )}
          {!allSet && (
            <button className="primary-btn" onClick={setAsDefault}>
              {t(waiting ? 'Open Settings again' : 'Set as default')}
            </button>
          )}
        </div>
      </div>
    </div>
  );
}
