// Updates the installer from src/fileTypes.json (the one list of file types Next Office opens):
// - src-tauri/tauri.conf.json: the installer claims Next Office's own formats (.nxslides, .nxsheet).
// - src-tauri/windows/hooks.nsh: the uninstaller removes the registration the app writes at start-up.
// Run by `npm run app:build`; a test fails if these files are out of date.
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const { types } = JSON.parse(readFileSync(join(root, 'src/fileTypes.json'), 'utf8'));

// Must match src-tauri/src/file_types.rs.
const APP_NAME = 'Next Office';
const PROGID_PREFIX = 'NextOffice';

export function fileAssociations() {
  return types.filter((t) => t.native).map((t) => ({ ext: [t.ext], name: t.name, description: t.name, role: 'Editor' }));
}

export function hooks(exeName) {
  const lines = types.flatMap((t) => [
    `    DeleteRegKey HKCU "Software\\Classes\\${PROGID_PREFIX}.${t.ext}"`,
    `    DeleteRegValue HKCU "Software\\Classes\\.${t.ext}\\OpenWithProgids" "${PROGID_PREFIX}.${t.ext}"`,
  ]);
  return [
    '; Generated from src/fileTypes.json by scripts/sync-file-types.mjs. Do not edit by hand.',
    '',
    '; Uninstalling removes what Next Office registered with Windows at start-up ("Open with" entries and',
    '; its page in Settings > Default apps). An update keeps it.',
    '!macro NSIS_HOOK_POSTUNINSTALL',
    '  ${If} $UpdateMode <> 1',
    ...lines,
    `    DeleteRegKey HKCU "Software\\Classes\\Applications\\${exeName}"`,
    `    DeleteRegKey HKCU "Software\\${PROGID_PREFIX}"`,
    `    DeleteRegValue HKCU "Software\\RegisteredApplications" "${APP_NAME}"`,
    "    System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'",
    '  ${EndIf}',
    '!macroend',
    '',
  ].join('\r\n');
}

const confPath = join(root, 'src-tauri/tauri.conf.json');

export function sync({ check = false } = {}) {
  const conf = JSON.parse(readFileSync(confPath, 'utf8'));
  const exeName = `${conf.mainBinaryName ?? conf.productName}.exe`;
  conf.bundle.fileAssociations = fileAssociations();
  conf.bundle.windows.nsis = { ...conf.bundle.windows.nsis, installerHooks: 'windows/hooks.nsh' };
  const files = [
    [confPath, JSON.stringify(conf, null, 2) + '\n'],
    [join(root, 'src-tauri/windows/hooks.nsh'), hooks(exeName)],
  ];
  const stale = files.filter(([p, text]) => {
    try {
      return readFileSync(p, 'utf8').replace(/\r\n/g, '\n') !== text.replace(/\r\n/g, '\n');
    } catch {
      return true;
    }
  });
  if (!check)
    for (const [p, text] of stale) {
      mkdirSync(dirname(p), { recursive: true });
      writeFileSync(p, text);
    }
  return stale.map(([p]) => p);
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  const changed = sync();
  console.log(changed.length ? `Updated:\n  ${changed.join('\n  ')}` : 'File types are in sync.');
}
