/**
 * Interface language: English or Arabic. English strings are the keys (`t('Save')`), so components stay
 * readable; `i18n.ar.ts` holds the Arabic for each. Arabic mirrors the interface (`<html dir="rtl">`)
 * and shows Arabic-Indic digits (٠١٢٣). Slide content is never translated or mirrored.
 */
import { create } from 'zustand';
import { AR, AR_PLURALS, type PluralForms } from './i18n.ar';
import { storage } from './storage';
import { setFallbackDir } from './richtext';

export type Lang = 'en' | 'ar';
/** 'auto' follows the Windows (or browser) display language. */
export type LangSetting = Lang | 'auto';

const KEY = 'next-office:lang';

let lang: Lang = 'en';

export const useLang = create<{ lang: Lang; setting: LangSetting }>(() => ({ lang: 'en', setting: 'auto' }));

const systemLang = (): Lang => (/^ar\b/i.test(globalThis.navigator?.language ?? '') ? 'ar' : 'en');

function apply(setting: LangSetting) {
  lang = setting === 'auto' ? systemLang() : setting;
  setFallbackDir(lang === 'ar' ? 'rtl' : 'ltr');
  if (typeof document !== 'undefined') {
    document.documentElement.lang = lang;
    document.documentElement.dir = lang === 'ar' ? 'rtl' : 'ltr';
  }
  useLang.setState({ lang, setting });
}

/** Call once at start-up, after storage has loaded. */
export function initLang() {
  const saved = storage.getItem(KEY);
  apply(saved === 'en' || saved === 'ar' ? saved : 'auto');
}

export function setLang(setting: LangSetting) {
  try {
    storage.setItem(KEY, setting);
  } catch {
    // Storage full or blocked: the choice still applies to this session.
  }
  apply(setting);
}

export const currentLang = () => lang;

/** Locale for dates and times: Arabic with Arabic-Indic digits, or the system's. */
export const dateLocale = () => (lang === 'ar' ? 'ar-EG' : undefined);

const AR_DIGITS = '٠١٢٣٤٥٦٧٨٩';
/** Western digits → Arabic-Indic in the Arabic interface. */
export const digits = (s: string | number) => (lang === 'ar' ? String(s).replace(/[0-9]/g, (d) => AR_DIGITS[+d]) : String(s));

/** A percentage in the interface's digits: 75% or ٧٥٪. */
export const percent = (n: number) => (lang === 'ar' ? `${digits(n)}٪` : `${n}%`);

/** Arabic-Indic (and Persian) digits typed into a number box → Western, so it parses. */
export const parseNum = (s: string) =>
  parseFloat(s.replace(/[٠-٩]/g, (d) => String(d.charCodeAt(0) - 0x660)).replace(/[۰-۹]/g, (d) => String(d.charCodeAt(0) - 0x6f0)).replace('٫', '.'));

/**
 * Translate an interface string. `{name}` placeholders are filled from `vars`: numbers are shown in the
 * interface's digits, other values (file names, user text) as they are.
 */
export function t(key: string, vars?: Record<string, string | number>): string {
  const template = lang === 'ar' ? digits(AR[key] ?? key) : key;
  if (!vars) return template;
  return template.replace(/\{(\w+)\}/g, (m, k: string) => {
    const v = vars[k];
    return v === undefined ? m : typeof v === 'number' ? digits(v) : v;
  });
}

/**
 * Count-dependent text: `tp(n, '{n} slide', '{n} slides')`. Arabic entries (keyed by the plural English)
 * give a form per category of `Intl.PluralRules('ar')`: zero, one, two, few, many, other.
 */
export function tp(n: number, one: string, other: string): string {
  if (lang === 'ar') {
    const forms = AR_PLURALS[other];
    if (forms) {
      const cat = new Intl.PluralRules('ar').select(n) as keyof PluralForms;
      return (forms[cat] ?? forms.other).replace('{n}', digits(n));
    }
  }
  return (n === 1 ? one : other).replace('{n}', String(n));
}

export const isRtlUi = () => lang === 'ar';
