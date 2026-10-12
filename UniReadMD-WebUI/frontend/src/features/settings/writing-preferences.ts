export const WRITING_KEY = 'unireadmd.writing.v1';
export const THEME_NAMES = ['paper', 'book', 'sepia', 'graphite', 'midnight'] as const;
export type WritingTheme = typeof THEME_NAMES[number];
export interface WritingPreferences {
  statUnit: 'words' | 'characters';
  readingSpeed: number;
  equationNumbering: boolean;
  lightTheme: WritingTheme;
  darkTheme: WritingTheme;
}

export function normalizeWritingPreferences(value: Partial<WritingPreferences> = {}):
  WritingPreferences {
  const theme = (input: unknown, fallback: WritingTheme) =>
    THEME_NAMES.includes(input as WritingTheme) ? input as WritingTheme : fallback;
  return {
    statUnit: value.statUnit === 'characters' ? 'characters' : 'words',
    readingSpeed: Number.isInteger(value.readingSpeed)
      && value.readingSpeed! >= 50 && value.readingSpeed! <= 2000 ? value.readingSpeed! : 220,
    equationNumbering: value.equationNumbering === true,
    lightTheme: theme(value.lightTheme, 'paper'), darkTheme: theme(value.darkTheme, 'graphite'),
  };
}

export function readWritingPreferences() {
  try {
    const saved = JSON.parse(localStorage.getItem(WRITING_KEY) || '{}');
    return normalizeWritingPreferences(saved && typeof saved === 'object' ? saved : {});
  } catch { return normalizeWritingPreferences(); }
}
