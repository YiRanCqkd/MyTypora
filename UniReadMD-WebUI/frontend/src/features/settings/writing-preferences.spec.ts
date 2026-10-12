import { describe, expect, it } from 'vitest';
import { normalizeWritingPreferences, readWritingPreferences, WRITING_KEY }
  from './writing-preferences';

describe('写作设置兼容', () => {
  it('旧配置缺少新字段时继续使用默认设置，损坏设置安全回退', () => {
    localStorage.setItem(WRITING_KEY, 'invalid');
    expect(readWritingPreferences().readingSpeed).toBe(220);
    expect(normalizeWritingPreferences({ readingSpeed: 0, lightTheme: 'invalid' as never }))
      .toMatchObject({ readingSpeed: 220, lightTheme: 'paper', statUnit: 'words' });
  });
  it('保存合法阅读速度、统计单位、公式和独立主题', () => {
    localStorage.setItem(WRITING_KEY, JSON.stringify({ readingSpeed: 300,
      statUnit: 'characters', equationNumbering: true,
      lightTheme: 'sepia', darkTheme: 'midnight',
    }));
    expect(readWritingPreferences()).toEqual({ readingSpeed: 300,
      statUnit: 'characters', equationNumbering: true,
      lightTheme: 'sepia', darkTheme: 'midnight',
    });
    localStorage.removeItem(WRITING_KEY);
  });
});
