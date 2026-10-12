import MarkdownIt from 'markdown-it';
import { afterEach, describe, expect, it } from 'vitest';
import plugin from '../../vendor/markdown-it-mathjax3';
import { WRITING_KEY } from '../settings/writing-preferences';

afterEach(() => localStorage.removeItem(WRITING_KEY));
describe('文档内公式兼容', () => {
  it('先收集标签再渲染，前向引用、自动与手动编号保持一致', () => {
    localStorage.setItem(WRITING_KEY, '{"equationNumbering":true}');
    const md = new MarkdownIt().use(plugin);
    const html = md.render('前向 $\\eqref{a}$\n\n$$\nx=1\\label{a}\n$$\n\n'
      + '$$\ny=2\\tag{B}\\label{b}\n$$\n\n引用 $\\ref{b}$');
    expect(html).toContain("href='#equation-a'");
    expect(html).toContain('>(1)</a>');
    expect(html).toContain('>B</a>');
    expect(html).not.toContain('katex-error');
    expect(md.render('$\\ref{a}$')).not.toContain('math-reference');
  });
  it('化学公式和文档内宏可用，切换文档不继承宏', () => {
    const md = new MarkdownIt().use(plugin);
    expect(md.render('$\\ce{H2O}$')).not.toContain('katex-error');
    expect(md.render('$$\n\\def\\foo{123}\n$$\n\n$\\foo$'))
      .not.toContain('<mtext>\\foo</mtext>');
    expect(md.render('$\\foo$')).toContain('<mtext>\\foo</mtext>');
  });
});
