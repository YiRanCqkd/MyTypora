import { describe, expect, it } from 'vitest';
import {
  buildInlineShortcutPatch,
  indentMarkdownBlockLines,
} from './markdown-shortcuts';

describe('markdown-shortcuts utils', () => {
  it('buildInlineShortcutPatch 应支持加粗与反向取消加粗', () => {
    const wrapPatch = buildInlineShortcutPatch('hello', 'bold');
    expect(wrapPatch).toEqual({
      insert: '**hello**',
      selectionStart: 2,
      selectionEnd: 7,
      changed: true,
    });

    const unwrapPatch = buildInlineShortcutPatch('**hello**', 'bold');
    expect(unwrapPatch).toEqual({
      insert: 'hello',
      selectionStart: 0,
      selectionEnd: 5,
      changed: true,
    });
  });

  it('buildInlineShortcutPatch 应支持斜体与空选区占位', () => {
    const wrapPatch = buildInlineShortcutPatch('title', 'italic');
    expect(wrapPatch.insert).toBe('*title*');
    expect(wrapPatch.selectionStart).toBe(1);
    expect(wrapPatch.selectionEnd).toBe(6);

    const emptyPatch = buildInlineShortcutPatch('', 'italic');
    expect(emptyPatch.insert).toBe('**');
    expect(emptyPatch.selectionStart).toBe(1);
    expect(emptyPatch.selectionEnd).toBe(1);
  });

  it('buildInlineShortcutPatch 应支持链接包裹与取消包裹', () => {
    const wrapPatch = buildInlineShortcutPatch('OpenAI', 'link');
    expect(wrapPatch).toEqual({
      insert: '[OpenAI](https://)',
      selectionStart: 9,
      selectionEnd: 17,
      changed: true,
    });

    const unwrapPatch = buildInlineShortcutPatch('[OpenAI](https://openai.com)', 'link');
    expect(unwrapPatch).toEqual({
      insert: 'OpenAI',
      selectionStart: 0,
      selectionEnd: 6,
      changed: true,
    });
  });

  it('indentMarkdownBlockLines 应支持缩进与反缩进', () => {
    const source = '- item 1\n  - item 2';
    const indentPatch = indentMarkdownBlockLines(source);
    expect(indentPatch).toEqual({
      text: '  - item 1\n    - item 2',
      changed: true,
    });

    const outdentPatch = indentMarkdownBlockLines(indentPatch.text, {
      outdent: true,
    });
    expect(outdentPatch).toEqual({
      text: source,
      changed: true,
    });
  });
});
