import MarkdownIt from 'markdown-it';
import { rememberDocumentWords } from './document-text';
import { highlightCode, headingSlug } from './code-highlight';
import { sanitizeMarkdownHtml } from './safe-html';
import 'katex/dist/katex.min.css';
import * as markdownItEmoji from 'markdown-it-emoji';
import markdownItFootnote from 'markdown-it-footnote';
import markdownItFrontMatter from 'markdown-it-front-matter';
import markdownItMark from 'markdown-it-mark';
import markdownItSub from 'markdown-it-sub';
import markdownItSup from 'markdown-it-sup';
import markdownItTaskLists from 'markdown-it-task-lists';
import markdownItTocDoneRight from 'markdown-it-toc-done-right';
import type { PluginWithParams } from 'markdown-it';
import { HighlightStyle } from '@codemirror/language';
import { tags } from '@lezer/highlight';
import type { createWorkbenchRuntime } from '../editor/workbench-runtime';
type CreateMarkdownRendererDependencies =
  ReturnType<typeof createWorkbenchRuntime>;

export function createMarkdownRenderer(input: CreateMarkdownRendererDependencies) {
  const { runtime,
  } = input;

  const SESSION_STORAGE_KEY = 'unireadmd.phase2.session.v1';

  const RECENT_FILES_LIMIT = 12;

  const CONTENT_ZOOM_DEFAULT = 100;

  const CONTENT_ZOOM_MIN = 50;

  const CONTENT_ZOOM_MAX = 200;

  const CONTENT_ZOOM_STEP = 10;

  const SIDEBAR_WIDTH_DEFAULT = 300;

  const SIDEBAR_WIDTH_MIN = 220;

  const SIDEBAR_WIDTH_MAX = 520;

  const MAIN_PREVIEW_WIDTH_DEFAULT_RATIO = 0.58;

  const MAIN_PREVIEW_WIDTH_MIN = 320;

  const MAIN_SOURCE_WIDTH_MIN = 320;

  const PREVIEW_INLINE_EDIT_SYNC_DELAY = 120;

  const SOURCE_MAPPED_BLOCK_TOKENS = [
    'paragraph_open',
    'heading_open',
    'blockquote_open',
    'list_item_open',
    'table_open',
    'code_block',
  ];

  const PREVIEW_TOKEN_LABELS: Record<string, string> = {
    paragraph_open: '段落',
    heading_open: '标题',
    blockquote_open: '引用块',
    list_item_open: '列表项',
    table_open: '表格',
    fence: '代码块',
    code_block: '代码块',
    math_block: '公式',
  };

  const sourceHighlightStyle = HighlightStyle.define([
    {
      tag: [
        tags.heading,
        tags.propertyName,
        tags.labelName,
      ],
      color: '#cebcca',
      fontWeight: '400',
      textDecoration: 'none',
    },
    {
      tag: tags.comment,
      color: '#9fb1ff',
    },
    {
      tag: [
        tags.string,
        tags.special(tags.string),
      ],
      color: '#a7a7d9',
    },
    {
      tag: [
        tags.number,
        tags.atom,
        tags.bool,
        tags.null,
      ],
      color: '#848695',
      fontStyle: 'italic',
    },
    {
      tag: [
        tags.link,
        tags.url,
      ],
      color: '#95b94b',
      textDecoration: 'none',
    },
  ]);

  const markdown = new MarkdownIt({
    highlight: highlightCode,
    html: true,
    linkify: true,
    breaks: true,
  });
  // 原始 HTML 不得伪造编辑源映射和图表能力属性。
  for (const rule of ['html_block', 'html_inline']) {
    markdown.renderer.rules[rule] = (tokens, index) => {
      return sanitizeMarkdownHtml(tokens[index].content, { trustedAttributes: false });
    };
  }

  function resolveMarkdownPlugin(pluginLike: unknown, exportKey?: string) {
    if (typeof pluginLike === 'function') {
      return pluginLike as PluginWithParams;
    }

    if (!pluginLike || typeof pluginLike !== 'object') {
      return null;
    }

    const record = pluginLike as Record<string, unknown>;
    if (exportKey && typeof record[exportKey] === 'function') {
      return record[exportKey] as PluginWithParams;
    }

    if (typeof record.default === 'function') {
      return record.default as PluginWithParams;
    }

    for (const candidateKey of ['full', 'light', 'bare']) {
      if (typeof record[candidateKey] === 'function') {
        return record[candidateKey] as PluginWithParams;
      }
    }

    return null;
  }

  function useMarkdownPlugin(pluginLike: unknown, ...params: unknown[]) {
    const plugin = resolveMarkdownPlugin(pluginLike);
    if (!plugin) {
      console.warn('[markdown] 插件加载失败：未找到可用导出。');
      return;
    }

    markdown.use(plugin as PluginWithParams, ...params);
  }

  useMarkdownPlugin(markdownItFrontMatter, (frontMatter: string) => {
    runtime.latestFrontMatter = frontMatter;
  });

  useMarkdownPlugin(markdownItFootnote);

  useMarkdownPlugin(markdownItSub);

  useMarkdownPlugin(markdownItSup);

  useMarkdownPlugin(markdownItMark);

  useMarkdownPlugin(markdownItTaskLists, {
    enabled: true,
    label: true,
    labelAfter: true,
  });

  useMarkdownPlugin(resolveMarkdownPlugin(markdownItEmoji, 'full'));

  useMarkdownPlugin(markdownItTocDoneRight, {
    slugify: headingSlug,
    level: [1, 2, 3, 4, 5, 6],
    containerClass: 'md-toc',
  });
  markdown.core.ruler.push('stable-heading-ids', (state) => {
    // 复用预览解析结果统计字数，避免大文档再次进行完整 Markdown 解析。
    rememberDocumentWords(state.src, state.tokens);
    const used = new Set<string>();
    state.tokens.forEach((token, index) => {
      if (token.type !== 'heading_open') return;
      const label = state.tokens[index + 1].children?.filter((child) => {
        return ['text', 'code_inline'].includes(child.type);
      }).map((child) => child.content).join('') || '';
      const base = headingSlug(label);
      let id = base;
      let suffix = 1;
      while (used.has(id)) id = `${base}-${suffix++}`;
      used.add(id);
      token.attrSet('id', id);
    });
  });

  return { SESSION_STORAGE_KEY,RECENT_FILES_LIMIT,CONTENT_ZOOM_DEFAULT,
    CONTENT_ZOOM_MIN,CONTENT_ZOOM_MAX,CONTENT_ZOOM_STEP,SIDEBAR_WIDTH_DEFAULT,SIDEBAR_WIDTH_MIN,
    SIDEBAR_WIDTH_MAX,MAIN_PREVIEW_WIDTH_DEFAULT_RATIO,MAIN_PREVIEW_WIDTH_MIN,
    MAIN_SOURCE_WIDTH_MIN,PREVIEW_INLINE_EDIT_SYNC_DELAY,SOURCE_MAPPED_BLOCK_TOKENS,
    PREVIEW_TOKEN_LABELS,sourceHighlightStyle,markdown,resolveMarkdownPlugin,useMarkdownPlugin,
  };
}
