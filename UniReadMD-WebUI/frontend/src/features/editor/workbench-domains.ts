import type { useExportFeature } from '../export/export-feature';
import type { useExportFoundation } from '../export/export-foundation';
import type { useDiagramFeature } from '../markdown/diagram-feature';
import type { useMarkdownFeature } from '../markdown/markdown-feature';
import type { useNativeCommandFeature } from '../settings/native-command-feature';
import type { useSettingsActions } from '../settings/settings-actions';
import type { useSettingsFoundation } from '../settings/settings-foundation';
import type { useSettingsSession } from '../settings/settings-session';
import type { useWorkspaceActions } from '../workspace/workspace-actions';
import type { useWorkspaceFoundation } from '../workspace/workspace-foundation';
import type { useWorkspacePresentation } from '../workspace/workspace-presentation';
import type { useCodeMirrorFeature } from './codemirror-feature';
import type { useEditorCommands } from './editor-commands';
import type { useEditorNavigation } from './editor-navigation';
import type { useEditorPresentation } from './editor-presentation';
import type { usePreviewEditingEventsFeature } from './preview-editing-events-feature';
import type { usePreviewEditingFeature } from './preview-editing-feature';
import type { useSourceEditingFeature } from './source-editing-feature';
import type { useTableEditingCore } from './table-editing-core';
import type { useTableEditingFeature } from './table-editing-feature';
import type { useWorkbenchFoundation } from './workbench-foundation';

export type EditorDomain = Readonly<{
  codeMirror: () => ReturnType<typeof useCodeMirrorFeature>;
  commands: () => ReturnType<typeof useEditorCommands>;
  events: () => ReturnType<typeof usePreviewEditingEventsFeature>;
  foundation: ReturnType<typeof useWorkbenchFoundation>;
  navigation: () => ReturnType<typeof useEditorNavigation>;
  presentation: () => ReturnType<typeof useEditorPresentation>;
  preview: () => ReturnType<typeof usePreviewEditingFeature>;
  runtime: () => void;
  source: () => ReturnType<typeof useSourceEditingFeature>;
  table: () => ReturnType<typeof useTableEditingCore>;
  tableFeature: () => ReturnType<typeof useTableEditingFeature>;
}>;

export type ExportDomain = Readonly<{
  feature: () => ReturnType<typeof useExportFeature>;
  foundation: () => ReturnType<typeof useExportFoundation>;
}>;

export type MarkdownDomain = Readonly<{
  diagrams: () => ReturnType<typeof useDiagramFeature>;
  feature: () => ReturnType<typeof useMarkdownFeature>;
}>;

export type SettingsDomain = Readonly<{
  actions: () => ReturnType<typeof useSettingsActions>;
  foundation: () => ReturnType<typeof useSettingsFoundation>;
  nativeCommands: () => ReturnType<typeof useNativeCommandFeature>;
  session: () => ReturnType<typeof useSettingsSession>;
}>;

export type WorkspaceDomain = Readonly<{
  actions: () => ReturnType<typeof useWorkspaceActions>;
  foundation: () => ReturnType<typeof useWorkspaceFoundation>;
  presentation: () => ReturnType<typeof useWorkspacePresentation>;
}>;

export type WorkbenchDomains = Readonly<{
  editor: EditorDomain;
  export: ExportDomain;
  markdown: MarkdownDomain;
  settings: SettingsDomain;
  workspace: WorkspaceDomain;
}>;

export type ResolveWorkbenchDomains = () => WorkbenchDomains;

export class WorkbenchDomainResolutionError extends Error {
  readonly code = 'ERR_WORKBENCH_DOMAINS_UNAVAILABLE';

  constructor() {
    super('工作台领域尚未完成装配');
    this.name = 'WorkbenchDomainResolutionError';
  }
}

export function createWorkbenchDomainResolver() {
  let domains: WorkbenchDomains | undefined;

  return {
    initialize(value: WorkbenchDomains) {
      domains = value;
    },
    resolve(): WorkbenchDomains {
      if (!domains) {
        throw new WorkbenchDomainResolutionError();
      }
      return domains;
    },
  };
}

type DeferredFeature<T> =
  | { readonly kind: 'pending' }
  | { readonly kind: 'ready'; readonly value: T };

export function createDeferredFeature<T>(factory: () => T) {
  let state: DeferredFeature<T> = { kind: 'pending' };

  return () => {
    if (state.kind === 'ready') {
      return state.value;
    }

    const value = factory();
    state = { kind: 'ready', value };
    return value;
  };
}
