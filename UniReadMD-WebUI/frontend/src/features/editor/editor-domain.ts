import type { ResolveWorkbenchDomains } from './workbench-domains';
import { createDeferredFeature } from './workbench-domains';
import { useCodeMirrorFeature } from './codemirror-feature';
import { useEditorCommands } from './editor-commands';
import { useEditorNavigation } from './editor-navigation';
import { useEditorPresentation } from './editor-presentation';
import { usePreviewEditingEventsFeature } from './preview-editing-events-feature';
import { usePreviewEditingFeature } from './preview-editing-feature';
import { useSourceEditingFeature } from './source-editing-feature';
import { useTableEditingCore } from './table-editing-core';
import { useTableEditingFeature } from './table-editing-feature';
import { useWorkbenchFoundation } from './workbench-foundation';
import { useWorkbenchRuntimeFeature } from './workbench-runtime-feature';

export function useEditorDomain(resolve: ResolveWorkbenchDomains) {
  const foundation = useWorkbenchFoundation();
  const presentation = createDeferredFeature(() => {
    const domains = resolve();
    return useEditorPresentation({
      editor: foundation,
      markdown: domains.markdown.feature(),
      settings: domains.settings.foundation(),
      workspace: domains.workspace.foundation(),
    });
  });
  const navigation = createDeferredFeature(() => {
    return useEditorNavigation({ editor: foundation, presentation: presentation() });
  });
  const preview = createDeferredFeature(() => {
    const domains = resolve();
    return usePreviewEditingFeature({
      editor: foundation,
      markdown: domains.markdown.feature(),
      navigation: navigation(),
      presentation: presentation(),
      settings: domains.settings.foundation(),
    });
  });
  const table = createDeferredFeature(() => {
    return useTableEditingCore({
      editor: foundation,
      navigation: navigation(),
      presentation: presentation(),
      preview: preview(),
    });
  });
  const events = createDeferredFeature(() => {
    const domains = resolve();
    return usePreviewEditingEventsFeature({
      editor: foundation,
      presentation: presentation(),
      preview: preview(),
      settings: domains.settings.foundation(),
      table: table(),
    });
  });
  const source = createDeferredFeature(() => {
    const domains = resolve();
    return useSourceEditingFeature({
      editor: foundation,
      navigation: navigation(),
      presentation: presentation(),
      preview: preview(),
      settings: domains.settings.foundation(),
      table: table(),
    });
  });
  const tableFeature = createDeferredFeature(() => {
    const domains = resolve();
    return useTableEditingFeature({
      editor: foundation,
      events: events(),
      navigation: navigation(),
      presentation: presentation(),
      preview: preview(),
      settings: domains.settings.foundation(),
      source: source(),
      table: table(),
      workspace: domains.workspace.presentation(),
    });
  });
  const codeMirror = createDeferredFeature(() => {
    const domains = resolve();
    return useCodeMirrorFeature({
      editor: foundation,
      markdown: domains.markdown.feature(),
      navigation: navigation(),
      presentation: presentation(),
      settings: domains.settings.foundation(),
      source: source(),
      workspace: domains.workspace.presentation(),
    });
  });
  const commands = createDeferredFeature(() => {
    const domains = resolve();
    return useEditorCommands({
      editor: foundation,
      exporting: domains.export.feature(),
      navigation: navigation(),
      presentation: presentation(),
      preview: preview(),
      settings: domains.settings.foundation(),
      settingsActions: domains.settings.actions(),
      settingsSession: domains.settings.session(),
      source: source(),
      table: table(),
      workspaceActions: domains.workspace.actions(),
      workspacePresentation: domains.workspace.presentation(),
    });
  });
  const runtime = createDeferredFeature(() => {
    const domains = resolve();
    return useWorkbenchRuntimeFeature({
      lifecycle: {
        editor: {
          codeMirror: codeMirror(),commands: commands(),foundation,navigation: navigation(),
          preview: preview(),
        },
        export: domains.export.feature(),
        markdown: {
          diagrams: domains.markdown.diagrams(),renderer: domains.markdown.feature(),
        },
        settings: {
          actions: domains.settings.actions(),foundation: domains.settings.foundation(),
          nativeCommands: domains.settings.nativeCommands(),session: domains.settings.session(),
        },
        workspace: {
          actions: domains.workspace.actions(),foundation: domains.workspace.foundation(),
          presentation: domains.workspace.presentation(),
        },
      },
      viewport: {
        editor: {
          codeMirror: codeMirror(),foundation,navigation: navigation(),
          presentation: presentation(),preview: preview(),source: source(),table: table(),
        },
        markdown: { diagrams: domains.markdown.diagrams() },
        settings: { foundation: domains.settings.foundation() },
        workspace: {
          actions: domains.workspace.actions(),foundation: domains.workspace.foundation(),
          presentation: domains.workspace.presentation(),
        },
      },
      watchers: {
        editor: {
          commands: commands(),foundation,presentation: presentation(),
        },
        settings: {
          foundation: domains.settings.foundation(),session: domains.settings.session(),
        },
        workspace: {
          foundation: domains.workspace.foundation(),presentation: domains.workspace.presentation(),
        },
      },
    });
  });

  return {
    codeMirror,commands,events,foundation,navigation,presentation,preview,runtime,source,table,
    tableFeature,
  };
}
