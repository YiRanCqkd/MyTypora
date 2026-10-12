import { beforeEach, describe, expect, it } from 'vitest';
import { defaultExportOptions, readExportOptions, setExportResources } from './export-options';

describe('导出设置迁移', () => {
  beforeEach(() => {
    localStorage.clear();
    setExportResources(defaultExportOptions());
  });
  it('缺失或损坏设置恢复默认，旧设置只覆盖已有字段', () => {
    expect(readExportOptions().imageWidth).toBe(1280);
    localStorage.setItem('unireadmd.export-options', '{');
    expect(readExportOptions().jpegQuality).toBe(92);
    localStorage.setItem('unireadmd.export-options', JSON.stringify({
      title: '标题', imageWidth: '坏值',
    }));
    expect(readExportOptions().title).toBe('标题');
    expect(readExportOptions().imageWidth).toBe(1280);
  });
  it('持久化的路径不能冒充本窗口选取资源，清除后不再传递', () => {
    localStorage.setItem('unireadmd.export-options', JSON.stringify({
      referenceDoc: 'D:/旧参考.docx',
    }));
    expect(readExportOptions().referenceDoc).toBe('');
    setExportResources({ ...defaultExportOptions(), referenceDoc: 'D:/新参考.docx' });
    expect(readExportOptions().referenceDoc).toBe('D:/新参考.docx');
    setExportResources(defaultExportOptions());
    expect(readExportOptions().referenceDoc).toBe('');
  });
});
