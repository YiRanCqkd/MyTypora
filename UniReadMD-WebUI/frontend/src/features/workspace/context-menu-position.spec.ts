import { describe, expect, it } from 'vitest';
import { contextMenuKeys, fitContextMenu, measuredContextMenuStyle } from './context-menu-position';

const bounds = { left: 300, top: 30, right: 1200, bottom: 750 };
describe('右键菜单实际尺寸定位', () => {
  it('空间足够时保持点击位置', () => {
    expect(fitContextMenu({ x: 400, y: 100 }, { width: 224, height: 380 }, bounds))
      .toEqual({ left: '400px', top: '100px' });
  });
  it('右下角菜单限制在可用区域中', () => {
    expect(fitContextMenu({ x: 1190, y: 740 }, { width: 224, height: 380 }, bounds))
      .toEqual({ left: '976px', top: '370px' });
  });
  it('菜单大于区域时从左上角开始，避免负坐标', () => {
    expect(fitContextMenu({ x: 0, y: 0 }, { width: 1400, height: 1000 }, bounds))
      .toEqual({ left: '300px', top: '30px' });
  });
  it('先限制最大高度，再读取真实大小定位', () => {
    const menu = document.createElement('div');
    menu.getBoundingClientRect = () => ({ width: 224, height: 100 } as DOMRect);
    expect(measuredContextMenuStyle(menu, { x: 400, y: 700 }, bounds))
      .toEqual({ left: '400px', top: '650px', maxHeight: '720px', visibility: 'visible' });
  });
  it('默认限制为窗口可视区域并保留八像素边距', () => {
    const menu = document.createElement('div');
    menu.getBoundingClientRect = () => ({ width: 224, height: 100 } as DOMRect);
    expect(measuredContextMenuStyle(menu, { x: -10, y: -10 }))
      .toMatchObject({ left: '8px', top: '8px', visibility: 'visible' });
  });
});

describe('右键菜单键盘导航', () => {
  function setup(html = '<button>甲</button><button disabled>禁用</button><button>乙</button>') {
    const menu = document.createElement('div');
    menu.innerHTML = html;
    document.body.replaceChildren(menu);
    menu.addEventListener('keydown', contextMenuKeys);
    return menu;
  }
  it('方向键跳过禁用项并循环，首尾键正确定位', () => {
    const menu = setup();
    const first = menu.querySelector('button')!;
    const last = menu.querySelectorAll('button')[2]!;
    first.focus();
    first.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    expect(document.activeElement).toBe(last);
    last.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    expect(document.activeElement).toBe(first);
    first.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowUp', bubbles: true }));
    expect(document.activeElement).toBe(last);
    last.dispatchEvent(new KeyboardEvent('keydown', { key: 'Home', bubbles: true }));
    expect(document.activeElement).toBe(first);
    first.dispatchEvent(new KeyboardEvent('keydown', { key: 'End', bubbles: true }));
    expect(document.activeElement).toBe(last);
  });
  it('输入框保留自身按键，普通按键及空菜单不处理', () => {
    const menu = setup('<input><button>应用</button>');
    const input = menu.querySelector('input')!;
    input.focus();
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    expect(document.activeElement).toBe(input);
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'x', bubbles: true }));
    expect(document.activeElement).toBe(input);
    const empty = setup('');
    empty.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    expect(document.activeElement).toBe(document.body);
  });
});
