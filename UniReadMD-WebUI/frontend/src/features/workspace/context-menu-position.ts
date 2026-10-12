export interface MenuBounds {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

export function contextMenuKeys(event: KeyboardEvent) {
  if (!['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(event.key)
    || event.target instanceof HTMLInputElement) return;
  const controls = [...(event.currentTarget as HTMLElement)
    .querySelectorAll<HTMLElement>('button:not(:disabled), input:not(:disabled)')];
  if (!controls.length) return;
  const current = controls.indexOf(document.activeElement as HTMLElement);
  const next = event.key === 'Home' ? 0 : event.key === 'End' ? controls.length - 1
    : (current + (event.key === 'ArrowDown' ? 1 : -1) + controls.length) % controls.length;
  event.preventDefault();
  controls[next]?.focus({ preventScroll: true });
}

export function fitContextMenu(
  point: { x: number; y: number }, size: { width: number; height: number }, bounds: MenuBounds,
) {
  return {
    left: `${Math.max(bounds.left, Math.min(point.x, bounds.right - size.width))}px`,
    top: `${Math.max(bounds.top, Math.min(point.y, bounds.bottom - size.height))}px`,
  };
}

// 先隐藏测量再显示，避免按预估高度跳位；菜单内部滚动时不重新定位。
export function measuredContextMenuStyle(
  menu: HTMLElement, point: { x: number; y: number }, bounds?: MenuBounds,
) {
  const area = bounds || { left: 8, top: 8,
    right: window.innerWidth - 8, bottom: window.innerHeight - 8 };
  menu.style.maxHeight = `${Math.max(0, area.bottom - area.top)}px`;
  const rect = menu.getBoundingClientRect();
  return { ...fitContextMenu(point, rect, area),
    maxHeight: menu.style.maxHeight, visibility: 'visible' };
}
