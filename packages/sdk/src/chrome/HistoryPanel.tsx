/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License").
 */

/**
 * History / Activity side panel — a self-contained, in-memory activity feed
 * of user-facing edits in the current session.
 *
 * Persistent version history stays HOST-OWNED. This panel subscribes to the
 * live `ICommandService.onCommandExecuted` stream, keeps a
 * bounded in-memory ring (last 100) of user-facing command executions, maps
 * each command id to a friendly label, and renders them newest-first.
 *
 * Ordering is derived from a monotonically incrementing sequence counter
 * (NOT wall-clock time) so nothing calls `Date.now()`/`new Date()` at module
 * load; a per-entry relative-age string is computed lazily at render time from
 * a capture timestamp taken inside the (post-mount) event handler.
 *
 * The app-only `useUI().toggleHistoryPanel` is replaced by the panel host's
 * `onClose`; the Univer facade + services are reached through the `api` prop's
 * `_injector`, matching StatusBar / FormulaBar precedent.
 */
import { useEffect, useMemo, useRef, useState, type CSSProperties } from 'react';
import { CommandType, ICommandService, type ICommandInfo } from '@univerjs/core';

import type { PanelComponentProps } from './extensions';
import { Icon } from './Icon';
import { PanelHeader } from './panel-shell';

const RING_CAP = 100;

interface ActivityEntry {
  /** Monotonic sequence — sole source of ordering (no wall-clock at module load). */
  seq: number;
  /** Command id that fired (e.g. `sheet.command.set-range-values`). */
  id: string;
  /** Friendly one-line label. */
  label: string;
  /** Icon glyph for the row. */
  icon: string;
  /** Capture time (ms). Taken inside the handler, post-mount — safe. */
  t: number;
}

/**
 * Map a user-facing COMMAND id to a friendly label + icon. Commands (not
 * mutations) are the user-intent layer, so one command == one activity row —
 * no de-duping a single paste into a dozen mutation rows. Anything not mapped
 * still records with a generic fallback label so the feed stays honest.
 */
const COMMAND_LABELS: Record<string, { label: string; icon: string }> = {
  'sheet.command.set-range-values': { label: '编辑了单元格', icon: 'edit' },
  'sheet.command.set-range-bold': { label: '切换了加粗', icon: 'format_bold' },
  'sheet.command.set-range-italic': { label: '切换了斜体', icon: 'format_italic' },
  'sheet.command.set-range-underline': { label: '切换了下划线', icon: 'format_underlined' },
  'sheet.command.set-range-stroke-through': {
    label: '切换了删除线',
    icon: 'strikethrough_s',
  },
  'sheet.command.set-range-font-family': { label: '更改了字体', icon: 'font_download' },
  'sheet.command.set-range-font-size': { label: '更改了字号', icon: 'format_size' },
  'sheet.command.set-range-text-color': { label: '更改了文字颜色', icon: 'format_color_text' },
  'sheet.command.set-background-color': { label: '更改了填充颜色', icon: 'format_color_fill' },
  'sheet.command.set-range-text-align': { label: '更改了对齐方式', icon: 'format_align_left' },
  'sheet.command.set-range-vertical-align': {
    label: '更改了对齐方式',
    icon: 'vertical_align_center',
  },
  'sheet.command.set-border-command': { label: '更改了边框', icon: 'border_all' },
  'sheet.command.clear-selection-content': { label: '清除了单元格内容', icon: 'backspace' },
  'sheet.command.clear-selection-format': { label: '清除了格式', icon: 'format_clear' },
  'sheet.command.clear-selection-all': { label: '清除了单元格', icon: 'clear_all' },
  'sheet.command.insert-row': { label: '插入了行', icon: 'add_row_above' },
  'sheet.command.insert-row-before': { label: '在上方插入了行', icon: 'add_row_above' },
  'sheet.command.insert-row-after': { label: '在下方插入了行', icon: 'add_row_below' },
  'sheet.command.insert-col': { label: '插入了列', icon: 'add_column_left' },
  'sheet.command.insert-col-before': { label: '在左侧插入了列', icon: 'add_column_left' },
  'sheet.command.insert-col-after': { label: '在右侧插入了列', icon: 'add_column_right' },
  'sheet.command.remove-row': { label: '删除了行', icon: 'delete' },
  'sheet.command.remove-col': { label: '删除了列', icon: 'delete' },
  'sheet.command.delete-range-move-up': { label: '删除了单元格', icon: 'delete' },
  'sheet.command.delete-range-move-left': { label: '删除了单元格', icon: 'delete' },
  'sheet.command.insert-range-move-down': { label: '插入了单元格', icon: 'add' },
  'sheet.command.insert-range-move-right': { label: '插入了单元格', icon: 'add' },
  'sheet.command.move-range': { label: '移动了单元格', icon: 'open_with' },
  'sheet.command.move-rows': { label: '移动了行', icon: 'swap_vert' },
  'sheet.command.move-cols': { label: '移动了列', icon: 'swap_horiz' },
  'sheet.command.add-worksheet-merge-all': { label: '合并了单元格', icon: 'cell_merge' },
  'sheet.command.remove-worksheet-merge': { label: '取消了单元格合并', icon: 'grid_on' },
  'sheet.command.set-worksheet-row-height': { label: '调整了行高', icon: 'height' },
  'sheet.command.set-worksheet-col-width': { label: '调整了列宽', icon: 'width_normal' },
  'sheet.command.set-row-hidden': { label: '隐藏了行', icon: 'visibility_off' },
  'sheet.command.set-specific-rows-visible': { label: '显示了行', icon: 'visibility' },
  'sheet.command.set-col-hidden': { label: '隐藏了列', icon: 'visibility_off' },
  'sheet.command.set-specific-cols-visible': { label: '显示了列', icon: 'visibility' },
  'sheet.command.set-frozen': { label: '更改了冻结窗格', icon: 'ac_unit' },
  'sheet.command.insert-sheet': { label: '添加了工作表', icon: 'add_box' },
  'sheet.command.remove-sheet': { label: '删除了工作表', icon: 'delete' },
  'sheet.command.set-worksheet-name': {
    label: '重命名了工作表',
    icon: 'drive_file_rename_outline',
  },
  'sheet.command.set-worksheet-order': { label: '调整了工作表顺序', icon: 'reorder' },
  'sheet.command.set-tab-color': { label: '更改了标签颜色', icon: 'palette' },
  'sheet.command.paste': { label: '粘贴', icon: 'content_paste' },
  'sheet.command.copy': { label: '复制', icon: 'content_copy' },
  'sheet.command.cut': { label: '剪切', icon: 'content_cut' },
  'sheet.command.add-table': { label: '插入了表格', icon: 'table' },
  'sheet.command.delete-table': { label: '删除了表格', icon: 'delete' },
  'sheet.command.numfmt.set.numfmt': { label: '更改了数字格式', icon: 'tag' },
  'univer.command.undo': { label: '撤销', icon: 'undo' },
  'univer.command.redo': { label: '重做', icon: 'redo' },
};

/**
 * Which command ids count as "user-facing activity". We only record
 * `CommandType.COMMAND` (intent) executions, and we skip pure selection /
 * navigation / scroll operations even though some of those are dispatched as
 * COMMANDs — they're noise, not edits. The allow-list above IS the primary
 * signal; unmapped COMMANDs still record with a fallback label unless they
 * match an obvious noise prefix.
 */
const NOISE_PREFIXES = [
  'sheet.operation.',
  'sheet.command.set-selection',
  'sheet.command.scroll',
  'sheet.command.set-zoom',
  'sheet.command.set-activate-cell-edit',
  'sheet.command.set-cell-edit-visible',
  'sheet.command.set-editor',
  'formula.',
  'univer.command.set-current-locale',
];

function isNoise(id: string): boolean {
  return NOISE_PREFIXES.some((p) => id.startsWith(p));
}

function labelFor(id: string): { label: string; icon: string } {
  const known = COMMAND_LABELS[id];
  if (known) return known;
  return { label: '更改', icon: 'bolt' };
}

function relativeTime(t: number, now: number): string {
  const delta = now - t;
  if (delta < 5_000) return '刚刚';
  if (delta < 60_000) return `${Math.floor(delta / 1000)} 秒前`;
  if (delta < 3_600_000) return `${Math.floor(delta / 60_000)} 分钟前`;
  return `${Math.floor(delta / 3_600_000)} 小时前`;
}

const headerStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 8,
  padding: '10px 12px',
  borderBottom: '1px solid var(--cs-chrome-border, #edeff3)',
};

export function HistoryPanel({ api, onClose }: PanelComponentProps) {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const univer = (api as any).univer;
  const [entries, setEntries] = useState<ActivityEntry[]>([]);
  // Ordering counter — monotonic, initialised at 0 (no wall-clock at load).
  const seqRef = useRef(0);
  // Re-render tick so relative times refresh even without new activity.
  const [nowTick, setNowTick] = useState(0);

  useEffect(() => {
    if (!univer) return;
    const injector = (univer as { _injector?: { get(t: unknown): unknown } })._injector;
    const cmd = injector?.get(ICommandService) as
      | {
          onCommandExecuted: (cb: (info: ICommandInfo, options?: unknown) => void) => {
            dispose: () => void;
          };
        }
      | undefined;
    if (!cmd) return;

    const sub = cmd.onCommandExecuted((info) => {
      // Only intent-level COMMANDs; mutations/operations would spam the feed.
      // `type` is optional on ICommandInfo but the dispatcher populates it —
      // treat an explicit non-COMMAND type as a reject, and let undefined
      // fall through to the id-based allow-list below.
      if (info.type !== undefined && info.type !== CommandType.COMMAND) return;
      if (isNoise(info.id)) return;
      // Ignore anything not on the allow-list unless it looks like a real
      // sheet command (keeps the feed to genuine edits, not internal churn).
      if (!COMMAND_LABELS[info.id] && !info.id.startsWith('sheet.command.')) return;

      const { label, icon } = labelFor(info.id);
      const seq = ++seqRef.current;
      const entry: ActivityEntry = { seq, id: info.id, label, icon, t: Date.now() };
      setEntries((prev) => {
        const next = prev.length >= RING_CAP ? prev.slice(prev.length - RING_CAP + 1) : prev;
        return [...next, entry];
      });
    });
    return () => sub.dispose();
  }, [univer]);

  // Refresh relative timestamps once a minute while the panel is open.
  useEffect(() => {
    const h = setInterval(() => setNowTick((n) => n + 1), 60_000);
    return () => clearInterval(h);
  }, []);

  // Newest-first. Ordering is purely by seq — the render-time `now` only
  // affects the human-readable age string, never the order.
  const sorted = useMemo(() => [...entries].sort((a, b) => b.seq - a.seq), [entries]);
  const now = useMemo(() => Date.now(), [nowTick, entries.length]);

  const empty = sorted.length === 0;

  return (
    <div
      data-testid="cs-history-panel"
      style={{ height: '100%', display: 'flex', flexDirection: 'column' }}
    >
      <PanelHeader icon="history" title="活动记录" count={sorted.length} onClose={onClose} />

      <div style={{ flex: 1, overflow: 'auto', padding: 12 }}>
        {empty ? (
          <div
            data-testid="cs-history-panel-empty"
            style={{ textAlign: 'center', opacity: 0.75, padding: '24px 8px' }}
          >
            <Icon name="history" size={40} style={{ opacity: 0.4 }} />
            <div style={{ fontWeight: 600, marginTop: 8 }}>暂无活动记录</div>
            <div style={{ fontSize: 13, marginTop: 4 }}>
              您在本次会话中进行的编辑会显示在这里。
            </div>
          </div>
        ) : (
          <ol
            role="list"
            style={{
              listStyle: 'none',
              margin: 0,
              padding: 0,
              display: 'flex',
              flexDirection: 'column',
              gap: 2,
            }}
          >
            {sorted.map((e) => (
              <li
                key={e.seq}
                data-testid="cs-history-row"
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: 10,
                  padding: '8px 8px',
                  borderRadius: 8,
                }}
              >
                <span
                  aria-hidden
                  style={{
                    display: 'inline-flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    width: 28,
                    height: 28,
                    flex: '0 0 auto',
                    borderRadius: 6,
                    background: 'var(--cs-chrome-subtle-bg, #f4f5f8)',
                    opacity: 0.9,
                  }}
                >
                  <Icon name={e.icon} size={16} />
                </span>
                <span style={{ flex: 1, minWidth: 0, fontSize: 13, fontWeight: 500 }}>
                  {e.label}
                </span>
                <time
                  dateTime={new Date(e.t).toISOString()}
                  style={{ opacity: 0.6, fontSize: 12, whiteSpace: 'nowrap' }}
                >
                  {relativeTime(e.t, now)}
                </time>
              </li>
            ))}
          </ol>
        )}
      </div>
    </div>
  );
}
