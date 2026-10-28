/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * MenuBar — the built-in dropdown menu row for `<CasualSheets chrome>`.
 *
 * Office / Google-Sheets-style horizontal menu strip (no logo, no title — the
 * host frames the editor with its own bar): File · Edit · View · Insert ·
 * Format · Data · Help. Each top-level button opens a dropdown of items that
 * either dispatch a verified Univer command / facade call through the
 * `CasualSheetsAPI` (no app context), or — for actions that need a dialog the
 * SDK doesn't ship — route through the optional `onDialogRequest` host hook.
 *
 * This mirrors `apps/web/src/shell/MenuBar.tsx`, but reimplements every action
 * INLINE against `api` / `api.univer` (the FUniver facade) + `api.executeCommand`.
 * The command ids + facade logic are copied verbatim from
 * `apps/web/src/shell/home-tab-actions.ts` + `tab-actions.ts` so behaviour
 * matches the real app exactly.
 *
 * Feature flags: pass `features` to hide a control or whole menu group when its
 * feature is disabled. Defaults to all-enabled. A control whose feature is
 * `false` does not render. An entire top-level menu that ends up with no
 * runnable items is dropped.
 *
 * Dialog routing: items marked `dialog: '<kind>'` (Format Cells, Insert Chart,
 * PivotTable, Find & Replace, Insert/Delete cells, Sparkline, Name Manager,
 * Goal Seek, Data Validation, Conditional Formatting, …) call
 * `onDialogRequest(kind, context)` so the host renders its own UI. When
 * `onDialogRequest` is not provided, those items are omitted (the SDK never
 * fakes a dialog).
 *
 * Self-contained: only one menu is open at a time; Escape and an outside
 * pointerdown close it.
 */

import { useEffect, useRef, useState, type CSSProperties, type ReactNode } from 'react';
import { BorderStyleTypes, BorderType } from '@univerjs/core';
import { openExternal } from './openExternal';
import type { FUniver } from '@univerjs/core/facade';
import type { CasualSheetsAPI } from '../sheets/api';
import { ensurePluginByName } from '../univer';
import { Icon } from './Icon';
import { ensureChromeFonts } from './fonts';
import { useDialogs } from './dialog-context';
import type { ChromeExtensions } from './extensions';
import {
  computeVisibleMenus,
  type MenuDef,
  type MenuId,
  type MenuItemDef,
  type RunFn,
} from './menu-model';

export type { MenuDialogKind } from './menu-model';

/**
 * The menu data model + pure gating engine live in `./menu-model` (univer-free,
 * so unit-testable). `RunFn` / `MenuItemDef` / `MenuDef` / `MenuId` /
 * `MenuDialogKind` are re-imported above; the `run` handlers that populate
 * `MENUS` are defined inline below against the `@univerjs/core` facade.
 */

/**
 * Pretty-print a `Ctrl+Shift+X` shortcut for the current platform. The SDK
 * doesn't import the app's `formatShortcut`, so this is a compact inline
 * equivalent: on macOS, swap Ctrl→⌘, Alt→⌥, Shift→⇧ and drop the `+`.
 */
function fmtShortcut(shortcut: string): string {
  const isMac =
    typeof navigator !== 'undefined' && /Mac|iPhone|iPad/i.test(navigator.platform || '');
  if (!isMac) return shortcut;
  return shortcut
    .split('+')
    .map((part) => {
      switch (part) {
        case 'Ctrl':
          return '⌘';
        case 'Alt':
          return '⌥';
        case 'Shift':
          return '⇧';
        default:
          return part;
      }
    })
    .join('');
}

/* ───────────────────────── inline action helpers ─────────────────────────
 * Every helper below is a faithful port of the corresponding function in
 * apps/web/src/shell/home-tab-actions.ts + tab-actions.ts, reimplemented over
 * the SDK's `api` (CasualSheetsAPI) / `api.univer` (FUniver) so the chrome
 * never imports app context. Command ids are identical to the app's.
 */

function fu(api: CasualSheetsAPI): FUniver {
  return api.univer;
}

function activeRange(api: CasualSheetsAPI) {
  return fu(api).getActiveWorkbook()?.getActiveSheet()?.getActiveRange() ?? null;
}

function activeSheet(api: CasualSheetsAPI) {
  return fu(api).getActiveWorkbook()?.getActiveSheet() ?? null;
}

/** Selection as A1 (e.g. `B2:D10`), or null. Used to seed dialog context. */
function selectionA1(api: CasualSheetsAPI): string | null {
  const range = activeRange(api);
  if (!range) return null;
  try {
    return range.getA1Notation();
  } catch {
    return null;
  }
}

/* ── File ─────────────────────────────────────────────────────────────── */

const importXlsx: RunFn = (api) => {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = '.xlsx,.xlsm';
  input.style.display = 'none';
  const cleanup = () => {
    input.onchange = null;
    input.oncancel = null;
    input.remove();
  };
  input.onchange = () => {
    void (async () => {
      try {
        const file = input.files?.[0];
        if (!file) return;
        await api.importXlsx(file);
      } catch (err) {
        window.alert(`导入 XLSX 失败：${err instanceof Error ? err.message : '请检查文件后重试。'}`);
      } finally {
        cleanup();
      }
    })();
  };
  input.oncancel = cleanup;
  document.body.appendChild(input);
  input.click();
};

const exportXlsx: RunFn = (api) => {
  void (async () => {
    let url: string | null = null;
    let link: HTMLAnchorElement | null = null;
    try {
      const blob = await api.exportXlsx();
      const workbookName = fu(api).getActiveWorkbook()?.getName().trim() || '工作簿';
      url = URL.createObjectURL(blob);
      link = document.createElement('a');
      link.href = url;
      link.download = `${workbookName.replace(/\.(xlsx|xlsm)$/i, '')}.xlsx`;
      link.style.display = 'none';
      document.body.appendChild(link);
      link.click();
    } catch (err) {
      window.alert(`导出 XLSX 失败：${err instanceof Error ? err.message : '请稍后重试。'}`);
    } finally {
      link?.remove();
      if (url) URL.revokeObjectURL(url);
    }
  })();
};

/* ── Edit ─────────────────────────────────────────────────────────────── */

const undo: RunFn = (api) => void api.executeCommand('univer.command.undo');
const redo: RunFn = (api) => void api.executeCommand('univer.command.redo');
const cut: RunFn = (api) => void api.executeCommand('univer.command.cut');
const copy: RunFn = (api) => void api.executeCommand('univer.command.copy');
const paste: RunFn = (api) => void api.executeCommand('univer.command.paste');
const pasteFormattingOnly: RunFn = (api) => void api.executeCommand('sheet.command.paste-format');

/* ── Insert (facade-driven; mirror tab-actions.ts) ─────────────────────── */

const insertRowAbove: RunFn = (api) => {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  if (!range || !sheet) return;
  sheet.insertRowBefore(range.getRow());
};

const insertRowBelow: RunFn = (api) => {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  if (!range || !sheet) return;
  sheet.insertRowAfter(range.getRow() + range.getHeight() - 1);
};

const insertColumnLeft: RunFn = (api) => {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  if (!range || !sheet) return;
  sheet.insertColumnBefore(range.getColumn());
};

const insertColumnRight: RunFn = (api) => {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  if (!range || !sheet) return;
  sheet.insertColumnAfter(range.getColumn() + range.getWidth() - 1);
};

const deleteSelectedRow: RunFn = (api) => {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  if (!range || !sheet) return;
  sheet.deleteRows(range.getRow(), range.getHeight());
};

const deleteSelectedColumn: RunFn = (api) => {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  if (!range || !sheet) return;
  sheet.deleteColumns(range.getColumn(), range.getWidth());
};

const insertNewSheet: RunFn = (api) => {
  fu(api).getActiveWorkbook()?.insertSheet();
};

const insertImage: RunFn = (api) => void api.executeCommand('sheet.command.insert-float-image');

const insertHyperlink: RunFn = (api) =>
  void api.executeCommand('sheet.operation.insert-hyper-link');

const insertComment: RunFn = (api) => void api.executeCommand('sheet.operation.show-comment-modal');

const insertTodayDate: RunFn = (api) => {
  const range = activeRange(api);
  if (!range) return;
  const today = new Date();
  const pad = (n: number) => n.toString().padStart(2, '0');
  const v = `${today.getFullYear()}-${pad(today.getMonth() + 1)}-${pad(today.getDate())}`;
  range.setValue({ v });
};

const insertCurrentTime: RunFn = (api) => {
  const range = activeRange(api);
  if (!range) return;
  const now = new Date();
  const pad = (n: number) => n.toString().padStart(2, '0');
  const v = `${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
  range.setValue({ v });
};

const insertTable: RunFn = (api) => {
  void (async () => {
    const range = activeRange(api);
    const sheet = activeSheet(api);
    const wb = fu(api).getActiveWorkbook();
    if (!range || !sheet || !wb) return;
    await ensurePluginByName('table');
    await api.executeCommand('sheet.command.add-table', {
      unitId: wb.getId(),
      subUnitId: sheet.getSheetId(),
      range: {
        startRow: range.getRow(),
        startColumn: range.getColumn(),
        endRow: range.getRow() + range.getHeight() - 1,
        endColumn: range.getColumn() + range.getWidth() - 1,
      },
    });
  })();
};

const autoFitColumns: RunFn = (api) => {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  if (!range || !sheet) return;
  const withAutoWidth = sheet as unknown as {
    setColumnAutoWidth?: (col: number, n: number) => unknown;
  };
  withAutoWidth.setColumnAutoWidth?.(range.getColumn(), range.getWidth());
};

const AUTO_FIT_ROW_CAP = 500;
const autoFitRows: RunFn = (api) => {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  if (!range || !sheet) return;
  const start = range.getRow();
  const count = Math.min(range.getHeight(), AUTO_FIT_ROW_CAP);
  for (let r = 0; r < count; r++) sheet.autoFitRow(start + r);
};

/* ── Format ───────────────────────────────────────────────────────────── */

// Patterns copied from home-tab-actions.ts NUMBER_FORMAT_PATTERNS.
const NUMBER_FORMAT_PATTERNS = {
  general: '',
  number: '#,##0.00',
  integer: '#,##0',
  currency: '"$"#,##0.00',
  accounting: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
  percent: '0.00%',
  date: 'yyyy-mm-dd',
  time: 'hh:mm:ss',
  scientific: '0.00E+00',
  text: '@',
} as const;
type NumberFormatKey = keyof typeof NUMBER_FORMAT_PATTERNS;

const NUM_FORMAT_ORDER: NumberFormatKey[] = [
  'general',
  'number',
  'integer',
  'currency',
  'accounting',
  'percent',
  'date',
  'time',
  'scientific',
  'text',
];

const NUM_FORMAT_SHORTCUT: Partial<Record<NumberFormatKey, string>> = {
  number: 'Ctrl+Shift+1',
  time: 'Ctrl+Shift+2',
  date: 'Ctrl+Shift+3',
  currency: 'Ctrl+Shift+4',
  percent: 'Ctrl+Shift+5',
  scientific: 'Ctrl+Shift+6',
};

function setNumberFormatByKey(api: CasualSheetsAPI, key: NumberFormatKey) {
  // setNumberFormat lives on the sheets-numfmt facade extension (runtime cast).
  const range = activeRange(api) as unknown as { setNumberFormat?: (p: string) => unknown } | null;
  range?.setNumberFormat?.(NUMBER_FORMAT_PATTERNS[key]);
}

const increaseDecimal: RunFn = (api) =>
  void api.executeCommand('sheet.command.numfmt.add.decimal.command');
const decreaseDecimal: RunFn = (api) =>
  void api.executeCommand('sheet.command.numfmt.subtract.decimal.command');
const clearFormat: RunFn = (api) => void api.executeCommand('sheet.command.clear-selection-format');

function applyBorders(api: CasualSheetsAPI, choice: 'all' | 'outside' | 'none') {
  const range = activeRange(api);
  if (!range) return;
  const type =
    choice === 'all' ? BorderType.ALL : choice === 'outside' ? BorderType.OUTSIDE : BorderType.NONE;
  const style = choice === 'none' ? BorderStyleTypes.NONE : BorderStyleTypes.THIN;
  range.setBorder(type, style, '#000000');
}

/* ── Format → Visibility (command-driven; mirror tab-actions.ts) ───────── */

function rowSpan(api: CasualSheetsAPI) {
  const wb = fu(api).getActiveWorkbook();
  const sheet = activeSheet(api);
  const range = activeRange(api);
  if (!wb || !sheet || !range) return null;
  const startRow = range.getRow();
  const endRow = startRow + range.getHeight() - 1;
  const maxCol = (sheet as unknown as { getMaxColumns?: () => number }).getMaxColumns?.() ?? 1;
  return { wb, sheet, startRow, endRow, maxCol };
}

function colSpan(api: CasualSheetsAPI) {
  const wb = fu(api).getActiveWorkbook();
  const sheet = activeSheet(api);
  const range = activeRange(api);
  if (!wb || !sheet || !range) return null;
  const startColumn = range.getColumn();
  const endColumn = startColumn + range.getWidth() - 1;
  const maxRow = (sheet as unknown as { getMaxRows?: () => number }).getMaxRows?.() ?? 1;
  return { wb, sheet, startColumn, endColumn, maxRow };
}

const hideSelectedRows: RunFn = (api) => {
  const s = rowSpan(api);
  if (!s) return;
  void api.executeCommand('sheet.command.set-rows-hidden', {
    unitId: s.wb.getId(),
    subUnitId: s.sheet.getSheetId(),
    ranges: [
      {
        startRow: s.startRow,
        endRow: s.endRow,
        startColumn: 0,
        endColumn: Math.max(0, s.maxCol - 1),
        rangeType: 1,
      },
    ],
  });
};

const unhideSelectedRows: RunFn = (api) => {
  const s = rowSpan(api);
  if (!s) return;
  void api.executeCommand('sheet.command.set-specific-rows-visible', {
    unitId: s.wb.getId(),
    subUnitId: s.sheet.getSheetId(),
    ranges: [
      {
        startRow: s.startRow,
        endRow: s.endRow,
        startColumn: 0,
        endColumn: Math.max(0, s.maxCol - 1),
        rangeType: 1,
      },
    ],
  });
};

const hideSelectedColumns: RunFn = (api) => {
  const s = colSpan(api);
  if (!s) return;
  void api.executeCommand('sheet.command.set-col-hidden', {
    unitId: s.wb.getId(),
    subUnitId: s.sheet.getSheetId(),
    ranges: [
      {
        startRow: 0,
        endRow: Math.max(0, s.maxRow - 1),
        startColumn: s.startColumn,
        endColumn: s.endColumn,
        rangeType: 2,
      },
    ],
  });
};

const unhideSelectedColumns: RunFn = (api) => {
  const s = colSpan(api);
  if (!s) return;
  // Univer's "show specific cols" id is asymmetric with the row variant.
  void api.executeCommand('sheet.command.set-col-visible-on-cols', {
    unitId: s.wb.getId(),
    subUnitId: s.sheet.getSheetId(),
    ranges: [
      {
        startRow: 0,
        endRow: Math.max(0, s.maxRow - 1),
        startColumn: s.startColumn,
        endColumn: s.endColumn,
        rangeType: 2,
      },
    ],
  });
};

/* ── View ─────────────────────────────────────────────────────────────── */

type FreezeCapableSheet = {
  setFrozenRows: (n: number) => unknown;
  setFrozenColumns: (n: number) => unknown;
};

const freezeFirstRow: RunFn = (api) => {
  (activeSheet(api) as unknown as FreezeCapableSheet | null)?.setFrozenRows(1);
};
const freezeFirstColumn: RunFn = (api) => {
  (activeSheet(api) as unknown as FreezeCapableSheet | null)?.setFrozenColumns(1);
};
const freezeAtSelection: RunFn = (api) =>
  void api.executeCommand('sheet.command.set-selection-frozen');
const unfreezePanes: RunFn = (api) => void api.executeCommand('sheet.command.cancel-frozen');

const toggleGridlines: RunFn = (api) => {
  const wb = fu(api).getActiveWorkbook();
  const sheet = activeSheet(api);
  if (!wb || !sheet) return;
  // BooleanNumber: 0 = hide, 1 = show. The app reads current state; without a
  // reactive UI store here we hide (matches the app's default click path).
  void api.executeCommand('sheet.command.toggle-gridlines', {
    unitId: wb.getId(),
    subUnitId: sheet.getSheetId(),
    showGridlines: 0,
  });
};

const toggleCommentPanel: RunFn = (api) =>
  void api.executeCommand('sheet.operation.toggle-comment-panel');

const jumpToFirstCell: RunFn = (api) => {
  activeSheet(api)?.getRange(0, 0).activate();
};

const switchToPreviousSheet: RunFn = (api) => switchSheetByDelta(api, -1);
const switchToNextSheet: RunFn = (api) => switchSheetByDelta(api, +1);

function switchSheetByDelta(api: CasualSheetsAPI, delta: -1 | 1) {
  const wb = fu(api).getActiveWorkbook();
  const active = wb?.getActiveSheet();
  if (!wb || !active) return;
  const sheets = wb.getSheets();
  const activeId = active.getSheetId();
  const idx = sheets.findIndex((s) => s.getSheetId() === activeId);
  if (idx < 0) return;
  const nextIdx = idx + delta;
  if (nextIdx < 0 || nextIdx >= sheets.length) return;
  wb.setActiveSheet(sheets[nextIdx]);
}

const showFormulas: RunFn = (api) => void api.executeCommand('sheet.command.set-show-formula', {});

/* ── Data ─────────────────────────────────────────────────────────────── */

const sortAsc: RunFn = (api) => sortRange(api, true);
const sortDesc: RunFn = (api) => sortRange(api, false);

function sortRange(api: CasualSheetsAPI, ascending: boolean) {
  const range = activeRange(api);
  if (!range) return;
  // sort() comes from the sheets-sort facade extension (runtime cast).
  const withSort = range as unknown as {
    sort?: (spec: { column: number; ascending: boolean }) => unknown;
  };
  withSort.sort?.({ column: range.getColumn(), ascending });
}

const toggleFilter: RunFn = (api) => {
  void (async () => {
    await ensurePluginByName('filter');
    const wb = fu(api).getActiveWorkbook();
    const sheet = activeSheet(api);
    const range = activeRange(api);
    if (!wb || !sheet || !range) return;
    const sheetWithFilter = sheet as unknown as { getFilter?: () => unknown };
    if (sheetWithFilter.getFilter?.()) {
      await api.executeCommand('sheet.command.remove-sheet-filter', {
        unitId: wb.getId(),
        subUnitId: sheet.getSheetId(),
      });
      return;
    }
    await api.executeCommand('sheet.command.set-filter-range', {
      unitId: wb.getId(),
      subUnitId: sheet.getSheetId(),
      range: {
        startRow: range.getRow(),
        startColumn: range.getColumn(),
        endRow: range.getRow() + range.getHeight() - 1,
        endColumn: range.getColumn() + range.getWidth() - 1,
      },
    });
  })();
};

const splitTextToColumns: RunFn = (api) =>
  void api.executeCommand('sheet.command.split-text-to-columns');

const forceRecalculate: RunFn = (api) =>
  void api.executeCommand('formula.mutation.set-formula-calculation-start', {
    forceCalculation: true,
  });

/* ───────────────────────────── menu structure ─────────────────────────── */

const NUMBER_FORMAT_LABELS: Record<NumberFormatKey, string> = {
  general: '常规',
  number: '数值',
  integer: '整数',
  currency: '货币',
  accounting: '会计专用',
  percent: '百分比',
  date: '日期',
  time: '时间',
  scientific: '科学记数',
  text: '文本',
};

// Mirrors the app's menu order. Dialog items resolve `dialog` against
// `onDialogRequest`; non-dialog items dispatch via the inline helpers above.
const MENUS: MenuDef[] = [
  {
    id: 'file',
    label: '文件',
    feature: 'file',
    items: [
      {
        kind: 'item',
        id: 'import-xlsx',
        label: '导入 XLSX…',
        icon: 'upload_file',
        run: importXlsx,
      },
      {
        kind: 'item',
        id: 'export-xlsx',
        label: '导出 XLSX',
        icon: 'download',
        run: exportXlsx,
      },
      { kind: 'separator', id: 'sep-file' },
      { kind: 'item', id: 'properties', label: '属性…', icon: 'info', dialog: 'properties' },
      {
        kind: 'item',
        id: 'about',
        label: '关于 Casual Sheets',
        icon: 'help_outline',
        dialog: 'about',
        feature: 'branding',
      },
    ],
  },
  {
    id: 'edit',
    label: '编辑',
    items: [
      { kind: 'item', id: 'undo', label: '撤销', icon: 'undo', shortcut: 'Ctrl+Z', run: undo },
      { kind: 'item', id: 'redo', label: '重做', icon: 'redo', shortcut: 'Ctrl+Y', run: redo },
      { kind: 'separator', id: 'sep-clip' },
      { kind: 'item', id: 'cut', label: '剪切', icon: 'content_cut', shortcut: 'Ctrl+X', run: cut },
      {
        kind: 'item',
        id: 'copy',
        label: '复制',
        icon: 'content_copy',
        shortcut: 'Ctrl+C',
        run: copy,
      },
      {
        kind: 'item',
        id: 'paste',
        label: '粘贴',
        icon: 'content_paste',
        shortcut: 'Ctrl+V',
        run: paste,
      },
      {
        kind: 'item',
        id: 'paste-format',
        label: '仅粘贴格式',
        icon: 'content_paste',
        shortcut: 'Ctrl+Shift+V',
        run: pasteFormattingOnly,
      },
      {
        kind: 'item',
        id: 'paste-special',
        label: '选择性粘贴…',
        icon: 'content_paste_go',
        shortcut: 'Ctrl+Alt+V',
        dialog: 'paste-special',
      },
      { kind: 'separator', id: 'sep-find' },
      {
        kind: 'item',
        id: 'find-replace',
        label: '查找和替换…',
        icon: 'search',
        shortcut: 'Ctrl+F',
        dialog: 'find-replace',
      },
      { kind: 'separator', id: 'sep-cells' },
      {
        kind: 'item',
        id: 'edit-insert-cells',
        label: '插入单元格…',
        icon: 'add_box',
        shortcut: 'Ctrl++',
        dialog: 'insert-cells',
      },
      {
        kind: 'item',
        id: 'edit-delete-cells',
        label: '删除单元格…',
        icon: 'indeterminate_check_box',
        shortcut: 'Ctrl+-',
        dialog: 'delete-cells',
      },
    ],
  },
  {
    id: 'view',
    label: '视图',
    items: [
      {
        kind: 'item',
        id: 'show-formulas',
        label: '显示公式',
        icon: 'description',
        shortcut: 'Ctrl+`',
        run: showFormulas,
      },
      {
        kind: 'item',
        id: 'toggle-gridlines',
        label: '切换网格线',
        icon: 'grid_on',
        run: toggleGridlines,
      },
      { kind: 'separator', id: 'sep-freeze' },
      {
        kind: 'item',
        id: 'freeze-row',
        label: '冻结首行',
        icon: 'border_horizontal',
        run: freezeFirstRow,
      },
      {
        kind: 'item',
        id: 'freeze-col',
        label: '冻结首列',
        icon: 'border_vertical',
        run: freezeFirstColumn,
      },
      {
        kind: 'item',
        id: 'freeze-selection',
        label: '冻结窗格（所选位置）',
        icon: 'grid_4x4',
        run: freezeAtSelection,
      },
      { kind: 'item', id: 'unfreeze', label: '取消冻结', icon: 'grid_off', run: unfreezePanes },
      { kind: 'separator', id: 'sep-nav' },
      {
        kind: 'item',
        id: 'jump-home',
        label: '跳转到 A1',
        icon: 'home',
        shortcut: 'Ctrl+Home',
        run: jumpToFirstCell,
      },
      {
        kind: 'item',
        id: 'prev-sheet',
        label: '上一工作表',
        icon: 'navigate_before',
        shortcut: 'Ctrl+PageUp',
        run: switchToPreviousSheet,
      },
      {
        kind: 'item',
        id: 'next-sheet',
        label: '下一工作表',
        icon: 'navigate_next',
        shortcut: 'Ctrl+PageDown',
        run: switchToNextSheet,
      },
      { kind: 'separator', id: 'sep-panels' },
      {
        kind: 'item',
        id: 'comments-panel',
        label: '批注面板',
        icon: 'forum',
        run: toggleCommentPanel,
      },
    ],
  },
  {
    id: 'insert',
    label: '插入',
    items: [
      {
        kind: 'item',
        id: 'new-sheet',
        label: '新建工作表',
        icon: 'add_box',
        shortcut: 'Shift+F11',
        run: insertNewSheet,
      },
      {
        kind: 'item',
        id: 'insert-table',
        label: '表格',
        icon: 'table_rows',
        shortcut: 'Ctrl+L',
        run: insertTable,
        feature: 'tables',
      },
      {
        kind: 'item',
        id: 'insert-chart',
        label: '图表…',
        icon: 'bar_chart',
        dialog: 'insert-chart',
        feature: 'charts',
      },
      {
        kind: 'item',
        id: 'insert-sparkline',
        label: '迷你图…',
        icon: 'show_chart',
        dialog: 'insert-sparkline',
        feature: 'sparklines',
      },
      {
        kind: 'item',
        id: 'insert-pivot',
        label: '数据透视表…',
        icon: 'pivot_table_chart',
        dialog: 'insert-pivot',
        feature: 'pivots',
      },
      { kind: 'separator', id: 'sep-objects' },
      { kind: 'item', id: 'insert-image', label: '图片…', icon: 'image', run: insertImage },
      {
        kind: 'item',
        id: 'insert-function',
        label: '函数…',
        icon: 'functions',
        shortcut: 'Shift+F3',
        dialog: 'insert-function',
      },
      {
        kind: 'item',
        id: 'insert-link',
        label: '超链接…',
        icon: 'link',
        shortcut: 'Ctrl+K',
        run: insertHyperlink,
      },
      {
        kind: 'item',
        id: 'insert-comment',
        label: '批注',
        icon: 'comment',
        shortcut: 'Shift+F2',
        run: insertComment,
      },
      { kind: 'separator', id: 'sep-rowcol' },
      {
        kind: 'submenu',
        id: 'insert-rowcol',
        label: '行和列',
        icon: 'grid_on',
        items: [
          {
            kind: 'item',
            id: 'insert-row-above',
            label: '在上方插入行',
            icon: 'vertical_align_top',
            run: insertRowAbove,
          },
          {
            kind: 'item',
            id: 'insert-row-below',
            label: '在下方插入行',
            icon: 'vertical_align_bottom',
            run: insertRowBelow,
          },
          {
            kind: 'item',
            id: 'insert-col-left',
            label: '在左侧插入列',
            icon: 'keyboard_tab_rtl',
            run: insertColumnLeft,
          },
          {
            kind: 'item',
            id: 'insert-col-right',
            label: '在右侧插入列',
            icon: 'keyboard_tab',
            run: insertColumnRight,
          },
        ],
      },
      { kind: 'separator', id: 'sep-autofit' },
      {
        kind: 'item',
        id: 'autofit-col',
        label: '自动调整列宽',
        icon: 'settings_ethernet',
        run: autoFitColumns,
      },
      {
        kind: 'item',
        id: 'autofit-row',
        label: '自动调整行高',
        icon: 'height',
        run: autoFitRows,
      },
      { kind: 'separator', id: 'sep-date' },
      {
        kind: 'item',
        id: 'insert-today',
        label: '当前日期',
        icon: 'today',
        shortcut: 'Ctrl+;',
        run: insertTodayDate,
      },
      {
        kind: 'item',
        id: 'insert-time',
        label: '当前时间',
        icon: 'schedule',
        shortcut: 'Ctrl+Shift+:',
        run: insertCurrentTime,
      },
    ],
  },
  {
    id: 'format',
    label: '格式',
    items: [
      {
        kind: 'item',
        id: 'format-cells',
        label: '设置单元格格式…',
        icon: 'format_shapes',
        shortcut: 'Ctrl+1',
        dialog: 'format-cells',
      },
      { kind: 'separator', id: 'sep-format-cells' },
      {
        kind: 'item',
        id: 'bold',
        label: '加粗',
        icon: 'format_bold',
        shortcut: 'Ctrl+B',
        run: (api) => void api.executeCommand('sheet.command.set-range-bold'),
      },
      {
        kind: 'item',
        id: 'italic',
        label: '斜体',
        icon: 'format_italic',
        shortcut: 'Ctrl+I',
        run: (api) => void api.executeCommand('sheet.command.set-range-italic'),
      },
      {
        kind: 'item',
        id: 'underline',
        label: '下划线',
        icon: 'format_underlined',
        shortcut: 'Ctrl+U',
        run: (api) => void api.executeCommand('sheet.command.set-range-underline'),
      },
      {
        kind: 'item',
        id: 'wrap-text',
        label: '自动换行',
        icon: 'wrap_text',
        run: (api) => void api.executeCommand('sheet.command.set-text-wrap', { value: 3 }),
      },
      { kind: 'separator', id: 'sep-numfmt' },
      {
        kind: 'submenu',
        id: 'num-format',
        label: '数字格式',
        icon: 'looks_one',
        items: NUM_FORMAT_ORDER.map<MenuItemDef>((k) => ({
          kind: 'item',
          id: `num-${k}`,
          label: NUMBER_FORMAT_LABELS[k],
          icon: 'looks_one',
          shortcut: NUM_FORMAT_SHORTCUT[k],
          run: (api) => setNumberFormatByKey(api, k),
        })),
      },
      {
        kind: 'item',
        id: 'decimal-up',
        label: '增加小数位数',
        icon: 'decimal_increase',
        run: increaseDecimal,
      },
      {
        kind: 'item',
        id: 'decimal-down',
        label: '减少小数位数',
        icon: 'decimal_decrease',
        run: decreaseDecimal,
      },
      { kind: 'separator', id: 'sep-borders' },
      {
        kind: 'submenu',
        id: 'borders',
        label: '边框',
        icon: 'border_all',
        items: [
          {
            kind: 'item',
            id: 'border-all',
            label: '所有框线',
            icon: 'border_all',
            run: (api) => applyBorders(api, 'all'),
          },
          {
            kind: 'item',
            id: 'border-outside',
            label: '外侧框线',
            icon: 'border_outer',
            run: (api) => applyBorders(api, 'outside'),
          },
          {
            kind: 'item',
            id: 'border-none',
            label: '无框线',
            icon: 'border_clear',
            run: (api) => applyBorders(api, 'none'),
          },
        ],
      },
      { kind: 'separator', id: 'sep-cond' },
      {
        kind: 'item',
        id: 'conditional-formatting',
        label: '条件格式…',
        icon: 'palette',
        dialog: 'conditional-formatting',
        feature: 'conditionalFormatting',
      },
      { kind: 'separator', id: 'sep-visibility' },
      {
        kind: 'submenu',
        id: 'visibility',
        label: '显示与隐藏',
        icon: 'visibility',
        items: [
          {
            kind: 'item',
            id: 'hide-row',
            label: '隐藏行',
            icon: 'visibility_off',
            shortcut: 'Ctrl+9',
            run: hideSelectedRows,
          },
          {
            kind: 'item',
            id: 'unhide-row',
            label: '取消隐藏行',
            icon: 'visibility',
            shortcut: 'Ctrl+Shift+9',
            run: unhideSelectedRows,
          },
          {
            kind: 'item',
            id: 'hide-col',
            label: '隐藏列',
            icon: 'visibility_off',
            shortcut: 'Ctrl+0',
            run: hideSelectedColumns,
          },
          {
            kind: 'item',
            id: 'unhide-col',
            label: '取消隐藏列',
            icon: 'visibility',
            shortcut: 'Ctrl+Shift+0',
            run: unhideSelectedColumns,
          },
        ],
      },
      { kind: 'separator', id: 'sep-clear' },
      {
        kind: 'item',
        id: 'clear-format',
        label: '清除格式',
        icon: 'format_clear',
        run: clearFormat,
      },
      { kind: 'separator', id: 'sep-delete' },
      {
        kind: 'item',
        id: 'delete-row',
        label: '删除行',
        icon: 'delete_sweep',
        run: deleteSelectedRow,
      },
      {
        kind: 'item',
        id: 'delete-col',
        label: '删除列',
        icon: 'folder_delete',
        run: deleteSelectedColumn,
      },
    ],
  },
  {
    id: 'data',
    label: '数据',
    items: [
      {
        kind: 'item',
        id: 'sort-asc',
        label: '升序排序',
        icon: 'arrow_upward',
        run: sortAsc,
      },
      {
        kind: 'item',
        id: 'sort-desc',
        label: '降序排序',
        icon: 'arrow_downward',
        run: sortDesc,
      },
      {
        kind: 'item',
        id: 'sort-custom',
        label: '自定义排序…',
        icon: 'sort',
        dialog: 'custom-sort',
      },
      {
        kind: 'item',
        id: 'toggle-filter',
        label: '切换筛选',
        icon: 'filter_alt',
        shortcut: 'Ctrl+Shift+L',
        run: toggleFilter,
        feature: 'filter',
      },
      { kind: 'separator', id: 'sep-tools' },
      {
        kind: 'item',
        id: 'data-validation',
        label: '数据验证…',
        icon: 'rule',
        dialog: 'data-validation',
        feature: 'dataValidation',
      },
      {
        kind: 'item',
        id: 'name-manager',
        label: '名称管理器…',
        icon: 'bookmark_add',
        shortcut: 'Ctrl+F3',
        dialog: 'name-manager',
      },
      {
        kind: 'item',
        id: 'goal-seek',
        label: '单变量求解…',
        icon: 'analytics',
        dialog: 'goal-seek',
      },
      { kind: 'separator', id: 'sep-clean' },
      {
        kind: 'item',
        id: 'text-to-columns',
        label: '分列',
        icon: 'splitscreen',
        run: splitTextToColumns,
      },
      {
        kind: 'item',
        id: 'recalculate',
        label: '重新计算',
        icon: 'autorenew',
        shortcut: 'F9',
        run: forceRecalculate,
      },
    ],
  },
  {
    id: 'help',
    label: '帮助',
    // Whole-menu gate: an embedded host passes `features={{ help: false }}` to
    // drop the Help menu entirely (no editor-branded surface). Individual
    // branding links below also carry `feature: 'branding'` so a host that
    // keeps Help (for keyboard shortcuts) can still hide the About / GitHub
    // links with `features={{ branding: false }}`.
    feature: 'help',
    items: [
      {
        kind: 'item',
        id: 'keyboard-shortcuts',
        label: '键盘快捷键',
        icon: 'info',
        shortcut: 'Ctrl+/',
        dialog: 'keyboard-shortcuts',
      },
      { kind: 'separator', id: 'sep-help', feature: 'branding' },
      {
        kind: 'item',
        id: 'about',
        label: '关于 Casual Sheets',
        icon: 'info',
        dialog: 'about',
        feature: 'branding',
      },
      {
        kind: 'item',
        id: 'github',
        label: '在 GitHub 上查看',
        icon: 'open_in_new',
        run: () => openExternal('https://github.com/CasualOffice/sheets'),
        feature: 'branding',
      },
    ],
  },
];

/* ───────────────────────────── styling ────────────────────────────────── */

const BAR_STYLE: CSSProperties = {
  position: 'relative',
  display: 'flex',
  alignItems: 'center',
  gap: 2,
  padding: '2px 6px',
  borderBottom: '1px solid var(--cs-chrome-border, #e6e9ee)',
  background: 'var(--cs-chrome-bg, #eef1f5)',
  flex: '0 0 auto',
  userSelect: 'none',
  font: 'inherit',
  fontSize: 13,
};

const MENU_BTN_STYLE: CSSProperties = {
  height: 26,
  display: 'inline-flex',
  alignItems: 'center',
  padding: '0 10px',
  border: 'none',
  borderRadius: 6,
  background: 'transparent',
  cursor: 'pointer',
  color: 'var(--cs-chrome-fg, #201f1e)',
  font: 'inherit',
  fontSize: 13,
};

const DROPDOWN_STYLE: CSSProperties = {
  position: 'absolute',
  top: '100%',
  minWidth: 220,
  marginTop: 2,
  padding: 4,
  display: 'flex',
  flexDirection: 'column',
  border: '1px solid var(--cs-chrome-border, #e6e9ee)',
  borderRadius: 8,
  background: 'var(--cs-chrome-input-bg, #fff)',
  boxShadow: '0 6px 20px rgba(0,0,0,0.16)',
  zIndex: 1000,
};

const ITEM_STYLE: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 10,
  width: '100%',
  height: 30,
  padding: '0 10px',
  border: 'none',
  borderRadius: 6,
  background: 'transparent',
  cursor: 'pointer',
  color: 'var(--cs-chrome-fg, #201f1e)',
  font: 'inherit',
  fontSize: 13,
  textAlign: 'left',
};

const SEPARATOR_STYLE: CSSProperties = {
  height: 1,
  margin: '4px 6px',
  background: 'var(--cs-chrome-border, #e6e9ee)',
};

const SHORTCUT_STYLE: CSSProperties = {
  marginLeft: 'auto',
  paddingLeft: 16,
  fontSize: 11,
  color: 'var(--cs-chrome-muted, #6b7280)',
};

const SUBMENU_PANEL_STYLE: CSSProperties = {
  position: 'absolute',
  top: -4,
  left: '100%',
  minWidth: 200,
  marginLeft: 2,
  padding: 4,
  display: 'flex',
  flexDirection: 'column',
  border: '1px solid var(--cs-chrome-border, #e6e9ee)',
  borderRadius: 8,
  background: 'var(--cs-chrome-input-bg, #fff)',
  boxShadow: '0 6px 20px rgba(0,0,0,0.16)',
  zIndex: 1001,
};

/* ───────────────────────────── component ──────────────────────────────── */

export interface MenuBarProps {
  /** Live API, or `null` until the editor is ready. */
  api: CasualSheetsAPI | null;
  /**
   * Per-feature toggles. A control / menu whose `feature` key is `false` is not
   * rendered. Unknown / omitted keys default to enabled.
   */
  features?: Record<string, boolean>;
  /**
   * Host chrome extensions — custom menu items appended to their target menu.
   * Dialogs/toolbar/panels are handled elsewhere; only `extensions.menu` is read
   * here.
   */
  extensions?: ChromeExtensions;
}

export function MenuBar({ api, features = {}, extensions }: MenuBarProps) {
  const dialogs = useDialogs();
  const [open, setOpen] = useState<MenuId | null>(null);
  const [openSubmenu, setOpenSubmenu] = useState<string | null>(null);
  const rootRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    ensureChromeFonts();
  }, []);

  // Close on Escape + on a pointerdown outside the menu bar.
  useEffect(() => {
    if (open === null) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        setOpen(null);
        setOpenSubmenu(null);
      }
    };
    const onDown = (e: PointerEvent) => {
      if (!rootRef.current?.contains(e.target as Node)) {
        setOpen(null);
        setOpenSubmenu(null);
      }
    };
    document.addEventListener('keydown', onKey);
    document.addEventListener('pointerdown', onDown, true);
    return () => {
      document.removeEventListener('keydown', onKey);
      document.removeEventListener('pointerdown', onDown, true);
    };
  }, [open]);

  const close = () => {
    setOpen(null);
    setOpenSubmenu(null);
  };

  const runItem = (item: Extract<MenuItemDef, { kind: 'item' }>) => {
    close();
    if (!api) return;
    if (item.dialog) {
      // Range-seeded dialogs get the current selection in A1 as context.
      const seeded =
        item.dialog === 'insert-chart' ||
        item.dialog === 'insert-pivot' ||
        item.dialog === 'insert-sparkline' ||
        item.dialog === 'insert-cells' ||
        item.dialog === 'delete-cells';
      // openDialog opens the SDK built-in by default; a host override / host-owned
      // kind / onDialogRequest wins per the resolution rules in DialogProvider.
      dialogs.openDialog(item.dialog, seeded ? selectionA1(api) : undefined);
      return;
    }
    item.run?.(api);
  };

  // Compute the visible menus once per render (feature + dialog-open gating +
  // host menu extensions). Extracted to `computeVisibleMenus` so the gating
  // contract is unit-testable without a DOM.
  const visibleMenus = computeVisibleMenus(MENUS, features, dialogs.canOpen, extensions?.menu);

  const renderItems = (items: MenuItemDef[]): ReactNode =>
    items.map((item) => {
      if (item.kind === 'separator') {
        return <div key={item.id} style={SEPARATOR_STYLE} role="separator" aria-hidden />;
      }
      if (item.kind === 'submenu') {
        const isSubOpen = openSubmenu === item.id;
        return (
          <div
            key={item.id}
            style={{ position: 'relative' }}
            onMouseEnter={() => setOpenSubmenu(item.id)}
            onMouseLeave={() => setOpenSubmenu((cur) => (cur === item.id ? null : cur))}
          >
            <button
              type="button"
              role="menuitem"
              data-testid={`cs-menuitem-${item.id}`}
              aria-haspopup="menu"
              aria-expanded={isSubOpen}
              disabled={!api}
              style={{ ...ITEM_STYLE, opacity: api ? 1 : 0.5 }}
              onMouseDown={(e) => e.preventDefault()}
            >
              {item.icon ? (
                <Icon name={item.icon} size={18} />
              ) : (
                <span style={{ width: 18 }} aria-hidden />
              )}
              <span>{item.label}</span>
              <Icon name="chevron_right" size={18} style={{ marginLeft: 'auto' }} />
            </button>
            {isSubOpen && (
              <div style={SUBMENU_PANEL_STYLE} role="menu" aria-label={item.label}>
                {renderItems(item.items)}
              </div>
            )}
          </div>
        );
      }
      return (
        <button
          key={item.id}
          type="button"
          role="menuitem"
          data-testid={`cs-menuitem-${item.id}`}
          disabled={!api}
          style={{ ...ITEM_STYLE, opacity: api ? 1 : 0.5 }}
          onMouseDown={(e) => {
            e.preventDefault();
            runItem(item);
          }}
          onMouseEnter={(e) => {
            setOpenSubmenu(null);
            e.currentTarget.style.background = 'var(--cs-chrome-hover, rgba(0,0,0,0.06))';
          }}
          onMouseLeave={(e) => {
            e.currentTarget.style.background = 'transparent';
          }}
        >
          {item.icon ? (
            <Icon name={item.icon} size={18} />
          ) : (
            <span style={{ width: 18 }} aria-hidden />
          )}
          <span>{item.label}</span>
          {item.shortcut && <span style={SHORTCUT_STYLE}>{fmtShortcut(item.shortcut)}</span>}
        </button>
      );
    });

  return (
    <div
      ref={rootRef}
      style={BAR_STYLE}
      data-testid="cs-menubar"
      role="menubar"
      aria-label="菜单栏"
    >
      {visibleMenus.map((menu) => {
        const isOpen = open === menu.id;
        return (
          <div key={menu.id} style={{ position: 'relative' }}>
            <button
              type="button"
              data-menu={menu.id}
              data-testid={`cs-menu-${menu.id}`}
              aria-haspopup="menu"
              aria-expanded={isOpen}
              style={{
                ...MENU_BTN_STYLE,
                background: isOpen ? 'var(--cs-chrome-active, #e6f3f7)' : 'transparent',
                color: isOpen ? 'var(--cs-chrome-active-fg, #0e7490)' : MENU_BTN_STYLE.color,
              }}
              // mousedown (not click) so the grid selection isn't lost first;
              // toggle the menu open/closed.
              onMouseDown={(e) => {
                e.preventDefault();
                setOpenSubmenu(null);
                setOpen((cur) => (cur === menu.id ? null : menu.id));
              }}
              onMouseEnter={(e) => {
                // Hover-to-switch once a menu is already open (Office behaviour).
                if (open !== null && open !== menu.id) {
                  setOpen(menu.id);
                  setOpenSubmenu(null);
                } else if (!isOpen)
                  e.currentTarget.style.background = 'var(--cs-chrome-hover, rgba(0,0,0,0.06))';
              }}
              onMouseLeave={(e) => {
                e.currentTarget.style.background = isOpen
                  ? 'var(--cs-chrome-active, #e6f3f7)'
                  : 'transparent';
              }}
            >
              {menu.label}
            </button>
            {isOpen && (
              <div style={DROPDOWN_STYLE} role="menu" aria-label={menu.label}>
                {renderItems(menu.items)}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}
