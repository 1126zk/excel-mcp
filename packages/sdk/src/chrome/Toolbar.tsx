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

import { useEffect, useRef, useState, type RefObject } from 'react';
import { BorderStyleTypes, BorderType, ICommandService } from '@univerjs/core';
import type { CasualSheetsAPI } from '../sheets/api';
import { ensurePluginByName } from '../univer';
import { Icon } from './Icon';
import { ensureChromeFonts } from './fonts';
import { useDialogs } from './dialog-context';
import type { ChromeExtensions, ToolbarExtension } from './extensions';
import {
  BigToolbarButton,
  BordersControl,
  RibbonGroup,
  ToolbarButton,
  ToolbarColorButton,
  ToolbarDropdown,
  ToolbarSelect,
} from './RibbonControls';

type HAlign = 'left' | 'center' | 'right' | 'unset';
type VAlign = 'top' | 'middle' | 'bottom' | 'unset';

interface ActiveStyle {
  ready: boolean;
  isBold: boolean;
  isItalic: boolean;
  isUnderline: boolean;
  isStrike: boolean;
  isWrapped: boolean;
  fontFamily: string;
  fontSize: number;
  fontColor: string;
  fillColor: string;
  align: HAlign;
  vAlign: VAlign;
  numberFormat: string;
  isMerged: boolean;
  isMultiCell: boolean;
  isFormatPainterActive: boolean;
}

const NO_STYLE: ActiveStyle = {
  ready: false,
  isBold: false,
  isItalic: false,
  isUnderline: false,
  isStrike: false,
  isWrapped: false,
  fontFamily: '',
  fontSize: 11,
  fontColor: '',
  fillColor: '',
  align: 'unset',
  vAlign: 'unset',
  numberFormat: '',
  isMerged: false,
  isMultiCell: false,
  isFormatPainterActive: false,
};

function readActiveStyle(api: CasualSheetsAPI): ActiveStyle {
  const wb = api.univer.getActiveWorkbook();
  const sheet = wb?.getActiveSheet();
  const selection = sheet?.getActiveRange();
  if (!sheet || !selection) return NO_STYLE;

  const row = selection.getRow();
  const column = selection.getColumn();
  const cell = sheet.getRange(row, column);
  const style = cell.getCellStyleData?.() as {
    bl?: number;
    it?: number;
    ul?: { s?: number };
    st?: { s?: number };
    tb?: number;
    ff?: string;
    fs?: number;
    cl?: { rgb?: string };
    bg?: { rgb?: string };
    ht?: number;
    vt?: number;
    n?: { pattern?: string };
  } | null;
  const width = selection.getWidth();
  const height = selection.getHeight();
  const isMerged = sheet.getMergedRanges().some((merged) => {
    const range = merged.getRange();
    return (
      range.startRow === row &&
      range.startColumn === column &&
      range.endRow === row + height - 1 &&
      range.endColumn === column + width - 1
    );
  });

  let isFormatPainterActive = false;
  try {
    type Injector = { get(token: unknown): { getStatus?: () => number } | undefined };
    const univer = api.univer as unknown as {
      _univer?: { __getInjector?: () => Injector };
      _injector?: Injector;
    };
    const injector = univer._univer?.__getInjector?.() ?? univer._injector;
    const status = injector?.get({ name: 'IFormatPainterService' })?.getStatus?.();
    isFormatPainterActive = status !== undefined && status !== 0;
  } catch {
    isFormatPainterActive = false;
  }

  return {
    ready: true,
    isBold: style?.bl === 1,
    isItalic: style?.it === 1,
    isUnderline: style?.ul?.s === 1,
    isStrike: style?.st?.s === 1,
    isWrapped: style?.tb === 3 || cell.getWrap(),
    fontFamily: style?.ff ?? '',
    fontSize: style?.fs ?? 11,
    fontColor: style?.cl?.rgb ?? '',
    fillColor: style?.bg?.rgb ?? '',
    align: style?.ht === 1 ? 'left' : style?.ht === 2 ? 'center' : style?.ht === 3 ? 'right' : 'unset',
    vAlign:
      style?.vt === 1 ? 'top' : style?.vt === 2 ? 'middle' : style?.vt === 3 ? 'bottom' : 'unset',
    numberFormat: style?.n?.pattern ?? '',
    isMerged,
    isMultiCell: width * height > 1,
    isFormatPainterActive,
  };
}

function activeRange(api: CasualSheetsAPI) {
  return api.univer.getActiveWorkbook()?.getActiveSheet()?.getActiveRange() ?? null;
}

function activeSheet(api: CasualSheetsAPI) {
  return api.univer.getActiveWorkbook()?.getActiveSheet() ?? null;
}

function undo(api: CasualSheetsAPI) {
  void api.executeCommand('univer.command.undo');
}

function redo(api: CasualSheetsAPI) {
  void api.executeCommand('univer.command.redo');
}

function cut(api: CasualSheetsAPI) {
  void api.executeCommand('univer.command.cut');
}

function copy(api: CasualSheetsAPI) {
  void api.executeCommand('univer.command.copy');
}

function paste(api: CasualSheetsAPI) {
  void api.executeCommand('univer.command.paste');
}

function startFormatPainter(api: CasualSheetsAPI) {
  void api.executeCommand('sheet.command.set-once-format-painter');
}

function toggleBold(api: CasualSheetsAPI, current: boolean) {
  activeRange(api)?.setFontWeight(current ? 'normal' : 'bold');
}

function toggleItalic(api: CasualSheetsAPI, current: boolean) {
  activeRange(api)?.setFontStyle(current ? 'normal' : 'italic');
}

function toggleUnderline(api: CasualSheetsAPI, current: boolean) {
  activeRange(api)?.setFontLine(current ? 'none' : 'underline');
}

function toggleStrikethrough(api: CasualSheetsAPI, current: boolean) {
  activeRange(api)?.setFontLine(current ? 'none' : 'line-through');
}

function setFontFamily(api: CasualSheetsAPI, family: string) {
  activeRange(api)?.setFontFamily(family || null);
}

function setFontSize(api: CasualSheetsAPI, size: number) {
  if (!Number.isFinite(size) || size <= 0) return;
  activeRange(api)?.setFontSize(size);
}

function setFontColor(api: CasualSheetsAPI, color: string) {
  activeRange(api)?.setFontColor(color || null);
}

function setFillColor(api: CasualSheetsAPI, color: string) {
  activeRange(api)?.setBackground(color);
}

function setAlignment(api: CasualSheetsAPI, alignment: 'left' | 'center' | 'right') {
  activeRange(api)?.setHorizontalAlignment(alignment === 'right' ? 'normal' : alignment);
}

function setVerticalAlignment(api: CasualSheetsAPI, alignment: 'top' | 'middle' | 'bottom') {
  activeRange(api)?.setVerticalAlignment(alignment);
}

function toggleWrap(api: CasualSheetsAPI, current: boolean) {
  activeRange(api)?.setWrap(!current);
}

function toggleMerge(api: CasualSheetsAPI, current: boolean) {
  const range = activeRange(api);
  if (!range) return;
  if (current) {
    range.breakApart();
  } else {
    if (range.getWidth() * range.getHeight() <= 1) return;
    range.merge();
    range.setHorizontalAlignment('center');
  }
}

type BorderChoice = 'all' | 'outside' | 'top' | 'bottom' | 'left' | 'right' | 'none';
type BorderWeight = 'thin' | 'medium' | 'thick';

const DEFAULT_BORDER_COLOR = '#000000';
const WEIGHT_TO_STYLE: Record<BorderWeight, BorderStyleTypes> = {
  thin: BorderStyleTypes.THIN,
  medium: BorderStyleTypes.MEDIUM,
  thick: BorderStyleTypes.THICK,
};

function setBorders(
  api: CasualSheetsAPI,
  choice: BorderChoice,
  color: string = DEFAULT_BORDER_COLOR,
  weight: BorderWeight = 'thin',
) {
  const range = activeRange(api);
  if (!range) return;
  const type =
    choice === 'all'
      ? BorderType.ALL
      : choice === 'outside'
        ? BorderType.OUTSIDE
        : choice === 'top'
          ? BorderType.TOP
          : choice === 'bottom'
            ? BorderType.BOTTOM
            : choice === 'left'
              ? BorderType.LEFT
              : choice === 'right'
                ? BorderType.RIGHT
                : BorderType.NONE;
  const style = choice === 'none' ? BorderStyleTypes.NONE : WEIGHT_TO_STYLE[weight];
  range.setBorder(type, style, color);
}

const NUMBER_FORMAT_PATTERNS = {
  general: '',
  integer: '#,##0',
  number: '#,##0.00',
  currency: '"$"#,##0.00',
  accounting: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
  percent: '0.00%',
  date: 'yyyy-mm-dd',
  time: 'hh:mm:ss',
  scientific: '0.00E+00',
  text: '@',
} as const;

type NumberFormatKey = keyof typeof NUMBER_FORMAT_PATTERNS;

const NUMBER_FORMAT_OPTIONS: { value: NumberFormatKey; label: string }[] = [
  { value: 'general', label: '常规' },
  { value: 'integer', label: '整数' },
  { value: 'number', label: '数值（两位小数）' },
  { value: 'currency', label: '货币' },
  { value: 'accounting', label: '会计专用' },
  { value: 'percent', label: '百分比' },
  { value: 'date', label: '日期（yyyy-mm-dd）' },
  { value: 'time', label: '时间（hh:mm:ss）' },
  { value: 'scientific', label: '科学记数' },
  { value: 'text', label: '文本' },
];

function detectFormatKey(pattern: string): NumberFormatKey {
  for (const key of Object.keys(NUMBER_FORMAT_PATTERNS) as NumberFormatKey[]) {
    if (NUMBER_FORMAT_PATTERNS[key] === pattern) return key;
  }
  return 'general';
}

function setNumberFormat(api: CasualSheetsAPI, pattern: string) {
  const range = activeRange(api) as unknown as { setNumberFormat?: (p: string) => unknown } | null;
  range?.setNumberFormat?.(pattern);
}

function setNumberFormatByKey(api: CasualSheetsAPI, key: NumberFormatKey) {
  setNumberFormat(api, NUMBER_FORMAT_PATTERNS[key]);
}

function increaseDecimal(api: CasualSheetsAPI) {
  void api.executeCommand('sheet.command.numfmt.add.decimal.command');
}

function decreaseDecimal(api: CasualSheetsAPI) {
  void api.executeCommand('sheet.command.numfmt.subtract.decimal.command');
}

type AutoFn = 'SUM' | 'AVERAGE' | 'COUNT' | 'MIN' | 'MAX';

function applyAutoFunction(api: CasualSheetsAPI, fn: AutoFn) {
  const sheet = activeSheet(api);
  const range = activeRange(api);
  if (!sheet || !range) return;
  if (range.getWidth() * range.getHeight() > 1) {
    const target = sheet.getRange(range.getRow() + range.getHeight(), range.getColumn());
    target.setValue({ f: `=${fn}(${range.getA1Notation()})` });
    target.activate();
  } else {
    sheet.getRange(range.getRow(), range.getColumn()).setValue({ f: `=${fn}()` });
  }
}

const TABLE_THEMES = [
  { id: 'table-default-0', label: '靛蓝' },
  { id: 'table-default-1', label: '青绿' },
  { id: 'table-default-2', label: '绿色' },
  { id: 'table-default-3', label: '紫色' },
  { id: 'table-default-4', label: '粉色' },
  { id: 'table-default-5', label: '红色' },
] as const;

type TableThemeId = (typeof TABLE_THEMES)[number]['id'];
type ActiveRange = NonNullable<ReturnType<typeof activeRange>>;
type TableBounds = { startRow: number; startColumn: number; endRow: number; endColumn: number };

function detectTableBounds(api: CasualSheetsAPI, range: ActiveRange): TableBounds | null {
  if (range.getWidth() * range.getHeight() > 1) {
    return {
      startRow: range.getRow(),
      startColumn: range.getColumn(),
      endRow: range.getRow() + range.getHeight() - 1,
      endColumn: range.getColumn() + range.getWidth() - 1,
    };
  }

  const sheet = activeSheet(api);
  if (!sheet) return null;
  const startRow = range.getRow();
  const startColumn = range.getColumn();
  const hasValue = (r: number, c: number) => {
    try {
      const value = sheet.getRange(r, c).getCellData()?.v;
      return value !== undefined && value !== null && value !== '';
    } catch {
      return false;
    }
  };

  if (!hasValue(startRow, startColumn)) {
    return { startRow, startColumn, endRow: startRow, endColumn: startColumn };
  }

  let top = startRow;
  while (top > 0 && hasValue(top - 1, startColumn)) top--;
  let bottom = startRow;
  while (bottom < sheet.getMaxRows() - 1 && hasValue(bottom + 1, startColumn)) bottom++;
  let left = startColumn;
  while (left > 0 && hasValue(top, left - 1)) left--;
  let right = startColumn;
  while (right < sheet.getMaxColumns() - 1 && hasValue(top, right + 1)) right++;
  return { startRow: top, startColumn: left, endRow: bottom, endColumn: right };
}

let formatAsTableInFlight = false;

async function formatAsTable(api: CasualSheetsAPI, themeId?: TableThemeId) {
  if (formatAsTableInFlight) return;
  formatAsTableInFlight = true;
  try {
    const wb = api.univer.getActiveWorkbook();
    const sheet = activeSheet(api);
    const range = activeRange(api);
    if (!wb || !sheet || !range) return;
    const bounds = detectTableBounds(api, range);
    if (!bounds) return;

    await ensurePluginByName('table');
    const existing = (wb as unknown as { getTableList?: () => unknown[] }).getTableList?.() ?? [];
    const overlaps = existing.some((entry) => {
      const table = entry as { subUnitId?: string; range?: TableBounds };
      return (
        table.subUnitId === sheet.getSheetId() &&
        table.range &&
        !(
          table.range.endRow < bounds.startRow ||
          table.range.startRow > bounds.endRow ||
          table.range.endColumn < bounds.startColumn ||
          table.range.startColumn > bounds.endColumn
        )
      );
    });
    if (overlaps) return;

    const tableSheet = sheet as unknown as {
      addTable?: (
        name: string,
        range: TableBounds,
        id: string,
        options?: { tableStyleId: TableThemeId },
      ) => unknown;
      setColumnAutoWidth?: (column: number, count: number) => unknown;
    };
    if (typeof tableSheet.addTable !== 'function') return;
    const id = `table-${Date.now().toString(36)}`;
    const result = tableSheet.addTable(
      `Table_${Date.now().toString(36)}`,
      bounds,
      id,
      themeId ? { tableStyleId: themeId } : undefined,
    );
    const after = () => {
      const cells =
        (bounds.endRow - bounds.startRow + 1) * (bounds.endColumn - bounds.startColumn + 1);
      if (cells <= 20_000) {
        tableSheet.setColumnAutoWidth?.(
          bounds.startColumn,
          bounds.endColumn - bounds.startColumn + 1,
        );
      }
    };
    if (result && typeof (result as Promise<boolean>).then === 'function') {
      await (result as Promise<boolean>).then(after).catch(() => {});
    } else {
      after();
    }
  } finally {
    formatAsTableInFlight = false;
  }
}

function sortRange(api: CasualSheetsAPI, ascending: boolean) {
  const range = activeRange(api);
  if (!range) return;
  const withSort = range as unknown as {
    sort?: (spec: { column: number; ascending: boolean }) => unknown;
  };
  withSort.sort?.({ column: range.getColumn(), ascending });
}

async function toggleFilter(api: CasualSheetsAPI) {
  await ensurePluginByName('filter');
  const wb = api.univer.getActiveWorkbook();
  const sheet = activeSheet(api);
  const range = activeRange(api);
  if (!wb || !sheet || !range) return;
  const existing = (sheet as unknown as { getFilter?: () => unknown }).getFilter?.();
  if (existing) {
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
}

function insertHyperlink(api: CasualSheetsAPI) {
  void api.executeCommand('sheet.operation.insert-hyper-link');
}

function insertComment(api: CasualSheetsAPI) {
  void api.executeCommand('sheet.operation.show-comment-modal');
}

const FONT_FAMILIES = [
  'Calibri',
  'Arial',
  'Helvetica',
  'Inter',
  'Times New Roman',
  'Georgia',
  'Verdana',
  'Courier New',
  'JetBrains Mono',
];
const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48];

function useScrollOverflow(ref: RefObject<HTMLElement | null>) {
  const [state, setState] = useState({ left: false, right: false });
  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    const update = () => {
      const left = el.scrollLeft > 0;
      const right = el.scrollLeft + el.clientWidth < el.scrollWidth - 2;
      setState((previous) =>
        previous.left === left && previous.right === right ? previous : { left, right },
      );
    };
    update();
    el.addEventListener('scroll', update, { passive: true });
    const observer = new ResizeObserver(update);
    observer.observe(el);
    window.addEventListener('resize', update);
    return () => {
      el.removeEventListener('scroll', update);
      observer.disconnect();
      window.removeEventListener('resize', update);
    };
  }, [ref]);
  return state;
}

export interface ToolbarProps {
  api: CasualSheetsAPI | null;
  features?: Record<string, boolean>;
  extensions?: ChromeExtensions;
}

function enabled(features: Record<string, boolean> | undefined, key: string): boolean {
  return features?.[key] !== false;
}

export function Toolbar({ api, features, extensions }: ToolbarProps) {
  const dialogs = useDialogs();
  const [state, setState] = useState<ActiveStyle>(NO_STYLE);
  const ready = Boolean(api) && state.ready;

  useEffect(() => {
    ensureChromeFonts();
  }, []);

  useEffect(() => {
    if (!api) {
      setState(NO_STYLE);
      return;
    }
    const refresh = () => setState(readActiveStyle(api));
    refresh();
    const injector = (api.univer as unknown as { _injector?: { get(t: unknown): unknown } })
      ._injector;
    const commandService = injector?.get(ICommandService) as
      | { onCommandExecuted: (cb: () => void) => { dispose: () => void } }
      | undefined;
    const subscription = commandService?.onCommandExecuted(refresh);
    return () => subscription?.dispose();
  }, [api]);

  const innerRef = useRef<HTMLDivElement>(null);
  const overflow = useScrollOverflow(innerRef);
  const scrollBy = (dx: number) => innerRef.current?.scrollBy({ left: dx, behavior: 'smooth' });
  const adjustFontSize = (delta: number) => {
    if (!api) return;
    const current = state.fontSize > 0 ? state.fontSize : 11;
    const next = Math.max(6, Math.min(72, current + delta));
    if (next !== current) setFontSize(api, next);
  };

  const familyValue = state.fontFamily || 'Calibri';
  const familyOptions =
    state.fontFamily && !FONT_FAMILIES.includes(state.fontFamily)
      ? [state.fontFamily, ...FONT_FAMILIES]
      : FONT_FAMILIES;
  const sizeValue = state.fontSize || 11;
  const sizeOptions =
    state.fontSize && !FONT_SIZES.includes(state.fontSize)
      ? [state.fontSize, ...FONT_SIZES].sort((a, b) => a - b)
      : FONT_SIZES;

  const showFont = enabled(features, 'font');
  const showFontStyle = enabled(features, 'font-style');
  const showColor = enabled(features, 'color');
  const showBorders = enabled(features, 'borders');
  const showAlignment = enabled(features, 'alignment');
  const showMerge = enabled(features, 'merge');
  const showAutoSum = enabled(features, 'autosum');
  const showInsertChart = enabled(features, 'insert-chart') && dialogs.canOpen('insert-chart');
  const showPivotTable = enabled(features, 'pivot-table') && dialogs.canOpen('insert-pivot');
  const showFormatTable = enabled(features, 'format-as-table');
  const showSortFilter = enabled(features, 'sort-filter');
  const showHyperlink = enabled(features, 'insert-hyperlink');
  const showComment = enabled(features, 'insert-comment');
  const showFormatCells = enabled(features, 'format-cells') && dialogs.canOpen('format-cells');
  const showClearFormat = enabled(features, 'clear-format');
  const showInsert =
    showAutoSum ||
    showInsertChart ||
    showPivotTable ||
    showFormatTable ||
    showSortFilter ||
    showHyperlink ||
    showComment ||
    showFormatCells ||
    showClearFormat ||
    Boolean(extensions?.toolbar?.length);

  const toolbarExt = (extensions?.toolbar ?? []).filter(
    (extension) => !extension.isVisible || (api ? extension.isVisible(api) : false),
  );
  const runExt = (extension: ToolbarExtension) => {
    if (!api) return;
    if (extension.command) void api.executeCommand(extension.command, extension.commandParams);
    else extension.onClick?.(api);
  };

  return (
    <nav className="toolbar cs-ribbon-root" data-testid="toolbar" aria-label="工具栏">
      {overflow.left && (
        <button
          type="button"
          className="toolbar__overflow toolbar__overflow--left"
          data-testid="toolbar-overflow-left"
          aria-label="向左滚动工具栏"
          tabIndex={-1}
          onClick={() => scrollBy(-240)}
        >
          <Icon name="chevron_left" size={16} />
        </button>
      )}
      {overflow.right && (
        <button
          type="button"
          className="toolbar__overflow toolbar__overflow--right"
          data-testid="toolbar-overflow-right"
          aria-label="向右滚动工具栏，查看更多按钮"
          tabIndex={-1}
          onClick={() => scrollBy(240)}
        >
          <Icon name="chevron_right" size={16} />
        </button>
      )}
      <div className="toolbar__inner" ref={innerRef}>
        {enabled(features, 'history') && (
          <RibbonGroup
            label="历史"
            row1={
              <ToolbarButton
                id="undo"
                label="撤销 (Ctrl+Z)"
                icon="undo"
                disabled={!ready}
                onClick={() => api && undo(api)}
              />
            }
            row2={
              <ToolbarButton
                id="redo"
                label="重做 (Ctrl+Y)"
                icon="redo"
                disabled={!ready}
                onClick={() => api && redo(api)}
              />
            }
          />
        )}

        {(enabled(features, 'clipboard') || enabled(features, 'format-painter')) && (
          <RibbonGroup
            label="剪贴板"
            lead={
              enabled(features, 'clipboard') ? (
                <BigToolbarButton
                  id="paste"
                  label="粘贴 (Ctrl+V)"
                  icon="content_paste"
                  disabled={!ready}
                  onClick={() => api && paste(api)}
                />
              ) : undefined
            }
            row1={
              enabled(features, 'clipboard') ? (
                <>
                  <ToolbarButton
                    id="cut"
                    label="剪切 (Ctrl+X)"
                    icon="content_cut"
                    disabled={!ready}
                    onClick={() => api && cut(api)}
                  />
                  <ToolbarButton
                    id="copy"
                    label="复制 (Ctrl+C)"
                    icon="content_copy"
                    disabled={!ready}
                    onClick={() => api && copy(api)}
                  />
                  <ToolbarButton
                    id="paste-special"
                    label="选择性粘贴… (Ctrl+Alt+V)"
                    icon="content_paste_go"
                    disabled={!ready || !dialogs.canOpen('paste-special')}
                    onClick={() => dialogs.openDialog('paste-special')}
                  />
                </>
              ) : undefined
            }
            row2={
              enabled(features, 'format-painter') ? (
                <ToolbarButton
                  id="format-painter"
                  label={state.isFormatPainterActive ? '格式刷（已启用）' : '格式刷'}
                  icon="format_paint"
                  pressed={state.isFormatPainterActive}
                  disabled={!ready}
                  onClick={() => api && startFormatPainter(api)}
                />
              ) : undefined
            }
          />
        )}

        {(showFont || showFontStyle || showColor || showBorders) && (
          <RibbonGroup
            label="字体"
            row1={
              showFont ? (
                <>
                  <ToolbarSelect
                    id="font-family"
                    label="字体"
                    value={familyValue}
                    options={familyOptions.map((family) => ({ value: family, label: family }))}
                    width={130}
                    disabled={!ready}
                    onChange={(value) => api && setFontFamily(api, value)}
                  />
                  <ToolbarSelect
                    id="font-size"
                    label="字号"
                    value={String(sizeValue)}
                    options={sizeOptions.map((size) => ({ value: String(size), label: String(size) }))}
                    width={56}
                    disabled={!ready}
                    onChange={(value) => api && setFontSize(api, Number(value))}
                  />
                  <ToolbarButton
                    id="font-size-up"
                    label="增大字号"
                    icon="text_increase"
                    disabled={!ready}
                    onClick={() => adjustFontSize(+1)}
                  />
                  <ToolbarButton
                    id="font-size-down"
                    label="减小字号"
                    icon="text_decrease"
                    disabled={!ready}
                    onClick={() => adjustFontSize(-1)}
                  />
                </>
              ) : undefined
            }
            row2={
              <>
                {showFontStyle && (
                  <>
                    <ToolbarButton
                      id="bold"
                      label="加粗 (Ctrl+B)"
                      icon="format_bold"
                      pressed={state.isBold}
                      disabled={!ready}
                      onClick={() => api && toggleBold(api, state.isBold)}
                    />
                    <ToolbarButton
                      id="italic"
                      label="斜体 (Ctrl+I)"
                      icon="format_italic"
                      pressed={state.isItalic}
                      disabled={!ready}
                      onClick={() => api && toggleItalic(api, state.isItalic)}
                    />
                    <ToolbarButton
                      id="underline"
                      label="下划线 (Ctrl+U)"
                      icon="format_underlined"
                      pressed={state.isUnderline}
                      disabled={!ready}
                      onClick={() => api && toggleUnderline(api, state.isUnderline)}
                    />
                    <ToolbarButton
                      id="strikethrough"
                      label="删除线"
                      icon="format_strikethrough"
                      pressed={state.isStrike}
                      disabled={!ready}
                      onClick={() => api && toggleStrikethrough(api, state.isStrike)}
                    />
                  </>
                )}
                {showColor && (
                  <>
                    <ToolbarColorButton
                      id="font-color"
                      label="文字颜色"
                      icon="format_color_text"
                      value={state.fontColor}
                      defaultColor="#000000"
                      disabled={!ready}
                      onChange={(color) => api && setFontColor(api, color)}
                    />
                    <ToolbarColorButton
                      id="fill-color"
                      label="填充颜色"
                      icon="format_color_fill"
                      value={state.fillColor}
                      defaultColor="#ffeb3b"
                      disabled={!ready}
                      onChange={(color) => api && setFillColor(api, color)}
                    />
                  </>
                )}
                {showBorders && (
                  <BordersControl
                    label="边框"
                    icon="border_all"
                    disabled={!ready}
                    defaultColor={DEFAULT_BORDER_COLOR}
                    items={[
                      { id: 'all', label: '所有框线', icon: 'border_all' },
                      { id: 'outside', label: '外侧框线', icon: 'border_outer' },
                      { id: 'top', label: '上框线', icon: 'border_top' },
                      { id: 'bottom', label: '下框线', icon: 'border_bottom' },
                      { id: 'left', label: '左框线', icon: 'border_left' },
                      { id: 'right', label: '右框线', icon: 'border_right' },
                      { id: 'none', label: '无框线', icon: 'border_clear' },
                    ]}
                    onDefault={(color, weight) => api && setBorders(api, 'all', color, weight)}
                    onChoose={(choice, color, weight) =>
                      api && setBorders(api, choice as BorderChoice, color, weight)
                    }
                  />
                )}
              </>
            }
          />
        )}

        {(showAlignment || showMerge) && (
          <RibbonGroup
            label="对齐方式"
            row1={
              showAlignment ? (
                <>
                  <ToolbarButton
                    id="align-left"
                    label="左对齐"
                    icon="format_align_left"
                    pressed={state.align === 'left'}
                    disabled={!ready}
                    onClick={() => api && setAlignment(api, 'left')}
                  />
                  <ToolbarButton
                    id="align-center"
                    label="居中对齐"
                    icon="format_align_center"
                    pressed={state.align === 'center'}
                    disabled={!ready}
                    onClick={() => api && setAlignment(api, 'center')}
                  />
                  <ToolbarButton
                    id="align-right"
                    label="右对齐"
                    icon="format_align_right"
                    pressed={state.align === 'right'}
                    disabled={!ready}
                    onClick={() => api && setAlignment(api, 'right')}
                  />
                  <ToolbarButton
                    id="wrap-text"
                    label="自动换行"
                    icon="wrap_text"
                    pressed={state.isWrapped}
                    disabled={!ready}
                    onClick={() => api && toggleWrap(api, state.isWrapped)}
                  />
                </>
              ) : undefined
            }
            row2={
              <>
                {showAlignment && (
                  <>
                    <ToolbarButton
                      id="align-top"
                      label="顶端对齐"
                      icon="vertical_align_top"
                      pressed={state.vAlign === 'top'}
                      disabled={!ready}
                      onClick={() => api && setVerticalAlignment(api, 'top')}
                    />
                    <ToolbarButton
                      id="align-middle"
                      label="垂直居中"
                      icon="vertical_align_center"
                      pressed={state.vAlign === 'middle'}
                      disabled={!ready}
                      onClick={() => api && setVerticalAlignment(api, 'middle')}
                    />
                    <ToolbarButton
                      id="align-bottom"
                      label="底端对齐"
                      icon="vertical_align_bottom"
                      pressed={state.vAlign === 'bottom'}
                      disabled={!ready}
                      onClick={() => api && setVerticalAlignment(api, 'bottom')}
                    />
                  </>
                )}
                {showMerge && (
                  <ToolbarButton
                    id="merge-cells"
                    label={state.isMerged ? '取消合并单元格' : '合并后居中'}
                    icon={state.isMerged ? 'call_split' : 'cell_merge'}
                    pressed={state.isMerged}
                    disabled={!ready || (!state.isMerged && !state.isMultiCell)}
                    onClick={() => api && toggleMerge(api, state.isMerged)}
                  />
                )}
              </>
            }
          />
        )}

        {enabled(features, 'number') && (
          <RibbonGroup
            label="数字"
            row1={
              <ToolbarSelect
                id="num-format"
                label="数字格式"
                value={detectFormatKey(state.numberFormat)}
                options={NUMBER_FORMAT_OPTIONS}
                width={108}
                disabled={!ready}
                onChange={(value) => api && setNumberFormatByKey(api, value as NumberFormatKey)}
              />
            }
            row2={
              <>
                <ToolbarButton
                  id="numfmt-currency"
                  label="货币"
                  icon="attach_money"
                  pressed={state.numberFormat === NUMBER_FORMAT_PATTERNS.currency}
                  disabled={!ready}
                  onClick={() => api && setNumberFormat(api, NUMBER_FORMAT_PATTERNS.currency)}
                />
                <ToolbarButton
                  id="numfmt-percent"
                  label="百分比"
                  icon="percent"
                  pressed={state.numberFormat === NUMBER_FORMAT_PATTERNS.percent}
                  disabled={!ready}
                  onClick={() => api && setNumberFormat(api, NUMBER_FORMAT_PATTERNS.percent)}
                />
                <ToolbarButton
                  id="decimal-up"
                  label="增加小数位数"
                  icon="add"
                  disabled={!ready}
                  onClick={() => api && increaseDecimal(api)}
                />
                <ToolbarButton
                  id="decimal-down"
                  label="减少小数位数"
                  icon="remove"
                  disabled={!ready}
                  onClick={() => api && decreaseDecimal(api)}
                />
              </>
            }
          />
        )}

        {showInsert && (
          <RibbonGroup
            label="插入"
            row1={
              <>
                {showAutoSum && (
                  <ToolbarDropdown
                    id="auto-sum"
                    label="自动求和"
                    icon="functions"
                    disabled={!ready}
                    items={[
                      { id: 'SUM', label: '求和', icon: 'functions' },
                      { id: 'AVERAGE', label: '平均值', icon: 'bar_chart' },
                      { id: 'COUNT', label: '计数', icon: 'numbers' },
                      { id: 'MAX', label: '最大值', icon: 'arrow_upward' },
                      { id: 'MIN', label: '最小值', icon: 'arrow_downward' },
                    ]}
                    onDefault={() => api && applyAutoFunction(api, 'SUM')}
                    onChoose={(id) => api && applyAutoFunction(api, id as AutoFn)}
                  />
                )}
                {showInsertChart && (
                  <ToolbarButton
                    id="insert-chart"
                    label="插入图表"
                    icon="bar_chart"
                    disabled={!ready}
                    onClick={() => dialogs.openDialog('insert-chart')}
                  />
                )}
                {showPivotTable && (
                  <ToolbarButton
                    id="insert-pivot"
                    label="插入数据透视表"
                    icon="pivot_table_chart"
                    disabled={!ready}
                    onClick={() => dialogs.openDialog('insert-pivot')}
                  />
                )}
                {showFormatTable && (
                  <ToolbarDropdown
                    id="format-as-table"
                    label="套用表格格式"
                    icon="table_rows"
                    disabled={!ready}
                    items={[
                      { id: 'plain', label: '普通（无样式）', icon: 'grid_on' },
                      ...TABLE_THEMES.map((theme) => ({
                        id: theme.id,
                        label: theme.label,
                        icon: 'table_chart',
                      })),
                    ]}
                    onDefault={() => api && void formatAsTable(api, 'table-default-0')}
                    onChoose={(id) => {
                      if (!api) return;
                      void formatAsTable(api, id === 'plain' ? undefined : (id as TableThemeId));
                    }}
                  />
                )}
                {toolbarExt.map((extension) => (
                  <ToolbarButton
                    key={extension.id}
                    id={`ext-${extension.id}`}
                    label={extension.label}
                    icon={extension.icon}
                    disabled={!ready}
                    onClick={() => runExt(extension)}
                  />
                ))}
              </>
            }
            row2={
              <>
                {showSortFilter && (
                  <ToolbarDropdown
                    id="sort-filter"
                    label="排序和筛选"
                    icon="filter_list"
                    disabled={!ready}
                    items={[
                      { id: 'sort-asc', label: '升序排序', icon: 'arrow_downward' },
                      { id: 'sort-desc', label: '降序排序', icon: 'arrow_upward' },
                      { id: 'toggle-filter', label: '切换筛选', icon: 'filter_alt' },
                    ]}
                    onDefault={() => api && state.isMultiCell && sortRange(api, true)}
                    onChoose={(id) => {
                      if (!api) return;
                      if (id === 'sort-asc') sortRange(api, true);
                      else if (id === 'sort-desc') sortRange(api, false);
                      else if (id === 'toggle-filter') void toggleFilter(api);
                    }}
                  />
                )}
                {showHyperlink && (
                  <ToolbarButton
                    id="insert-hyperlink"
                    label="插入超链接 (Ctrl+K)"
                    icon="link"
                    disabled={!ready}
                    onClick={() => api && insertHyperlink(api)}
                  />
                )}
                {showComment && (
                  <ToolbarButton
                    id="insert-comment"
                    label="插入批注"
                    icon="add_comment"
                    disabled={!ready}
                    onClick={() => api && insertComment(api)}
                  />
                )}
                {showFormatCells && (
                  <ToolbarButton
                    id="format-cells"
                    label="设置单元格格式…"
                    icon="format_shapes"
                    disabled={!ready}
                    onClick={() => dialogs.openDialog('format-cells')}
                  />
                )}
                {showClearFormat && (
                  <ToolbarButton
                    id="clear-format"
                    label="清除格式"
                    icon="format_clear"
                    disabled={!ready}
                    onClick={() => api && void api.executeCommand('sheet.command.clear-selection-format')}
                  />
                )}
              </>
            }
          />
        )}
      </div>
    </nav>
  );
}
