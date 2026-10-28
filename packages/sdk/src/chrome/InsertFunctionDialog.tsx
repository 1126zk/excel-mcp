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
 * InsertFunctionDialog — the SDK chrome's built-in Insert Function picker.
 *
 * A searchable, category-grouped list of common spreadsheet functions. Picking a
 * function seeds the active cell with `=NAME(` so the user can type arguments,
 * then focuses the grid. Applying is grounded in the real Sheets FUniver facade:
 *   - `api.univer.getActiveWorkbook()?.getActiveSheet()?.getActiveRange()` gives
 *     the live FRange (see FormatCellsDialog `activeRange`);
 *   - `FRange.setValue('=NAME(')` writes the formula stub into the range's
 *     top-left cell (verified in `@univerjs/sheets/lib/types/facade/f-range.d.ts`
 *     L814 `setValue(value: CellValue | ICellData): FRange`);
 *   - `FRange.getA1Notation()` (f-range.d.ts L1220) labels the target cell;
 *   - `api.focus()` returns keyboard focus to the grid for argument entry.
 *
 * Note on the caret: the installed FRange facade has no "enter cell edit mode
 * with the caret parked after the open-paren" method — `setValue` commits the
 * cell content and Univer re-parses it. So on choose we write `=NAME(` and focus
 * the grid; the user double-clicks / F2s to continue typing arguments. That's the
 * closest real, installed behaviour (recorded in the dialog's limitations).
 *
 * Mounted by `<DialogHost>` when `openDialog('insert-function')` is called and no
 * host override is registered.
 */

import { useMemo, useState, type CSSProperties } from 'react';
import type { DialogComponentProps } from './extensions';
import type { CasualSheetsAPI } from '../sheets/api';
import { Dialog } from './Dialog';
import {
  DIALOG_BTN_PRIMARY_STYLE,
  DIALOG_BTN_SECONDARY_STYLE,
  DIALOG_INPUT_STYLE,
} from './dialog-styles';

/** A single function entry in the picker. */
interface FunctionEntry {
  /** Canonical function name (upper-case), inserted verbatim. */
  name: string;
  /** Category bucket for grouping. */
  category: string;
  /** One-line description shown under the list / in the preview. */
  description: string;
  /** Argument-signature hint shown in the preview, e.g. `SUM(value1, [value2, …])`. */
  syntax: string;
}

/**
 * The common-function catalog. Deliberately a focused, Google-Sheets-style
 * "common functions" set rather than the full library — this is a quick-insert
 * picker, not a reference. Grouped by `category`.
 */
const FUNCTIONS: FunctionEntry[] = [
  // Math
  {
    name: 'SUM',
    category: 'Math',
    description: '计算一组数值的总和。',
    syntax: 'SUM(数值1, [数值2, …])',
  },
  {
    name: 'PRODUCT',
    category: 'Math',
    description: '计算一组数值的乘积。',
    syntax: 'PRODUCT(因数1, [因数2, …])',
  },
  {
    name: 'ROUND',
    category: 'Math',
    description: '将数字四舍五入到指定的小数位数。',
    syntax: 'ROUND(数值, [小数位数])',
  },
  {
    name: 'ABS',
    category: 'Math',
    description: '计算数字的绝对值。',
    syntax: 'ABS(数值)',
  },
  {
    name: 'MOD',
    category: 'Math',
    description: '计算除法运算的余数。',
    syntax: 'MOD(被除数, 除数)',
  },
  {
    name: 'POWER',
    category: 'Math',
    description: '计算数字的指定次幂。',
    syntax: 'POWER(底数, 指数)',
  },
  // Statistical
  {
    name: 'AVERAGE',
    category: 'Statistical',
    description: '计算一组数值的算术平均值。',
    syntax: 'AVERAGE(数值1, [数值2, …])',
  },
  {
    name: 'COUNT',
    category: 'Statistical',
    description: '计算区域中数值的个数。',
    syntax: 'COUNT(值1, [值2, …])',
  },
  {
    name: 'COUNTA',
    category: 'Statistical',
    description: '计算区域中非空值的个数。',
    syntax: 'COUNTA(值1, [值2, …])',
  },
  {
    name: 'MAX',
    category: 'Statistical',
    description: '计算一组数值中的最大值。',
    syntax: 'MAX(数值1, [数值2, …])',
  },
  {
    name: 'MIN',
    category: 'Statistical',
    description: '计算一组数值中的最小值。',
    syntax: 'MIN(数值1, [数值2, …])',
  },
  {
    name: 'MEDIAN',
    category: 'Statistical',
    description: '计算一组数值的中位数。',
    syntax: 'MEDIAN(数值1, [数值2, …])',
  },
  // Logical
  {
    name: 'IF',
    category: 'Logical',
    description: '条件为真时返回一个值，为假时返回另一个值。',
    syntax: 'IF(条件, 条件为真时的值, 条件为假时的值)',
  },
  {
    name: 'IFS',
    category: 'Logical',
    description: '测试多个条件，并返回第一个匹配条件对应的值。',
    syntax: 'IFS(条件1, 值1, [条件2, 值2, …])',
  },
  {
    name: 'AND',
    category: 'Logical',
    description: '所有参数均为真时返回真。',
    syntax: 'AND(逻辑值1, [逻辑值2, …])',
  },
  {
    name: 'OR',
    category: 'Logical',
    description: '任一参数为真时返回真。',
    syntax: 'OR(逻辑值1, [逻辑值2, …])',
  },
  {
    name: 'IFERROR',
    category: 'Logical',
    description: '公式出错时返回备用值。',
    syntax: 'IFERROR(值, 出错时的值)',
  },
  // Lookup
  {
    name: 'VLOOKUP',
    category: 'Lookup',
    description: '在列中搜索关键字，并返回同一行中的值。',
    syntax: 'VLOOKUP(搜索关键字, 区域, 索引, [是否已排序])',
  },
  {
    name: 'HLOOKUP',
    category: 'Lookup',
    description: '在行中搜索关键字，并返回同一列中的值。',
    syntax: 'HLOOKUP(搜索关键字, 区域, 索引, [是否已排序])',
  },
  {
    name: 'INDEX',
    category: 'Lookup',
    description: '返回区域中指定行和列偏移位置的值。',
    syntax: 'INDEX(引用, [行], [列])',
  },
  {
    name: 'MATCH',
    category: 'Lookup',
    description: '返回值在区域中的位置。',
    syntax: 'MATCH(搜索关键字, 区域, [搜索类型])',
  },
  {
    name: 'XLOOKUP',
    category: 'Lookup',
    description: '搜索区域并返回对应的值。',
    syntax: 'XLOOKUP(搜索关键字, 查找区域, 返回区域)',
  },
  // Text
  {
    name: 'CONCATENATE',
    category: 'Text',
    description: '将多个字符串首尾连接。',
    syntax: 'CONCATENATE(字符串1, [字符串2, …])',
  },
  {
    name: 'LEFT',
    category: 'Text',
    description: '返回字符串最左侧的字符。',
    syntax: 'LEFT(字符串, [字符数])',
  },
  {
    name: 'RIGHT',
    category: 'Text',
    description: '返回字符串最右侧的字符。',
    syntax: 'RIGHT(字符串, [字符数])',
  },
  {
    name: 'MID',
    category: 'Text',
    description: '返回字符串中间的字符。',
    syntax: 'MID(字符串, 起始位置, 长度)',
  },
  { name: 'LEN', category: 'Text', description: '计算字符串的长度。', syntax: 'LEN(字符串)' },
  {
    name: 'TRIM',
    category: 'Text',
    description: '删除首尾空格和重复空格。',
    syntax: 'TRIM(字符串)',
  },
  // Date
  { name: 'TODAY', category: 'Date', description: '返回当前日期。', syntax: 'TODAY()' },
  { name: 'NOW', category: 'Date', description: '返回当前日期和时间。', syntax: 'NOW()' },
  {
    name: 'DATE',
    category: 'Date',
    description: '根据年、月、日生成日期。',
    syntax: 'DATE(年, 月, 日)',
  },
  {
    name: 'DATEDIF',
    category: 'Date',
    description: '按指定单位计算两个日期之间的差值。',
    syntax: 'DATEDIF(开始日期, 结束日期, 单位)',
  },
];

/** Ordered category list, so groups render in a stable, sensible order. */
const CATEGORY_ORDER = ['Math', 'Statistical', 'Logical', 'Lookup', 'Text', 'Date'];
const CATEGORY_LABELS: Record<string, string> = {
  Math: '数学',
  Statistical: '统计',
  Logical: '逻辑',
  Lookup: '查找',
  Text: '文本',
  Date: '日期',
};

/** The active FRange, or null when there is no selection. */
function activeRange(api: CasualSheetsAPI) {
  return api.univer.getActiveWorkbook()?.getActiveSheet()?.getActiveRange() ?? null;
}

/**
 * Insert `=NAME(` into the top-left cell of the active range via the facade,
 * then focus the grid for argument entry. Returns false when there's no
 * selection to write into.
 */
function insertFunction(api: CasualSheetsAPI, fn: FunctionEntry): boolean {
  const range = activeRange(api);
  if (!range) return false;

  // Zero-arg functions (TODAY / NOW) are complete on their own; for the rest,
  // leave the open-paren so the user continues typing arguments.
  const zeroArg = fn.syntax.endsWith('()');
  const formula = zeroArg ? `=${fn.name}()` : `=${fn.name}(`;

  // FRange.setValue writes into the range's top-left cell (f-range.d.ts L814).
  range.setValue(formula);

  // Return keyboard focus to the grid so the user can keep typing arguments.
  api.focus();
  return true;
}

const SEARCH_STYLE: CSSProperties = {
  ...DIALOG_INPUT_STYLE,
  width: '100%',
  marginBottom: 12,
};

const RANGE_NOTE_STYLE: CSSProperties = {
  fontSize: 12,
  color: 'var(--cs-chrome-muted, #605e5c)',
  marginBottom: 12,
};

const LIST_STYLE: CSSProperties = {
  maxHeight: 280,
  overflow: 'auto',
  border: '1px solid var(--cs-chrome-border, #edeff3)',
  borderRadius: 8,
};

const GROUP_HEADER_STYLE: CSSProperties = {
  position: 'sticky',
  top: 0,
  padding: '6px 12px',
  fontSize: 11,
  fontWeight: 600,
  letterSpacing: 0.4,
  textTransform: 'uppercase',
  color: 'var(--cs-chrome-muted, #605e5c)',
  background: 'var(--cs-chrome-input-bg, #fff)',
  borderBottom: '1px solid var(--cs-chrome-border, #edeff3)',
};

function itemStyle(selected: boolean): CSSProperties {
  return {
    display: 'block',
    width: '100%',
    textAlign: 'left',
    padding: '8px 12px',
    border: 'none',
    borderBottom: '1px solid var(--cs-chrome-border, #f3f4f6)',
    background: selected ? 'var(--cs-chrome-hover-bg, #eef6f8)' : 'transparent',
    color: 'var(--cs-chrome-fg, #201f1e)',
    font: 'inherit',
    fontSize: 13,
    cursor: 'pointer',
  };
}

const ITEM_NAME_STYLE: CSSProperties = {
  fontWeight: 600,
  fontFamily: 'var(--cs-chrome-mono, ui-monospace, SFMono-Regular, Menlo, monospace)',
};

const ITEM_DESC_STYLE: CSSProperties = {
  fontSize: 12,
  color: 'var(--cs-chrome-muted, #605e5c)',
  marginTop: 2,
};

const PREVIEW_STYLE: CSSProperties = {
  marginTop: 12,
  padding: '10px 12px',
  borderRadius: 8,
  background: 'var(--cs-chrome-subtle-bg, #f6f8fa)',
  fontSize: 12,
  color: 'var(--cs-chrome-fg, #201f1e)',
};

const PREVIEW_SYNTAX_STYLE: CSSProperties = {
  fontFamily: 'var(--cs-chrome-mono, ui-monospace, SFMono-Regular, Menlo, monospace)',
  fontWeight: 600,
  marginBottom: 4,
};

export function InsertFunctionDialog({ api, onClose }: DialogComponentProps) {
  const [query, setQuery] = useState('');
  const [selectedName, setSelectedName] = useState<string | null>(null);

  // Label the target cell once for the header hint.
  const rangeLabel = useMemo(() => {
    const fRange = activeRange(api) as unknown as { getA1Notation?: () => string } | null;
    return fRange?.getA1Notation?.() ?? null;
  }, [api]);

  const hasSelection = activeRange(api) !== null;

  // Filter by name / description / category, case-insensitive.
  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return FUNCTIONS;
    return FUNCTIONS.filter(
      (fn) =>
        fn.name.toLowerCase().includes(q) ||
        fn.description.toLowerCase().includes(q) ||
        fn.category.toLowerCase().includes(q) ||
        (CATEGORY_LABELS[fn.category] ?? '').toLowerCase().includes(q),
    );
  }, [query]);

  // Group the filtered set by category, preserving CATEGORY_ORDER.
  const grouped = useMemo(() => {
    const byCat = new Map<string, FunctionEntry[]>();
    for (const fn of filtered) {
      const list = byCat.get(fn.category);
      if (list) list.push(fn);
      else byCat.set(fn.category, [fn]);
    }
    return CATEGORY_ORDER.filter((c) => byCat.has(c)).map((c) => ({
      category: c,
      items: byCat.get(c)!,
    }));
  }, [filtered]);

  const selected = useMemo(
    () => FUNCTIONS.find((fn) => fn.name === selectedName) ?? null,
    [selectedName],
  );

  const choose = (fn: FunctionEntry) => {
    if (insertFunction(api, fn)) onClose();
  };

  const insertSelected = () => {
    if (selected && insertFunction(api, selected)) onClose();
  };

  return (
    <Dialog
      title="插入函数"
      onClose={onClose}
      width={460}
      data-testid="cs-insert-function-dialog"
      footer={
        <>
          <button type="button" style={DIALOG_BTN_SECONDARY_STYLE} onClick={onClose}>
            取消
          </button>
          <button
            type="button"
            style={DIALOG_BTN_PRIMARY_STYLE}
            data-testid="cs-insert-function-insert"
            disabled={!hasSelection || !selected}
            onClick={insertSelected}
          >
            插入
          </button>
        </>
      }
    >
      {hasSelection ? (
        <div style={RANGE_NOTE_STYLE} data-testid="cs-insert-function-range">
          插入到 <strong>{rangeLabel ?? '活动单元格'}</strong>
        </div>
      ) : (
        <div style={RANGE_NOTE_STYLE} data-testid="cs-insert-function-no-selection">
          请先选择一个单元格，然后重新打开此对话框。
        </div>
      )}

      <input
        style={SEARCH_STYLE}
        data-testid="cs-insert-function-search"
        type="search"
        placeholder="搜索函数（名称、类别或说明）"
        value={query}
        autoFocus
        onChange={(e) => setQuery(e.target.value)}
      />

      <div
        style={LIST_STYLE}
        role="listbox"
        aria-label="函数"
        data-testid="cs-insert-function-list"
      >
        {grouped.length === 0 && (
          <div style={{ padding: '12px', fontSize: 13, color: 'var(--cs-chrome-muted, #605e5c)' }}>
            没有与&ldquo;{query}&rdquo;匹配的函数。
          </div>
        )}
        {grouped.map((group) => (
          <div key={group.category}>
            <div style={GROUP_HEADER_STYLE}>{CATEGORY_LABELS[group.category] ?? group.category}</div>
            {group.items.map((fn) => (
              <button
                key={fn.name}
                type="button"
                role="option"
                aria-selected={selectedName === fn.name}
                data-testid={`cs-insert-function-item-${fn.name}`}
                style={itemStyle(selectedName === fn.name)}
                onClick={() => setSelectedName(fn.name)}
                onDoubleClick={() => choose(fn)}
              >
                <div style={ITEM_NAME_STYLE}>{fn.name}</div>
                <div style={ITEM_DESC_STYLE}>{fn.description}</div>
              </button>
            ))}
          </div>
        ))}
      </div>

      {selected && (
        <div style={PREVIEW_STYLE} data-testid="cs-insert-function-preview">
          <div style={PREVIEW_SYNTAX_STYLE}>{selected.syntax}</div>
          <div>{selected.description}</div>
        </div>
      )}
    </Dialog>
  );
}
