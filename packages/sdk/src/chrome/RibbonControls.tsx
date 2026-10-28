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

import {
  Children,
  cloneElement,
  isValidElement,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  type ChangeEvent,
  type CSSProperties,
  type FocusEvent as ReactFocusEvent,
  type MouseEvent as ReactMouseEvent,
  type ReactElement,
  type ReactNode,
  type RefObject,
} from 'react';
import { createPortal } from 'react-dom';
import { Icon } from './Icon';

type PopoverProps = {
  anchorRef: RefObject<HTMLElement | null>;
  onClose: () => void;
  className?: string;
  ['data-testid']?: string;
  children: ReactNode;
};

const PORTAL_TOKEN_NAMES = [
  '--cs-chrome-bg',
  '--cs-chrome-fg',
  '--cs-chrome-muted',
  '--cs-chrome-border',
  '--cs-chrome-input-bg',
  '--cs-chrome-hover',
  '--cs-chrome-active',
  '--cs-chrome-active-fg',
] as const;

function Popover({ anchorRef, onClose, className, children, ...rest }: PopoverProps) {
  const ref = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState<{ top: number; left: number }>({ top: 0, left: 0 });
  const [portalTokens, setPortalTokens] = useState<CSSProperties>({});
  const [theme, setTheme] = useState<string | null>(null);

  useLayoutEffect(() => {
    const anchor = anchorRef.current;
    const rect = anchor?.getBoundingClientRect();
    if (rect) setPos({ top: rect.bottom + 2, left: rect.left });
    if (anchor) {
      setTheme(anchor.closest('[data-theme]')?.getAttribute('data-theme') ?? null);
      const computed = getComputedStyle(anchor);
      const tokens: Record<string, string> = {};
      for (const name of PORTAL_TOKEN_NAMES) {
        const value = computed.getPropertyValue(name);
        if (value) tokens[name] = value;
      }
      setPortalTokens(tokens);
    }
  }, [anchorRef]);

  useEffect(() => {
    const onDocPointerDown = (e: PointerEvent) => {
      if (
        !ref.current?.contains(e.target as Node) &&
        !anchorRef.current?.contains(e.target as Node)
      ) {
        onClose();
      }
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    document.addEventListener('pointerdown', onDocPointerDown, true);
    document.addEventListener('keydown', onKey);
    return () => {
      document.removeEventListener('pointerdown', onDocPointerDown, true);
      document.removeEventListener('keydown', onKey);
    };
  }, [onClose, anchorRef]);

  return createPortal(
    <div
      ref={ref}
      className={`menu ${className ?? ''}`}
      style={{ top: pos.top, left: pos.left, ...portalTokens }}
      role="menu"
      data-theme={theme ?? undefined}
      data-testid={rest['data-testid']}
    >
      {children}
    </div>,
    document.body,
  );
}

type TooltipProps = {
  label: string;
  shortcut?: string;
  children: ReactElement;
  side?: 'top' | 'bottom';
  delay?: number;
};

function Tooltip({ label, shortcut, children, side = 'bottom', delay = 150 }: TooltipProps) {
  const child = Children.only(children);
  const [coords, setCoords] = useState<{ top: number; left: number } | null>(null);
  const [clampedLeft, setClampedLeft] = useState<number | null>(null);
  const tipRef = useRef<HTMLSpanElement>(null);
  const timerRef = useRef<number | null>(null);

  useLayoutEffect(() => {
    if (!coords || !tipRef.current) return;
    const half = tipRef.current.offsetWidth / 2;
    const min = half + 8;
    const max = window.innerWidth - half - 8;
    const next = Math.min(Math.max(coords.left, min), max);
    if (next !== clampedLeft) setClampedLeft(next);
  }, [coords, clampedLeft]);

  useEffect(
    () => () => {
      if (timerRef.current !== null) window.clearTimeout(timerRef.current);
    },
    [],
  );

  if (!isValidElement(child) || !label) return <>{children}</>;

  const place = (el: HTMLElement) => {
    const rect = el.getBoundingClientRect();
    const left = rect.left + rect.width / 2;
    if (side === 'top') {
      setCoords({ top: rect.top - 6, left });
    } else {
      setCoords({ top: rect.bottom + 6, left });
    }
  };

  const childProps = (child as any).props as {
    onMouseEnter?: (e: ReactMouseEvent<HTMLElement>) => void;
    onMouseLeave?: (e: ReactMouseEvent<HTMLElement>) => void;
    onFocus?: (e: ReactFocusEvent<HTMLElement>) => void;
    onBlur?: (e: ReactFocusEvent<HTMLElement>) => void;
  };

  const show = (target: HTMLElement) => {
    if (timerRef.current !== null) window.clearTimeout(timerRef.current);
    timerRef.current = window.setTimeout(() => {
      place(target);
      timerRef.current = null;
    }, delay);
  };

  const hide = () => {
    if (timerRef.current !== null) {
      window.clearTimeout(timerRef.current);
      timerRef.current = null;
    }
    setCoords(null);
    setClampedLeft(null);
  };

  const onMouseEnter = (e: ReactMouseEvent<HTMLElement>) => {
    show(e.currentTarget);
    childProps.onMouseEnter?.(e);
  };
  const onMouseLeave = (e: ReactMouseEvent<HTMLElement>) => {
    hide();
    childProps.onMouseLeave?.(e);
  };
  const onFocus = (e: ReactFocusEvent<HTMLElement>) => {
    show(e.currentTarget);
    childProps.onFocus?.(e);
  };
  const onBlur = (e: ReactFocusEvent<HTMLElement>) => {
    hide();
    childProps.onBlur?.(e);
  };

  const cloned = cloneElement(child, {
    onMouseEnter,
    onMouseLeave,
    onFocus,
    onBlur,
  } as any);

  return (
    <>
      {cloned}
      {coords &&
        createPortal(
          <span
            ref={tipRef}
            role="tooltip"
            data-testid="tooltip"
            className={`tooltip tooltip--${side}${shortcut ? ' tooltip--with-shortcut' : ''}`}
            style={{ top: coords.top, left: clampedLeft ?? coords.left }}
          >
            <span className="tooltip__label">{label}</span>
            {shortcut && (
              <span className="tooltip__shortcut" data-testid="tooltip-shortcut">
                {shortcut}
              </span>
            )}
          </span>,
          document.body,
        )}
    </>
  );
}

export function RibbonGroup({
  label,
  rows,
  lead,
  row1,
  row2,
  children,
}: {
  label: string;
  rows?: boolean;
  lead?: ReactNode;
  row1?: ReactNode;
  row2?: ReactNode;
  children?: ReactNode;
}) {
  const explicit = row1 !== undefined || row2 !== undefined || lead !== undefined;
  return (
    <div
      className="ribbon__group"
      role="group"
      aria-label={label}
      data-testid={`ribbon-group-${label.toLowerCase()}`}
    >
      {explicit ? (
        <div className="ribbon__group-body ribbon__group-body--two-row">
          {lead && <div className="ribbon__group-lead">{lead}</div>}
          <div className="ribbon__group-stack">
            <div className="ribbon__group-row" data-row="1">{row1}</div>
            <div className="ribbon__group-row" data-row="2">{row2}</div>
          </div>
        </div>
      ) : (
        <div className={`ribbon__group-body${rows ? ' ribbon__group-body--rows' : ''}`}>
          {children}
        </div>
      )}
      <div className="ribbon__group-label" aria-hidden="true">{label}</div>
    </div>
  );
}

export function BigToolbarButton({
  id,
  label,
  icon,
  disabled,
  hasChevron,
  onClick,
}: {
  id: string;
  label: string;
  icon: string;
  disabled?: boolean;
  hasChevron?: boolean;
  onClick?: () => void;
}) {
  const split = splitShortcut(label);
  return (
    <Tooltip label={split.label} shortcut={split.shortcut}>
      <button
        type="button"
        className="btn btn--big"
        data-testid={`ribbon-btn-${id}`}
        aria-label={label}
        disabled={disabled}
        onClick={onClick}
      >
        <Icon name={icon} size={20} />
        <span className="btn__big-label">{split.label}</span>
        {hasChevron && (
          <Icon
            name="arrow_drop_down"
            size={16}
            style={{ fontSize: 14, color: 'var(--cs-ribbon-text-secondary)', lineHeight: 1 }}
          />
        )}
      </button>
    </Tooltip>
  );
}

export function RibbonRow({ children }: { children: ReactNode }) {
  return <div className="ribbon__row">{children}</div>;
}

type ToolbarButtonProps = {
  id: string;
  label: string;
  icon: string;
  pressed?: boolean;
  disabled?: boolean;
  onClick?: () => void;
};

function splitShortcut(label: string): { label: string; shortcut?: string } {
  const m = label.match(/^(.*?)\s*[（(]((?:Ctrl|Cmd|Alt|Shift|F\d{1,2}|Enter|Tab|Esc)[^）)]*)[）)]\s*$/);
  if (!m) return { label };
  return { label: m[1], shortcut: m[2] };
}

export function ToolbarButton({
  id,
  label,
  icon,
  pressed,
  disabled,
  onClick,
}: ToolbarButtonProps) {
  const split = splitShortcut(label);
  return (
    <Tooltip label={split.label} shortcut={split.shortcut}>
      <button
        type="button"
        className="btn btn--icon"
        data-testid={`ribbon-btn-${id}`}
        aria-label={label}
        aria-pressed={pressed}
        disabled={disabled}
        onClick={onClick}
      >
        <Icon name={icon} size={16} />
      </button>
    </Tooltip>
  );
}

type SelectProps = {
  id: string;
  label: string;
  value: string;
  options: { value: string; label: string }[];
  width?: number;
  disabled?: boolean;
  onChange: (v: string) => void;
};

export function ToolbarSelect({
  id,
  label,
  value,
  options,
  width = 110,
  disabled,
  onChange,
}: SelectProps) {
  return (
    <Tooltip label={label}>
      <select
        data-testid={`ribbon-select-${id}`}
        aria-label={label}
        disabled={disabled}
        value={value}
        style={{ width }}
        className="ribbon__select"
        onChange={(e: ChangeEvent<HTMLSelectElement>) => onChange(e.target.value)}
      >
        {options.map((o) => (
          <option key={o.value} value={o.value}>
            {o.label}
          </option>
        ))}
      </select>
    </Tooltip>
  );
}

type ColorButtonProps = {
  id: string;
  label: string;
  icon: string;
  value: string;
  defaultColor: string;
  disabled?: boolean;
  onChange: (color: string) => void;
};

type DropdownItem = { id: string; label: string; icon: string };

const MENU_ICON_STYLE: CSSProperties = {
  color: 'var(--cs-ribbon-text-secondary)',
  flexShrink: 0,
};

export function ToolbarDropdown({
  id,
  label,
  icon,
  items,
  disabled,
  onChoose,
  onDefault,
}: {
  id: string;
  label: string;
  icon: string;
  items: DropdownItem[];
  disabled?: boolean;
  onChoose: (itemId: string) => void;
  onDefault?: () => void;
}) {
  const anchorRef = useRef<HTMLDivElement>(null);
  const [open, setOpen] = useState(false);

  return (
    <span
      ref={anchorRef}
      className="btn-split"
      data-testid={`ribbon-dropdown-${id}`}
    >
      <Tooltip label={label}>
        <button
          type="button"
          className="btn btn--icon btn-split__icon"
          data-testid={`ribbon-dropdown-${id}-apply`}
          aria-label={label}
          disabled={disabled}
          onClick={() => {
            onDefault?.();
            if (open) setOpen(false);
          }}
        >
          <Icon name={icon} size={16} />
        </button>
      </Tooltip>
      <button
        type="button"
        className="btn btn--icon btn-split__caret"
        data-testid={`ribbon-dropdown-${id}-caret`}
        aria-haspopup="menu"
        aria-expanded={open}
        aria-label={`${label}选项`}
        disabled={disabled}
        onClick={() => setOpen((v) => !v)}
      >
        <Icon name="arrow_drop_down" size={16} />
      </button>
      {open && (
        <Popover
          anchorRef={anchorRef}
          onClose={() => setOpen(false)}
          data-testid={`ribbon-dropdown-${id}-popover`}
        >
          {items.map((item) => (
            <button
              key={item.id}
              type="button"
              className="menu__item"
              role="menuitem"
              data-testid={`ribbon-dropdown-${id}-item-${item.id}`}
              onClick={() => {
                onChoose(item.id);
                setOpen(false);
              }}
            >
              <Icon name={item.icon} size={16} style={MENU_ICON_STYLE} />
              <span>{item.label}</span>
            </button>
          ))}
        </Popover>
      )}
    </span>
  );
}

const BORDER_COLOR_PRESETS = [
  '#000000', '#666666', '#9aa0a6', '#d2d6dc',
  '#d93025', '#e8710a', '#f9ab00', '#188038',
  '#1a73e8', '#7627bb', '#a142f4', '#e91e63',
];

export type BorderWeightChoice = 'thin' | 'medium' | 'thick';

const BORDER_WEIGHT_OPTIONS: Array<{ id: BorderWeightChoice; label: string; px: number }> = [
  { id: 'thin', label: '细', px: 1 },
  { id: 'medium', label: '中', px: 2 },
  { id: 'thick', label: '粗', px: 3 },
];

export function BordersControl({
  label,
  icon,
  items,
  disabled,
  defaultColor,
  defaultWeight = 'thin',
  onChoose,
  onDefault,
}: {
  label: string;
  icon: string;
  items: DropdownItem[];
  disabled?: boolean;
  defaultColor: string;
  defaultWeight?: BorderWeightChoice;
  onChoose: (itemId: string, color: string, weight: BorderWeightChoice) => void;
  onDefault?: (color: string, weight: BorderWeightChoice) => void;
}) {
  const anchorRef = useRef<HTMLDivElement>(null);
  const [open, setOpen] = useState(false);
  const [color, setColor] = useState(defaultColor);
  const [weight, setWeight] = useState<BorderWeightChoice>(defaultWeight);

  return (
    <span
      ref={anchorRef}
      className="btn-split"
      data-testid="ribbon-dropdown-borders"
    >
      <Tooltip label={`${label}（${weight}，${color}）`}>
        <button
          type="button"
          className="btn btn--icon btn-split__icon btn-color__icon"
          data-testid="ribbon-dropdown-borders-apply"
          aria-label={label}
          disabled={disabled}
          onClick={() => {
            onDefault?.(color, weight);
            if (open) setOpen(false);
          }}
        >
          <Icon name={icon} size={16} />
          <span
            className="btn-color__swatch"
            style={{ background: color }}
            aria-hidden="true"
          />
        </button>
      </Tooltip>
      <button
        type="button"
        className="btn btn--icon btn-split__caret"
        data-testid="ribbon-dropdown-borders-caret"
        aria-haspopup="menu"
        aria-expanded={open}
        aria-label={`${label}选项`}
        disabled={disabled}
        onClick={() => setOpen((v) => !v)}
      >
        <Icon name="arrow_drop_down" size={16} />
      </button>
      {open && (
        <Popover
          anchorRef={anchorRef}
          onClose={() => setOpen(false)}
          data-testid="ribbon-dropdown-borders-popover"
        >
          {items.map((item) => (
            <button
              key={item.id}
              type="button"
              className="menu__item"
              role="menuitem"
              data-testid={`ribbon-dropdown-borders-item-${item.id}`}
              onClick={() => {
                onChoose(item.id, color, weight);
                setOpen(false);
              }}
            >
              <Icon name={item.icon} size={16} style={MENU_ICON_STYLE} />
              <span>{item.label}</span>
            </button>
          ))}
          <div className="menu__divider" />
          <div className="menu__color-row" data-testid="ribbon-dropdown-borders-weight-row">
            <div className="menu__color-label">线条粗细</div>
            <div className="menu__color-swatches" role="group" aria-label="边框线条粗细">
              {BORDER_WEIGHT_OPTIONS.map((opt) => (
                <button
                  key={opt.id}
                  type="button"
                  className={`menu__weight-swatch${weight === opt.id ? ' menu__weight-swatch--active' : ''}`}
                  data-testid={`ribbon-dropdown-borders-weight-${opt.id}`}
                  aria-label={`将边框粗细设置为${opt.label}`}
                  aria-pressed={weight === opt.id}
                  title={opt.label}
                  onClick={() => setWeight(opt.id)}
                >
                  <span
                    className="menu__weight-swatch-line"
                    style={{ background: color, height: `${opt.px}px` }}
                    aria-hidden="true"
                  />
                </button>
              ))}
            </div>
          </div>
          <div className="menu__divider" />
          <div className="menu__color-row" data-testid="ribbon-dropdown-borders-color-row">
            <div className="menu__color-label">线条颜色</div>
            <div className="menu__color-swatches" role="group" aria-label="边框线条颜色">
              {BORDER_COLOR_PRESETS.map((preset) => (
                <button
                  key={preset}
                  type="button"
                  className={`menu__color-swatch${color === preset ? ' menu__color-swatch--active' : ''}`}
                  style={{ background: preset }}
                  data-testid={`ribbon-dropdown-borders-color-${preset.slice(1)}`}
                  aria-label={`将边框颜色设置为 ${preset}`}
                  aria-pressed={color === preset}
                  onClick={() => setColor(preset)}
                />
              ))}
              <label
                className="menu__color-swatch menu__color-swatch--custom"
                title="自定义颜色"
              >
                <Icon name="palette" size={16} />
                <input
                  type="color"
                  value={color}
                  data-testid="ribbon-dropdown-borders-color-custom"
                  onChange={(e) => setColor(e.target.value)}
                />
              </label>
            </div>
          </div>
        </Popover>
      )}
    </span>
  );
}

export function ToolbarColorButton({
  id,
  label,
  icon,
  value,
  defaultColor,
  disabled,
  onChange,
}: ColorButtonProps) {
  const applied = value || defaultColor;
  return (
    <span className="btn-color" data-testid={`ribbon-color-${id}`}>
      <Tooltip label={label}>
        <button
          type="button"
          className="btn btn--icon btn-color__icon"
          data-testid={`ribbon-color-${id}-apply`}
          aria-label={label}
          disabled={disabled}
          onClick={() => onChange(applied)}
        >
          <Icon name={icon === 'format_color_fill' ? 'ink_highlighter' : icon} size={16} />
          <span className="btn-color__swatch" style={{ background: applied }} aria-hidden="true" />
        </button>
      </Tooltip>
      <label className="btn-color__picker" title={`${label} — 选择颜色`}>
        <Icon name="arrow_drop_down" size={16} />
        <input
          type="color"
          data-testid={`ribbon-color-${id}-input`}
          value={applied || '#000000'}
          disabled={disabled}
          onChange={(e: ChangeEvent<HTMLInputElement>) => onChange(e.target.value)}
          tabIndex={-1}
        />
      </label>
    </span>
  );
}
