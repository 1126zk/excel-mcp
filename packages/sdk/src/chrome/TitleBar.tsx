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

import { useEffect, useState, type CSSProperties, type ReactNode } from 'react';
import { ICommandService } from '@univerjs/core';
import type { CasualSheetsAPI } from '../sheets/api';
import { Icon } from './Icon';
import { ensureChromeFonts } from './fonts';

const BAR_STYLE: CSSProperties = {
  display: 'flex',
  alignItems: 'stretch',
  gap: 12,
  height: 64,
  minHeight: 64,
  padding: '0 16px',
  boxSizing: 'border-box',
  background: 'var(--cs-chrome-bg, #eef1f5)',
  color: 'var(--cs-chrome-fg, #201f1e)',
  flex: '0 0 auto',
  userSelect: 'none',
  letterSpacing: 0,
  zIndex: 4,
};

const BRAND_STYLE: CSSProperties = {
  display: 'inline-flex',
  alignItems: 'center',
  flex: '0 0 auto',
};

const CENTER_STYLE: CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  justifyContent: 'center',
  flex: '1 1 auto',
  minWidth: 0,
  padding: '4px 0',
  gap: 1,
};

const CENTER_TOP_STYLE: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  minWidth: 0,
  height: 26,
  fontSize: 14,
  lineHeight: 1,
};

const CENTER_BOTTOM_STYLE: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  minWidth: 0,
  height: 28,
  fontSize: 13,
  lineHeight: 1,
};

const NAME_STYLE: CSSProperties = {
  display: 'inline-flex',
  alignItems: 'center',
  maxWidth: '100%',
  height: 22,
  marginLeft: -6,
  padding: '0 6px',
  overflow: 'hidden',
  textOverflow: 'ellipsis',
  whiteSpace: 'nowrap',
  borderRadius: 6,
  fontSize: 14,
  fontWeight: 500,
  lineHeight: 1,
};

const ACTIONS_STYLE: CSSProperties = {
  display: 'inline-flex',
  alignItems: 'center',
  flex: '0 0 auto',
  gap: 10,
};

const THEME_BUTTON_STYLE: CSSProperties = {
  display: 'inline-flex',
  alignItems: 'center',
  justifyContent: 'center',
  width: 32,
  height: 32,
  padding: 0,
  flex: '0 0 auto',
  border: '1px solid transparent',
  borderRadius: 6,
  background: 'transparent',
  color: 'var(--cs-chrome-fg, #201f1e)',
  cursor: 'pointer',
};

const MENU_LAYOUT_CSS = `
[data-testid="casual-sheets-title-bar"] .titlebar__center-bottom > [data-testid="cs-menubar"] {
  height: 100% !important;
  margin: 0 !important;
  padding: 0 !important;
  flex: 1 1 auto !important;
  gap: 4px !important;
  border-bottom: 0 !important;
  background: transparent !important;
}
[data-testid="casual-sheets-title-bar"] .titlebar__center-bottom [data-menu] {
  height: 24px !important;
  margin: 0 !important;
  padding: 0 6px !important;
  font-size: 13px !important;
  line-height: 1 !important;
}
[data-testid="casual-sheets-title-bar"] .titlebar__center-bottom > [data-testid="cs-menubar"] > div:first-child > [data-menu] {
  margin-left: -6px !important;
}
[data-testid="casual-sheets-title-bar"] .titlebar__filename:hover {
  background: var(--cs-chrome-hover, rgba(0,0,0,0.06));
}
`;

export interface TitleBarProps {
  api: CasualSheetsAPI | null;
  appearance: 'light' | 'dark';
  onToggleAppearance: () => void;
  menu?: ReactNode;
}

function readWorkbookName(api: CasualSheetsAPI | null): string {
  if (!api) return '未命名工作簿';
  const activeName = api.univer.getActiveWorkbook()?.getName().trim();
  return activeName || api.getContent()?.name?.trim() || '未命名工作簿';
}

export function TitleBar({ api, appearance, onToggleAppearance, menu }: TitleBarProps) {
  const [workbookName, setWorkbookName] = useState(() => readWorkbookName(api));
  const themeLabel = appearance === 'dark' ? '切换到浅色主题' : '切换到深色主题';

  useEffect(() => {
    ensureChromeFonts();
  }, []);

  useEffect(() => {
    if (!api) {
      setWorkbookName('未命名工作簿');
      return;
    }
    const refresh = () => setWorkbookName(readWorkbookName(api));
    refresh();
    const injector = (api.univer as unknown as { _injector?: { get(t: unknown): unknown } })
      ._injector;
    const cmd = injector?.get(ICommandService) as
      | { onCommandExecuted: (cb: () => void) => { dispose: () => void } }
      | undefined;
    const sub = cmd?.onCommandExecuted(refresh);
    return () => sub?.dispose();
  }, [api]);

  return (
    <header
      className="titlebar"
      style={BAR_STYLE}
      data-testid="casual-sheets-title-bar"
      role="banner"
    >
      <style>{MENU_LAYOUT_CSS}</style>
      <span className="titlebar__brand" style={BRAND_STYLE} title="Casual Sheets">
        <svg
          className="titlebar__brand-icon"
          width={28}
          height={36}
          viewBox="0 0 48 48"
          fill="none"
          xmlns="http://www.w3.org/2000/svg"
          role="img"
          aria-label="Casual Sheets"
        >
          <defs>
            <linearGradient
              id="cs-titlebar-logo-gradient"
              x1="6"
              y1="4"
              x2="42"
              y2="44"
              gradientUnits="userSpaceOnUse"
            >
              <stop stopColor="#1aa3c4" />
              <stop offset="1" stopColor="#0a5266" />
            </linearGradient>
          </defs>
          <rect x="3" y="3" width="42" height="42" rx="12" fill="url(#cs-titlebar-logo-gradient)" />
          <path
            d="M3 15C3 8.373 8.373 3 15 3H33C39.627 3 45 8.373 45 15V18H3V15Z"
            fill="#ffffff"
            fillOpacity="0.12"
          />
          <rect x="13" y="13" width="10" height="10" rx="2.5" fill="#ffffff" fillOpacity="0.5" />
          <rect x="25" y="13" width="10" height="10" rx="2.5" fill="#ffffff" fillOpacity="0.5" />
          <rect x="13" y="25" width="10" height="10" rx="2.5" fill="#ffffff" fillOpacity="0.4" />
          <rect x="25" y="25" width="10" height="10" rx="2.5" fill="#ffffff" />
        </svg>
      </span>
      <div className="titlebar__center" style={CENTER_STYLE}>
        <div className="titlebar__center-top" style={CENTER_TOP_STYLE}>
          <span
            className="titlebar__filename"
            style={NAME_STYLE}
            title={workbookName}
            data-testid="cs-workbook-name"
          >
            {workbookName}
          </span>
        </div>
        <div className="titlebar__center-bottom" style={CENTER_BOTTOM_STYLE}>
          {menu}
        </div>
      </div>
      <div className="titlebar__actions" style={ACTIONS_STYLE}>
        <button
          type="button"
          title={themeLabel}
          aria-label={themeLabel}
          data-action="toggle-appearance"
          data-testid="cs-title-toggle-appearance"
          style={THEME_BUTTON_STYLE}
          onMouseDown={(e) => e.preventDefault()}
          onClick={onToggleAppearance}
          onMouseEnter={(e) => {
            e.currentTarget.style.background = 'var(--cs-chrome-hover, rgba(0,0,0,0.06))';
          }}
          onMouseLeave={(e) => {
            e.currentTarget.style.background = 'transparent';
          }}
        >
          <Icon name={appearance === 'dark' ? 'light_mode' : 'dark_mode'} size={20} />
        </button>
      </div>
    </header>
  );
}
