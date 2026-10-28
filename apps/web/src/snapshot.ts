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

import { LocaleType, type IWorkbookData } from '@univerjs/core';
import appPkg from '../package.json';

const UNIVER_DEP = (appPkg.dependencies as Record<string, string>)['@univerjs/core'];
export const UNIVER_VERSION = UNIVER_DEP.replace(/^[~^]/, '');

export const INITIAL_ROWS = 1024;
export const INITIAL_COLUMNS = 26;

export function emptyWorkbook(): IWorkbookData {
  const nowIso = new Date().toISOString();
  return {
    id: `wb-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`,
    rev: 1,
    name: '未命名工作簿',
    appVersion: UNIVER_VERSION,
    locale: LocaleType.ZH_CN,
    styles: {},
    sheetOrder: ['sheet-1'],
    sheets: {
      'sheet-1': {
        id: 'sheet-1',
        name: '工作表1',
        cellData: {},
        rowCount: INITIAL_ROWS,
        columnCount: INITIAL_COLUMNS,
      },
    },
    custom: {
      properties: { createdAt: nowIso, modifiedAt: nowIso },
    },
  };
}
