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
 * Default zh-CN and en-US locale bundles for the React and in-iframe runtimes.
 *
 * The FULL feature set (`lazyPlugins={true}`) includes the base editor
 * (render + formula engines, UI, docs/docs-ui, sheets, sheets-ui,
 * sheets-formula-ui, numfmt) PLUS the lazily-registered feature plugins
 * (tables, sort, filter, conditional formatting, data validation, drawing,
 * hyperlinks, notes, thread comments, find/replace). Univer's `LocaleService`
 * resolves every UI string against this map; without a plugin's strings its UI
 * renders raw i18n keys (e.g. the comment panel showed
 * `thread-comment-ui.editor.reply` instead of "Reply"). So we MUST seed strings
 * for every plugin the runtime can register — matching the reference app's
 * `apps/web/src/locale.ts` set.
 *
 * Hosts can replace these bundles through the `locales` prop.
 */

import { LocaleType, Tools } from '@univerjs/core';

import UniverUIEnUS from '@univerjs/ui/locale/en-US';
import UniverDocsUIEnUS from '@univerjs/docs-ui/locale/en-US';
import UniverSheetsEnUS from '@univerjs/sheets/locale/en-US';
import UniverSheetsUIEnUS from '@univerjs/sheets-ui/locale/en-US';
import UniverSheetsFormulaUIEnUS from '@univerjs/sheets-formula-ui/locale/en-US';
import UniverSheetsNumfmtUIEnUS from '@univerjs/sheets-numfmt-ui/locale/en-US';
import UniverUIZhCN from '@univerjs/ui/locale/zh-CN';
import UniverDocsUIZhCN from '@univerjs/docs-ui/locale/zh-CN';
import UniverSheetsZhCN from '@univerjs/sheets/locale/zh-CN';
import UniverSheetsUIZhCN from '@univerjs/sheets-ui/locale/zh-CN';
import UniverSheetsFormulaUIZhCN from '@univerjs/sheets-formula-ui/locale/zh-CN';
import UniverSheetsNumfmtUIZhCN from '@univerjs/sheets-numfmt-ui/locale/zh-CN';
// Feature-plugin strings — these load lazily, so without their locales the
// feature UIs (comment panel, table dialog, filter, sort, CF, data validation,
// images, hyperlinks, notes, find/replace) render raw i18n keys.
import UniverSheetsTableUIEnUS from '@univerjs/sheets-table-ui/locale/en-US';
import UniverSheetsSortUIEnUS from '@univerjs/sheets-sort-ui/locale/en-US';
import UniverSheetsFilterUIEnUS from '@univerjs/sheets-filter-ui/locale/en-US';
import UniverSheetsCfUIEnUS from '@univerjs/sheets-conditional-formatting-ui/locale/en-US';
import UniverSheetsTableUIZhCN from '@univerjs/sheets-table-ui/locale/zh-CN';
import UniverSheetsSortUIZhCN from '@univerjs/sheets-sort-ui/locale/zh-CN';
import UniverSheetsFilterUIZhCN from '@univerjs/sheets-filter-ui/locale/zh-CN';
import UniverSheetsCfUIZhCN from '@univerjs/sheets-conditional-formatting-ui/locale/zh-CN';
// The DV *UI* package ships only the panel chrome (`sheets-data-validation-ui.*`).
// Validator titles, operator names, rule names and error messages live in the
// base `@univerjs/data-validation` + `@univerjs/sheets-data-validation` packages;
// without them the DV Type/Operator selectors and cell errors render raw i18n keys.
import UniverDvEnUS from '@univerjs/data-validation/locale/en-US';
import UniverSheetsDvEnUS from '@univerjs/sheets-data-validation/locale/en-US';
import UniverSheetsDvUIEnUS from '@univerjs/sheets-data-validation-ui/locale/en-US';
import UniverDvZhCN from '@univerjs/data-validation/locale/zh-CN';
import UniverSheetsDvZhCN from '@univerjs/sheets-data-validation/locale/zh-CN';
import UniverSheetsDvUIZhCN from '@univerjs/sheets-data-validation-ui/locale/zh-CN';
// Same base-vs-UI split: the `-ui` packages above ship only panel chrome, while
// error toasts + generated labels (filter range errors, table name/column
// prefixes + validation, hyperlink ref errors) live in the base packages'
// namespaces. Without the base locales they render raw i18n keys.
import UniverSheetsFilterEnUS from '@univerjs/sheets-filter/locale/en-US';
import UniverSheetsHyperLinkEnUS from '@univerjs/sheets-hyper-link/locale/en-US';
import UniverSheetsTableEnUS from '@univerjs/sheets-table/locale/en-US';
import UniverSheetsDrawingUIEnUS from '@univerjs/sheets-drawing-ui/locale/en-US';
import UniverDrawingUIEnUS from '@univerjs/drawing-ui/locale/en-US';
import UniverSheetsHyperLinkUIEnUS from '@univerjs/sheets-hyper-link-ui/locale/en-US';
import UniverSheetsNoteUIEnUS from '@univerjs/sheets-note-ui/locale/en-US';
import UniverThreadCommentUIEnUS from '@univerjs/thread-comment-ui/locale/en-US';
import UniverSheetsThreadCommentUIEnUS from '@univerjs/sheets-thread-comment-ui/locale/en-US';
import UniverFindReplaceEnUS from '@univerjs/find-replace/locale/en-US';
import UniverSheetsFilterZhCN from '@univerjs/sheets-filter/locale/zh-CN';
import UniverSheetsHyperLinkZhCN from '@univerjs/sheets-hyper-link/locale/zh-CN';
import UniverSheetsTableZhCN from '@univerjs/sheets-table/locale/zh-CN';
import UniverSheetsDrawingUIZhCN from '@univerjs/sheets-drawing-ui/locale/zh-CN';
import UniverDrawingUIZhCN from '@univerjs/drawing-ui/locale/zh-CN';
import UniverSheetsHyperLinkUIZhCN from '@univerjs/sheets-hyper-link-ui/locale/zh-CN';
import UniverSheetsNoteUIZhCN from '@univerjs/sheets-note-ui/locale/zh-CN';
import UniverThreadCommentUIZhCN from '@univerjs/thread-comment-ui/locale/zh-CN';
import UniverSheetsThreadCommentUIZhCN from '@univerjs/sheets-thread-comment-ui/locale/zh-CN';
import UniverFindReplaceZhCN from '@univerjs/find-replace/locale/zh-CN';

const enUS = Tools.deepMerge(
  {},
  UniverUIEnUS,
  UniverDocsUIEnUS,
  UniverSheetsEnUS,
  UniverSheetsUIEnUS,
  UniverSheetsFormulaUIEnUS,
  UniverSheetsNumfmtUIEnUS,
  UniverSheetsTableUIEnUS,
  UniverSheetsSortUIEnUS,
  UniverSheetsFilterUIEnUS,
  UniverSheetsCfUIEnUS,
  UniverDvEnUS,
  UniverSheetsDvEnUS,
  UniverSheetsDvUIEnUS,
  UniverSheetsFilterEnUS,
  UniverSheetsHyperLinkEnUS,
  UniverSheetsTableEnUS,
  UniverDrawingUIEnUS,
  UniverSheetsDrawingUIEnUS,
  UniverSheetsHyperLinkUIEnUS,
  UniverSheetsNoteUIEnUS,
  UniverThreadCommentUIEnUS,
  UniverSheetsThreadCommentUIEnUS,
  UniverFindReplaceEnUS,
);

const zhCN = Tools.deepMerge(
  {},
  UniverUIZhCN,
  UniverDocsUIZhCN,
  UniverSheetsZhCN,
  UniverSheetsUIZhCN,
  UniverSheetsFormulaUIZhCN,
  UniverSheetsNumfmtUIZhCN,
  UniverSheetsTableUIZhCN,
  UniverSheetsSortUIZhCN,
  UniverSheetsFilterUIZhCN,
  UniverSheetsCfUIZhCN,
  UniverDvZhCN,
  UniverSheetsDvZhCN,
  UniverSheetsDvUIZhCN,
  UniverSheetsFilterZhCN,
  UniverSheetsHyperLinkZhCN,
  UniverSheetsTableZhCN,
  UniverDrawingUIZhCN,
  UniverSheetsDrawingUIZhCN,
  UniverSheetsHyperLinkUIZhCN,
  UniverSheetsNoteUIZhCN,
  UniverThreadCommentUIZhCN,
  UniverSheetsThreadCommentUIZhCN,
  UniverFindReplaceZhCN,
);

export const EMBED_LOCALES = {
  [LocaleType.ZH_CN]: zhCN,
  [LocaleType.EN_US]: enUS,
};
