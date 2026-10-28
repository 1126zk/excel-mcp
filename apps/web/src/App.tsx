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

import { useState } from 'react';
import { CasualSheets } from '@casualoffice/sheets/sheets';
import { emptyWorkbook } from './snapshot';

export function App() {
  const [initialData] = useState(emptyWorkbook);
  const [error, setError] = useState<string | null>(null);

  return (
    <div style={{ position: 'fixed', inset: 0 }}>
      {error && (
        <pre
          role="alert"
          style={{
            position: 'absolute',
            inset: 0,
            zIndex: 10000,
            margin: 0,
            padding: 24,
            overflow: 'auto',
            color: '#b00020',
            background: '#fff',
            whiteSpace: 'pre-wrap',
          }}
        >
          {error}
        </pre>
      )}
      <CasualSheets
        initialData={initialData}
        chrome="full"
        onReady={(api) => console.info('Casual Sheets ready', api)}
        onChange={(snapshot) => console.debug('Workbook changed (memory only)', snapshot)}
        onSave={(snapshot) => console.info('Save requested (memory only)', snapshot)}
        onError={(cause) => {
          console.error('Casual Sheets failed', cause);
          setError(cause.stack ?? cause.message);
        }}
      />
    </div>
  );
}
