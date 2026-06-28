import { SafeXmlNode } from '../../parser/XmlParser';
import { formatValue } from './format';
import { parseOoxmlBoolValue } from './ooxml';
import type { DataTableInfo } from './types';

export function parseDataTable(plotArea: SafeXmlNode): { showKeys: boolean } | undefined {
  const dTable = plotArea.child('dTable');
  if (!dTable.exists()) return undefined;
  const showKeys = parseOoxmlBoolValue(dTable.child('showKeys').attr('val'), true);
  return { showKeys };
}

export function buildDataTableElement(
  info: DataTableInfo,
  seriesColors?: string[],
): HTMLTableElement {
  const table = document.createElement('table');
  table.style.width = '100%';
  table.style.borderCollapse = 'collapse';
  table.style.fontSize = '10px';
  table.style.marginTop = '8px';

  const { seriesArr, showKeys, formatCode } = info;
  const categories = seriesArr.find((s) => s.categories.length > 0)?.categories || [];
  const fc = formatCode;

  const thead = document.createElement('thead');
  const headerRow = document.createElement('tr');
  const emptyTh = document.createElement('th');
  emptyTh.style.border = '1px solid #ccc';
  emptyTh.style.padding = '2px 6px';
  emptyTh.style.textAlign = 'left';
  emptyTh.style.fontWeight = 'bold';
  headerRow.appendChild(emptyTh);
  for (let i = 0; i < categories.length; i++) {
    const th = document.createElement('th');
    th.style.border = '1px solid #ccc';
    th.style.padding = '2px 6px';
    th.style.textAlign = 'right';
    th.style.fontWeight = 'bold';
    th.textContent = categories[i] ?? '';
    headerRow.appendChild(th);
  }
  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  for (let si = 0; si < seriesArr.length; si++) {
    const s = seriesArr[si];
    const tr = document.createElement('tr');
    const nameTd = document.createElement('td');
    nameTd.style.border = '1px solid #ccc';
    nameTd.style.padding = '2px 6px';
    nameTd.style.textAlign = 'left';
    nameTd.style.fontWeight = 'bold';
    if (showKeys && seriesColors && seriesColors[si]) {
      const key = document.createElement('span');
      key.style.display = 'inline-block';
      key.style.width = '8px';
      key.style.height = '8px';
      key.style.marginRight = '4px';
      key.style.verticalAlign = 'middle';
      key.style.backgroundColor = seriesColors[si];
      nameTd.appendChild(key);
    }
    nameTd.appendChild(document.createTextNode(s.name || ''));
    tr.appendChild(nameTd);
    for (let ci = 0; ci < categories.length; ci++) {
      const td = document.createElement('td');
      td.style.border = '1px solid #ccc';
      td.style.padding = '2px 6px';
      td.style.textAlign = 'right';
      const val = s.values[ci];
      td.textContent =
        val !== undefined && !s.blankIndices?.has(ci) ? formatValue(val, fc ?? s.formatCode) : '';
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);

  return table;
}
