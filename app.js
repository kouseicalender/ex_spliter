document.getElementById('excel-input').addEventListener('change', handleExcelUpload);

function handleExcelUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // ヘッダーを探して列位置を特定
    const header = rows[0];
    const colIndex = {
      month: header.indexOf("月"),
      day: header.indexOf("日"),
      jp: header.indexOf("（下）１行目"),
      en: header.indexOf("（下）２行目"),
      source: header.indexOf("（下）３行目")
    };

    const result = rows.slice(1).map(row => {
      const month = row[colIndex.month];
      const day = row[colIndex.day];
      const jp = row[colIndex.jp];
      const en = row[colIndex.en];
      const source = row[colIndex.source];
      if (!month || !day || !jp || !en || !source) return null;
      return {
        date: `${month}月${day}日`,
        jp: jp.toString().trim(),
        en: en.toString().trim(),
        source: source.toString().trim()
      };
    }).filter(x => x);

    displayTable(result);
  };

  reader.readAsArrayBuffer(file);
}

function displayTable(data) {
  const output = document.getElementById('output');
  output.innerHTML = '';

  const table = document.createElement('table');
  const header = table.insertRow();
  ['日付', '日本語', '英語', '出典'].forEach(text => {
    const th = document.createElement('th');
    th.textContent = text;
    header.appendChild(th);
  });

  data.forEach(row => {
    const tr = document.createElement('tr');
    [row.date, row.jp, row.en, row.source].forEach(text => {
      const td = document.createElement('td');
      td.textContent = text;
      tr.appendChild(td);
    });
    table.appendChild(tr);
  });

  output.appendChild(table);
}
