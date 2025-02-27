// ① 対象のパネルから複数件のデータを抽出する関数
function extractDataFromPanels() {
  // 売上データのパネルを全て取得
  const panels = document.querySelectorAll('.co-expansion-panel.js-accordion-content.u-pt-300.u-pb-300');
  const results = [];
  
  panels.forEach(panel => {
    const rows = panel.querySelectorAll('.js-accordion-body .lo-grid.co-breakdown-table-row');
    const data = {};
    
    rows.forEach(row => {
      const cells = row.querySelectorAll('.lo-grid-cell, .lo-u-auto');
      if (cells.length >= 2) {
        const label = cells[0].textContent.trim();
        let value = cells[1].textContent.trim();
        
        if (label.includes("注文番号")) {
          const a = cells[1].querySelector('a');
          data.orderNumber = a ? a.textContent.trim() : value;
        } else if (label.includes("注文日時")) {
          data.orderDate = value;
        } else if (label.includes("小計")) {
          data.subtotal = value.replace(/¥|\s|,/g, '');
        } else if (label.includes("手数料")) {
          data.fee = value.replace(/¥|\s|,/g, '');
        }
      }
    });
    
    // 少なくとも注文番号と注文日時があれば有効なデータと判断
    if (data.orderNumber && data.orderDate) {
      results.push(data);
    }
  });
  return results;
}

// ② 日本語の日付（例："2022年10月29日 04時55分"）を "YYYY/MM/DD" に変換する関数
function formatDate(jpDateStr) {
  const match = jpDateStr.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
  if (match) {
    const year = match[1];
    const month = match[2].padStart(2, '0');
    const day = match[3].padStart(2, '0');
    return `${year}/${month}/${day}`;
  }
  return jpDateStr;
}

// ③ 複数件のデータをまとめたExcelファイルを生成してダウンロードする関数
function generateExcel() {
  const dataArray = extractDataFromPanels();
  if (!dataArray || dataArray.length === 0) {
    alert("抽出できる売上データが見つかりませんでした。");
    return;
  }
  
  const rows = [];
  // ヘッダー行
  rows.push(["収支区分", "発生日", "勘定科目", "税区分", "金額", "取引先", "備考"]);
  
  // 各データごとに収入行と手数料行を追加
  dataArray.forEach(data => {
    const date = formatDate(data.orderDate);
    rows.push([
      "収入",
      date,
      "売上高",
      "課対仕入10%",
      data.subtotal || "",
      "Booth",
      data.orderNumber || ""
    ]);
    rows.push([
      "",
      "",
      "支払手数料",
      "課対仕入10%",
      data.fee ? "-" + data.fee : "",
      "Booth",
      ""
    ]);
  });
  
  // SheetJSでワークブックとワークシートを生成
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  
  // ワークブックをExcelファイル用のバイナリ配列に変換
  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([wbout], { type: "application/octet-stream" });
  const url = URL.createObjectURL(blob);
  
  // ダウンロード用のリンクを作成し自動クリック
  const a = document.createElement("a");
  a.download = document.title + ".xlsx"; // ページタイトルをファイル名に利用
  a.href = url;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ④ ページ上にExcelエクスポート用のボタンを追加する関数（左下に配置）
function addExportButton() {
  const btn = document.createElement("button");
  btn.textContent = "Excelエクスポート";
  btn.style.position = "fixed";
  btn.style.bottom = "20px";
  btn.style.left = "20px";
  btn.style.zIndex = "10000";
  btn.style.padding = "10px 20px";
  btn.style.backgroundColor = "#007bff";
  btn.style.color = "#fff";
  btn.style.border = "none";
  btn.style.borderRadius = "4px";
  btn.style.cursor = "pointer";
  btn.addEventListener("click", generateExcel);
  document.body.appendChild(btn);
}

// ⑤ 現在のURLが /sales/{年}/{月} の形式（末尾スラッシュはオプション）の場合のみボタンを追加
if (window.location.pathname.match(/^\/sales\/\d{4}\/\d{1,2}\/?$/)) {
  addExportButton();
}
