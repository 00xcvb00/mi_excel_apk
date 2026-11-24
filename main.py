import flet_webview as fwv
import flet as ft
import base64
import json
import os

HTML_APP = """<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
<link rel='stylesheet' href='https://jsdelivr.net' />
<link rel='stylesheet' href='https://jsdelivr.net' />
<link rel='stylesheet' href='https://jsdelivr.net' />
<link rel='stylesheet' href='https://jsdelivr.net' />
<script src="https://jsdelivr.net"></script>
<script src="https://jsdelivr.net"></script>
<script src="https://jsdelivr.net"></script>
<script src="https://jsdelivr.net"></script>
<style>
  html, body { margin: 0; padding: 0; width: 100%; height: 100%; overflow: hidden; font-family: sans-serif; background: #f3f3f3; }
  #topbar { height: 42px; background: #1e7145; color: white; display: flex; align-items: center; gap: 6px; padding: 0 8px; overflow-x: auto; }
  #topbar button, #topbar label { background: #2b8a57; color: white; border: 1px solid #45a873; border-radius: 4px; padding: 5px 10px; font-size: 12px; font-weight: bold; white-space: nowrap; cursor: pointer; }
  #txtStatus { font-size: 11px; color: #d4f7e2; margin-left: auto; white-space: nowrap; }
  #luckysheet { position: absolute; width: 100%; top: 42px; bottom: 0; left: 0; }
</style>
</head>
<body>
<div id="topbar">
  <label>Abrir Excel<input type="file" accept=".xlsx,.xls" style="display:none" onchange="loadExcel(event)"></label>
  <label>Cargar TXT<input type="file" accept=".txt,.csv" style="display:none" onchange="loadTxt(event)"></label>
  <button onclick="exportXlsx()">Guardar XLSX</button>
  <button onclick="exportPdf()">Exportar PDF</button>
  <span id="txtStatus">Excel Desktop (Sin Macros)</span>
</div>
<div id="luckysheet"></div>

<script>
  let txtLines = [];
  let txtIdx = 0;
  let currentFileName = "Documento_Editado.xlsx";

  function createGrid(sheetsData) {
    luckysheet.destroy();
    luckysheet.create({
      container: 'luckysheet',
      lang: 'es',
      showinfobar: false,
      showtoolbar: true,
      showsheetbar: true,
      showstatisticBar: true,
      allowEdit: true,
      data: sheetsData || [{ name: "Hoja1", celldata: [], row: 60, column: 20 }],
      hook: {
        cellMousedown: function(cell, pos) {
          if (txtLines.length > 0 && txtIdx < txtLines.length) {
            luckysheet.setCellValue(pos.r, pos.c, txtLines[txtIdx]);
            txtIdx++;
            document.getElementById('txtStatus').innerText = "Quedan " + (txtLines.length - txtIdx) + " líneas";
          }
        }
      }
    });
  }

  function loadExcel(e) {
    const file = e.target.files;
    if (!file) return;
    currentFileName = "Editado_" + file.name;
    const reader = new FileReader();
    reader.onload = function(evt) {
      const wb = XLSX.read(evt.target.result, { type: 'binary', cellStyles: true });
      const sheets = wb.SheetNames.map((name, i) => {
        const ws = wb.Sheets[name];
        const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        let celldata = [];
        for (let r = 0; r < json.length; r++) {
          for (let c = 0; c < json[r].length; c++) {
            if (json[r][c] !== "") {
              let val = json[r][c];
              let obj = { v: val, m: String(val) };
              if (typeof val === "string" && val.startsWith("=")) obj.f = val;
              celldata.push({ r: r, c: c, v: obj });
            }
          }
        }
        return { name: name, status: i === 0 ? 1 : 0, order: i, celldata: celldata, row: Math.max(json.length + 15, 60), column: 20 };
      });
      createGrid(sheets);
      document.getElementById('txtStatus').innerText = file.name;
    };
    reader.readAsBinaryString(file);
  }

  function loadTxt(e) {
    const file = e.target.files;
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function(evt) {
      txtLines = evt.target.result.split(/\\r?\\n/).map(l => l.trim()).filter(l => l.length > 0);
      txtIdx = 0;
      document.getElementById('txtStatus').innerText = txtLines.length + " líneas listas (toca celdas)";
    };
    reader.readAsText(file);
  }

  function exportXlsx() {
    const allSheets = luckysheet.getAllSheets();
    const wb = XLSX.utils.book_new();
    allSheets.forEach(s => {
      const aoa = (s.data || []).map(row => (row || []).map(cell => cell ? (cell.f || cell.v || "") : ""));
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      XLSX.utils.book_append_sheet(wb, ws, s.name.substring(0, 31));
    });
    XLSX.writeFile(wb, currentFileName);
  }

  function exportPdf() {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: "landscape" });
    const sheet = luckysheet.getSheetData();
    let y = 15;
    doc.setFontSize(10);
    for (let r = 0; r < Math.min(sheet.length, 40); r++) {
      let rowStr = (sheet[r] || []).slice(0, 10).map(c => c ? (c.m || c.v || "") : "").join(" | ");
      if (rowStr.replace(/\\|/g, "").trim() !== "") {
        doc.text(rowStr.substring(0, 130), 10, y);
        y += 7;
        if (y > 190) { doc.addPage(); y = 15; }
      }
    }
    doc.save(currentFileName.replace(/\\.xlsx?$/i, "") + ".pdf");
  }

  createGrid();
</script>
</body>
</html>"""

def main(page: ft.Page):
    page.title = "Excel Desktop Editor"
    page.padding = 0
    page.spacing = 0
    wv = fwv.WebView(
        url="/index.html",
        expand=True,
        enable_javascript=True,
    )
    page.add(wv)

ft.run(main)
