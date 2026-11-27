"use strict";

{
  const workbook = XLSX.utils.book_new();
  const rows = [
    ["範例線材加工表"],
    ["機型:示範設備", "料號:DEMO-001", "版次:1"],
    ["NO", "線名", "前接頭", "剝線長", "線材代號／規格", "長度", "轉接頭", "管徑", "長度", "轉接頭", "線名", "剝線長", "後接頭", "備註"],
    ["1", "馬達電源線", "接頭 A", "20 mm", "DEMO-CABLE-4C", "1.20 m", "", "", "", "", "馬達端子", "25 mm", "端子 B", "示範標記 A1～A4"],
    ["2", "感測器訊號線", "插頭 C", "12 mm", "DEMO-SIGNAL-3C", "850 mm", "", "", "", "", "訊號端子", "15 mm", "端子 D", "示範標記 S1～S3"],
  ];
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), "範例線表");
  window.demoWorkbook = {
    name: "ProjectTrace 範例線表.xlsx",
    base64: XLSX.write(workbook, { bookType: "xlsx", type: "base64" }),
  };
}
