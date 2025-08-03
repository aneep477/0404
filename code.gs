function doPost(e) {
  var data = e.parameter;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Laporan");
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Laporan");
    sheet.appendRow(["Nama Pelajar", "No. Angka Giliran", "Catatan Pensyarah", "Tarikh", "Tandatangan"]);
  }

  var rowData = [
    data.studentName,
    data.agiiran,
    data.notes,
    data.date,
    data.signature
  ];
  sheet.appendRow(rowData);

  var pdfContent = `
    \\documentclass{article}
    \\usepackage[utf8]{inputenc}
    \\usepackage[a4paper, margin=1in]{geometry}
    \\begin{document}
    \\section*{Laporan Pelajar LAMPT-04-04}
    \\textbf{Nama Pelajar:} ${data.studentName}
    \\textbf{No. Angka Giliran:} ${data.agiiran}
    \\textbf{Catatan Pensyarah:} ${data.notes}
    \\textbf{Tarikh:} ${data.date}
    \\textbf{Tandatangan:} ${data.signature ? data.signature : "Tiada tandatangan"}
    \\end{document}
  `;

  var pdfBlob = generatePDF(pdfContent);
  var fileName = data.studentName.replace(/ /g, "_") + "_Laporan.pdf";
  DriveApp.createFile(pdfBlob).setName(fileName);

  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    message: "Laporan telah disimpan dan PDF dijana."
  })).setMimeType(ContentService.MimeType.JSON);
}

function generatePDF(latexContent) {
  var tempFile = DriveApp.createFile("temp.tex", latexContent);
  var pdf = DriveApp.getFileById(tempFile.getId()).getAs("application/pdf");
  tempFile.setTrashed(true);
  return pdf;
}
