import formidable from "formidable";
import { readFile } from "fs/promises";
import XLSX from "xlsx";
import PDFDocument from "pdfkit";

export const config = {
  api: {
    bodyParser: false,
  },
};

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).send("Método no permitido");
    return;
  }

  const form = formidable();
  form.parse(req, async (err, fields, files) => {
    if (err || !files.archivo) {
      res.status(400).send("Archivo no válido");
      return;
    }

    try {
      const filePath = files.archivo[0].filepath;
      const data = await readFile(filePath);
      const workbook = XLSX.read(data, { type: "buffer" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];

      // Filas desde la 10 (range: 9), columnas: A, C, F, H
      const json = XLSX.utils.sheet_to_json(sheet, { header: 1, range: 9 });
      const table = json.map(row => [
        row[0] || "",
        row[2] || "",
        row[5] || "",
        row[7] || "",
      ]).filter(r => r.some(x => x));

      res.setHeader("Content-Type", "application/pdf");
      res.setHeader("Content-Disposition", "attachment; filename=stock.pdf");
      const doc = new PDFDocument({ size: "A4", margin: 40 });

      // FONDO oscuro
      doc.rect(0, 0, doc.page.width, doc.page.height).fill("#181c23");
      doc.fillColor("#00b7c2")
        .fontSize(20)
        .font("Helvetica-Bold")
        .text("Stock de Productos", { align: "center" });
      doc.moveDown(1.2);

      // Encabezados
      const headers = ["Código", "Descripción", "Rubro", "Stock"];
      const startX = 48, startY = doc.y + 10;
      const cellW = [70, 220, 90, 60];
      const cellH = 28;

      doc.fontSize(13).font("Helvetica-Bold");
      headers.forEach((h, i) => {
        doc
          .fillColor("#282c34")
          .rect(startX + cellW.slice(0, i).reduce((a, b) => a + b, 0), startY, cellW[i], cellH)
          .fill("#00b7c2");
        doc
          .fillColor("#f4f4f4")
          .text(h, startX + cellW.slice(0, i).reduce((a, b) => a + b, 0) + 8, startY + 7, { width: cellW[i] - 16, align: "left" });
      });

      let y = startY + cellH;
      doc.font("Helvetica").fontSize(11);
      table.forEach((row, rowIdx) => {
        if (y > doc.page.height - 60) {
          doc.addPage();
          y = 50;
        }
        row.forEach((cell, i) => {
          doc
            .fillColor("#23272e")
            .rect(startX + cellW.slice(0, i).reduce((a, b) => a + b, 0), y, cellW[i], cellH)
            .fill(rowIdx % 2 === 0 ? "#23272e" : "#181c23");
          doc
            .fillColor("#f4f4f4")
            .text(String(cell), startX + cellW.slice(0, i).reduce((a, b) => a + b, 0) + 8, y + 7, { width: cellW[i] - 16, align: "left", ellipsis: true });
        });
        y += cellH;
      });

      doc.end();
      doc.pipe(res);
    } catch (e) {
      res.status(500).send("Error procesando archivo");
    }
  });
}
