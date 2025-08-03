import formidable from 'formidable'
import { readFile } from 'fs/promises'
import XLSX from 'xlsx'
import PDFDocument from 'pdfkit'

export const config = {
  api: {
    bodyParser: false
  }
}

const css = `
:root {
  --bg: #11151a;
  --fg: #f4f4f4;
  --accent: #00b7c2;
  --card: #181c23;
  --border: #282c34;
  --button-bg: #222c37;
  --button-hover: #00b7c2;
  --button-fg: #fff;
}
html, body {
  height: 100%;
  background: var(--bg);
  color: var(--fg);
  font-family: 'Segoe UI', 'Arial', sans-serif;
  margin: 0;
  padding: 0;
}
main {
  max-width: 400px;
  margin: 6vh auto;
  padding: 2rem;
  background: var(--card);
  border-radius: 18px;
  box-shadow: 0 8px 24px #0008;
  border: 1px solid var(--border);
  display: flex;
  flex-direction: column;
  align-items: center;
}
h1 {
  margin-top: 0;
  font-weight: 700;
  letter-spacing: 0.01em;
  color: var(--accent);
  font-size: 2.2rem;
}
form {
  width: 100%;
  display: flex;
  flex-direction: column;
  gap: 1.2rem;
  margin-top: 1.5rem;
  margin-bottom: 1rem;
}
input[type="file"] {
  color-scheme: dark;
  background: var(--button-bg);
  border: 1px solid var(--border);
  border-radius: 8px;
  color: var(--fg);
  padding: 7px;
}
button {
  background: var(--button-bg);
  color: var(--button-fg);
  border: none;
  border-radius: 8px;
  padding: 0.9em 1.2em;
  font-size: 1rem;
  font-weight: bold;
  cursor: pointer;
  transition: background .2s, color .2s;
}
button:hover {
  background: var(--button-hover);
  color: #101520;
}
#status {
  margin-top: 0.6em;
  font-size: 1.08em;
  color: var(--accent);
  font-weight: 500;
  min-height: 1.2em;
}
`

const html = `
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>Excel a PDF - Modo Oscuro</title>
  <link rel="stylesheet" href="/style.css" />
</head>
<body>
  <main>
    <h1>Excel a PDF</h1>
    <form id="form" enctype="multipart/form-data">
      <input type="file" name="archivo" id="archivo" accept=".xls,.xlsx" required />
      <button type="submit">Convertir a PDF</button>
    </form>
    <div id="status"></div>
  </main>
  <script>
    const form = document.getElementById('form');
    const status = document.getElementById('status');
    form.onsubmit = async (e) => {
      e.preventDefault();
      status.innerText = "Procesando...";
      const data = new FormData(form);
      const res = await fetch('/generar-pdf', { method: 'POST', body: data });
      if (res.ok) {
        const blob = await res.blob();
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = 'stock.pdf';
        link.click();
        status.innerText = "¡PDF generado!";
      } else {
        status.innerText = "Error al generar PDF";
      }
    }
  </script>
</body>
</html>
`

export default async function handler (req, res) {
  // Servir CSS embebido
  if (req.method === 'GET' && req.url === '/style.css') {
    res.setHeader('Content-Type', 'text/css')
    res.status(200).send(css)
    return
  }

  // Servir HTML embebido en /
  if (req.method === 'GET' && (req.url === '/' || req.url.startsWith('/?'))) {
    res.setHeader('Content-Type', 'text/html')
    res.status(200).send(html)
    return
  }

  // Función para crear PDF desde excel
  if (req.method === 'POST' && req.url.startsWith('/generar-pdf')) {
    const form = formidable()
    form.parse(req, async (err, fields, files) => {
      if (err || !files.archivo) {
        res.status(400).send('Archivo no válido')
        return
      }
      try {
        const filePath = files.archivo[0].filepath
        const data = await readFile(filePath)
        const workbook = XLSX.read(data, { type: 'buffer' })
        const sheet = workbook.Sheets[workbook.SheetNames[0]]
        // Desde fila 10, columnas A, C, F, H
        const json = XLSX.utils.sheet_to_json(sheet, { header: 1, range: 9 })
        const table = json
          .map(row => [row[0] || '', row[2] || '', row[5] || '', row[7] || ''])
          .filter(r => r.some(x => x))

        res.setHeader('Content-Type', 'application/pdf')
        res.setHeader('Content-Disposition', 'attachment; filename=stock.pdf')
        const doc = new PDFDocument({ size: 'A4', margin: 40 })
        doc.rect(0, 0, doc.page.width, doc.page.height).fill('#181c23')
        doc
          .fillColor('#00b7c2')
          .fontSize(20)
          .font('Helvetica-Bold')
          .text('Stock de Productos', { align: 'center' })
        doc.moveDown(1.2)

        const headers = ['Código', 'Descripción', 'Rubro', 'Stock']
        const startX = 48,
          startY = doc.y + 10
        const cellW = [70, 220, 90, 60]
        const cellH = 28

        doc.fontSize(13).font('Helvetica-Bold')
        headers.forEach((h, i) => {
          doc
            .fillColor('#282c34')
            .rect(
              startX + cellW.slice(0, i).reduce((a, b) => a + b, 0),
              startY,
              cellW[i],
              cellH
            )
            .fill('#00b7c2')
          doc
            .fillColor('#f4f4f4')
            .text(
              h,
              startX + cellW.slice(0, i).reduce((a, b) => a + b, 0) + 8,
              startY + 7,
              { width: cellW[i] - 16, align: 'left' }
            )
        })

        let y = startY + cellH
        doc.font('Helvetica').fontSize(11)
        table.forEach((row, rowIdx) => {
          if (y > doc.page.height - 60) {
            doc.addPage()
            y = 50
          }
          row.forEach((cell, i) => {
            doc
              .fillColor('#23272e')
              .rect(
                startX + cellW.slice(0, i).reduce((a, b) => a + b, 0),
                y,
                cellW[i],
                cellH
              )
              .fill(rowIdx % 2 === 0 ? '#23272e' : '#181c23')
            doc
              .fillColor('#f4f4f4')
              .text(
                String(cell),
                startX + cellW.slice(0, i).reduce((a, b) => a + b, 0) + 8,
                y + 7,
                { width: cellW[i] - 16, align: 'left', ellipsis: true }
              )
          })
          y += cellH
        })

        doc.end()
        doc.pipe(res)
      } catch (e) {
        res.status(500).send('Error procesando archivo')
      }
    })
    return
  }

  // Si no es ninguna de las anteriores, 404
  res.status(404).send('Not found')
}
