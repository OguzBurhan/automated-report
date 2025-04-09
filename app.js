// app.js
const express = require('express');
const multer  = require('multer');
const XLSX = require('xlsx');
const app = express();
const port = 3000;

// Set up Multer with in-memory storage.
const upload = multer({ storage: multer.memoryStorage() });

// Serve a simple UI for file upload.
app.get('/', (req, res) => {
  res.send(`
    <!DOCTYPE html>
    <html>
    <head>
      <title>XLSX File Processor</title>
      <style>
        body {
          font-family: Arial, sans-serif;
          background-color: #f4f4f4;
          margin: 40px auto;
          max-width: 600px;
          padding: 20px;
          border: 1px solid #ccc;
          border-radius: 5px;
        }
        h2 { text-align: center; }
        form { text-align: center; margin-top: 20px; }
        .form-control { margin: 15px 0; }
        label { font-weight: bold; }
        input[type="file"] { margin: 10px 0; }
        input[type="submit"] {
          padding: 10px 20px;
          background-color: #007bff;
          color: #fff;
          border: none;
          border-radius: 4px;
          cursor: pointer;
        }
      </style>
    </head>
    <body>
      <h2>XLSX File Processor</h2>
      <form action="/process" method="post" enctype="multipart/form-data">
        <div class="form-control">
          <label>Source XLSX File:</label>
          <input type="file" name="sourceFile" accept=".xlsx" required />
        </div>
        <div class="form-control">
          <label>Template XLSX File:</label>
          <input type="file" name="templateFile" accept=".xlsx" required />
        </div>
        <input type="submit" value="Upload and Process" />
      </form>
    </body>
    </html>
  `);
});

// POST endpoint to process the XLSX files.
app.post('/process', upload.fields([
  { name: 'sourceFile', maxCount: 1 },
  { name: 'templateFile', maxCount: 1 }
]), (req, res) => {
  try {
    // Verify that both files are provided.
    if (!req.files.sourceFile || !req.files.templateFile) {
      return res.status(400).send("Both files are required.");
    }

    // Convert file buffers to XLSX workbooks.
    const sourceBuffer = req.files.sourceFile[0].buffer;
    const templateBuffer = req.files.templateFile[0].buffer;

    const sourceWorkbook = XLSX.read(sourceBuffer, { type: 'buffer' });
    const sourceSheetName = sourceWorkbook.SheetNames[0];
    const sourceSheet = sourceWorkbook.Sheets[sourceSheetName];
    // Convert sheet to an array-of-arrays.
    let sourceData = XLSX.utils.sheet_to_json(sourceSheet, { header: 1 });

    const templateWorkbook = XLSX.read(templateBuffer, { type: 'buffer' });
    const templateSheetName = templateWorkbook.SheetNames[0];
    const templateSheet = templateWorkbook.Sheets[templateSheetName];
    let templateData = XLSX.utils.sheet_to_json(templateSheet, { header: 1 });

    // Ensure template has at least 63 rows; if not, pad with empty arrays.
    while (templateData.length < 63) {
      templateData.push([]);
    }

    // ---------------- Extract Data from Source ----------------

    // Section 1: rows 2 to 9 (array indexes 1 to 8), extract only column C (index 2)
    let extractedSection1 = [];
    for (let i = 1; i < 9; i++) {
      const row = sourceData[i] || [];
      extractedSection1.push(row[2] || "");
    }

    // Section 2: rows 13 to 25 (indexes 12 to 24), extract columns B, C, D (indexes 1, 2, 3)
    let extractedSection2 = [];
    for (let i = 12; i < 25; i++) {
      const row = sourceData[i] || [];
      extractedSection2.push({
        B: row[1] || "",
        C: row[2] || "",
        D: row[3] || ""
      });
    }

    // ---------------- Paste Extracted Data into Template ----------------

    // Utility function: pad an array to a given length.
    function padArray(arr, length) {
      while (arr.length < length) {
        arr.push("");
      }
      return arr;
    }

    // Paste Section 1 into template rows 40–47 (indexes 39–46) into column D (index 3).
    for (let j = 0; j < extractedSection1.length; j++) {
      const destIndex = 39 + j;
      if (templateData[destIndex] === undefined) {
        templateData[destIndex] = [];
      }
      templateData[destIndex] = padArray(templateData[destIndex], 4);
      templateData[destIndex][3] = extractedSection1[j];
    }

    // Paste Section 2 into template rows 51–63 (indexes 50–62) into columns C, D, E (indexes 2, 3, 4).
    const maxSection2Rows = 13;
    for (let k = 0; k < Math.min(extractedSection2.length, maxSection2Rows); k++) {
      const destIndex = 50 + k;
      if (templateData[destIndex] === undefined) {
        templateData[destIndex] = [];
      }
      templateData[destIndex] = padArray(templateData[destIndex], 5);
      templateData[destIndex][2] = extractedSection2[k].B;
      templateData[destIndex][3] = extractedSection2[k].C;
      templateData[destIndex][4] = extractedSection2[k].D;
    }

    // ---------------- Write Out New XLSX File ----------------

    const newSheet = XLSX.utils.aoa_to_sheet(templateData);
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newSheet, templateSheetName);

    // Write workbook to a buffer.
    const outputBuffer = XLSX.write(newWorkbook, { type: 'buffer', bookType: 'xlsx' });

    // Return the XLSX file as a download.
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=final_output.xlsx');
    res.send(outputBuffer);

  } catch (error) {
    console.error("Error processing XLSX files:", error);
    res.status(500).send("Error processing files.");
  }
});

app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});
