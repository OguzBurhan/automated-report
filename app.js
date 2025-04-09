// app.js
const express = require('express');
const multer  = require('multer');
const XLSX = require('xlsx');
const serverless = require('serverless-http');

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
    if (!req.files.sourceFile || !req.files.templateFile) {
      return res.status(400).send("Both files are required.");
    }

    // Read file buffers.
    const sourceBuffer = req.files.sourceFile[0].buffer;
    const templateBuffer = req.files.templateFile[0].buffer;

    // Read the workbooks.
    const sourceWorkbook = XLSX.read(sourceBuffer, { type: 'buffer' });
    const templateWorkbook = XLSX.read(templateBuffer, { type: 'buffer' });

    // Get the first sheet names.
    const sourceSheetName = sourceWorkbook.SheetNames[0];
    const templateSheetName = templateWorkbook.SheetNames[0];

    const sourceSheet = sourceWorkbook.Sheets[sourceSheetName];
    const templateSheet = templateWorkbook.Sheets[templateSheetName];

    // --- Optimize Data Extraction by Specifying Ranges ---
    // For Section 1 (rows 2–9), we assume columns A:Z cover the data.
    const sourceDataSec1 = XLSX.utils.sheet_to_json(sourceSheet, { 
      header: 1, 
      range: "A2:Z9" 
    });
    // For Section 2 (rows 13–25)
    const sourceDataSec2 = XLSX.utils.sheet_to_json(sourceSheet, { 
      header: 1, 
      range: "A13:Z25" 
    });

    // Read the full template data.
    let templateData = XLSX.utils.sheet_to_json(templateSheet, { header: 1 });
    while (templateData.length < 63) {
      templateData.push([]);
    }

    // ---------------- Extract Data from Source ----------------

    // Section 1: from sourceDataSec1, get only column C (index 2).
    let extractedSection1 = sourceDataSec1.map(row => row[2] || "");

    // Section 2: from sourceDataSec2, get columns B, C, D (indexes 1,2,3).
    let extractedSection2 = sourceDataSec2.map(row => ({
      B: row[1] || "",
      C: row[2] || "",
      D: row[3] || ""
    }));

    // ---------------- Paste Data into Template ----------------

    function padArray(arr, length) {
      while (arr.length < length) {
        arr.push("");
      }
      return arr;
    }

    // Paste Section 1 into template rows 40–47 (indexes 39–46) into column D (index 3)
    extractedSection1.forEach((value, idx) => {
      const destIndex = 39 + idx;
      if (!templateData[destIndex]) templateData[destIndex] = [];
      templateData[destIndex] = padArray(templateData[destIndex], 4);
      templateData[destIndex][3] = value;
    });

    // Paste Section 2 into template rows 51–63 (indexes 50–62) into columns C, D, E (indexes 2,3,4)
    const maxSection2Rows = 13;
    extractedSection2.slice(0, maxSection2Rows).forEach((obj, idx) => {
      const destIndex = 50 + idx;
      if (!templateData[destIndex]) templateData[destIndex] = [];
      templateData[destIndex] = padArray(templateData[destIndex], 5);
      templateData[destIndex][2] = obj.B;
      templateData[destIndex][3] = obj.C;
      templateData[destIndex][4] = obj.D;
    });

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

// For local testing: start the server when not in production.
if (process.env.NODE_ENV !== 'production') {
  app.listen(port, () => {
    console.log(`Server is running on http://localhost:${port}`);
  });
}

// Export the Express app wrapped with serverless-http for Vercel.
module.exports = serverless(app);
