const express = require('express');
const multer = require('multer');
const path = require('path');
const cors = require('cors'); 
const fs = require('fs');
const XLSX = require('xlsx');
const app = express();
const port = 3000;

app.use(cors());

// Error handling middleware
app.use((err, req, res, next) => {
  console.error(err.stack);
  if (err.code === 'LIMIT_FILE_SIZE') {
    return res.status(413).send('File size limit exceeded.');
  } else if (err.code === 'LIMIT_UNEXPECTED_FILE') {
    return res.status(422).send('Invalid file field.');
  } else {
    return res.status(500).send('Something went wrong!');
  }
});

// Multer configuration for file upload
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
      cb(null, 'uploads/') // Destination folder for storing uploaded files
    },
    filename: function (req, file, cb) {
      cb(null, file.originalname); // Use original filename
    }
  });
  
  const upload = multer({ storage: storage });
  
  async function modifyAndExportData(filePath) {
    return new Promise((resolve, reject) => {
      const workbook = XLSX.readFile(filePath);
      const worksheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[worksheetName];
      const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
  
      // Modify data as needed
      // For example, modify the first cell of the first row
      if (data.length > 0 && data[0].length > 0) {
        data[0][0] = "Modified: " + data[0][0];
      }
  
      // Create a new workbook with the modified data
      const modifiedWorksheet = XLSX.utils.aoa_to_sheet(data);
      const modifiedWorkbook = { Sheets: { 'data': modifiedWorksheet }, SheetNames: ['data'] };
  
      // Write the modified workbook to a buffer
      const buffer = XLSX.write(modifiedWorkbook, { type: 'buffer', bookType: 'xlsx' });
  
      // Construct the file path for the modified file
      const modifiedFilePath = path.join("C:/Users/Varshini/Desktop/ang-poc/api/", 'temp', 'modified_file.xlsx');
  
      // Write the buffer to a file
      fs.writeFile(modifiedFilePath, buffer, (err) => {
        if (err) {
          console.error('Error writing file:', err);
          reject(err);
        } else {
          console.log('File written successfully:', modifiedFilePath);
          resolve(modifiedFilePath);
        }
      });
    });
  }
  
  
  // POST endpoint to handle file upload and modification
  app.post('/upload', upload.single('file'), async (req, res) => {
    try {
      if (!req.file) {
        return res.status(400).send('No files were uploaded.');
      }
     // console.log('Uploaded file:', req.file.path);
      const modifiedFilePath = await modifyAndExportData(req.file.path);
      res.download(modifiedFilePath, 'modified_file.xlsx', (err) => {
        if (err) {
          throw err;
        }
        // Clean up temporary file after download
        fs.unlink(modifiedFilePath, (err) => {
          if (err) {
            console.error('Error cleaning up temporary file:', err);
          }
        });
      });
    } catch (error) {
      console.error('Error processing file:', error);
      res.status(500).send('Error processing file.');
    }
  });

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});
