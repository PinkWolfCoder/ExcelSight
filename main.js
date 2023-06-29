const { app, BrowserWindow, ipcMain } = require('electron');
const fs = require('fs');
const path = require('path');
const ExcelController = require('./excelmanipulation');

let mainWindow;

app.commandLine.appendSwitch('touch-events', 'enabled');
app.disableHardwareAcceleration();

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
      enableRemoteModule: true,
    },
  });

  mainWindow.loadFile(path.join(__dirname, 'index.html'));

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

app.on('ready', createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});

ipcMain.handle('load-excel', (event, filePath) => {
  const fullPath = path.resolve(__dirname, filePath);
  console.log(`Loading Excel file from path: ${fullPath}`);
  return new Promise((resolve, reject) => {
    fs.access(fullPath, fs.constants.F_OK, (err) => {
      if (err) {
        console.error(`The file ${filePath} does not exist.`);
        event.sender.send('load-result', false); // Inform renderer of failure
        reject({ success: false, message: 'File not found' });
      } else {
        console.log(`Excel file loaded successfully: ${fullPath}`);
        
        let excelController;
        try {
          excelController = new ExcelController(fullPath);
        } catch (error) {
          console.error('Error creating ExcelController:', error);
          reject({ success: false, message: 'Error creating ExcelController' });
          return;
        }

        // Call readEmptyRow method
        const rowData = excelController.readEmptyRow();
        console.log('Empty row data:', rowData);

        event.sender.send('load-result', true); // Inform renderer of success
        event.sender.send('row-data', rowData);
        resolve({ success: true, message: 'File loaded successfully' });
      }
    });
  });
});


ipcMain.on('read-row', (event, rowNumber) => {
  console.log(`Reading row ${rowNumber} from Excel...`);
  const filePath = path.resolve(__dirname, 'data.xlsx');
  let excelController;
  try {
    excelController = new ExcelController(filePath);
  } catch (error) {
    console.error('Error creating ExcelController:', error);
    return;
  }
  
  excelController.readRow(parseInt(rowNumber))
    .then((rowData) => {
      console.log('Read data:', rowData);
      event.sender.send('row-data', rowData); // Send the data to the renderer process
    })
    .catch((error) => {
      console.error('Error reading row:', error);
    })
    .finally(() => {
      excelController.close();
    });
});


ipcMain.handle('append-row', (event, data) => {
  console.log('Appending data to Excel:', data);
  const filePath = path.resolve(__dirname, 'data.xlsx');
  let excelController;
  try {
    excelController = new ExcelController(filePath);
    excelController.appendRow([data.col1, data.col2, data.col3]);
    console.log('Data appended successfully');
    return { success: true, message: 'Data appended successfully' };
  } catch (error) {
    console.error('Error appending data:', error);
    return { success: false, message: 'Error appending data' };
  } finally {
    if (excelController) {
      excelController.close();
    }
  }
});

ipcMain.on('exit-app', () => {
  app.quit();
});
