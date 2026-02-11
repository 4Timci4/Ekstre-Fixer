const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');

// Force Turkish locale so date inputs show dd.MM.yyyy
app.commandLine.appendSwitch('lang', 'tr');

// Hot reload in development
if (!app.isPackaged) {
  require('electron-reload')(__dirname, {
    electron: path.join(__dirname, 'node_modules', '.bin', 'electron.cmd'),
    hardResetMethod: 'exit'
  });
}

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 480,
    height: 620,
    minWidth: 360,
    minHeight: 480,
    resizable: true,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    },
    icon: path.join(__dirname, 'icon.ico')
  });

  mainWindow.loadFile('index.html');
  mainWindow.setMenuBarVisibility(false);
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

// IPC Handlers
ipcMain.handle('select-files', async (event, defaultPath) => {
  const opts = {
    title: 'Excel dosyaları seçin',
    filters: [{ name: 'Excel files', extensions: ['xlsx', 'xls'] }],
    properties: ['openFile', 'multiSelections']
  };
  if (defaultPath) opts.defaultPath = defaultPath;
  const result = await dialog.showOpenDialog(mainWindow, opts);
  return result.filePaths;
});

ipcMain.handle('save-file', async (event, defaultName, defaultDir) => {
  const defaultPath = defaultDir ? path.join(defaultDir, defaultName) : defaultName;
  const result = await dialog.showSaveDialog(mainWindow, {
    title: 'İşlenmiş Dosyayı Kaydet',
    defaultPath: defaultPath,
    filters: [{ name: 'Excel files', extensions: ['xlsx'] }]
  });
  return result.filePath;
});

ipcMain.handle('select-directory', async (event, defaultPath) => {
  const opts = {
    title: 'İşlenmiş Dosyaların Kaydedileceği Klasörü Seçin',
    properties: ['openDirectory']
  };
  if (defaultPath) opts.defaultPath = defaultPath;
  const result = await dialog.showOpenDialog(mainWindow, opts);
  return result.filePaths[0];
});

ipcMain.handle('open-folder', async (event, folderPath) => {
  shell.openPath(folderPath);
});

ipcMain.handle('select-folder', async (event, { title, defaultPath }) => {
  const opts = {
    title: title || 'Klasör Seçin',
    properties: ['openDirectory']
  };
  if (defaultPath) opts.defaultPath = defaultPath;
  const result = await dialog.showOpenDialog(mainWindow, opts);
  return result.filePaths[0] || null;
});

ipcMain.handle('show-message', async (event, { type, title, message, buttons }) => {
  const result = await dialog.showMessageBox(mainWindow, {
    type: type || 'info',
    title: title,
    message: message,
    buttons: buttons || ['OK']
  });
  return result.response;
});
