const path = require('path');
const {app, BrowserWindow } = require('electron');

const isDev = process.env.NODE_ENV !=='development';

//app.disableHardwareAcceleration()



function createMainWindow() {
    const mainWindow = new BrowserWindow({
        title: 'UBS Billing Report Checker',
        // width: isDev ? 1000 : 500,     
        width: 700,    
        height: 550,
        autoHideMenuBar: true,
        maxWidth: 700,
        maxHeight: 550,
        resizable: false
    });

    //open devtools if in dev env
    if(isDev){
        mainWindow.webContents.openDevTools();
    } 

    mainWindow.loadFile(path.join(__dirname, './renderer/index.html')); 
    
}

app.whenReady().then(() => {
    createMainWindow();

    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) {
          createWindow()
        }
      });
    });
    
    app.on('window-all-closed', () => {
      if (process.platform !== 'darwin') {
        app.quit()
      }
    });