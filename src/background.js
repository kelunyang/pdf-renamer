'use strict'

import _ from 'lodash';
import path from 'path';
import fs from 'fs-extra';
import * as pdfjsLib from 'pdfjs-dist/legacy/build/pdf.js';
if(process.env.NODE_ENV === 'production') {
  pdfjsLib.GlobalWorkerOptions.workerSrc = path.join(__dirname, 'pdf.worker.js');
} else {
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'pdfjs-dist/legacy/build/pdf.worker.js';
}
import { app, protocol, BrowserWindow, ipcMain, shell } from 'electron'
import { createProtocol } from 'vue-cli-plugin-electron-builder/lib'
import installExtension, { VUEJS_DEVTOOLS } from 'electron-devtools-installer'
const isDevelopment = process.env.NODE_ENV !== 'production'

// Scheme must be registered before the app is ready
protocol.registerSchemesAsPrivileged([
  { scheme: 'app', privileges: { secure: true, standard: true } }
])

async function createWindow() {
  // Create the browser window.
  const win = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      
      // Use pluginOptions.nodeIntegration, leave this alone
      // See nklayman.github.io/vue-cli-plugin-electron-builder/guide/security.html#node-integration for more info
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js'),
    }
  })

  if (process.env.WEBPACK_DEV_SERVER_URL) {
    // Load the url of the dev server if in development mode
    await win.loadURL(process.env.WEBPACK_DEV_SERVER_URL)
    if (!process.env.IS_TEST) win.webContents.openDevTools()
  } else {
    createProtocol('app')
    // Load the index.html when not in development
    win.loadURL('app://./index.html')
  }
  ipcMain.on("scanFolder", (event, arg)  => {
    let dir = path.dirname(arg);
    let files = fs.readdirSync(dir);
    files = _.filter(files, (file) => {
      return file.match(/^.*\.(pdf|PDF)$/g);
    });
    files = _.map(files, (file) => {
      return {
        name: path.basename(file),
        fullpath: dir + "/" + file
      }
    })
    win.webContents.send("scanFolder", {
      pdfs: files,
      folder: dir
    });
  });
  ipcMain.on("openLink",async (event, arg)  => {
    shell.openExternal(arg);
  });
  ipcMain.on("renamePDF",async (event, arg)  => {
    win.webContents.send("scanStatus", arg.pdf.name + "複製中...");
    await fs.ensureDir(arg.folder + "/export/");
    await fs.copy(arg.pdf.fullpath, arg.folder + "/export/" + arg.pdf.rule.result + ".pdf");
    win.webContents.send("renamePDF", arg.pdf);
  });
  ipcMain.on("scanPDF",async (event, arg)  => {
    let pdf = await (pdfjsLib.getDocument(arg.pdf.fullpath)).promise;
    let textConents = [];
    let matchRule = undefined;
    let meta = await pdf.getMetadata();
    win.webContents.send("scanStatus", "已開啟：" + arg.pdf.name);
    for(let i=1; i<= pdf.numPages; i++) {
      let page = await pdf.getPage(i);
      win.webContents.send("scanStatus", arg.pdf.name + "第" + i + "頁");
      let pageContent = "";
      let contents = await page.getTextContent();
      if(contents.items.length > 0) {
        let strs = _.map(contents.items, (content) => {
          return content.str;
        });
        pageContent = _.join(strs, "");
      }
      textConents.push(pageContent);
    }
    let textContent = _.join(textConents, "");
    for(let i=0; i<arg.rules.length; i++) {
      let rule = arg.rules[i];
      let match = [];
      let type = rule.type ? "內容" : "中介";
      win.webContents.send("scanStatus", arg.pdf.name + "的[" + type + "]是否有：" + rule.keyword);
      if(rule.type) {
        match = [...textContent.matchAll(new RegExp(rule.keyword, 'g'))];
      } else {
        let metaData = "";
        if(meta.info.Author !== undefined) metaData += meta.info.Author;
        if(meta.info.Title !== undefined) metaData += meta.info.Title;
        if(meta.info.Subject !== undefined) metaData += meta.info.Subject;
        if(meta.info.Keywords !== undefined) metaData += meta.info.Keywords;
        match = [...metaData.matchAll(new RegExp(rule.keyword, 'g'))];
      }
      if(match.length === rule.times) {
        matchRule = rule;
        break;
      }
    }
    win.webContents.send("scanStatus", arg.pdf.name + "檢查完成！");
    win.webContents.send("scanPDF", {
      pdf: arg.pdf,
      rule: matchRule
    });
  });
}

// Quit when all windows are closed.
app.on('window-all-closed', () => {
  // On macOS it is common for applications and their menu bar
  // to stay active until the user quits explicitly with Cmd + Q
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

app.on('activate', () => {
  // On macOS it's common to re-create a window in the app when the
  // dock icon is clicked and there are no other windows open.
  if (BrowserWindow.getAllWindows().length === 0) createWindow()
})

// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
// Some APIs can only be used after this event occurs.
app.on('ready', async () => {
  if (isDevelopment && !process.env.IS_TEST) {
    // Install Vue Devtools
    try {
      await installExtension(VUEJS_DEVTOOLS)
    } catch (e) {
      console.error('Vue Devtools failed to install:', e.toString())
    }
  }
  createWindow()
})

// Exit cleanly on request from parent process in development mode.
if (isDevelopment) {
  if (process.platform === 'win32') {
    process.on('message', (data) => {
      if (data === 'graceful-exit') {
        app.quit()
      }
    })
  } else {
    process.on('SIGTERM', () => {
      app.quit()
    })
  }
}
