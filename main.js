const { app, BrowserWindow, ipcMain } = require('electron')
const path = require('path')
const fs = require('fs')

const APP_ID = 'com.ghzplugin.nexhotel'
const WINDOW_CHROME_COLOR = '#071a3d'
const WINDOW_OVERLAY_COLOR = '#00000000'
const WINDOW_OVERLAY_HEIGHT = 18
const ICON_PATH = path.join(__dirname, 'assets', 'nexhotel-icon.ico')

function ensureDataDir() {
  const dataDir = process.env.NEXHOTEL_DATA_DIR || path.join(app.getPath('userData'), 'data')
  fs.mkdirSync(dataDir, { recursive: true })
  return dataDir
}

function createWindow() {
  const dataDir = ensureDataDir()
  const win = new BrowserWindow({
    width: 1540,
    height: 950,
    minWidth: 1240,
    minHeight: 780,
    autoHideMenuBar: true,
    title: 'NexHotel',
    icon: ICON_PATH,
    backgroundColor: WINDOW_CHROME_COLOR,
    titleBarStyle: 'hidden',
    titleBarOverlay: {
      color: WINDOW_OVERLAY_COLOR,
      symbolColor: '#ffeec1',
      height: WINDOW_OVERLAY_HEIGHT
    },
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
      additionalArguments: [`--data-dir=${dataDir}`]
    }
  })

  win.loadFile('index.html')

  if (process.env.NEXHOTEL_DEBUG === '1') {
    win.webContents.on('console-message', (_event, level, message, line, sourceId) => {
      console.log(`[renderer:${level}] ${message} (${sourceId}:${line})`)
    })

    win.webContents.on('did-finish-load', async () => {
      const result = await win.webContents.executeJavaScript(`
        (() => {
          const results = []
          document.querySelectorAll('.nav-btn[data-page]').forEach((button) => {
            button.click()
            const page = button.dataset.page
            results.push({
              page,
              active: Boolean(document.getElementById('page-' + page)?.classList.contains('active'))
            })
          })
          return {
            buttons: document.querySelectorAll('.nav-btn').length,
            failed: results.filter((item) => !item.active),
            lastPage: document.querySelector('.page.active')?.id || ''
          }
        })()
      `)
      console.log('[nexhotel-debug]', JSON.stringify(result))

      if (process.env.NEXHOTEL_DEBUG_FLOW === '1') {
        const flow = await win.webContents.executeJavaScript(`
          (async () => {
            const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms))
            const today = new Date().toISOString().slice(0, 10)
            const tomorrowDate = new Date()
            tomorrowDate.setDate(tomorrowDate.getDate() + 1)
            const tomorrow = tomorrowDate.toISOString().slice(0, 10)
            const click = (id) => {
              const el = document.getElementById(id)
              if (!el) throw new Error('Elemento nao encontrado: ' + id)
              el.click()
            }
            const findButton = (selector, label) => {
              const button = Array.from(document.querySelectorAll(selector)).find((item) =>
                String(item.textContent || '').includes(label)
              )
              if (!button) throw new Error('Botao nao encontrado: ' + label)
              return button
            }

            window.open = () => ({
              document: {
                open() {},
                write() {},
                close() {}
              }
            })

            document.querySelector('[data-page="quartos"]').click()
            document.getElementById('roomNumber').value = 'TST-101'
            document.getElementById('roomCategory').value = 'STANDARD'
            document.getElementById('roomRate').value = '100'
            document.getElementById('roomBaseStatus').value = 'DISPONIVEL'
            click('btnSaveRoom')
            await sleep(150)

            findButton('#tbRooms button', 'Hospedar').click()
            await sleep(150)
            document.getElementById('hostNewName').value = 'Cliente Teste Debug'
            document.getElementById('hostNewDoc').value = '00000000000'
            document.getElementById('hostCheckin').value = today
            document.getElementById('hostCheckout').value = tomorrow
            document.getElementById('hostDaily').value = '100'
            click('hostDialogOk')
            await sleep(250)

            const hostedAfterCheckin = document.querySelectorAll('#tbHosted tr').length
            if (!document.getElementById('page-hospedados').classList.contains('active')) {
              throw new Error('Aba hospedados nao ficou ativa apos check-in rapido.')
            }

            findButton('#tbHosted button', 'Adicionar pagamento').click()
            await sleep(150)
            document.getElementById('quickPaymentAmount').value = '100'
            document.getElementById('quickPaymentMethod').value = 'PIX'
            click('paymentDialogOk')
            await sleep(250)

            findButton('#tbHosted button', 'Fechar conta').click()
            await sleep(150)
            click('appDialogOk')
            await sleep(350)

            const noHosted = Boolean(document.querySelector('#tbHosted .empty'))
            return {
              hostedAfterCheckin,
              noHostedAfterClose: noHosted,
              activePage: document.querySelector('.page.active')?.id || '',
              receiptAudit: Array.from(document.querySelectorAll('#tbAudit tr')).some((row) =>
                String(row.textContent || '').includes('Emissao de recibo')
              )
            }
          })()
        `)
        console.log('[nexhotel-debug-flow]', JSON.stringify(flow))
      }

      app.quit()
    })
  }
}

app.setAppUserModelId(APP_ID)

require('./js/ghz-backend')({
  app, ipcMain, getDataDir: ensureDataDir,
  appId: 'nexhotel',
  manifestUrl: 'https://raw.githubusercontent.com/GhuzzBeatz/NexHotel/master/update-manifest.json'
})

app.whenReady().then(createWindow)

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit()
})
