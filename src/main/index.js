const path = require('path')
const fs = require('fs')
const https = require('https')
const { app, BrowserWindow, ipcMain, shell, Menu } = require('electron')
const mysql = require('mysql2/promise')
const xlsx = require('xlsx')

const pad2 = (value) => String(value).padStart(2, '0')

const formatDateParts = (year, month, day) => `${pad2(day)}/${pad2(month)}/${year}`

const computeBarcode23i = (codice) => {
  const codiceLoc = String(codice ?? '').trim().padStart(11, '0').slice(-11)
  let sommaPari = 0
  let sommaDispari = 0

  for (let i = 0; i < codiceLoc.length; i += 1) {
    const digit = Number(codiceLoc[i] || 0)
    const posizione = i + 1
    if (posizione % 2 === 0) {
      sommaPari += digit
    } else {
      sommaDispari += digit
    }
  }

  let sommaResto = (sommaPari * 11) + sommaDispari
  let somma = Math.floor(sommaResto / 100)
  sommaResto %= 100

  somma += Math.floor(sommaResto / 10)
  sommaResto %= 10

  somma += sommaResto
  const checkDigit = somma % 10

  const codiceAr23i = `${codiceLoc}-${checkDigit}`
  const barcode23i = `${codiceLoc}${checkDigit}`
  return { codiceLoc, codiceAr23i, barcode23i }
}

const tryParseDate = (raw) => {
  if (raw instanceof Date && !Number.isNaN(raw.valueOf())) {
    return formatDateParts(raw.getFullYear(), raw.getMonth() + 1, raw.getDate())
  }

  if (typeof raw === 'string') {
    const trimmed = raw.trim()
    const match = trimmed.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/)
    if (match) {
      const day = Number(match[1])
      const month = Number(match[2])
      const year = match[3].length === 2 ? Number(`20${match[3]}`) : Number(match[3])
      if (day && month && year) {
        return formatDateParts(year, month, day)
      }
    }

    const parsed = new Date(trimmed)
    if (!Number.isNaN(parsed.valueOf())) {
      return formatDateParts(parsed.getFullYear(), parsed.getMonth() + 1, parsed.getDate())
    }
  }

  return null
}

const formatDateCell = (cell, date1904) => {
  if (!cell) return null

  if (cell.t === 'd') {
    return tryParseDate(cell.v)
  }

  if (cell.t === 'n') {
    const parts = xlsx.SSF.parse_date_code(cell.v, { date1904 })
    if (parts && parts.y && parts.m && parts.d) {
      return formatDateParts(parts.y, parts.m, parts.d)
    }
  }

  return tryParseDate(cell.v)
}

const dbConfig = {
  host: '82.223.30.31',
  port: 3306,
  user: 'mediaprint',
  password: 'M3d1aPr1ntDB@',
  database: 'ProvinciaAmbiente',
}

const dbTable = 'PIMP'

const baseDbColumns = [
  'id_pianificato',
  'REC_NOMINATIVO',
  'REC_NOMINATIVO2',
  'REC_INDIRIZZO',
  'REC_CIVICO',
  'REC_SUBCIVICO',
  'REC_CAP',
  'REC_LOCALITA',
  'REC_PROVINCIA',
  "CODICE IMPIANTO'",
  "PROGRESSIVO LETTERA'",
  'DATALETTERA',
  'ESITO',
  "CODICE VECCHIO'",
  'DATA_PREVISTA',
  'ORA_PREVISTA',
  'IMP_NOMINATIVO',
  'IMP_INDIRIZZO',
  'IMP_CIVICO',
  'IMP_SUBCIVICO',
  'IMP_LOCALITA',
  'CODELINE',
  'AUTODICHIARATO',
  'ISPEZIONABILE',
  'CATEGORIA',
  'ESITOULTIMA',
  'DATAULTIMAVERIFICA',
  'Codice lavorazione',
  'Progressivo',
]

const pimpArDbColumns = [
  'CODICE_AR',
  'CODICE_AR_23i',
  'BARCODE23i',
  'OMOLOGAZIONE',
]

const pimpDbColumns = [
  'OMOLOGAZIONE',
]

const jasperConfig = {
  url: 'https://jaspersoft.mediaprint.it/jasperserver/rest_v2/reports/Mediaprint/Clienti/Provincia_Ambiente/Layout/PIMP.pdf',
  user: 'jasperadmin',
  password: 'jasperadmin',
}

const normalizeHeader = (value) => String(value ?? '').trim()

const readWorkbookRows = (filePath) => {
  const workbook = xlsx.readFile(filePath, { cellStyles: true })
  const sheetName = workbook.SheetNames[0]
  if (!sheetName) {
    return []
  }

  const sheet = workbook.Sheets[sheetName]
  const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: '' })
  if (rows.length === 0) {
    return []
  }

  const headers = rows[0].map((header) => normalizeHeader(header))
  return rows.slice(1).map((row) => {
    const record = {}
    headers.forEach((header, index) => {
      if (header) {
        record[header] = row[index]
      }
    })
    return record
  })
}

function createWindow() {
  const isMac = process.platform === 'darwin'
  const win = new BrowserWindow({
    width: 1280,
    height: 820,
    minWidth: 1024,
    autoHideMenuBar: false,
    icon: path.join(__dirname, '../img/favicon.ico'),
    webPreferences: {
      preload: path.join(__dirname, '../preload/index.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  })

  win.loadFile(path.join(__dirname, '../renderer/index.html'))

  const menuTemplate = [
    ...(isMac
      ? [
          {
            label: app.name,
            submenu: [
              { role: 'about' },
              { type: 'separator' },
              { role: 'services' },
              { type: 'separator' },
              { role: 'hide' },
              { role: 'hideOthers' },
              { role: 'unhide' },
              { type: 'separator' },
              { role: 'quit' },
            ],
          },
        ]
      : []),
    {
      label: 'File',
      submenu: [
        isMac ? { role: 'close' } : { role: 'quit' },
      ],
    },
    {
      label: 'Visualizza',
      submenu: [
        { role: 'reload' },
        { role: 'forceReload' },
        { role: 'toggleDevTools' },
        { type: 'separator' },
        { role: 'resetZoom' },
        { role: 'zoomIn' },
        { role: 'zoomOut' },
        { type: 'separator' },
        { role: 'togglefullscreen' },
      ],
    },
  ]

  const menu = Menu.buildFromTemplate(menuTemplate)
  Menu.setApplicationMenu(menu)
  win.setMenuBarVisibility(true)
}

app.whenReady().then(createWindow)

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow()
  }
})

ipcMain.handle('excel:convert', async (_event, { filePath, metadata }) => {
  if (!filePath) {
    throw new Error('Impossibile trovare il file selezionato')
  }

  const ext = path.extname(filePath).toLowerCase()
  if (ext !== '.xlsx') {
    throw new Error('È richiesto un file .xlsx')
  }

  const workbook = xlsx.readFile(filePath, { cellStyles: true })

  const date1904 = workbook.Workbook?.WBProps?.date1904

  const workCode = metadata?.numero ? String(metadata.numero) : ''
  const isPimpAr = metadata?.operation === 'PIMP-AR'
  const baseBarcode = metadata?.barcode ? String(metadata.barcode).trim() : ''
  const baseBarcodeValue = baseBarcode ? Number(baseBarcode) : null

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName]
    const range = sheet['!ref'] ? xlsx.utils.decode_range(sheet['!ref']) : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } }
    const headerRow = range.s.r
    const dateColumns = new Set()
    const startDataRow = headerRow + 1
    const codeColumn = range.e.c + 1
    const progressColumn = range.e.c + 2
    const arCodeColumn = range.e.c + 3
    const arCode23Column = range.e.c + 4
    const barcode23Column = range.e.c + 5
    const omologazioneColumn = isPimpAr ? range.e.c + 6 : range.e.c + 3

    for (let col = range.s.c; col <= range.e.c; ++col) {
      const headerAddress = xlsx.utils.encode_cell({ r: headerRow, c: col })
      const headerCell = sheet[headerAddress]
      if (!headerCell || headerCell.v == null) continue
      const headerText = String(headerCell.v).toLowerCase()
      if (headerText.includes('data')) {
        dateColumns.add(col)
      }
    }

    const headerCodeAddress = xlsx.utils.encode_cell({ r: headerRow, c: codeColumn })
    const headerProgressAddress = xlsx.utils.encode_cell({ r: headerRow, c: progressColumn })
    sheet[headerCodeAddress] = { t: 's', v: 'Codice lavorazione' }
    sheet[headerProgressAddress] = { t: 's', v: 'Progressivo' }
    if (isPimpAr) {
      const headerArCode = xlsx.utils.encode_cell({ r: headerRow, c: arCodeColumn })
      const headerArCode23 = xlsx.utils.encode_cell({ r: headerRow, c: arCode23Column })
      const headerBarcode23 = xlsx.utils.encode_cell({ r: headerRow, c: barcode23Column })
      const headerOmolog = xlsx.utils.encode_cell({ r: headerRow, c: omologazioneColumn })
      sheet[headerArCode] = { t: 's', v: 'CODICE_AR' }
      sheet[headerArCode23] = { t: 's', v: 'CODICE_AR_23i' }
      sheet[headerBarcode23] = { t: 's', v: 'BARCODE23i' }
      sheet[headerOmolog] = { t: 's', v: 'OMOLOGAZIONE' }
    } else {
      const headerOmolog = xlsx.utils.encode_cell({ r: headerRow, c: omologazioneColumn })
      sheet[headerOmolog] = { t: 's', v: 'OMOLOGAZIONE' }
    }

    for (let row = range.s.r; row <= range.e.r; ++row) {
      for (let col = range.s.c; col <= range.e.c; ++col) {
        const address = xlsx.utils.encode_cell({ r: row, c: col })
        const cell = sheet[address]
        if (!cell) continue

        let textValue = cell.v != null ? String(cell.v) : ''
        if (dateColumns.has(col) && row !== headerRow) {
          const formattedDate = formatDateCell(cell, date1904)
          if (formattedDate) {
            textValue = formattedDate
          }
        }
        cell.t = 's'
        cell.v = textValue
        cell.w = textValue
      }
    }

    for (let row = startDataRow; row <= range.e.r; ++row) {
      const recordIndex = row - headerRow
      const codeAddress = xlsx.utils.encode_cell({ r: row, c: codeColumn })
      const progressAddress = xlsx.utils.encode_cell({ r: row, c: progressColumn })
      sheet[codeAddress] = { t: 's', v: workCode }
      sheet[progressAddress] = { t: 'n', v: recordIndex }

      if (isPimpAr && baseBarcodeValue != null && !Number.isNaN(baseBarcodeValue)) {
        const progressiveOffset = recordIndex - 1
        const codiceAr = String(baseBarcodeValue + progressiveOffset)
        const { codiceLoc, codiceAr23i, barcode23i } = computeBarcode23i(codiceAr)
        const arCodeAddress = xlsx.utils.encode_cell({ r: row, c: arCodeColumn })
        const arCode23Address = xlsx.utils.encode_cell({ r: row, c: arCode23Column })
        const barcode23Address = xlsx.utils.encode_cell({ r: row, c: barcode23Column })
        const omologAddress = xlsx.utils.encode_cell({ r: row, c: omologazioneColumn })
        sheet[arCodeAddress] = { t: 's', v: codiceLoc }
        sheet[arCode23Address] = { t: 's', v: codiceAr23i }
        sheet[barcode23Address] = { t: 's', v: barcode23i }
        sheet[omologAddress] = { t: 's', v: 'DCOCC0015' }
      } else if (!isPimpAr) {
        const omologAddress = xlsx.utils.encode_cell({ r: row, c: omologazioneColumn })
        sheet[omologAddress] = { t: 's', v: 'DCOOS2065' }
      }
    }

    const newRange = {
      s: range.s,
      e: { r: range.e.r, c: omologazioneColumn },
    }
    sheet['!ref'] = xlsx.utils.encode_range(newRange)
  })

  const outputPath = path.join(
    path.dirname(filePath),
    `${path.basename(filePath, ext)}-testo.xlsx`,
  )

  xlsx.writeFile(workbook, outputPath, { bookType: 'xlsx', cellStyles: true })

  return { outputPath }
})

ipcMain.handle('excel:reveal', async (_event, { filePath }) => {
  if (filePath) {
    shell.showItemInFolder(filePath)
  }
})

ipcMain.handle('excel:import-db', async (_event, { filePath, operation }) => {
  if (!filePath) {
    throw new Error('Impossibile trovare il file elaborato.')
  }

  const ext = path.extname(filePath).toLowerCase()
  if (ext !== '.xlsx') {
    throw new Error('È richiesto un file .xlsx')
  }

  const records = readWorkbookRows(filePath)
  if (!records.length) {
    return { inserted: 0 }
  }

  const columns =
    operation === 'PIMP-AR'
      ? [...baseDbColumns, ...pimpArDbColumns]
      : [...baseDbColumns, ...pimpDbColumns]

  const values = records.map((record) =>
    columns.map((column) => {
      const value = record[column]
      return value === '' ? null : value
    }),
  )

  const columnList = columns.map((column) => `\`${column}\``).join(', ')
  const placeholders = `(${columns.map(() => '?').join(', ')})`
  const sql = `INSERT INTO \`${dbTable}\` (${columnList}) VALUES ${values
    .map(() => placeholders)
    .join(', ')}`

  const connection = await mysql.createConnection(dbConfig)
  try {
    await connection.query(sql, values.flat())
  } finally {
    await connection.end()
  }

  return { inserted: values.length }
})

ipcMain.handle('excel:truncate-db', async () => {
  const connection = await mysql.createConnection(dbConfig)
  try {
    await connection.query(`TRUNCATE TABLE \`${dbTable}\``)
  } finally {
    await connection.end()
  }
  return { ok: true }
})

const downloadPdf = (url, authHeader, destination) =>
  new Promise((resolve, reject) => {
    const request = https.get(
      url,
      {
        headers: {
          Authorization: authHeader,
        },
      },
      (response) => {
        if (response.statusCode !== 200) {
          const status = response.statusCode || 'unknown'
          response.resume()
          reject(new Error(`Download Jasper fallito (HTTP ${status}).`))
          return
        }
        const fileStream = fs.createWriteStream(destination)
        response.pipe(fileStream)
        fileStream.on('finish', () => fileStream.close(resolve))
      },
    )

    request.on('error', (error) => reject(error))
  })

ipcMain.handle('jasper:download-pdf', async (_event, { filePath }) => {
  if (!filePath) {
    throw new Error('Impossibile determinare la cartella di destinazione.')
  }

  const dir = path.dirname(filePath)
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-')
  const savedPath = path.join(dir, `PIMP-${timestamp}.pdf`)
  const authHeader = `Basic ${Buffer.from(
    `${jasperConfig.user}:${jasperConfig.password}`,
  ).toString('base64')}`

  await downloadPdf(jasperConfig.url, authHeader, savedPath)
  shell.showItemInFolder(savedPath)
  return { savedPath }
})
