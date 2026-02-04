const operationDescription = {
  PIMP: 'Elabora i dati PIMP nativi (Numero a 6 cifre + riferimento cliente a 3 cifre).',
  'PIMP-AR': 'Modalità AR: abilitata per tener conto delle eventuali variazioni (solo file in ingresso).',
}

const form = document.getElementById('pimp-form')
const operationSelect = document.getElementById('operation')
const operationHelp = document.getElementById('operation-help')
const pimpFields = document.getElementById('pimp-fields')
const barcodeField = document.getElementById('barcode-field')
const numeroInput = document.getElementById('numero')
const clienteInput = document.getElementById('cliente')
const barcodeInput = document.getElementById('barcode')
const fileInput = document.getElementById('file-input')
const submitBtn = document.getElementById('submit-btn')
const openFolderBtn = document.getElementById('open-folder')
const importDbBtn = document.getElementById('import-db')
const truncateDbBtn = document.getElementById('truncate-db')
const downloadPdfBtn = document.getElementById('download-pdf')
const resetBtn = document.getElementById('reset-form')
const statusRow = document.getElementById('status-row')
const statusAlert = document.getElementById('status-alert')
const outputRow = document.getElementById('output-row')
const outputPathEl = document.getElementById('output-path')

const electronBridge = window.electronApi || {}
let outputPath = ''

const setStatus = (type, message) => {
  statusAlert.className = `alert alert-${type}`
  statusAlert.textContent = message
  statusRow.classList.remove('d-none')
}

const clearStatus = () => {
  statusRow.classList.add('d-none')
  statusAlert.textContent = ''
}

const updateOperation = () => {
  const operation = operationSelect.value
  operationHelp.textContent = operationDescription[operation] || ''
  pimpFields.classList.remove('d-none')
  numeroInput.required = true
  clienteInput.required = true
  if (operation === 'PIMP-AR') {
    barcodeField.classList.remove('d-none')
    barcodeInput.required = true
  } else {
    barcodeField.classList.add('d-none')
    barcodeInput.required = false
  }
}

const setProcessing = (isProcessing) => {
  submitBtn.disabled = isProcessing
  submitBtn.innerHTML = isProcessing ? '<i class="fa-solid fa-spinner"></i>' : '<i class="fa-solid fa-play"></i> <i class="fa-sharp fa-regular fa-file-excel"></i>'
}

form.addEventListener('submit', async (event) => {
  event.preventDefault()
  clearStatus()
  outputPath = ''
  outputRow.classList.add('d-none')
  openFolderBtn.classList.add('d-none')
  importDbBtn.classList.add('d-none')
  downloadPdfBtn.classList.add('d-none')

  const operation = operationSelect.value
  const file = fileInput.files?.[0]

  if (!file) {
    setStatus('danger', 'Carica un file .xlsx prima di avviare l’elaborazione.')
    return
  }

  if (!/^[0-9]{6}$/.test(numeroInput.value)) {
    setStatus('danger', 'Il Numero deve essere composto da esattamente 6 cifre.')
    return
  }
  if (!/^[0-9]{3}$/.test(clienteInput.value)) {
    setStatus('danger', 'Il riferimento cliente deve contenere 3 cifre.')
    return
  }
  if (operation === 'PIMP-AR' && !barcodeInput.value.trim()) {
    setStatus('danger', 'Inserisci il barcode postale iniziale.')
    return
  }

  if (!file.name.toLowerCase().endsWith('.xlsx')) {
    setStatus('danger', 'Sono accettati solo file con estensione .xlsx.')
    return
  }

  if (!electronBridge.convertWorkbook) {
    setStatus('danger', 'Electron API non disponibile: avvia l’app con Electron.')
    return
  }

  setProcessing(true)
  try {
    const response = await electronBridge.convertWorkbook({
      filePath: file.path,
      metadata: {
        operation,
        numero: numeroInput.value.trim(),
        cliente: clienteInput.value.trim(),
        barcode: barcodeInput.value.trim(),
      },
    })
    outputPath = response.outputPath
    setStatus('success', 'Elaborazione completata: il file con celle testuali è pronto.')
    outputPathEl.textContent = `Output: ${outputPath}`
    outputRow.classList.remove('d-none')
    openFolderBtn.classList.remove('d-none')
    importDbBtn.classList.remove('d-none')
    downloadPdfBtn.classList.remove('d-none')
  } catch (error) {
    setStatus('danger', error?.message ?? 'Si è verificato un errore durante l’elaborazione.')
  } finally {
    setProcessing(false)
  }
})

openFolderBtn.addEventListener('click', () => {
  if (outputPath && electronBridge.revealFile) {
    electronBridge.revealFile({ filePath: outputPath })
  }
})

importDbBtn.addEventListener('click', async () => {
  if (!outputPath) {
    setStatus('danger', 'Nessun file elaborato da caricare.')
    return
  }
  if (!electronBridge.importWorkbookToDb) {
    setStatus('danger', 'Electron API non disponibile: avvia l’app con Electron.')
    return
  }

  setProcessing(true)
  try {
    const response = await electronBridge.importWorkbookToDb({
      filePath: outputPath,
      operation: operationSelect.value,
    })
    const inserted = response?.inserted ?? 0
    setStatus('success', `Caricamento su DB completato. Righe inserite: ${inserted}.`)
  } catch (error) {
    setStatus('danger', error?.message ?? 'Errore durante il caricamento su DB.')
  } finally {
    setProcessing(false)
  }
})

truncateDbBtn.addEventListener('click', async () => {
  if (!electronBridge.truncateTable) {
    setStatus('danger', 'Electron API non disponibile: avvia l’app con Electron.')
    return
  }

  const confirmDelete = window.confirm('Confermi lo svuotamento della tabella PIMP?')
  if (!confirmDelete) return

  setProcessing(true)
  try {
    await electronBridge.truncateTable()
    setStatus('success', 'Tabella PIMP svuotata correttamente.')
  } catch (error) {
    setStatus('danger', error?.message ?? 'Errore durante lo svuotamento della tabella.')
  } finally {
    setProcessing(false)
  }
})

downloadPdfBtn.addEventListener('click', async () => {
  if (!outputPath) {
    setStatus('danger', 'Nessun file elaborato per determinare la cartella di destinazione.')
    return
  }
  if (!electronBridge.downloadJasperPdf) {
    setStatus('danger', 'Electron API non disponibile: avvia l’app con Electron.')
    return
  }

  setProcessing(true)
  try {
    const response = await electronBridge.downloadJasperPdf({ filePath: outputPath })
    const savedPath = response?.savedPath
    if (savedPath) {
      setStatus('success', `PDF salvato: ${savedPath}`)
    } else {
      setStatus('success', 'PDF scaricato correttamente.')
    }
  } catch (error) {
    setStatus('danger', error?.message ?? 'Errore durante il download del PDF.')
  } finally {
    setProcessing(false)
  }
})

operationSelect.addEventListener('change', updateOperation)
resetBtn.addEventListener('click', () => {
  form.reset()
  outputPath = ''
  clearStatus()
  outputRow.classList.add('d-none')
  updateOperation()
  setProcessing(false)
})
updateOperation()
