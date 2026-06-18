const __req = (() => {
  try {
    if (typeof require === 'function') return require
  } catch (e) {}
  return null
})()

const fs = __req ? __req('fs') : null
const path = __req ? __req('path') : null

function getDataDir() {
  try {
    const arg = (process.argv || []).find((a) => String(a).startsWith('--data-dir='))
    if (arg) return arg.replace('--data-dir=', '')
  } catch (e) {}
  return path ? path.join(process.cwd(), 'data') : 'data'
}

function filePath(name) {
  const dataDir = getDataDir()
  if (fs && !fs.existsSync(dataDir)) fs.mkdirSync(dataDir, { recursive: true })
  return path.join(dataDir, name)
}

function readJSON(name, fallback) {
  try {
    const p = filePath(name)
    if (!fs.existsSync(p)) return fallback
    const raw = fs.readFileSync(p, 'utf8').trim()
    return raw ? JSON.parse(raw) : fallback
  } catch (e) {
    return fallback
  }
}

function writeJSON(name, value) {
  try {
    fs.writeFileSync(filePath(name), JSON.stringify(value, null, 2), 'utf8')
    return true
  } catch (e) {
    return false
  }
}

function uid(prefix = 'id') {
  return `${prefix}_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`
}

function fmtMoney(v) {
  const n = Number(v || 0)
  return n.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })
}

function toNumber(v) {
  return Number(String(v || '0').replace(/\./g, '').replace(',', '.')) || 0
}

function todayISO() {
  return new Date().toISOString().slice(0, 10)
}

function daysBetween(a, b) {
  const d1 = new Date(a + 'T00:00:00')
  const d2 = new Date(b + 'T00:00:00')
  return Math.ceil((d2 - d1) / 86400000)
}

function parseMonthYear(monthYear) {
  const [year, month] = String(monthYear || '').split('-').map(Number)
  if (!year || !month) return null
  return { year, month }
}

window.store = {
  readJSON,
  writeJSON,
  uid,
  fmtMoney,
  toNumber,
  todayISO,
  daysBetween,
  parseMonthYear
}

