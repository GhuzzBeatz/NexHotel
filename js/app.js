const DB_FILE = "nexhotel_db.json"
const XLSX = require("xlsx")
const appFs = require("fs")
const appPath = require("path")
const appOs = require("os")
const electronShell = require("electron").shell

const dbDefault = {
  guests: [],
  rooms: [],
  reservations: [],
  finance: [],
  staff: [],
  stockItems: [],
  stockMovements: [],
  serviceOrders: [],
  timeEntries: [],
  bankAccounts: [],
  notifications: [],
  auditLogs: [],
  guestRequests: []
}

let db = loadDb()
let latestReportRows = []
let portalReservationId = null
let guestLookupMatches = []
let hostGuestLookupMatches = []
let dialogResolver = null
let transferDialogResolver = null
let hostDialogRoomId = ""
let paymentDialogReservationId = ""

function loadDb() {
  const loaded = window.store.readJSON(DB_FILE, null)
  return loaded && typeof loaded === "object"
    ? { ...dbDefault, ...loaded }
    : structuredClone(dbDefault)
}

function saveDb() {
  const ok = window.store.writeJSON(DB_FILE, db)
  if (!ok) toast("Falha ao salvar dados locais.")
  return ok
}

function byId(list, id) {
  return list.find((x) => String(x.id) === String(id))
}

function normal(v) {
  return String(v || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
}

function digitsOnly(v) {
  return String(v || "").replace(/\D/g, "")
}

function fieldText(v) {
  return String(v || "").trim()
}

function normalizeHeader(v) {
  return normal(v).replace(/[^a-z0-9]/g, "")
}

function firstFilled(...values) {
  return values.map(fieldText).find(Boolean) || ""
}

function guestLabel(guest) {
  if (!guest) return ""
  const doc = guest.doc ? `Doc: ${guest.doc}` : "Sem documento"
  const phone = guest.phone ? `Tel: ${guest.phone}` : ""
  return [guest.name, doc, phone].filter(Boolean).join(" | ")
}

function guestSearchText(guest) {
  return normal(`${guest.name} ${guest.doc} ${guest.phone} ${guest.email} ${guest.city} ${digitsOnly(guest.doc)} ${digitsOnly(guest.phone)}`)
}

function sameGuestIdentity(a, b) {
  const aDoc = digitsOnly(a.doc)
  const bDoc = digitsOnly(b.doc)
  if (aDoc && bDoc && aDoc === bDoc) return true

  const aEmail = normal(a.email)
  const bEmail = normal(b.email)
  if (aEmail && bEmail && aEmail === bEmail) return true

  const aPhone = digitsOnly(a.phone)
  const bPhone = digitsOnly(b.phone)
  if (aPhone && bPhone && aPhone === bPhone) return true

  const aCity = normal(a.city)
  const bCity = normal(b.city)
  return Boolean(normal(a.name) && aCity && bCity && normal(a.name) === normal(b.name) && aCity === bCity)
}

function findDuplicateGuest(candidate, ignoreId = "") {
  return db.guests.find((guest) => String(guest.id) !== String(ignoreId) && sameGuestIdentity(guest, candidate))
}

function excelDateToISO(value) {
  if (!value) return ""
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10)
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value)
    if (!parsed) return ""
    return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`
  }
  const raw = fieldText(value)
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw
  const br = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/)
  if (br) return `${br[3]}-${String(br[2]).padStart(2, "0")}-${String(br[1]).padStart(2, "0")}`
  return raw
}

function makeGuestFromRow(row) {
  const byHeader = {}
  Object.entries(row || {}).forEach(([key, value]) => {
    byHeader[normalizeHeader(key)] = value
  })

  return {
    id: window.store.uid("guest"),
    name: firstFilled(byHeader.nome, byHeader.nomecompleto, byHeader.cliente, byHeader.hospede),
    doc: firstFilled(byHeader.documento, byHeader.doc, byHeader.cpf, byHeader.rg),
    phone: firstFilled(byHeader.telefone, byHeader.tel, byHeader.whatsapp, byHeader.celular),
    email: firstFilled(byHeader.email, byHeader.emailcliente),
    birth: excelDateToISO(firstFilled(byHeader.datanascimento, byHeader.nascimento, byHeader.aniversario)),
    city: firstFilled(byHeader.cidade, byHeader.municipio),
    notes: firstFilled(byHeader.observacoes, byHeader.observacao, byHeader.obs)
  }
}

function guestExportRows() {
  return db.guests.map((guest) => ({
    Nome: guest.name || "",
    Documento: guest.doc || "",
    Telefone: guest.phone || "",
    Email: guest.email || "",
    "Data Nascimento": guest.birth || "",
    Cidade: guest.city || "",
    Observacoes: guest.notes || ""
  }))
}

function saveWorkbookToDownloads(workbook, fileName) {
  const downloads = appPath.join(appOs.homedir(), "Downloads")
  if (!appFs.existsSync(downloads)) appFs.mkdirSync(downloads, { recursive: true })
  const outPath = appPath.join(downloads, fileName)
  XLSX.writeFile(workbook, outPath)
  return outPath
}

function escapeHtml(text) {
  return String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;")
}

function toNumber(v) {
  return window.store.toNumber(v)
}

function money(v) {
  return window.store.fmtMoney(Number(v || 0))
}

function todayISO() {
  return window.store.todayISO()
}

function addDaysISO(isoDate, days) {
  const date = new Date(`${isoDate}T00:00:00`)
  date.setDate(date.getDate() + days)
  return date.toISOString().slice(0, 10)
}

function toLocalDateTimeInput(date = new Date()) {
  const d = new Date(date.getTime() - date.getTimezoneOffset() * 60000)
  return d.toISOString().slice(0, 16)
}

function fmtDate(isoDate) {
  if (!isoDate) return "-"
  const value = String(isoDate).slice(0, 10)
  const [y, m, d] = value.split("-")
  if (!y || !m || !d) return "-"
  return `${d}/${m}/${y}`
}

function fmtDateTime(isoDateTime) {
  if (!isoDateTime) return "-"
  const dt = new Date(isoDateTime)
  if (Number.isNaN(dt.getTime())) return String(isoDateTime)
  return dt.toLocaleString("pt-BR")
}

function toast(msg) {
  const el = document.getElementById("toast")
  el.textContent = msg
  el.hidden = false
  clearTimeout(toast.timer)
  toast.timer = setTimeout(() => {
    el.hidden = true
  }, 2200)
}

function closeAppDialog(result) {
  const dialog = document.getElementById("appDialog")
  dialog.hidden = true
  if (dialogResolver) {
    dialogResolver(result)
    dialogResolver = null
  }
}

function askConfirm(message, title = "Confirmar acao", okText = "Confirmar", cancelText = "Cancelar") {
  return new Promise((resolve) => {
    const dialog = document.getElementById("appDialog")
    const titleEl = document.getElementById("appDialogTitle")
    const messageEl = document.getElementById("appDialogMessage")
    const okBtn = document.getElementById("appDialogOk")
    const cancelBtn = document.getElementById("appDialogCancel")

    dialogResolver = resolve
    titleEl.textContent = title
    messageEl.textContent = message
    okBtn.textContent = okText
    cancelBtn.textContent = cancelText
    dialog.hidden = false
    okBtn.focus()
  })
}

function closeTransferDialog(roomId = "") {
  const dialog = document.getElementById("transferDialog")
  if (dialog) dialog.hidden = true
  if (transferDialogResolver) {
    transferDialogResolver(roomId)
    transferDialogResolver = null
  }
}

function askTransferRoom(reservation) {
  return new Promise((resolve) => {
    const dialog = document.getElementById("transferDialog")
    const infoEl = document.getElementById("transferDialogInfo")
    const selectEl = document.getElementById("transferRoomSelect")
    const helpEl = document.getElementById("transferDialogHelp")
    const okBtn = document.getElementById("transferDialogOk")
    const cancelBtn = document.getElementById("transferDialogCancel")
    const guest = byId(db.guests, reservation.guestId)
    const currentRoom = byId(db.rooms, reservation.roomId)
    const rooms = availableRoomsForTransfer(reservation)

    transferDialogResolver = resolve
    infoEl.textContent =
      `${reservation.code} - ${guest?.name || "Hospede"}\n` +
      `Quarto atual: ${currentRoom?.number || "-"}\n` +
      `Periodo: ${fmtDate(reservation.checkin)} ate ${fmtDate(reservation.checkout)}`

    if (!rooms.length) {
      selectEl.innerHTML = '<option value="">Nenhum quarto livre para este periodo</option>'
      helpEl.textContent = "Cadastre/libere outro quarto ou ajuste as reservas antes de transferir."
      okBtn.disabled = true
    } else {
      selectEl.innerHTML = rooms
        .map((room) => {
          const label = `Quarto ${room.number} - ${room.category} - ${money(room.rate)}`
          return `<option value="${room.id}">${escapeHtml(label)}</option>`
        })
        .join("")
      helpEl.textContent = "A diaria original sera mantida para nao alterar a conta do hospede."
      okBtn.disabled = false
    }

    dialog.hidden = false
    if (rooms.length) selectEl.focus()
    else cancelBtn.focus()
  })
}

function closeHostDialog() {
  document.getElementById("hostDialog").hidden = true
  hostDialogRoomId = ""
}

function resetHostDialog(room) {
  document.getElementById("hostGuestSearch").value = ""
  document.getElementById("hostGuestId").value = ""
  document.getElementById("hostGuestResults").classList.remove("active")
  clearFields(["hostNewName", "hostNewDoc", "hostNewPhone", "hostNewEmail", "hostExtra", "hostPaymentAmount", "hostNotes"])
  document.getElementById("hostCheckin").value = todayISO()
  document.getElementById("hostCheckout").value = addDaysISO(todayISO(), 1)
  document.getElementById("hostAdults").value = "1"
  document.getElementById("hostChildren").value = "0"
  document.getElementById("hostDaily").value = toNumber(room?.rate).toString().replace(".", ",")
  document.getElementById("hostPaymentMethod").value = "PIX"
  document.getElementById("hostPaymentAccount").selectedIndex = 0
}

function submitDirectHosting() {
  const room = byId(db.rooms, hostDialogRoomId)
  if (!room) return toast("Quarto invalido.")
  if (roomLiveStatus(room) !== "DISPONIVEL") return toast("Quarto nao esta disponivel para hospedar.")

  let guestId = document.getElementById("hostGuestId").value
  const newGuest = {
    id: window.store.uid("guest"),
    name: fieldText(document.getElementById("hostNewName").value),
    doc: fieldText(document.getElementById("hostNewDoc").value),
    phone: fieldText(document.getElementById("hostNewPhone").value),
    email: fieldText(document.getElementById("hostNewEmail").value),
    birth: "",
    city: "",
    notes: "Cadastrado no check-in rapido"
  }

  if (!guestId) {
    if (!newGuest.name) return toast("Selecione um cliente ou informe o nome do novo hospede.")
    const duplicate = findDuplicateGuest(newGuest)
    if (duplicate) return toast(`Cliente ja cadastrado: ${duplicate.name}. Busque e selecione no campo acima.`)
    db.guests.push(newGuest)
    guestId = newGuest.id
  }

  const checkin = document.getElementById("hostCheckin").value || todayISO()
  const checkout = document.getElementById("hostCheckout").value
  if (!checkout) return toast("Informe a data de check-out.")
  if (window.store.daysBetween(checkin, checkout) < 1) return toast("Check-out deve ser apos check-in.")
  if (roomHasReservationConflict(room.id, checkin, checkout)) return toast("Conflito de periodo para este quarto.")

  const code = reservationCode()
  const reservation = {
    id: window.store.uid("res"),
    code,
    guestId,
    roomId: room.id,
    checkin,
    checkout,
    adults: toNumber(document.getElementById("hostAdults").value) || 1,
    children: toNumber(document.getElementById("hostChildren").value) || 0,
    dailyRate: toNumber(document.getElementById("hostDaily").value) || toNumber(room.rate),
    extra: toNumber(document.getElementById("hostExtra").value),
    status: "HOSPEDADO",
    channel: "BALCAO",
    notes: fieldText(document.getElementById("hostNotes").value),
    createdAt: new Date().toISOString(),
    checkedInAt: new Date().toISOString()
  }
  db.reservations.push(reservation)

  const paymentAmount = toNumber(document.getElementById("hostPaymentAmount").value)
  if (paymentAmount > 0) {
    addReservationPayment(
      reservation.id,
      paymentAmount,
      document.getElementById("hostPaymentMethod").value,
      document.getElementById("hostPaymentAccount").value,
      `Pagamento inicial no check-in rapido - ${code}`,
      "CHECKIN_RAPIDO"
    )
  }

  const guest = byId(db.guests, guestId)
  addNotification("Hospede hospedado", `${guest?.name || "Hospede"} entrou no quarto ${room.number}.`, "INFO", "RESERVA", reservation.id)
  closeHostDialog()
  saveAndRefresh("Check-in rapido", `${code} | Quarto ${room.number} | ${guest?.name || "Hospede"}`)
  showPage("hospedados")
  toast("Hospedagem criada com sucesso.")
}

function closePaymentDialog() {
  document.getElementById("paymentDialog").hidden = true
  paymentDialogReservationId = ""
}

function submitQuickPayment() {
  const reservation = byId(db.reservations, paymentDialogReservationId)
  if (!reservation) return toast("Reserva nao encontrada para pagamento.")
  const amount = toNumber(document.getElementById("quickPaymentAmount").value)
  if (amount <= 0) return toast("Informe um valor de pagamento valido.")
  const method = document.getElementById("quickPaymentMethod").value
  const accountId = document.getElementById("quickPaymentAccount").value
  const description = fieldText(document.getElementById("quickPaymentDesc").value) || `Pagamento recebido da reserva ${reservation.code}`
  const entry = addReservationPayment(reservation.id, amount, method, accountId, description, "PAGAMENTO_RAPIDO")
  if (!entry) return toast("Nao foi possivel registrar o pagamento.")

  addNotification("Pagamento recebido", `${reservation.code}: ${money(amount)} recebido.`, "INFO", "FINANCE", entry.id)
  closePaymentDialog()
  saveAndRefresh("Pagamento de hospedagem", `${reservation.code} | ${money(amount)} | ${method}`)
  toast("Pagamento adicionado.")
}

function bindDialogEvents() {
  document.getElementById("appDialogOk").addEventListener("click", () => closeAppDialog(true))
  document.getElementById("appDialogCancel").addEventListener("click", () => closeAppDialog(false))
  document.getElementById("appDialog").addEventListener("click", (event) => {
    if (event.target.id === "appDialog") closeAppDialog(false)
  })
  document.getElementById("transferDialogOk").addEventListener("click", () => {
    closeTransferDialog(document.getElementById("transferRoomSelect").value)
  })
  document.getElementById("transferDialogCancel").addEventListener("click", () => closeTransferDialog(""))
  document.getElementById("transferDialog").addEventListener("click", (event) => {
    if (event.target.id === "transferDialog") closeTransferDialog("")
  })
  document.getElementById("hostDialogOk").addEventListener("click", submitDirectHosting)
  document.getElementById("hostDialogCancel").addEventListener("click", closeHostDialog)
  document.getElementById("hostDialog").addEventListener("click", (event) => {
    if (event.target.id === "hostDialog") closeHostDialog()
  })
  document.getElementById("paymentDialogOk").addEventListener("click", submitQuickPayment)
  document.getElementById("paymentDialogCancel").addEventListener("click", closePaymentDialog)
  document.getElementById("paymentDialog").addEventListener("click", (event) => {
    if (event.target.id === "paymentDialog") closePaymentDialog()
  })
  document.getElementById("hostGuestSearch").addEventListener("input", () => {
    document.getElementById("hostGuestId").value = ""
    renderHostGuestLookup()
  })
  document.getElementById("hostGuestSearch").addEventListener("focus", renderHostGuestLookup)
  document.getElementById("hostGuestSearch").addEventListener("blur", hideHostGuestLookupSoon)
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && !document.getElementById("appDialog").hidden) closeAppDialog(false)
    if (event.key === "Escape" && !document.getElementById("transferDialog").hidden) closeTransferDialog("")
    if (event.key === "Escape" && !document.getElementById("hostDialog").hidden) closeHostDialog()
    if (event.key === "Escape" && !document.getElementById("paymentDialog").hidden) closePaymentDialog()
  })
}

function showPage(page) {
  document.querySelectorAll(".page").forEach((el) => el.classList.remove("active"))
  document.querySelectorAll(".nav-btn").forEach((el) => el.classList.remove("active"))
  document.getElementById(`page-${page}`)?.classList.add("active")
  document.querySelector(`.nav-btn[data-page="${page}"]`)?.classList.add("active")
}

function ensureLogoFallback() {
  const logo = document.getElementById("brandLogo")
  const fallback = document.getElementById("brandFallback")
  if (!logo || !fallback) return
  logo.addEventListener("error", () => {
    logo.style.display = "none"
    fallback.style.display = "grid"
  })
}

function addAudit(action, details = "") {
  db.auditLogs.unshift({
    id: window.store.uid("audit"),
    when: new Date().toISOString(),
    action,
    details
  })
  if (db.auditLogs.length > 4000) db.auditLogs.pop()
}

function addNotification(title, message, level = "INFO", source = "SYSTEM", refId = "") {
  db.notifications.unshift({
    id: window.store.uid("ntf"),
    when: new Date().toISOString(),
    title,
    message,
    level,
    source,
    refId,
    read: false
  })
  if (db.notifications.length > 2500) db.notifications.pop()
}

function reservationCode() {
  const token = `${Date.now()}`.slice(-6)
  return `RSV-${token}`
}

function orderCode() {
  return `OS-${String(db.serviceOrders.length + 1).padStart(4, "0")}`
}

function reservationNights(reservation) {
  const days = window.store.daysBetween(reservation.checkin, reservation.checkout)
  return Math.max(1, days)
}

function reservationDailyRate(reservation) {
  if (toNumber(reservation.dailyRate) > 0) return toNumber(reservation.dailyRate)
  const room = byId(db.rooms, reservation.roomId)
  return room ? toNumber(room.rate) : 0
}

function reservationTotal(reservation) {
  return reservationNights(reservation) * reservationDailyRate(reservation) + toNumber(reservation.extra)
}

function reservationPaid(reservationId) {
  return db.finance
    .filter((x) => isFinanceActive(x) && x.type === "RECEITA" && String(x.reservationId) === String(reservationId))
    .reduce((sum, x) => sum + toNumber(x.amount), 0)
}

function reservationPending(reservation) {
  return Math.max(0, reservationTotal(reservation) - reservationPaid(reservation.id))
}

function roomLiveStatus(room) {
  if (room.baseStatus === "MANUTENCAO") return "MANUTENCAO"
  const hasHosting = db.reservations.some(
    (r) => String(r.roomId) === String(room.id) && r.status === "HOSPEDADO"
  )
  return hasHosting ? "OCUPADO" : "DISPONIVEL"
}

function activeHostingForRoom(roomId) {
  return db.reservations.find((r) => String(r.roomId) === String(roomId) && r.status === "HOSPEDADO")
}

function hostedReservations() {
  return db.reservations
    .filter((r) => r.status === "HOSPEDADO")
    .slice()
    .sort((a, b) => String(a.checkout).localeCompare(String(b.checkout)))
}

function isFinanceActive(entry) {
  return String(entry?.status || "ATIVO").toUpperCase() !== "CANCELADO"
}

function activeFinanceEntries() {
  return db.finance.filter(isFinanceActive)
}

function financeSignedAmount(entry) {
  return entry.type === "DESPESA" ? -toNumber(entry.amount) : toNumber(entry.amount)
}

function reservationPaymentMethods(reservationId) {
  const methods = activeFinanceEntries()
    .filter((entry) => entry.type === "RECEITA" && String(entry.reservationId) === String(reservationId))
    .map((entry) => entry.method)
    .filter(Boolean)
  return [...new Set(methods)].join(", ") || "-"
}

function addReservationPayment(reservationId, amount, method, accountId = "", description = "Pagamento da hospedagem", origin = "PAGAMENTO_RAPIDO") {
  const reservation = byId(db.reservations, reservationId)
  if (!reservation || amount <= 0) return null
  const entry = {
    id: window.store.uid("fin"),
    type: "RECEITA",
    category: "pagamento_hospedagem",
    method,
    accountId: accountId || "",
    reservationId: reservation.id,
    date: todayISO(),
    amount,
    description,
    status: "ATIVO",
    origin,
    createdAt: new Date().toISOString()
  }
  db.finance.push(entry)
  return entry
}

function roomHasReservationConflict(roomId, checkin, checkout, ignoreReservationId = "") {
  return db.reservations.some((reservation) => {
    if (String(reservation.id) === String(ignoreReservationId)) return false
    if (String(reservation.roomId) !== String(roomId)) return false
    if (["CANCELADA", "CHECKOUT"].includes(reservation.status)) return false
    return !(checkout <= reservation.checkin || checkin >= reservation.checkout)
  })
}

function availableRoomsForTransfer(reservation) {
  return db.rooms
    .filter((room) => String(room.id) !== String(reservation.roomId))
    .filter((room) => room.baseStatus !== "MANUTENCAO")
    .filter((room) => !roomHasReservationConflict(room.id, reservation.checkin, reservation.checkout, reservation.id))
    .sort((a, b) => String(a.number || "").localeCompare(String(b.number || ""), "pt-BR", { numeric: true }))
}

function statusTag(type, value) {
  const v = String(value || "").toUpperCase()
  if (type === "reservation") {
    if (v === "HOSPEDADO" || v === "CHECKOUT") return `<span class="tag ok">${value}</span>`
    if (v === "CANCELADA") return `<span class="tag bad">${value}</span>`
    if (v === "AGENDADA") return `<span class="tag info">${value}</span>`
    return `<span class="tag warn">${value}</span>`
  }
  if (type === "room") {
    if (v === "DISPONIVEL") return `<span class="tag ok">${value}</span>`
    if (v === "OCUPADO") return `<span class="tag warn">${value}</span>`
    return `<span class="tag bad">${value}</span>`
  }
  if (type === "level") {
    if (v === "INFO") return `<span class="tag info">${value}</span>`
    if (v === "WARN") return `<span class="tag warn">${value}</span>`
    return `<span class="tag bad">${value}</span>`
  }
  if (type === "staff") {
    if (v === "ATIVO") return `<span class="tag ok">${value}</span>`
    return `<span class="tag bad">${value}</span>`
  }
  if (type === "finance") {
    if (v === "CANCELADO") return `<span class="tag bad">CANCELADO</span>`
    return `<span class="tag ok">ATIVO</span>`
  }
  return `<span class="tag info">${value}</span>`
}

function getSystemAlerts() {
  const alerts = []

  const lowStock = db.stockItems.filter((item) => toNumber(item.qty) <= toNumber(item.minQty))
  lowStock.forEach((item) => {
    alerts.push({
      type: "Estoque",
      message: `${item.name} abaixo do minimo (${item.qty}/${item.minQty})`,
      priority: "ALTA"
    })
  })

  db.reservations
    .filter((r) => r.status === "HOSPEDADO" || r.status === "CONFIRMADA")
    .forEach((r) => {
      const d = window.store.daysBetween(todayISO(), r.checkout)
      if (d >= 0 && d <= 3) {
        const guest = byId(db.guests, r.guestId)
        alerts.push({
          type: "Check-out",
          message: `${r.code} - ${guest?.name || "Hospede"} em ${d} dia(s)`,
          priority: d === 0 ? "URGENTE" : "MEDIA"
        })
      }
    })

  db.reservations
    .filter((r) => r.status === "CHECKOUT")
    .forEach((r) => {
      const pending = reservationPending(r)
      if (pending > 0) {
        alerts.push({
          type: "Financeiro",
          message: `${r.code} com saldo pendente ${money(pending)}`,
          priority: "ALTA"
        })
      }
    })

  if (!alerts.length) {
    alerts.push({
      type: "Sistema",
      message: "Sem alertas operacionais no momento.",
      priority: "BAIXA"
    })
  }
  return alerts
}

function fillSelects() {
  const guestOptions = db.guests.length
    ? db.guests.map((g) => `<option value="${g.id}">${escapeHtml(g.name)}</option>`).join("")
    : '<option value="">Cadastre clientes primeiro</option>'

  const roomOptions = db.rooms.length
    ? db.rooms.map((r) => `<option value="${r.id}">Quarto ${escapeHtml(r.number)} - ${escapeHtml(r.category)}</option>`).join("")
    : '<option value="">Cadastre quartos primeiro</option>'

  const accountOptions = db.bankAccounts.length
    ? db.bankAccounts.map((a) => `<option value="${a.id}">${escapeHtml(a.name)}</option>`).join("")
    : '<option value="">Sem conta cadastrada</option>'

  const reservationOptions = db.reservations.length
    ? [`<option value="">Sem vinculo</option>`]
        .concat(
          db.reservations.map((r) => {
            const g = byId(db.guests, r.guestId)
            return `<option value="${r.id}">${escapeHtml(r.code)} - ${escapeHtml(g?.name || "Hospede")}</option>`
          })
        )
        .join("")
    : '<option value="">Sem reserva cadastrada</option>'

  const itemOptions = db.stockItems.length
    ? db.stockItems.map((i) => `<option value="${i.id}">${escapeHtml(i.name)}</option>`).join("")
    : '<option value="">Cadastre itens primeiro</option>'

  const activeStaff = db.staff.filter((s) => s.status === "ATIVO")
  const staffOptions = activeStaff.length
    ? activeStaff.map((s) => `<option value="${s.id}">${escapeHtml(s.name)}</option>`).join("")
    : '<option value="">Cadastre funcionarios ativos</option>'

  document.getElementById("resRoom").innerHTML = roomOptions
  document.getElementById("finAccount").innerHTML = accountOptions
  document.getElementById("finReservation").innerHTML = reservationOptions
  document.getElementById("quickPaymentAccount").innerHTML = accountOptions
  document.getElementById("hostPaymentAccount").innerHTML = accountOptions
  document.getElementById("movItem").innerHTML = itemOptions
  document.getElementById("pointStaff").innerHTML = staffOptions
  document.getElementById("osRoom").innerHTML = roomOptions
  document.getElementById("osGuest").innerHTML = guestOptions
  document.getElementById("osAssignee").innerHTML = staffOptions
  document.getElementById("staffPanelUser").innerHTML = staffOptions
}

function renderGuestLookup() {
  const input = document.getElementById("resGuestSearch")
  const box = document.getElementById("resGuestResults")
  if (!input || !box) return

  const q = normal(input.value)
  guestLookupMatches = (q ? db.guests.filter((guest) => guestSearchText(guest).includes(q)) : db.guests)
    .slice()
    .sort((a, b) => String(a.name || "").localeCompare(String(b.name || ""), "pt-BR"))
    .slice(0, 50)

  if (!guestLookupMatches.length) {
    box.innerHTML = '<button class="lookup-option" type="button"><strong>Nenhum cliente encontrado</strong><span>Cadastre ou importe clientes na aba Clientes.</span></button>'
    box.classList.add("active")
    return
  }

  box.innerHTML = guestLookupMatches
    .map((guest) => {
      return `
        <button class="lookup-option" type="button" onclick="selectReservationGuest('${guest.id}')">
          <strong>${escapeHtml(guest.name)}</strong>
          <span>${escapeHtml([guest.doc || "Sem documento", guest.phone || "", guest.email || ""].filter(Boolean).join(" | "))}</span>
        </button>`
    })
    .join("")
  box.classList.add("active")
}

function hideGuestLookupSoon() {
  setTimeout(() => {
    document.getElementById("resGuestResults")?.classList.remove("active")
  }, 160)
}

function renderHostGuestLookup() {
  const input = document.getElementById("hostGuestSearch")
  const box = document.getElementById("hostGuestResults")
  if (!input || !box) return

  const q = normal(input.value)
  hostGuestLookupMatches = (q ? db.guests.filter((guest) => guestSearchText(guest).includes(q)) : db.guests)
    .slice()
    .sort((a, b) => String(a.name || "").localeCompare(String(b.name || ""), "pt-BR"))
    .slice(0, 50)

  if (!hostGuestLookupMatches.length) {
    box.innerHTML = '<button class="lookup-option" type="button"><strong>Nenhum cliente encontrado</strong><span>Preencha o cadastro rapido abaixo.</span></button>'
    box.classList.add("active")
    return
  }

  box.innerHTML = hostGuestLookupMatches
    .map((guest) => {
      return `
        <button class="lookup-option" type="button" onclick="selectHostGuest('${guest.id}')">
          <strong>${escapeHtml(guest.name)}</strong>
          <span>${escapeHtml([guest.doc || "Sem documento", guest.phone || "", guest.email || ""].filter(Boolean).join(" | "))}</span>
        </button>`
    })
    .join("")
  box.classList.add("active")
}

function hideHostGuestLookupSoon() {
  setTimeout(() => {
    document.getElementById("hostGuestResults")?.classList.remove("active")
  }, 160)
}

function renderDashboard() {
  const rooms = db.rooms.length
  const occupied = db.rooms.filter((r) => roomLiveStatus(r) === "OCUPADO").length
  const activeReservations = db.reservations.filter((r) =>
    ["AGENDADA", "CONFIRMADA", "HOSPEDADO"].includes(r.status)
  ).length
  const month = todayISO().slice(0, 7)
  const financeEntries = activeFinanceEntries()
  const monthRevenue = financeEntries
    .filter((x) => x.type === "RECEITA" && String(x.date || "").slice(0, 7) === month)
    .reduce((sum, x) => sum + toNumber(x.amount), 0)
  const balance = financeEntries.reduce((sum, x) => sum + financeSignedAmount(x), 0)

  document.getElementById("kpiRooms").textContent = rooms
  document.getElementById("kpiOccupied").textContent = occupied
  document.getElementById("kpiReservations").textContent = activeReservations
  document.getElementById("kpiRevenue").textContent = money(monthRevenue)
  document.getElementById("kpiBalance").textContent = money(balance)

  const checkouts = db.reservations
    .filter((r) => ["HOSPEDADO", "CONFIRMADA"].includes(r.status))
    .map((r) => ({ reservation: r, days: window.store.daysBetween(todayISO(), r.checkout) }))
    .filter((x) => x.days >= 0 && x.days <= 3)
    .sort((a, b) => a.days - b.days)

  const tbCheckout = document.getElementById("tbDashCheckout")
  if (!checkouts.length) {
    tbCheckout.innerHTML = '<tr><td colspan="5" class="empty">Sem check-outs nos proximos 3 dias.</td></tr>'
  } else {
    tbCheckout.innerHTML = checkouts
      .map(({ reservation, days }) => {
        const guest = byId(db.guests, reservation.guestId)
        const room = byId(db.rooms, reservation.roomId)
        return `
          <tr>
            <td>${escapeHtml(reservation.code)}</td>
            <td>${escapeHtml(guest?.name || "N/A")}</td>
            <td>${escapeHtml(room?.number || "N/A")}</td>
            <td>${fmtDate(reservation.checkout)} (${days} dia[s])</td>
            <td>${money(reservationPending(reservation))}</td>
          </tr>`
      })
      .join("")
  }

  const tbAlerts = document.getElementById("tbDashAlerts")
  const alerts = getSystemAlerts()
  tbAlerts.innerHTML = alerts
    .map((a) => {
      const lvl = a.priority === "URGENTE" || a.priority === "ALTA" ? "CRITICAL" : a.priority === "MEDIA" ? "WARN" : "INFO"
      return `
        <tr>
          <td>${escapeHtml(a.type)}</td>
          <td>${escapeHtml(a.message)}</td>
          <td>${statusTag("level", lvl === "CRITICAL" ? "CRITICO" : lvl === "WARN" ? "ATENCAO" : "OK")}</td>
        </tr>`
    })
    .join("")
}

function renderRooms() {
  const q = normal(document.getElementById("searchRoom").value)
  const rooms = q
    ? db.rooms.filter((r) => {
        const hosting = activeHostingForRoom(r.id)
        const guest = hosting ? byId(db.guests, hosting.guestId) : null
        return normal(`${r.number} ${r.category} ${roomLiveStatus(r)} ${r.notes || ""} ${guest?.name || ""}`).includes(q)
      })
    : db.rooms
  const tb = document.getElementById("tbRooms")
  if (!rooms.length) {
    tb.innerHTML = '<tr><td colspan="7" class="empty">Nenhum quarto cadastrado.</td></tr>'
    return
  }
  tb.innerHTML = rooms
    .map((r) => {
      const status = roomLiveStatus(r)
      const hosting = activeHostingForRoom(r.id)
      const guest = hosting ? byId(db.guests, hosting.guestId) : null
      const pending = hosting ? reservationPending(hosting) : 0
      const guestInfo = hosting
        ? `
          <div class="room-guest">
            <strong>${escapeHtml(guest?.name || "Hospede")}</strong>
            <span>${escapeHtml(hosting.code)} | Falta ${money(pending)}</span>
          </div>`
        : '<span class="muted-line">Sem hospede no quarto</span>'
      const actions = hosting
        ? `
          <button class="btn" onclick="showHostedReservation('${hosting.id}')">Ver hospedagem</button>
          <button class="btn success" onclick="openPaymentDialog('${hosting.id}')">Adicionar pagamento</button>
          <button class="btn" onclick="printReceipt('${hosting.id}')">Gerar recibo</button>`
        : status === "DISPONIVEL"
          ? `<button class="btn primary" onclick="openDirectHosting('${r.id}')">Hospedar</button>`
          : ""
      return `
        <tr>
          <td>${escapeHtml(r.number)}</td>
          <td>${escapeHtml(r.category)}</td>
          <td>${money(r.rate)}</td>
          <td>${statusTag("room", status)}</td>
          <td>${guestInfo}</td>
          <td>${escapeHtml(r.notes || "-")}</td>
          <td>
            <span class="inline-actions">
              ${actions}
              <button class="btn" onclick="toggleRoomMaintenance('${r.id}')">${r.baseStatus === "MANUTENCAO" ? "Liberar" : "Manutencao"}</button>
              <button class="btn" onclick="delRoom('${r.id}')">Excluir</button>
            </span>
          </td>
        </tr>`
    })
    .join("")
}

function renderGuests() {
  const q = normal(document.getElementById("searchGuest").value)
  const rows = q
    ? db.guests.filter((g) =>
        normal(`${g.name} ${g.doc} ${g.phone} ${g.email} ${g.city}`).includes(q)
      )
    : db.guests
  const tb = document.getElementById("tbGuests")
  if (!rows.length) {
    tb.innerHTML = '<tr><td colspan="6" class="empty">Nenhum cliente cadastrado.</td></tr>'
    return
  }
  tb.innerHTML = rows
    .map(
      (g) => `
      <tr>
        <td>${escapeHtml(g.name)}</td>
        <td>${escapeHtml(g.doc || "-")}</td>
        <td>${escapeHtml(g.phone || "-")}</td>
        <td>${escapeHtml(g.email || "-")}</td>
        <td>${escapeHtml(g.city || "-")}</td>
        <td>
          <span class="inline-actions">
            <button class="btn" onclick="editGuest('${g.id}')">Editar</button>
            <button class="btn" onclick="delGuest('${g.id}')">Excluir</button>
          </span>
        </td>
      </tr>`
    )
    .join("")
}

function renderReservations() {
  const q = normal(document.getElementById("searchReservation").value)
  const rows = q
    ? db.reservations.filter((r) => {
        const guest = byId(db.guests, r.guestId)
        const room = byId(db.rooms, r.roomId)
        return normal(
          `${r.code} ${guest?.name || ""} ${room?.number || ""} ${r.status} ${r.channel || ""}`
        ).includes(q)
      })
    : db.reservations

  const tb = document.getElementById("tbReservations")
  if (!rows.length) {
    tb.innerHTML = '<tr><td colspan="9" class="empty">Nenhuma reserva cadastrada.</td></tr>'
    return
  }

  tb.innerHTML = rows
    .map((r) => {
      const guest = byId(db.guests, r.guestId)
      const room = byId(db.rooms, r.roomId)
      const paid = reservationPaid(r.id)
      const pending = reservationPending(r)
      const lastTransfer = Array.isArray(r.roomTransfers) ? r.roomTransfers[0] : null
      const transferInfo = lastTransfer
        ? `<div class="muted-line">Transferido de ${escapeHtml(lastTransfer.fromRoom || "-")} em ${fmtDateTime(lastTransfer.when)}</div>`
        : ""
      return `
      <tr>
        <td>${escapeHtml(r.code)}</td>
        <td>${escapeHtml(guest?.name || "N/A")}</td>
        <td>${escapeHtml(room?.number || "N/A")}${transferInfo}</td>
        <td>${fmtDate(r.checkin)} ate ${fmtDate(r.checkout)}</td>
        <td>
          <select onchange="updateReservationStatus('${r.id}', this.value)">
            ${["AGENDADA", "CONFIRMADA", "HOSPEDADO", "CHECKOUT", "CANCELADA"]
              .map((s) => `<option value="${s}" ${s === r.status ? "selected" : ""}>${s}</option>`)
              .join("")}
          </select>
        </td>
        <td>${money(reservationTotal(r))}</td>
        <td>${money(paid)}</td>
        <td>${money(pending)}</td>
        <td>
          <span class="inline-actions">
            <button class="btn" onclick="copyReservationCode('${escapeHtml(r.code)}')">Copiar codigo</button>
            ${r.status === "HOSPEDADO" ? `<button class="btn" onclick="transferReservationRoom('${r.id}')">Transferir quarto</button>` : ""}
            ${r.status === "HOSPEDADO" ? `<button class="btn success" onclick="openPaymentDialog('${r.id}')">Adicionar pagamento</button>` : ""}
            ${r.status === "HOSPEDADO" ? `<button class="btn primary" onclick="closeReservationAccount('${r.id}')">Fechar conta</button>` : ""}
            ${["HOSPEDADO", "CHECKOUT"].includes(r.status) ? `<button class="btn" onclick="printReceipt('${r.id}')">Recibo</button>` : ""}
            <button class="btn" onclick="openReservationInPortal('${r.id}')">Portal</button>
            <button class="btn" onclick="delReservation('${r.id}')">Excluir</button>
          </span>
        </td>
      </tr>`
    })
    .join("")
}

function renderHosted() {
  const q = normal(document.getElementById("searchHosted").value)
  const rows = hostedReservations().filter((reservation) => {
    if (!q) return true
    const guest = byId(db.guests, reservation.guestId)
    const room = byId(db.rooms, reservation.roomId)
    return normal(`${reservation.code} ${guest?.name || ""} ${room?.number || ""}`).includes(q)
  })

  const totals = hostedReservations().reduce(
    (acc, reservation) => {
      acc.total += reservationTotal(reservation)
      acc.paid += reservationPaid(reservation.id)
      acc.pending += reservationPending(reservation)
      if (reservation.checkout === todayISO()) acc.checkoutToday += 1
      return acc
    },
    { total: 0, paid: 0, pending: 0, checkoutToday: 0 }
  )

  document.getElementById("hostedCount").textContent = hostedReservations().length
  document.getElementById("hostedTotal").textContent = money(totals.total)
  document.getElementById("hostedPaid").textContent = money(totals.paid)
  document.getElementById("hostedPending").textContent = money(totals.pending)
  document.getElementById("hostedCheckoutToday").textContent = totals.checkoutToday

  const tb = document.getElementById("tbHosted")
  if (!rows.length) {
    tb.innerHTML = '<tr><td colspan="8" class="empty">Nenhum cliente hospedado no momento.</td></tr>'
    return
  }

  tb.innerHTML = rows
    .map((reservation) => {
      const guest = byId(db.guests, reservation.guestId)
      const room = byId(db.rooms, reservation.roomId)
      const total = reservationTotal(reservation)
      const paid = reservationPaid(reservation.id)
      const pending = reservationPending(reservation)
      return `
        <tr>
          <td>${escapeHtml(reservation.code)}</td>
          <td>${escapeHtml(guest?.name || "N/A")}<div class="muted-line">${escapeHtml(guest?.phone || guest?.doc || "")}</div></td>
          <td>${escapeHtml(room?.number || "N/A")}<div class="muted-line">${escapeHtml(room?.category || "")}</div></td>
          <td>${fmtDate(reservation.checkin)} ate ${fmtDate(reservation.checkout)}<div class="muted-line">${reservationNights(reservation)} diaria(s)</div></td>
          <td>${money(total)}</td>
          <td>${money(paid)}</td>
          <td>${pending > 0 ? `<strong>${money(pending)}</strong>` : '<span class="tag ok">Quitado</span>'}</td>
          <td>
            <span class="inline-actions">
              <button class="btn success" onclick="openPaymentDialog('${reservation.id}')">Adicionar pagamento</button>
              <button class="btn" onclick="transferReservationRoom('${reservation.id}')">Transferir quarto</button>
              <button class="btn primary" onclick="closeReservationAccount('${reservation.id}')">Fechar conta</button>
              <button class="btn" onclick="printReceipt('${reservation.id}')">Gerar recibo</button>
              <button class="btn" onclick="openReservationInPortal('${reservation.id}')">Portal</button>
            </span>
          </td>
        </tr>`
    })
    .join("")
}

function renderFinance() {
  const activeEntries = activeFinanceEntries()
  const revenue = activeEntries
    .filter((x) => x.type === "RECEITA")
    .reduce((sum, x) => sum + toNumber(x.amount), 0)
  const expense = activeEntries
    .filter((x) => x.type === "DESPESA")
    .reduce((sum, x) => sum + toNumber(x.amount), 0)
  const balance = revenue - expense

  document.getElementById("finRevenue").textContent = money(revenue)
  document.getElementById("finExpense").textContent = money(expense)
  document.getElementById("finBalance").textContent = money(balance)

  const tb = document.getElementById("tbFinance")
  if (!db.finance.length) {
    tb.innerHTML = '<tr><td colspan="10" class="empty">Nenhum lancamento financeiro.</td></tr>'
    return
  }
  tb.innerHTML = db.finance
    .slice()
    .sort((a, b) => String(b.date).localeCompare(String(a.date)))
    .map((x) => {
      const account = byId(db.bankAccounts, x.accountId)
      const reservation = byId(db.reservations, x.reservationId)
      const active = isFinanceActive(x)
      const cancelledInfo = x.cancelledAt ? `<div class="muted-line">Cancelado em ${fmtDateTime(x.cancelledAt)}</div>` : ""
      return `
      <tr${active ? "" : ' class="row-cancelled"'}>
        <td>${fmtDate(x.date)}</td>
        <td>${statusTag("level", x.type === "RECEITA" ? "RECEITA" : "DESPESA")}</td>
        <td>${escapeHtml(x.category || "-")}</td>
        <td>${escapeHtml(x.method || "-")}</td>
        <td>${escapeHtml(account?.name || "-")}</td>
        <td>${escapeHtml(reservation?.code || "-")}</td>
        <td>${escapeHtml(x.description || "-")}</td>
        <td>${money(x.amount)}</td>
        <td>${statusTag("finance", x.status || "ATIVO")}${cancelledInfo}</td>
        <td>${active ? `<button class="btn danger" onclick="cancelFinanceEntry('${x.id}')">Cancelar</button>` : "-"}</td>
      </tr>`
    })
    .join("")
}

function renderStock() {
  const tbItems = document.getElementById("tbStockItems")
  if (!db.stockItems.length) {
    tbItems.innerHTML = '<tr><td colspan="8" class="empty">Nenhum item em estoque.</td></tr>'
  } else {
    tbItems.innerHTML = db.stockItems
      .map((item) => {
        const low = toNumber(item.qty) <= toNumber(item.minQty)
        return `
          <tr>
            <td>${escapeHtml(item.name)}</td>
            <td>${escapeHtml(item.qty)}</td>
            <td>${escapeHtml(item.minQty)}</td>
            <td>${escapeHtml(item.unit)}</td>
            <td>${money(item.cost)}</td>
            <td>${escapeHtml(item.supplier || "-")}</td>
            <td>${low ? '<span class="tag bad">Baixo</span>' : '<span class="tag ok">Normal</span>'}</td>
            <td><button class="btn" onclick="delStockItem('${item.id}')">Excluir</button></td>
          </tr>`
      })
      .join("")
  }

  const tbMoves = document.getElementById("tbStockMoves")
  if (!db.stockMovements.length) {
    tbMoves.innerHTML = '<tr><td colspan="5" class="empty">Sem movimentos de estoque.</td></tr>'
  } else {
    tbMoves.innerHTML = db.stockMovements
      .slice()
      .sort((a, b) => String(b.date).localeCompare(String(a.date)))
      .map((m) => {
        const item = byId(db.stockItems, m.itemId)
        return `
          <tr>
            <td>${fmtDate(m.date)}</td>
            <td>${escapeHtml(item?.name || "-")}</td>
            <td>${escapeHtml(m.type)}</td>
            <td>${escapeHtml(m.qty)}</td>
            <td>${escapeHtml(m.reason || "-")}</td>
          </tr>`
      })
      .join("")
  }
}

function renderStaff() {
  const q = normal(document.getElementById("searchStaff").value)
  const rows = q
    ? db.staff.filter((s) =>
        normal(`${s.name} ${s.role} ${s.phone} ${s.email} ${s.shift} ${s.status}`).includes(q)
      )
    : db.staff
  const tb = document.getElementById("tbStaff")
  if (!rows.length) {
    tb.innerHTML = '<tr><td colspan="6" class="empty">Nenhum funcionario cadastrado.</td></tr>'
    return
  }
  tb.innerHTML = rows
    .map((s) => {
      return `
      <tr>
        <td>${escapeHtml(s.name)}</td>
        <td>${escapeHtml(s.role || "-")}</td>
        <td>${escapeHtml(s.shift || "-")}</td>
        <td>${statusTag("staff", s.status)}</td>
        <td>${escapeHtml(s.phone || "-")} ${s.email ? " | " + escapeHtml(s.email) : ""}</td>
        <td><button class="btn" onclick="delStaff('${s.id}')">Excluir</button></td>
      </tr>`
    })
    .join("")
}

function renderPoint() {
  const tb = document.getElementById("tbPoint")
  if (!db.timeEntries.length) {
    tb.innerHTML = '<tr><td colspan="4" class="empty">Nenhum registro de ponto.</td></tr>'
    return
  }
  tb.innerHTML = db.timeEntries
    .slice()
    .sort((a, b) => String(b.when).localeCompare(String(a.when)))
    .map((p) => {
      const staff = byId(db.staff, p.staffId)
      return `
      <tr>
        <td>${fmtDateTime(p.when)}</td>
        <td>${escapeHtml(staff?.name || "-")}</td>
        <td>${escapeHtml(p.type)}</td>
        <td>${escapeHtml(p.note || "-")}</td>
      </tr>`
    })
    .join("")
}

function renderOrders() {
  const q = normal(document.getElementById("searchOrder").value)
  const rows = q
    ? db.serviceOrders.filter((o) => {
        const staff = byId(db.staff, o.assigneeId)
        const room = byId(db.rooms, o.roomId)
        return normal(
          `${o.code} ${o.category} ${o.status} ${o.priority} ${staff?.name || ""} ${room?.number || ""}`
        ).includes(q)
      })
    : db.serviceOrders
  const tb = document.getElementById("tbOrders")
  if (!rows.length) {
    tb.innerHTML = '<tr><td colspan="8" class="empty">Nenhuma ordem de servico cadastrada.</td></tr>'
    return
  }

  tb.innerHTML = rows
    .slice()
    .sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt)))
    .map((o) => {
      const room = byId(db.rooms, o.roomId)
      const staff = byId(db.staff, o.assigneeId)
      return `
      <tr>
        <td>${escapeHtml(o.code)}</td>
        <td>${escapeHtml(o.category)}</td>
        <td>${escapeHtml(room?.number || "-")}</td>
        <td>${escapeHtml(staff?.name || "-")}</td>
        <td>${o.deadline ? fmtDateTime(o.deadline) : "-"}</td>
        <td>
          <select onchange="updateOrderStatus('${o.id}', this.value)">
            ${["ABERTA", "EM_ANDAMENTO", "FINALIZADA", "CANCELADA"]
              .map((s) => `<option value="${s}" ${s === o.status ? "selected" : ""}>${s}</option>`)
              .join("")}
          </select>
        </td>
        <td>${escapeHtml(o.priority)}</td>
        <td><button class="btn" onclick="delOrder('${o.id}')">Excluir</button></td>
      </tr>`
    })
    .join("")
}

function accountBalance(accountId) {
  const account = byId(db.bankAccounts, accountId)
  if (!account) return 0
  const flow = activeFinanceEntries()
    .filter((x) => String(x.accountId) === String(accountId))
    .reduce((sum, x) => sum + financeSignedAmount(x), 0)
  return toNumber(account.initialBalance) + flow
}

function renderAccounts() {
  const tb = document.getElementById("tbAccounts")
  if (!db.bankAccounts.length) {
    tb.innerHTML = '<tr><td colspan="6" class="empty">Nenhuma conta cadastrada.</td></tr>'
    return
  }
  tb.innerHTML = db.bankAccounts
    .map((a) => {
      const balance = accountBalance(a.id)
      return `
      <tr>
        <td>${escapeHtml(a.name)}</td>
        <td>${escapeHtml(a.bank || "-")}</td>
        <td>${escapeHtml(a.agency || "-")}</td>
        <td>${escapeHtml(a.number || "-")}</td>
        <td>${money(balance)}</td>
        <td><button class="btn" onclick="delAccount('${a.id}')">Excluir</button></td>
      </tr>`
    })
    .join("")
}

function buildReportRows() {
  const start = document.getElementById("repStart").value
  const end = document.getElementById("repEnd").value
  const statusFilter = document.getElementById("repStatus").value
  const categoryFilter = normal(document.getElementById("repCategory").value)

  const reservationRows = db.reservations
    .filter((r) => {
      if (statusFilter && r.status !== statusFilter) return false
      if (start && String(r.checkin) < start) return false
      if (end && String(r.checkout) > end) return false
      return true
    })
    .map((r) => {
      const g = byId(db.guests, r.guestId)
      return {
        type: "Reserva",
        reference: `${r.code} - ${g?.name || "Hospede"}`,
        date: r.checkin,
        status: r.status,
        category: r.channel || "-",
        value: reservationTotal(r)
      }
    })

  const financeRows = activeFinanceEntries()
    .filter((f) => {
      if (start && String(f.date) < start) return false
      if (end && String(f.date) > end) return false
      if (categoryFilter && !normal(f.category).includes(categoryFilter)) return false
      return true
    })
    .map((f) => {
      return {
        type: "Financeiro",
        reference: f.description || "-",
        date: f.date,
        status: f.type,
        category: f.category || "-",
        value: financeSignedAmount(f)
      }
    })

  return reservationRows.concat(financeRows).sort((a, b) => String(b.date).localeCompare(String(a.date)))
}

function renderReport(rows) {
  const tb = document.getElementById("tbReport")
  if (!rows.length) {
    tb.innerHTML = '<tr><td colspan="6" class="empty">Nenhum dado para o periodo/filtro selecionado.</td></tr>'
    return
  }
  tb.innerHTML = rows
    .map((r) => {
      const value = r.type === "Financeiro" && r.value < 0 ? `- ${money(Math.abs(r.value))}` : money(r.value)
      return `
        <tr>
          <td>${escapeHtml(r.type)}</td>
          <td>${escapeHtml(r.reference)}</td>
          <td>${fmtDate(r.date)}</td>
          <td>${escapeHtml(r.status)}</td>
          <td>${escapeHtml(r.category)}</td>
          <td>${escapeHtml(value)}</td>
        </tr>`
    })
    .join("")
}

function renderReportSummary(rows) {
  const reserves = rows.filter((x) => x.type === "Reserva")
  const finance = rows.filter((x) => x.type === "Financeiro")
  const financeBalance = finance.reduce((sum, x) => sum + x.value, 0)
  document.getElementById("repSummary").textContent =
    `Reservas: ${reserves.length} | Lancamentos: ${finance.length} | Saldo financeiro no periodo: ${money(financeBalance)}`
}

function renderNotifications() {
  const tb = document.getElementById("tbNotifications")
  if (!db.notifications.length) {
    tb.innerHTML = '<tr><td colspan="6" class="empty">Sem notificacoes registradas.</td></tr>'
    return
  }
  tb.innerHTML = db.notifications
    .map((n) => {
      return `
      <tr>
        <td>${fmtDateTime(n.when)}</td>
        <td>${escapeHtml(n.title)}</td>
        <td>${escapeHtml(n.message)}</td>
        <td>${statusTag("level", n.level)}</td>
        <td>${n.read ? '<span class="tag ok">Lida</span>' : '<span class="tag warn">Nova</span>'}</td>
        <td>${n.read ? "-" : `<button class="btn" onclick="markNotificationRead('${n.id}')">Marcar lida</button>`}</td>
      </tr>`
    })
    .join("")
}

function renderAudit() {
  const q = normal(document.getElementById("searchAudit").value)
  const rows = q
    ? db.auditLogs.filter((a) => normal(`${a.action} ${a.details}`).includes(q))
    : db.auditLogs
  const tb = document.getElementById("tbAudit")
  if (!rows.length) {
    tb.innerHTML = '<tr><td colspan="3" class="empty">Sem eventos de auditoria.</td></tr>'
    return
  }
  tb.innerHTML = rows
    .map(
      (a) => `
      <tr>
        <td>${fmtDateTime(a.when)}</td>
        <td>${escapeHtml(a.action)}</td>
        <td>${escapeHtml(a.details || "-")}</td>
      </tr>`
    )
    .join("")
}

function renderGuestPortal(reservationId = portalReservationId) {
  const panel = document.getElementById("guestPanelContent")
  if (!reservationId) {
    panel.textContent = "Nenhuma reserva selecionada."
    return
  }
  const reservation = byId(db.reservations, reservationId)
  if (!reservation) {
    panel.textContent = "Reserva nao encontrada."
    return
  }
  const guest = byId(db.guests, reservation.guestId)
  const room = byId(db.rooms, reservation.roomId)
  const paid = reservationPaid(reservation.id)
  const total = reservationTotal(reservation)
  const pending = total - paid
  const financeEntries = db.finance
    .filter((entry) => String(entry.reservationId) === String(reservation.id))
    .slice()
    .sort((a, b) => String(b.date).localeCompare(String(a.date)))
  const financeHtml = financeEntries.length
    ? `
      <div class="portal-ledger">
        <strong>Financeiro da reserva:</strong>
        ${financeEntries
          .map((entry) => {
            const active = isFinanceActive(entry)
            const signal = entry.type === "DESPESA" ? "-" : "+"
            return `
              <div class="ledger-row">
                <span>${fmtDate(entry.date)} | ${escapeHtml(entry.category || "-")} | ${signal} ${money(entry.amount)}</span>
                <span>${statusTag("finance", entry.status || "ATIVO")}</span>
                ${active ? `<button class="btn danger" onclick="cancelFinanceEntry('${entry.id}')">Cancelar</button>` : ""}
              </div>`
          })
          .join("")}
      </div>`
    : '<div class="muted-line">Nenhum lancamento financeiro vinculado a esta reserva.</div>'
  panel.innerHTML = `
    <strong>Reserva:</strong> ${escapeHtml(reservation.code)}<br>
    <strong>Hospede:</strong> ${escapeHtml(guest?.name || "-")}<br>
    <strong>Quarto:</strong> ${escapeHtml(room?.number || "-")} (${escapeHtml(room?.category || "-")})<br>
    <strong>Status:</strong> ${escapeHtml(reservation.status)}<br>
    <strong>Periodo:</strong> ${fmtDate(reservation.checkin)} ate ${fmtDate(reservation.checkout)}<br>
    <strong>Total:</strong> ${money(total)} | <strong>Pago:</strong> ${money(paid)} | <strong>Pendente:</strong> ${money(pending)}
    ${financeHtml}
  `
}

function renderStaffPanel() {
  const staffId = document.getElementById("staffPanelUser").value
  const tbOrders = document.getElementById("tbStaffOpenOrders")
  const tbPoint = document.getElementById("tbStaffRecentPoint")

  if (!staffId) {
    tbOrders.innerHTML = '<tr><td colspan="6" class="empty">Selecione um funcionario.</td></tr>'
    tbPoint.innerHTML = '<tr><td colspan="3" class="empty">Selecione um funcionario.</td></tr>'
    return
  }

  const myOrders = db.serviceOrders.filter(
    (o) => String(o.assigneeId) === String(staffId) && !["FINALIZADA", "CANCELADA"].includes(o.status)
  )
  if (!myOrders.length) {
    tbOrders.innerHTML = '<tr><td colspan="6" class="empty">Sem ordens em aberto.</td></tr>'
  } else {
    tbOrders.innerHTML = myOrders
      .map((o) => {
        const room = byId(db.rooms, o.roomId)
        return `
        <tr>
          <td>${escapeHtml(o.code)}</td>
          <td>${escapeHtml(o.category)}</td>
          <td>${escapeHtml(room?.number || "-")}</td>
          <td>${escapeHtml(o.status)}</td>
          <td>${escapeHtml(o.priority)}</td>
          <td><button class="btn" onclick="updateOrderStatus('${o.id}', 'FINALIZADA')">Concluir</button></td>
        </tr>`
      })
      .join("")
  }

  const myPoints = db.timeEntries
    .filter((p) => String(p.staffId) === String(staffId))
    .slice()
    .sort((a, b) => String(b.when).localeCompare(String(a.when)))
    .slice(0, 10)

  if (!myPoints.length) {
    tbPoint.innerHTML = '<tr><td colspan="3" class="empty">Sem registros recentes.</td></tr>'
  } else {
    tbPoint.innerHTML = myPoints
      .map(
        (p) => `
        <tr>
          <td>${fmtDateTime(p.when)}</td>
          <td>${escapeHtml(p.type)}</td>
          <td>${escapeHtml(p.note || "-")}</td>
        </tr>`
      )
      .join("")
  }
}

function refreshAll() {
  fillSelects()
  renderDashboard()
  renderRooms()
  renderGuests()
  renderReservations()
  renderHosted()
  renderFinance()
  renderStock()
  renderStaff()
  renderPoint()
  renderOrders()
  renderAccounts()
  renderReport(latestReportRows)
  renderReportSummary(latestReportRows)
  renderNotifications()
  renderAudit()
  renderGuestPortal()
  renderStaffPanel()
}

function saveAndRefresh(action, details = "") {
  if (action) addAudit(action, details)
  saveDb()
  latestReportRows = buildReportRows()
  refreshAll()
}

function clearFields(ids) {
  ids.forEach((id) => {
    const el = document.getElementById(id)
    if (!el) return
    if (el.tagName === "SELECT") el.selectedIndex = 0
    else el.value = ""
  })
}

function bindNav() {
  document.querySelectorAll(".nav-btn").forEach((btn) => {
    btn.addEventListener("click", () => showPage(btn.dataset.page))
  })
}

function bindExternalLinks() {
  document.querySelectorAll("[data-open-external]").forEach((link) => {
    link.addEventListener("click", (event) => {
      event.preventDefault()
      const url = link.getAttribute("href")
      if (url) electronShell.openExternal(url)
    })
  })
}

function bindRoomEvents() {
  document.getElementById("btnSaveRoom").addEventListener("click", () => {
    const number = document.getElementById("roomNumber").value.trim()
    const category = document.getElementById("roomCategory").value
    const rate = toNumber(document.getElementById("roomRate").value)
    const baseStatus = document.getElementById("roomBaseStatus").value
    const notes = document.getElementById("roomNotes").value.trim()
    if (!number || rate <= 0) return toast("Informe numero e diaria valida para o quarto.")
    if (db.rooms.some((r) => normal(r.number) === normal(number))) return toast("Ja existe quarto com esse numero.")
    db.rooms.push({
      id: window.store.uid("room"),
      number,
      category,
      rate,
      baseStatus,
      notes
    })
    addNotification("Quarto cadastrado", `Quarto ${number} criado.`, "INFO", "ROOM")
    saveAndRefresh("Cadastro de quarto", `Quarto ${number}`)
    clearFields(["roomNumber", "roomRate", "roomNotes"])
  })
  document.getElementById("searchRoom").addEventListener("input", renderRooms)
}

function setGuestFormMode(guest = null) {
  document.getElementById("guestEditId").value = guest?.id || ""
  document.getElementById("guestName").value = guest?.name || ""
  document.getElementById("guestDoc").value = guest?.doc || ""
  document.getElementById("guestPhone").value = guest?.phone || ""
  document.getElementById("guestEmail").value = guest?.email || ""
  document.getElementById("guestBirth").value = guest?.birth || ""
  document.getElementById("guestCity").value = guest?.city || ""
  document.getElementById("guestNotes").value = guest?.notes || ""
  document.getElementById("btnSaveGuest").textContent = guest ? "Atualizar cliente" : "Salvar cliente"
  document.getElementById("btnCancelGuestEdit").hidden = !guest
}

function currentGuestFormData() {
  return {
    id: document.getElementById("guestEditId").value || window.store.uid("guest"),
    name: fieldText(document.getElementById("guestName").value),
    doc: fieldText(document.getElementById("guestDoc").value),
    phone: fieldText(document.getElementById("guestPhone").value),
    email: fieldText(document.getElementById("guestEmail").value),
    birth: document.getElementById("guestBirth").value,
    city: fieldText(document.getElementById("guestCity").value),
    notes: fieldText(document.getElementById("guestNotes").value)
  }
}

async function importGuestsFromWorkbook(workbook) {
  const sheetName = workbook.SheetNames[0]
  const sheet = workbook.Sheets[sheetName]
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false })
  const valid = []
  const skipped = []

  rows.forEach((row, index) => {
    const guest = makeGuestFromRow(row)
    if (!guest.name) {
      skipped.push(`Linha ${index + 2}: sem nome`)
      return
    }
    const duplicate = findDuplicateGuest(guest) || valid.find((item) => sameGuestIdentity(item, guest))
    if (duplicate) {
      skipped.push(`Linha ${index + 2}: duplicado (${guest.name})`)
      return
    }
    valid.push(guest)
  })

  if (!valid.length) {
    toast(`Nenhum cliente novo encontrado. Ignorados: ${skipped.length}.`)
    return
  }

  const preview = valid.slice(0, 5).map((guest) => `- ${guest.name}`).join("\n")
  const message = `Importar ${valid.length} cliente(s)?\n\nPrimeiros registros:\n${preview}\n\nIgnorados por duplicidade/erro: ${skipped.length}`
  if (!(await askConfirm(message, "Importar clientes", "Importar", "Cancelar"))) return

  db.guests.push(...valid)
  addNotification("Clientes importados", `${valid.length} cliente(s) importados por planilha.`, "INFO", "GUEST")
  saveAndRefresh("Importacao de clientes", `${valid.length} importados | ${skipped.length} ignorados`)
  document.getElementById("guestImportStatus").textContent = `Importados: ${valid.length}. Ignorados: ${skipped.length}.`
  toast("Importacao concluida.")
}

function exportGuestsWorkbook(fileName, rows = guestExportRows()) {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.json_to_sheet(rows)
  XLSX.utils.book_append_sheet(workbook, sheet, "Clientes")
  return saveWorkbookToDownloads(workbook, fileName)
}

function bindGuestEvents() {
  document.getElementById("btnSaveGuest").addEventListener("click", () => {
    const guest = currentGuestFormData()
    if (!guest.name) return toast("Informe o nome do cliente.")

    const editId = document.getElementById("guestEditId").value
    const duplicate = findDuplicateGuest(guest, editId)
    if (duplicate) return toast(`Cliente duplicado: ${duplicate.name}.`)

    if (editId) {
      const current = byId(db.guests, editId)
      if (!current) return toast("Cliente nao encontrado para edicao.")
      Object.assign(current, guest, { id: editId })
      addNotification("Cliente atualizado", `${guest.name} atualizado.`, "INFO", "GUEST")
      saveAndRefresh("Edicao de cliente", guest.name)
    } else {
      db.guests.push(guest)
      addNotification("Novo cliente", `${guest.name} cadastrado no sistema.`, "INFO", "GUEST")
      saveAndRefresh("Cadastro de cliente", guest.name)
    }

    setGuestFormMode(null)
  })

  document.getElementById("btnCancelGuestEdit").addEventListener("click", () => {
    setGuestFormMode(null)
  })

  document.getElementById("btnImportGuests").addEventListener("click", () => {
    document.getElementById("guestImportFile").click()
  })

  document.getElementById("guestImportFile").addEventListener("change", (event) => {
    const file = event.target.files?.[0]
    event.target.value = ""
    if (!file) return
    if (!file.name.toLowerCase().endsWith(".xlsx")) return toast("Selecione uma planilha .xlsx.")

    const reader = new FileReader()
    reader.onload = async (loadEvent) => {
      try {
        const workbook = XLSX.read(new Uint8Array(loadEvent.target.result), { type: "array", cellDates: true })
        await importGuestsFromWorkbook(workbook)
      } catch (error) {
        toast("Nao foi possivel ler a planilha.")
      }
    }
    reader.readAsArrayBuffer(file)
  })

  document.getElementById("btnExportGuests").addEventListener("click", () => {
    const fileName = `nexhotel-clientes-${todayISO()}.xlsx`
    const outPath = exportGuestsWorkbook(fileName)
    addAudit("Exportacao de clientes", outPath)
    saveDb()
    renderAudit()
    toast(`Clientes exportados: ${outPath}`)
  })

  document.getElementById("btnDownloadGuestTemplate").addEventListener("click", () => {
    const sample = [{
      Nome: "Maria Silva",
      Documento: "12345678900",
      Telefone: "11999999999",
      Email: "maria@email.com",
      "Data Nascimento": "1990-01-15",
      Cidade: "Sao Paulo",
      Observacoes: "Cliente exemplo"
    }]
    const outPath = exportGuestsWorkbook("modelo-importacao-clientes-nexhotel.xlsx", sample)
    toast(`Modelo salvo: ${outPath}`)
  })

  document.getElementById("searchGuest").addEventListener("input", renderGuests)
}

function bindReservationEvents() {
  document.getElementById("btnSaveReservation").addEventListener("click", () => {
    const guestId = document.getElementById("resGuestId").value
    const roomId = document.getElementById("resRoom").value
    const checkin = document.getElementById("resCheckin").value
    const checkout = document.getElementById("resCheckout").value
    if (!guestId || !roomId || !checkin || !checkout) return toast("Pesquise e selecione o hospede, quarto e periodo.")
    if (window.store.daysBetween(checkin, checkout) < 1) return toast("Check-out deve ser apos check-in.")

    const room = byId(db.rooms, roomId)
    if (!room) return toast("Quarto invalido.")
    if (room.baseStatus === "MANUTENCAO") return toast("Quarto em manutencao.")

    if (roomHasReservationConflict(roomId, checkin, checkout)) return toast("Conflito de periodo para este quarto.")

    const code = reservationCode()
    const reservation = {
      id: window.store.uid("res"),
      code,
      guestId,
      roomId,
      checkin,
      checkout,
      adults: toNumber(document.getElementById("resAdults").value) || 1,
      children: toNumber(document.getElementById("resChildren").value) || 0,
      dailyRate: toNumber(document.getElementById("resDaily").value),
      extra: toNumber(document.getElementById("resExtra").value),
      status: document.getElementById("resStatus").value,
      channel: document.getElementById("resChannel").value,
      notes: document.getElementById("resNotes").value.trim(),
      createdAt: new Date().toISOString()
    }
    db.reservations.push(reservation)
    addNotification("Nova reserva", `${code} criada com status ${reservation.status}.`, "INFO", "RESERVA", reservation.id)
    saveAndRefresh("Criacao de reserva", code)
    clearFields(["resGuestSearch", "resGuestId", "resCheckin", "resCheckout", "resDaily", "resExtra", "resNotes"])
  })

  document.getElementById("resGuestSearch").addEventListener("input", () => {
    document.getElementById("resGuestId").value = ""
    renderGuestLookup()
  })
  document.getElementById("resGuestSearch").addEventListener("focus", renderGuestLookup)
  document.getElementById("resGuestSearch").addEventListener("blur", hideGuestLookupSoon)
  document.getElementById("searchReservation").addEventListener("input", renderReservations)
  document.getElementById("searchHosted").addEventListener("input", renderHosted)
}

function bindFinanceEvents() {
  document.getElementById("btnSaveFinance").addEventListener("click", () => {
    const type = document.getElementById("finType").value
    const category = document.getElementById("finCategory").value.trim() || "geral"
    const method = document.getElementById("finMethod").value
    const accountId = document.getElementById("finAccount").value
    const reservationId = document.getElementById("finReservation").value
    const date = document.getElementById("finDate").value || todayISO()
    const amount = toNumber(document.getElementById("finAmount").value)
    const description = document.getElementById("finDesc").value.trim()
    if (amount <= 0) return toast("Informe valor valido para o lancamento.")
    db.finance.push({
      id: window.store.uid("fin"),
      type,
      category,
      method,
      accountId: accountId || "",
      reservationId: reservationId || "",
      date,
      amount,
      description,
      status: "ATIVO",
      origin: "FINANCEIRO",
      createdAt: new Date().toISOString()
    })
    addNotification(
      "Lancamento financeiro",
      `${type} de ${money(amount)} registrado.`,
      type === "DESPESA" ? "WARN" : "INFO",
      "FINANCE"
    )
    saveAndRefresh("Lancamento financeiro", `${type} ${money(amount)} ${category}`)
    clearFields(["finAmount", "finDesc"])
  })

  document.getElementById("btnCloseCash").addEventListener("click", () => {
    const today = todayISO()
    const entries = activeFinanceEntries().filter((x) => x.date === today)
    const rev = entries.filter((x) => x.type === "RECEITA").reduce((s, x) => s + toNumber(x.amount), 0)
    const exp = entries.filter((x) => x.type === "DESPESA").reduce((s, x) => s + toNumber(x.amount), 0)
    addAudit("Fechamento de caixa", `Data ${today} | Receita ${money(rev)} | Despesa ${money(exp)} | Saldo ${money(rev - exp)}`)
    addNotification("Fechamento de caixa", `Saldo de hoje: ${money(rev - exp)}.`, "INFO", "FINANCE")
    saveDb()
    renderAudit()
    renderNotifications()
    toast("Fechamento de caixa registrado na auditoria.")
  })
}

function bindStockEvents() {
  document.getElementById("btnSaveItem").addEventListener("click", () => {
    const name = document.getElementById("itemName").value.trim()
    const unit = document.getElementById("itemUnit").value.trim() || "un"
    const minQty = toNumber(document.getElementById("itemMin").value)
    const qty = toNumber(document.getElementById("itemQty").value)
    const cost = toNumber(document.getElementById("itemCost").value)
    const supplier = document.getElementById("itemSupplier").value.trim()
    if (!name) return toast("Informe nome do item de estoque.")
    db.stockItems.push({
      id: window.store.uid("stk"),
      name,
      unit,
      minQty,
      qty,
      cost,
      supplier
    })
    addNotification("Item de estoque", `${name} cadastrado.`, "INFO", "STOCK")
    saveAndRefresh("Cadastro de item de estoque", name)
    clearFields(["itemName", "itemUnit", "itemMin", "itemQty", "itemCost", "itemSupplier"])
  })

  document.getElementById("btnSaveMovement").addEventListener("click", () => {
    const itemId = document.getElementById("movItem").value
    const type = document.getElementById("movType").value
    const qty = toNumber(document.getElementById("movQty").value)
    const reason = document.getElementById("movReason").value.trim() || "-"
    const date = document.getElementById("movDate").value || todayISO()
    if (!itemId || qty <= 0) return toast("Selecione item e informe quantidade.")
    const item = byId(db.stockItems, itemId)
    if (!item) return
    if (type === "SAIDA" && toNumber(item.qty) < qty) return toast("Estoque insuficiente para saida.")

    item.qty = type === "ENTRADA" ? toNumber(item.qty) + qty : toNumber(item.qty) - qty
    db.stockMovements.push({
      id: window.store.uid("mov"),
      itemId,
      type,
      qty,
      reason,
      date
    })

    if (toNumber(item.qty) <= toNumber(item.minQty)) {
      addNotification("Estoque baixo", `${item.name} atingiu nivel minimo.`, "WARN", "STOCK", item.id)
    }

    saveAndRefresh("Movimento de estoque", `${type} ${qty} de ${item.name}`)
    clearFields(["movQty", "movReason"])
  })
}

function bindStaffEvents() {
  document.getElementById("btnSaveStaff").addEventListener("click", () => {
    const name = document.getElementById("staffName").value.trim()
    if (!name) return toast("Informe nome do funcionario.")
    db.staff.push({
      id: window.store.uid("stf"),
      name,
      role: document.getElementById("staffRole").value.trim(),
      phone: document.getElementById("staffPhone").value.trim(),
      email: document.getElementById("staffEmail").value.trim(),
      shift: document.getElementById("staffShift").value.trim(),
      status: document.getElementById("staffStatus").value
    })
    addNotification("Equipe", `${name} adicionado na equipe.`, "INFO", "STAFF")
    saveAndRefresh("Cadastro de funcionario", name)
    clearFields(["staffName", "staffRole", "staffPhone", "staffEmail", "staffShift"])
  })
  document.getElementById("searchStaff").addEventListener("input", renderStaff)
}

function bindPointEvents() {
  document.getElementById("btnSavePoint").addEventListener("click", () => {
    const staffId = document.getElementById("pointStaff").value
    const type = document.getElementById("pointType").value
    const when = document.getElementById("pointDateTime").value
    const note = document.getElementById("pointNote").value.trim()
    if (!staffId || !when) return toast("Selecione funcionario e data/hora.")
    db.timeEntries.push({
      id: window.store.uid("pt"),
      staffId,
      type,
      when: new Date(when).toISOString(),
      note
    })
    const staff = byId(db.staff, staffId)
    saveAndRefresh("Registro de ponto", `${staff?.name || "Funcionario"} - ${type}`)
    clearFields(["pointNote"])
  })
}

function bindOrderEvents() {
  document.getElementById("btnSaveOrder").addEventListener("click", () => {
    const category = document.getElementById("osCategory").value
    const priority = document.getElementById("osPriority").value
    const status = document.getElementById("osStatus").value
    const roomId = document.getElementById("osRoom").value
    const guestId = document.getElementById("osGuest").value
    const assigneeId = document.getElementById("osAssignee").value
    const deadlineRaw = document.getElementById("osDeadline").value
    const description = document.getElementById("osDesc").value.trim()
    if (!roomId || !assigneeId || !description) return toast("Informe quarto, responsavel e descricao da OS.")
    const code = orderCode()
    db.serviceOrders.push({
      id: window.store.uid("os"),
      code,
      category,
      priority,
      status,
      roomId,
      guestId: guestId || "",
      assigneeId,
      description,
      deadline: deadlineRaw ? new Date(deadlineRaw).toISOString() : "",
      createdAt: new Date().toISOString()
    })
    const staff = byId(db.staff, assigneeId)
    addNotification("Nova OS", `${code} atribuida para ${staff?.name || "equipe"}.`, "WARN", "OS")
    saveAndRefresh("Criacao de OS", `${code} ${category}`)
    clearFields(["osDesc", "osDeadline"])
  })
  document.getElementById("searchOrder").addEventListener("input", renderOrders)
}

function bindAccountEvents() {
  document.getElementById("btnSaveAccount").addEventListener("click", () => {
    const name = document.getElementById("accName").value.trim()
    if (!name) return toast("Informe nome da conta bancaria.")
    db.bankAccounts.push({
      id: window.store.uid("acc"),
      name,
      bank: document.getElementById("accBank").value.trim(),
      agency: document.getElementById("accAgency").value.trim(),
      number: document.getElementById("accNumber").value.trim(),
      initialBalance: toNumber(document.getElementById("accInitialBalance").value)
    })
    saveAndRefresh("Cadastro de conta bancaria", name)
    clearFields(["accName", "accBank", "accAgency", "accNumber", "accInitialBalance"])
  })
}

function bindReportEvents() {
  document.getElementById("btnGenerateReport").addEventListener("click", () => {
    latestReportRows = buildReportRows()
    renderReport(latestReportRows)
    renderReportSummary(latestReportRows)
    addAudit("Geracao de relatorio", `Registros: ${latestReportRows.length}`)
    saveDb()
    renderAudit()
  })
  document.getElementById("btnPrintReport").addEventListener("click", () => {
    if (!latestReportRows.length) return toast("Gere um relatorio antes de imprimir.")
    const lines = latestReportRows
      .map((r) => `${r.type} | ${r.reference} | ${fmtDate(r.date)} | ${r.status} | ${r.category} | ${money(r.value)}`)
      .join("\n")
    const html = `
      <!doctype html>
      <html>
      <head><meta charset="utf-8"><title>Relatorio NexHotel</title></head>
      <body style="font-family: Arial, sans-serif; padding:20px; white-space:pre-wrap;">
        <h1>Relatorio NexHotel</h1>
        <pre>${escapeHtml(lines)}</pre>
        <script>setTimeout(()=>window.print(),200)</script>
      </body>
      </html>`
    const w = window.open("", "_blank", "width=950,height=750")
    if (!w) return toast("Nao foi possivel abrir a janela de impressao.")
    w.document.open()
    w.document.write(html)
    w.document.close()
  })
}

function bindNotificationEvents() {
  document.getElementById("btnMarkAllRead").addEventListener("click", () => {
    db.notifications.forEach((n) => {
      n.read = true
    })
    saveAndRefresh("Notificacoes", "Todas marcadas como lidas")
  })
}

function bindAuditEvents() {
  document.getElementById("searchAudit").addEventListener("input", renderAudit)
}

function bindGuestPortalEvents() {
  document.getElementById("btnGuestLookup").addEventListener("click", () => {
    const code = normal(document.getElementById("guestPortalCode").value.trim())
    if (!code) return toast("Informe o codigo da reserva.")
    const reservation = db.reservations.find((r) => normal(r.code) === code)
    if (!reservation) return toast("Reserva nao encontrada.")
    portalReservationId = reservation.id
    renderGuestPortal(portalReservationId)
    showPage("hospede")
  })

  document.getElementById("btnGuestRequest").addEventListener("click", () => {
    if (!portalReservationId) return toast("Selecione a reserva no portal do hospede.")
    const reservation = byId(db.reservations, portalReservationId)
    if (!reservation) return toast("Reserva invalida.")
    const type = document.getElementById("guestRequestType").value
    const value = toNumber(document.getElementById("guestRequestValue").value)
    const text = document.getElementById("guestRequestText").value.trim()
    if (!text) return toast("Escreva a mensagem da solicitacao.")

    const request = {
      id: window.store.uid("req"),
      reservationId: reservation.id,
      type,
      text,
      value,
      when: new Date().toISOString()
    }
    db.guestRequests.unshift(request)

    if (type === "PAGAMENTO" && value > 0) {
      db.finance.push({
        id: window.store.uid("fin"),
        type: "RECEITA",
        category: "pagamento_hospede",
        method: "PIX",
        accountId: "",
        reservationId: reservation.id,
        date: todayISO(),
        amount: value,
        description: `Pagamento informado no portal: ${text}`,
        status: "ATIVO",
        origin: "APP_HOSPEDE_PAGAMENTO",
        createdAt: new Date().toISOString()
      })
    } else if (type === "PEDIDO" && value > 0) {
      reservation.extra = toNumber(reservation.extra) + value
      db.finance.push({
        id: window.store.uid("fin"),
        type: "RECEITA",
        category: "consumo_hospede",
        method: "INTERNO",
        accountId: "",
        reservationId: reservation.id,
        date: todayISO(),
        amount: value,
        description: `Consumo solicitado no portal: ${text}`,
        status: "ATIVO",
        origin: "APP_HOSPEDE_CONSUMO",
        affectsReservationExtra: true,
        createdAt: new Date().toISOString()
      })
    } else {
      db.serviceOrders.push({
        id: window.store.uid("os"),
        code: orderCode(),
        category: type,
        priority: "MEDIA",
        status: "ABERTA",
        roomId: reservation.roomId,
        guestId: reservation.guestId,
        assigneeId: db.staff.find((s) => s.status === "ATIVO")?.id || "",
        description: `Portal hospede: ${text}`,
        deadline: "",
        createdAt: new Date().toISOString()
      })
    }

    addNotification("Solicitacao de hospede", `${reservation.code} - ${type}`, "WARN", "PORTAL", reservation.id)
    saveAndRefresh("Solicitacao no portal do hospede", `${reservation.code} - ${type}`)
    clearFields(["guestRequestValue", "guestRequestText"])
    renderGuestPortal(portalReservationId)
  })
}

function bindStaffPanelEvents() {
  document.getElementById("staffPanelUser").addEventListener("change", renderStaffPanel)

  document.getElementById("btnPanelCheckin").addEventListener("click", () => {
    const staffId = document.getElementById("staffPanelUser").value
    if (!staffId) return toast("Selecione funcionario no painel.")
    db.timeEntries.push({
      id: window.store.uid("pt"),
      staffId,
      type: "ENTRADA",
      when: new Date().toISOString(),
      note: "Registro rapido no painel"
    })
    const staff = byId(db.staff, staffId)
    saveAndRefresh("Ponto no painel", `${staff?.name || "-"} entrada`)
  })

  document.getElementById("btnPanelCheckout").addEventListener("click", () => {
    const staffId = document.getElementById("staffPanelUser").value
    if (!staffId) return toast("Selecione funcionario no painel.")
    db.timeEntries.push({
      id: window.store.uid("pt"),
      staffId,
      type: "SAIDA",
      when: new Date().toISOString(),
      note: "Registro rapido no painel"
    })
    const staff = byId(db.staff, staffId)
    saveAndRefresh("Ponto no painel", `${staff?.name || "-"} saida`)
  })
}

function setDefaults() {
  document.getElementById("finDate").value = todayISO()
  document.getElementById("movDate").value = todayISO()
  document.getElementById("pointDateTime").value = toLocalDateTimeInput(new Date())
  document.getElementById("repStart").value = todayISO().slice(0, 8) + "01"
  document.getElementById("repEnd").value = todayISO()
}

window.selectReservationGuest = function selectReservationGuest(guestId) {
  const guest = byId(db.guests, guestId)
  if (!guest) return
  document.getElementById("resGuestId").value = guest.id
  document.getElementById("resGuestSearch").value = guestLabel(guest)
  document.getElementById("resGuestResults").classList.remove("active")
}

window.selectHostGuest = function selectHostGuest(guestId) {
  const guest = byId(db.guests, guestId)
  if (!guest) return
  document.getElementById("hostGuestId").value = guest.id
  document.getElementById("hostGuestSearch").value = guestLabel(guest)
  document.getElementById("hostGuestResults").classList.remove("active")
  clearFields(["hostNewName", "hostNewDoc", "hostNewPhone", "hostNewEmail"])
}

window.editGuest = function editGuest(guestId) {
  const guest = byId(db.guests, guestId)
  if (!guest) return toast("Cliente nao encontrado.")
  setGuestFormMode(guest)
  showPage("clientes")
  document.getElementById("guestName").focus()
}

window.openDirectHosting = function openDirectHosting(roomId) {
  const room = byId(db.rooms, roomId)
  if (!room) return toast("Quarto nao encontrado.")
  if (roomLiveStatus(room) !== "DISPONIVEL") return toast("Quarto nao esta disponivel para hospedar.")

  hostDialogRoomId = room.id
  resetHostDialog(room)
  document.getElementById("hostDialogInfo").textContent =
    `Quarto ${room.number} - ${room.category}\n` +
    `Diaria sugerida: ${money(room.rate)}\n` +
    "Selecione um cliente ja cadastrado ou cadastre um novo abaixo."
  document.getElementById("hostDialog").hidden = false
  document.getElementById("hostGuestSearch").focus()
}

window.openPaymentDialog = function openPaymentDialog(reservationId) {
  const reservation = byId(db.reservations, reservationId)
  if (!reservation) return toast("Reserva nao encontrada.")
  const guest = byId(db.guests, reservation.guestId)
  const room = byId(db.rooms, reservation.roomId)
  const total = reservationTotal(reservation)
  const paid = reservationPaid(reservation.id)
  const pending = reservationPending(reservation)

  paymentDialogReservationId = reservation.id
  document.getElementById("paymentDialogInfo").textContent =
    `${reservation.code} - ${guest?.name || "Hospede"}\n` +
    `Quarto: ${room?.number || "-"} | Total: ${money(total)} | Pago: ${money(paid)} | Falta: ${money(pending)}`
  document.getElementById("quickPaymentAmount").value = pending > 0 ? pending.toFixed(2).replace(".", ",") : ""
  document.getElementById("quickPaymentMethod").value = "PIX"
  document.getElementById("quickPaymentAccount").selectedIndex = 0
  document.getElementById("quickPaymentDesc").value = `Pagamento recebido da reserva ${reservation.code}`
  document.getElementById("paymentDialog").hidden = false
  document.getElementById("quickPaymentAmount").focus()
}

window.showHostedReservation = function showHostedReservation(reservationId) {
  const reservation = byId(db.reservations, reservationId)
  if (!reservation) return toast("Hospedagem nao encontrada.")
  document.getElementById("searchHosted").value = reservation.code
  renderHosted()
  showPage("hospedados")
}

window.toggleRoomMaintenance = function toggleRoomMaintenance(roomId) {
  const room = byId(db.rooms, roomId)
  if (!room) return
  if (activeHostingForRoom(roomId)) return toast("Nao coloque em manutencao um quarto com hospede hospedado.")
  room.baseStatus = room.baseStatus === "MANUTENCAO" ? "DISPONIVEL" : "MANUTENCAO"
  saveAndRefresh("Alteracao de quarto", `Quarto ${room.number} => ${room.baseStatus}`)
}

window.delRoom = async function delRoom(roomId) {
  if (!(await askConfirm("Excluir quarto?", "Excluir quarto", "Excluir", "Cancelar"))) return
  if (db.reservations.some((r) => String(r.roomId) === String(roomId))) {
    return toast("Quarto possui reservas vinculadas.")
  }
  const room = byId(db.rooms, roomId)
  db.rooms = db.rooms.filter((r) => String(r.id) !== String(roomId))
  saveAndRefresh("Exclusao de quarto", room?.number || roomId)
}

window.delGuest = async function delGuest(guestId) {
  if (!(await askConfirm("Excluir cliente?", "Excluir cliente", "Excluir", "Cancelar"))) return
  if (db.reservations.some((r) => String(r.guestId) === String(guestId))) {
    return toast("Cliente possui reservas vinculadas.")
  }
  const guest = byId(db.guests, guestId)
  db.guests = db.guests.filter((g) => String(g.id) !== String(guestId))
  saveAndRefresh("Exclusao de cliente", guest?.name || guestId)
}

window.updateReservationStatus = function updateReservationStatus(resId, newStatus) {
  const reservation = byId(db.reservations, resId)
  if (!reservation) return
  reservation.status = newStatus
  if (newStatus === "CHECKOUT") {
    addNotification("Reserva em check-out", `${reservation.code} finalizada.`, "INFO", "RESERVA", reservation.id)
  }
  saveAndRefresh("Alteracao de reserva", `${reservation.code} => ${newStatus}`)
}

window.copyReservationCode = function copyReservationCode(code) {
  navigator.clipboard
    .writeText(code)
    .then(() => toast("Codigo da reserva copiado."))
    .catch(() => toast("Nao foi possivel copiar o codigo."))
}

window.openReservationInPortal = function openReservationInPortal(resId) {
  portalReservationId = resId
  renderGuestPortal(portalReservationId)
  showPage("hospede")
}

window.cancelFinanceEntry = async function cancelFinanceEntry(financeId) {
  const entry = byId(db.finance, financeId)
  if (!entry) return toast("Lancamento financeiro nao encontrado.")
  if (!isFinanceActive(entry)) return toast("Este lancamento ja esta cancelado.")

  const reservation = entry.reservationId ? byId(db.reservations, entry.reservationId) : null
  const reservationText = reservation ? `\nReserva vinculada: ${reservation.code}` : ""
  const message =
    `Cancelar este lancamento?\n\n` +
    `Tipo: ${entry.type}\n` +
    `Valor: ${money(entry.amount)}\n` +
    `Descricao: ${entry.description || "-"}${reservationText}\n\n` +
    `O registro ficara no historico como CANCELADO e nao entrara nos totais.`

  if (!(await askConfirm(message, "Cancelar lancamento", "Cancelar lancamento", "Voltar"))) return

  entry.status = "CANCELADO"
  entry.cancelledAt = new Date().toISOString()
  entry.cancelReason = "Cancelado pelo usuario"

  const shouldReverseExtra =
    reservation &&
    (entry.affectsReservationExtra ||
      entry.origin === "APP_HOSPEDE_CONSUMO" ||
      (entry.category === "consumo_hospede" && String(entry.description || "").startsWith("Consumo solicitado no portal:")))

  if (shouldReverseExtra) {
    const previousExtra = toNumber(reservation.extra)
    reservation.extra = Math.max(0, previousExtra - toNumber(entry.amount))
    entry.reversedReservationExtra = previousExtra - reservation.extra
  }

  addNotification(
    "Lancamento cancelado",
    `${entry.type} de ${money(entry.amount)} foi cancelado.`,
    "WARN",
    "FINANCE",
    entry.id
  )
  saveAndRefresh(
    "Cancelamento financeiro",
    `${entry.type} ${money(entry.amount)} | ${entry.description || "-"}${reservation ? ` | Reserva ${reservation.code}` : ""}`
  )
  toast("Lancamento cancelado e removido dos totais.")
}

window.transferReservationRoom = async function transferReservationRoom(resId) {
  const reservation = byId(db.reservations, resId)
  if (!reservation) return toast("Reserva nao encontrada.")
  if (reservation.status !== "HOSPEDADO") return toast("A transferencia so fica disponivel para reserva hospedada.")

  const fromRoom = byId(db.rooms, reservation.roomId)
  const newRoomId = await askTransferRoom(reservation)
  if (!newRoomId) return

  const toRoom = byId(db.rooms, newRoomId)
  if (!toRoom) return toast("Quarto de destino invalido.")
  if (toRoom.baseStatus === "MANUTENCAO") return toast("Quarto de destino esta em manutencao.")
  if (roomHasReservationConflict(toRoom.id, reservation.checkin, reservation.checkout, reservation.id)) {
    return toast("Quarto de destino possui conflito neste periodo.")
  }

  const keptDailyRate = reservationDailyRate(reservation)
  reservation.dailyRate = keptDailyRate
  reservation.roomTransfers = Array.isArray(reservation.roomTransfers) ? reservation.roomTransfers : []
  reservation.roomTransfers.unshift({
    id: window.store.uid("trf"),
    when: new Date().toISOString(),
    fromRoomId: reservation.roomId,
    toRoomId: toRoom.id,
    fromRoom: fromRoom?.number || "",
    toRoom: toRoom.number || "",
    keptDailyRate
  })
  reservation.roomId = toRoom.id

  addNotification(
    "Transferencia de quarto",
    `${reservation.code}: quarto ${fromRoom?.number || "-"} para ${toRoom.number}.`,
    "INFO",
    "RESERVA",
    reservation.id
  )
  saveAndRefresh(
    "Transferencia de quarto",
    `${reservation.code} | ${fromRoom?.number || "-"} => ${toRoom.number} | Diaria mantida ${money(keptDailyRate)}`
  )
  toast("Hospede transferido de quarto.")
}

function receiptHtml(reservation) {
  const guest = byId(db.guests, reservation.guestId)
  const room = byId(db.rooms, reservation.roomId)
  const total = reservationTotal(reservation)
  const paid = reservationPaid(reservation.id)
  const pending = reservationPending(reservation)
  const payments = activeFinanceEntries()
    .filter((entry) => entry.type === "RECEITA" && String(entry.reservationId) === String(reservation.id))
    .slice()
    .sort((a, b) => String(a.date).localeCompare(String(b.date)))
  const paymentRows = payments.length
    ? payments
        .map(
          (entry) => `
            <tr>
              <td>${fmtDate(entry.date)}</td>
              <td>${escapeHtml(entry.method || "-")}</td>
              <td>${escapeHtml(entry.description || "-")}</td>
              <td>${money(entry.amount)}</td>
            </tr>`
        )
        .join("")
    : '<tr><td colspan="4">Nenhum pagamento registrado.</td></tr>'

  return `
    <!doctype html>
    <html lang="pt-BR">
    <head>
      <meta charset="utf-8">
      <title>Recibo ${escapeHtml(reservation.receiptNumber || reservation.code)}</title>
      <style>
        body { font-family: Arial, sans-serif; color: #111827; margin: 0; padding: 28px; }
        .receipt { max-width: 820px; margin: 0 auto; border: 1px solid #d1d5db; padding: 26px; }
        .top { display: flex; justify-content: space-between; gap: 20px; border-bottom: 2px solid #111827; padding-bottom: 14px; margin-bottom: 18px; }
        h1 { margin: 0; font-size: 28px; letter-spacing: 0.04em; }
        .note { color: #92400e; font-weight: 700; margin-top: 6px; }
        .muted { color: #4b5563; font-size: 13px; }
        .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px 24px; margin: 18px 0; }
        .box { border: 1px solid #e5e7eb; padding: 10px; border-radius: 8px; }
        .label { color: #6b7280; font-size: 12px; text-transform: uppercase; margin-bottom: 4px; }
        .value { font-weight: 700; }
        table { width: 100%; border-collapse: collapse; margin-top: 12px; }
        th, td { border-bottom: 1px solid #e5e7eb; padding: 8px; text-align: left; font-size: 13px; }
        th { background: #f3f4f6; }
        .totals { margin-top: 18px; display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; }
        .text { margin-top: 22px; line-height: 1.6; }
        .signatures { display: grid; grid-template-columns: 1fr 1fr; gap: 42px; margin-top: 52px; }
        .line { border-top: 1px solid #111827; text-align: center; padding-top: 8px; }
        @media print { body { padding: 0; } .receipt { border: 0; } }
      </style>
    </head>
    <body>
      <div class="receipt">
        <div class="top">
          <div>
            <h1>RECIBO DE HOSPEDAGEM</h1>
            <div class="note">Recibo simples, sem valor fiscal</div>
            <div class="muted">Emitido pelo sistema NexHotel</div>
          </div>
          <div>
            <div><strong>Recibo:</strong> ${escapeHtml(reservation.receiptNumber || "-")}</div>
            <div><strong>Reserva:</strong> ${escapeHtml(reservation.code)}</div>
            <div><strong>Emissao:</strong> ${fmtDateTime(reservation.receiptIssuedAt || new Date().toISOString())}</div>
          </div>
        </div>

        <div class="grid">
          <div class="box"><div class="label">Estabelecimento</div><div class="value">________________________________________</div></div>
          <div class="box"><div class="label">Responsavel</div><div class="value">________________________________________</div></div>
          <div class="box"><div class="label">Hospede</div><div class="value">${escapeHtml(guest?.name || "-")}</div></div>
          <div class="box"><div class="label">Documento</div><div class="value">${escapeHtml(guest?.doc || "-")}</div></div>
          <div class="box"><div class="label">Quarto</div><div class="value">${escapeHtml(room?.number || "-")} - ${escapeHtml(room?.category || "-")}</div></div>
          <div class="box"><div class="label">Periodo</div><div class="value">${fmtDate(reservation.checkin)} ate ${fmtDate(reservation.checkout)} (${reservationNights(reservation)} diaria(s))</div></div>
          <div class="box"><div class="label">Diaria</div><div class="value">${money(reservationDailyRate(reservation))}</div></div>
          <div class="box"><div class="label">Formas de pagamento</div><div class="value">${escapeHtml(reservationPaymentMethods(reservation.id))}</div></div>
        </div>

        <table>
          <thead>
            <tr><th>Data</th><th>Metodo</th><th>Descricao</th><th>Valor</th></tr>
          </thead>
          <tbody>${paymentRows}</tbody>
        </table>

        <div class="totals">
          <div class="box"><div class="label">Total</div><div class="value">${money(total)}</div></div>
          <div class="box"><div class="label">Valor pago</div><div class="value">${money(paid)}</div></div>
          <div class="box"><div class="label">Saldo pendente</div><div class="value">${money(pending)}</div></div>
        </div>

        <div class="text">
          Recebemos de <strong>${escapeHtml(guest?.name || "hospede")}</strong> o valor de
          <strong>${money(paid)}</strong>, referente a hospedagem da reserva
          <strong>${escapeHtml(reservation.code)}</strong>. Este documento e um recibo simples de controle interno/comercial
          e nao substitui nota fiscal.
        </div>

        <div class="signatures">
          <div class="line">Assinatura do responsavel</div>
          <div class="line">Assinatura do hospede</div>
        </div>
      </div>
      <script>setTimeout(() => window.print(), 250)</script>
    </body>
    </html>`
}

window.printReceipt = function printReceipt(reservationId) {
  const reservation = byId(db.reservations, reservationId)
  if (!reservation) return toast("Reserva nao encontrada para recibo.")
  if (!reservation.receiptNumber) reservation.receiptNumber = `REC-${String(Date.now()).slice(-8)}`
  reservation.receiptIssuedAt = new Date().toISOString()
  const html = receiptHtml(reservation)
  const w = window.open("", "_blank", "width=920,height=780")
  if (!w) return toast("Nao foi possivel abrir o recibo.")
  w.document.open()
  w.document.write(html)
  w.document.close()
  addAudit("Emissao de recibo", `${reservation.receiptNumber} | ${reservation.code}`)
  saveDb()
  refreshAll()
}

window.closeReservationAccount = async function closeReservationAccount(reservationId) {
  const reservation = byId(db.reservations, reservationId)
  if (!reservation) return toast("Reserva nao encontrada.")
  if (reservation.status !== "HOSPEDADO") return toast("Apenas reservas hospedadas podem ser fechadas por aqui.")

  const pending = reservationPending(reservation)
  if (pending > 0) {
    if (await askConfirm(
      `Ainda falta receber ${money(pending)} desta hospedagem.\n\nDeseja adicionar um pagamento agora?`,
      "Conta com pendencia",
      "Adicionar pagamento",
      "Voltar"
    )) {
      window.openPaymentDialog(reservation.id)
    }
    return
  }

  if (!(await askConfirm(
    `Fechar a conta da reserva ${reservation.code}?\n\nO status mudara para CHECKOUT e o quarto sera liberado.`,
    "Fechar conta",
    "Fechar conta",
    "Voltar"
  ))) return

  reservation.status = "CHECKOUT"
  reservation.closedAt = new Date().toISOString()
  addNotification("Conta fechada", `${reservation.code} finalizada e quitada.`, "INFO", "RESERVA", reservation.id)
  saveAndRefresh("Fechamento de conta", `${reservation.code} | Total ${money(reservationTotal(reservation))}`)
  toast("Conta fechada. Gerando recibo simples.")
  window.printReceipt(reservation.id)
}

window.delReservation = async function delReservation(resId) {
  if (!(await askConfirm("Excluir reserva?", "Excluir reserva", "Excluir", "Cancelar"))) return
  const reservation = byId(db.reservations, resId)
  db.reservations = db.reservations.filter((r) => String(r.id) !== String(resId))
  db.finance = db.finance.filter((f) => String(f.reservationId) !== String(resId))
  db.guestRequests = db.guestRequests.filter((q) => String(q.reservationId) !== String(resId))
  saveAndRefresh("Exclusao de reserva", reservation?.code || resId)
}

window.delStockItem = async function delStockItem(itemId) {
  if (!(await askConfirm("Excluir item de estoque?", "Excluir item", "Excluir", "Cancelar"))) return
  const used = db.stockMovements.some((m) => String(m.itemId) === String(itemId))
  if (used) return toast("Item possui movimentos registrados.")
  const item = byId(db.stockItems, itemId)
  db.stockItems = db.stockItems.filter((x) => String(x.id) !== String(itemId))
  saveAndRefresh("Exclusao de item de estoque", item?.name || itemId)
}

window.delStaff = async function delStaff(staffId) {
  if (!(await askConfirm("Excluir funcionario?", "Excluir funcionario", "Excluir", "Cancelar"))) return
  const hasLinks =
    db.serviceOrders.some((o) => String(o.assigneeId) === String(staffId)) ||
    db.timeEntries.some((p) => String(p.staffId) === String(staffId))
  if (hasLinks) return toast("Funcionario possui ponto/ordens vinculadas.")
  const staff = byId(db.staff, staffId)
  db.staff = db.staff.filter((s) => String(s.id) !== String(staffId))
  saveAndRefresh("Exclusao de funcionario", staff?.name || staffId)
}

window.updateOrderStatus = function updateOrderStatus(orderId, status) {
  const order = byId(db.serviceOrders, orderId)
  if (!order) return
  order.status = status
  if (status === "FINALIZADA") {
    addNotification("OS finalizada", `${order.code} concluida.`, "INFO", "OS", order.id)
  }
  saveAndRefresh("Alteracao de OS", `${order.code} => ${status}`)
}

window.delOrder = async function delOrder(orderId) {
  if (!(await askConfirm("Excluir ordem de servico?", "Excluir OS", "Excluir", "Cancelar"))) return
  const order = byId(db.serviceOrders, orderId)
  db.serviceOrders = db.serviceOrders.filter((o) => String(o.id) !== String(orderId))
  saveAndRefresh("Exclusao de OS", order?.code || orderId)
}

window.delAccount = async function delAccount(accountId) {
  if (!(await askConfirm("Excluir conta bancaria?", "Excluir conta bancaria", "Excluir", "Cancelar"))) return
  if (db.finance.some((f) => String(f.accountId) === String(accountId))) {
    return toast("Conta possui lancamentos financeiros vinculados.")
  }
  const account = byId(db.bankAccounts, accountId)
  db.bankAccounts = db.bankAccounts.filter((a) => String(a.id) !== String(accountId))
  saveAndRefresh("Exclusao de conta bancaria", account?.name || accountId)
}

window.markNotificationRead = function markNotificationRead(notificationId) {
  const n = byId(db.notifications, notificationId)
  if (!n) return
  n.read = true
  saveAndRefresh("Notificacao", `Marcada como lida ${notificationId}`)
}

function bootstrap() {
  ensureLogoFallback()
  bindDialogEvents()
  bindNav()
  bindExternalLinks()
  bindRoomEvents()
  bindGuestEvents()
  bindReservationEvents()
  bindFinanceEvents()
  bindStockEvents()
  bindStaffEvents()
  bindPointEvents()
  bindOrderEvents()
  bindAccountEvents()
  bindReportEvents()
  bindNotificationEvents()
  bindAuditEvents()
  bindGuestPortalEvents()
  bindStaffPanelEvents()
  setDefaults()
  latestReportRows = buildReportRows()
  refreshAll()
}

bootstrap()
