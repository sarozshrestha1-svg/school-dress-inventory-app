const ITEMS = ["3 Piece Set", "Track Only", "T-Shirt Only", "Suit Only", "Custom Item"];
const SALES_ITEMS = ["3 Piece Set", "Track Only", "T-Shirt Only", "Suit Only", "Custom"];
const SIZES = ["16", "18", "20", "22", "24", "26", "28", "30"];
const PENDING_ITEMS = ["T-Shirt Left", "Track Left", "Suit Left", "Nothing Pending"];
const DEFAULT_API_URL = "https://script.google.com/macros/s/AKfycbxD6f-dmyZTvdAghqZSrmE7NLl4lQR2ifiM0iyce7nfHHerUoZVwCuSrer0CfuVBgcJ1A/exec";

const state = {
  stock: [],
  sales: [],
  credits: [],
  pending: [],
  notifications: [],
  dashboard: {},
  reportRows: [],
  reportHeaders: []
};

const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => Array.from(document.querySelectorAll(selector));

document.addEventListener("DOMContentLoaded", () => {
  fillStaticOptions();
  setDefaultDates();
  bindEvents();
  clearOldConnection();
  loadAllData();
});

function fillStaticOptions() {
  fillSelect($("#itemFilter"), ITEMS);
  fillSelect(document.querySelector("#inventoryForm [name='itemName']"), ITEMS);
  fillSelect(document.querySelector("#salesForm [name='itemName']"), SALES_ITEMS);
  fillSelect(document.querySelector("#inventoryForm [name='size']"), SIZES);
  fillSelect(document.querySelector("#salesForm [name='size']"), SIZES);
  fillSelect(document.querySelector("#salesForm [name='itemLeftToGive']"), PENDING_ITEMS);
  fillSelect($("#sizeFilter"), SIZES);
}

function fillSelect(select, values) {
  values.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    select.appendChild(option);
  });
}

function setDefaultDates() {
  const today = new Date().toISOString().slice(0, 10);
  ["dateReceived", "saleDate"].forEach((name) => {
    const input = document.querySelector(`[name='${name}']`);
    if (input) input.value = today;
  });
  $("#reportDate").value = today;
}

function bindEvents() {
  $$(".tab-button").forEach((button) => {
    button.addEventListener("click", () => showPage(button.dataset.page));
  });
  $("#syncButton").addEventListener("click", loadAllData);
  $("#inventoryForm").addEventListener("submit", submitInventory);
  $("#salesForm").addEventListener("submit", submitSale);
  $("#runReport").addEventListener("click", renderReport);
  $("#exportPdf").addEventListener("click", () => window.print());
  $("#exportExcel").addEventListener("click", exportReportExcel);
  ["searchText", "schoolFilter", "itemFilter", "sizeFilter", "dateFilter"].forEach((id) => {
    $(`#${id}`).addEventListener("input", renderAll);
  });
  document.querySelector("#inventoryForm [name='itemName']").addEventListener("change", toggleCustomFields);
  document.querySelector("#salesForm [name='itemName']").addEventListener("change", toggleCustomFields);
  ["quantitySold", "sellingPrice", "additionalCost", "amountPaid"].forEach((name) => {
    document.querySelector(`#salesForm [name='${name}']`).addEventListener("input", updateRemainingAmount);
  });
}

function showPage(pageId) {
  $$(".tab-button").forEach((button) => button.classList.toggle("active", button.dataset.page === pageId));
  $$(".page").forEach((page) => page.classList.toggle("active", page.id === pageId));
}

function getApiUrl() {
  return DEFAULT_API_URL;
}

async function api(action, payload = {}) {
  const url = getApiUrl();
  if (!url) throw new Error("Apps Script URL is not connected.");
  return jsonp(url, { action, ...payload });
}

async function saveEntry(action, payload = {}) {
  try {
    return await api(action, payload);
  } catch (error) {
    if (!isConnectionError(error)) throw error;
    await sendWithMobileFetch(action, payload);
    return { ok: true, message: "Saved using mobile backup connection.", backup: true };
  }
}

function clearOldConnection() {
  try {
    localStorage.removeItem("schoolDressApiUrl");
  } catch (error) {
    // Some mobile browsers block local storage in private mode; the app does not need it.
  }
}

function jsonp(url, payload) {
  return new Promise((resolve, reject) => {
    const callbackName = `schoolDressCallback_${Date.now()}_${Math.random().toString(36).slice(2)}`;
    const script = document.createElement("script");
    const timeout = setTimeout(() => {
      cleanup();
      reject(new Error("Connection timed out. Check the Apps Script Web App URL."));
    }, 30000);
    const params = new URLSearchParams({
      callback: callbackName,
      payload: JSON.stringify(payload)
    });

    function cleanup() {
      clearTimeout(timeout);
      script.remove();
      delete window[callbackName];
    }

    window[callbackName] = (data) => {
      cleanup();
      if (!data.ok) reject(new Error(data.message || "Request failed."));
      else resolve(data);
    };
    script.onerror = () => {
      cleanup();
      reject(new Error(`Could not connect to Google Apps Script. Open this link once on this phone, then return to the app: ${DEFAULT_API_URL}`));
    };
    script.src = `${url}${url.includes("?") ? "&" : "?"}${params.toString()}`;
    document.body.appendChild(script);
  });
}

function isConnectionError(error) {
  const message = String(error && error.message ? error.message : error);
  return message.includes("Could not connect") || message.includes("Connection timed out");
}

async function sendWithMobileFetch(action, payload) {
  if (!window.fetch) {
    throw new Error("This mobile browser cannot send data to Google Sheets. Please update Chrome and try again.");
  }
  const params = new URLSearchParams({
    mobileSave: String(Date.now()),
    payload: JSON.stringify({ action, ...payload })
  });
  await fetch(`${getApiUrl()}?${params.toString()}`, {
    method: "GET",
    mode: "no-cors",
    cache: "no-store",
    credentials: "omit",
    redirect: "follow"
  });
}

async function loadAllData(options = {}) {
  const silent = Boolean(options.silent);
  try {
    if (!silent) showToast("Loading latest records...");
    const data = await api("getAllData");
    Object.assign(state, {
      stock: data.stock || [],
      sales: data.sales || [],
      credits: data.credits || [],
      pending: data.pending || [],
      notifications: data.notifications || [],
      dashboard: data.dashboard || {}
    });
    renderAll();
    $("#lastUpdated").textContent = `Synced ${new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })}`;
    if (!silent) showToast("Records updated.");
  } catch (error) {
    showToast(error.message);
  }
}

async function submitInventory(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const submitButton = form.querySelector("button[type='submit']");
  const payload = formToObject(form);
  payload.itemName = normalizeCustomItem(payload.itemName, payload.customItem);
  try {
    setSubmitting(submitButton, true, "Saving...");
    const result = await saveEntry("addInventory", { inventory: payload });
    resetEntryForm(form);
    if (!result.backup) await loadAllData({ silent: true });
    showToast(result.backup ? "Inventory sent to Google Sheets. Form cleared." : "Inventory saved. Form cleared and stock refreshed.");
  } catch (error) {
    showToast(error.message);
  } finally {
    setSubmitting(submitButton, false);
  }
}

async function submitSale(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const submitButton = form.querySelector("button[type='submit']");
  const payload = formToObject(form);
  payload.itemName = normalizeCustomItem(payload.itemName, payload.customItem);
  updateRemainingAmount();
  payload.remainingAmount = form.remainingAmount.value;
  try {
    setSubmitting(submitButton, true, "Saving...");
    const result = await saveEntry("addSale", { sale: payload });
    resetEntryForm(form);
    if (!result.backup) await loadAllData({ silent: true });
    showToast(result.backup ? "Sale sent to Google Sheets. Form cleared." : "Sale saved. Form cleared and stock refreshed.");
  } catch (error) {
    showToast(error.message);
  } finally {
    setSubmitting(submitButton, false);
  }
}

function resetEntryForm(form) {
  form.reset();
  setDefaultDates();
  toggleCustomFields();
  const remainingField = form.querySelector("[name='remainingAmount']");
  if (remainingField) remainingField.value = "";
  form.scrollIntoView({ behavior: "smooth", block: "start" });
}

function setSubmitting(button, isSubmitting, label) {
  if (!button) return;
  if (!button.dataset.defaultText) button.dataset.defaultText = button.textContent;
  button.disabled = isSubmitting;
  button.textContent = isSubmitting ? label : button.dataset.defaultText;
}

function formToObject(form) {
  return Object.fromEntries(new FormData(form).entries());
}

function normalizeCustomItem(itemName, customItem) {
  if (itemName.toLowerCase().startsWith("custom") && customItem.trim()) return customItem.trim();
  return itemName;
}

function toggleCustomFields() {
  const inventoryItem = document.querySelector("#inventoryForm [name='itemName']").value;
  const salesItem = document.querySelector("#salesForm [name='itemName']").value;
  $("#customInventoryWrap").classList.toggle("hidden", inventoryItem !== "Custom Item");
  $("#customSalesWrap").classList.toggle("hidden", salesItem !== "Custom");
}

function updateRemainingAmount() {
  const form = $("#salesForm");
  const qty = number(form.quantitySold.value);
  const price = number(form.sellingPrice.value);
  const extra = number(form.additionalCost.value);
  const paid = number(form.amountPaid.value);
  const total = qty * price + extra;
  const remaining = Math.max(total - paid, 0);
  form.remainingAmount.value = remaining.toFixed(2);
  form.creditSale.value = remaining > 0 ? "Yes" : "No";
}

function renderAll() {
  refreshFilterOptions();
  renderDashboard();
  renderStockTable();
  renderCreditTable();
  renderPendingTable();
  renderReport();
}

function refreshFilterOptions() {
  const itemSelect = $("#itemFilter");
  const currentItem = itemSelect.value;
  const allItems = new Set(["", ...ITEMS]);
  [...state.stock, ...state.sales, ...state.pending].forEach((row) => {
    const value = row.itemName || row.itemSold;
    if (value) allItems.add(value);
  });
  itemSelect.innerHTML = "";
  allItems.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value || "All items";
    itemSelect.appendChild(option);
  });
  itemSelect.value = allItems.has(currentItem) ? currentItem : "";
}

function renderDashboard() {
  const filteredStock = filterRows(state.stock);
  const filteredSales = filterRows(state.sales);
  const filteredCredits = filterRows(state.credits);
  const filteredPending = filterRows(state.pending);
  $("#totalInventory").textContent = sum(filteredStock, "availableQuantity");
  $("#totalSales").textContent = money(sum(filteredSales, "totalBill"));
  $("#totalCredit").textContent = money(sum(filteredCredits, "remainingAmount"));
  $("#pendingCount").textContent = filteredPending.length;
  $("#lowStockCount").textContent = filteredStock.filter((row) => number(row.availableQuantity) > 0 && number(row.availableQuantity) <= 5).length;
  $("#outStockCount").textContent = filteredStock.filter((row) => number(row.availableQuantity) === 0).length;
  $("#creditPageTotal").textContent = `Outstanding: ${money(sum(filteredCredits, "remainingAmount"))}`;
  $("#pendingPageTotal").textContent = `${filteredPending.length} pending`;
}

function renderStockTable() {
  const rows = filterRows(state.stock);
  fillTable($("#stockTable"), rows, (row) => [
    row.schoolName,
    row.itemName,
    row.size,
    row.availableQuantity,
    statusBadge(row.status)
  ]);
}

function renderCreditTable() {
  const rows = filterRows(state.credits);
  fillTable($("#creditTable"), rows, (row) => [
    row.studentName,
    row.schoolName,
    money(row.totalBill),
    money(row.amountPaid),
    money(row.remainingAmount),
    formatDate(row.saleDate)
  ]);
}

function renderPendingTable() {
  const rows = filterRows(state.pending);
  fillTable($("#pendingTable"), rows, (row) => [
    row.studentName,
    row.schoolName,
    row.itemSold,
    row.pendingItem,
    formatDate(row.saleDate)
  ]);
}

function renderReport() {
  const type = $("#reportType").value;
  const date = $("#reportDate").value || new Date().toISOString().slice(0, 10);
  const { title, headers, rows } = buildReport(type, date);
  state.reportHeaders = headers;
  state.reportRows = rows;
  $("#reportTitle").textContent = title;
  const thead = $("#reportTable thead");
  const tbody = $("#reportTable tbody");
  thead.innerHTML = `<tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr>`;
  fillTable(tbody, rows, (row) => headers.map((header) => row[header] ?? ""));
}

function buildReport(type, date) {
  if (type === "inventory") {
    return report("Inventory Report", ["School", "Item", "Size", "Available", "Status"], filterRows(state.stock).map((row) => ({
      School: row.schoolName,
      Item: row.itemName,
      Size: row.size,
      Available: row.availableQuantity,
      Status: row.status
    })));
  }
  if (type === "credit") {
    return report("Credit Report", ["Student", "School", "Total Bill", "Paid", "Remaining", "Date"], filterRows(state.credits).map((row) => ({
      Student: row.studentName,
      School: row.schoolName,
      "Total Bill": money(row.totalBill),
      Paid: money(row.amountPaid),
      Remaining: money(row.remainingAmount),
      Date: formatDate(row.saleDate)
    })));
  }
  if (type === "pending") {
    return report("Pending Delivery Report", ["Student", "School", "Item Sold", "Pending Item", "Date"], filterRows(state.pending).map((row) => ({
      Student: row.studentName,
      School: row.schoolName,
      "Item Sold": row.itemSold,
      "Pending Item": row.pendingItem,
      Date: formatDate(row.saleDate)
    })));
  }

  const salesRows = filterRows(state.sales).filter((row) => isInReportPeriod(row.saleDate, date, type));
  const title = type === "dailySales" ? `Daily Sales: ${date}` : type === "weeklySales" ? `Weekly Sales: Week of ${date}` : `Monthly Sales: ${date.slice(0, 7)}`;
  return report(title, ["Date", "Student", "School", "Item", "Size", "Qty", "Total Bill", "Paid", "Remaining"], salesRows.map((row) => ({
    Date: formatDate(row.saleDate),
    Student: row.studentName,
    School: row.schoolName,
    Item: row.itemName,
    Size: row.size,
    Qty: row.quantitySold,
    "Total Bill": money(row.totalBill),
    Paid: money(row.amountPaid),
    Remaining: money(row.remainingAmount)
  })));
}

function report(title, headers, rows) {
  return { title, headers, rows };
}

function isInReportPeriod(rowDate, selectedDate, type) {
  const date = normalizeDate(rowDate);
  if (!date) return false;
  if (type === "dailySales") return date === selectedDate;
  if (type === "monthlySales") return date.slice(0, 7) === selectedDate.slice(0, 7);
  const current = new Date(`${date}T00:00:00`);
  const selected = new Date(`${selectedDate}T00:00:00`);
  const day = selected.getDay();
  const mondayOffset = day === 0 ? -6 : 1 - day;
  const start = new Date(selected);
  start.setDate(selected.getDate() + mondayOffset);
  const end = new Date(start);
  end.setDate(start.getDate() + 6);
  return current >= start && current <= end;
}

function filterRows(rows) {
  const search = $("#searchText").value.trim().toLowerCase();
  const school = $("#schoolFilter").value.trim().toLowerCase();
  const item = $("#itemFilter").value;
  const size = $("#sizeFilter").value;
  const date = $("#dateFilter").value;
  return rows.filter((row) => {
    const haystack = Object.values(row).join(" ").toLowerCase();
    const rowDate = normalizeDate(row.saleDate || row.dateReceived || row.date || "");
    return (!search || haystack.includes(search)) &&
      (!school || String(row.schoolName || row.school || "").toLowerCase().includes(school)) &&
      (!item || row.itemName === item || row.itemSold === item) &&
      (!size || String(row.size) === size) &&
      (!date || rowDate === date);
  });
}

function fillTable(tbody, rows, mapper) {
  if (!rows.length) {
    tbody.innerHTML = `<tr><td colspan="12">No records found.</td></tr>`;
    return;
  }
  tbody.innerHTML = rows.map((row) => `<tr>${mapper(row).map((cell) => `<td>${cell}</td>`).join("")}</tr>`).join("");
}

function statusBadge(status) {
  const lower = String(status || "In Stock").toLowerCase();
  const className = lower.includes("out") ? "out" : lower.includes("low") ? "low" : "ok";
  return `<span class="status ${className}">${escapeHtml(status || "In Stock")}</span>`;
}

function exportReportExcel() {
  if (!state.reportRows.length) return showToast("Generate a report first.");
  const headerRow = state.reportHeaders.map(escapeHtml).join("</th><th>");
  const bodyRows = state.reportRows.map((row) => {
    return `<tr>${state.reportHeaders.map((header) => `<td>${escapeHtml(row[header] ?? "")}</td>`).join("")}</tr>`;
  }).join("");
  const html = `<table><thead><tr><th>${headerRow}</th></tr></thead><tbody>${bodyRows}</tbody></table>`;
  const blob = new Blob([html], { type: "application/vnd.ms-excel" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `${$("#reportType").value}-${new Date().toISOString().slice(0, 10)}.xls`;
  link.click();
  URL.revokeObjectURL(link.href);
}

function showToast(message) {
  const toast = $("#toast");
  toast.textContent = message;
  toast.classList.add("show");
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => toast.classList.remove("show"), 2600);
}

function sum(rows, field) {
  return rows.reduce((total, row) => total + number(row[field]), 0);
}

function number(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function money(value) {
  return number(value).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function normalizeDate(value) {
  if (!value) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(String(value))) return String(value);
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  return date.toISOString().slice(0, 10);
}

function formatDate(value) {
  return normalizeDate(value) || value || "";
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;"
  }[char]));
}
