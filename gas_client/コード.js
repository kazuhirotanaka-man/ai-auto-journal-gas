/**
 * AI仕訳生成ツール (配布用クライアント)
 * ※ このスクリプトは非常に薄いラッパーです。
 * 実際の処理はすべてライブラリ（JournalLib）へ移譲します。
 */

function onOpen(e) {
  JournalLib.onOpenHook(e);
}

// === トップメニューからの呼び出し用 ===
function processNewReceipts() {
  JournalLib.processNewReceipts();
}

function showSidebar() {
  JournalLib.showSidebar();
}

function showDialog() {
  JournalLib.showDialog();
}

function exportCsvHandler() {
  JournalLib.exportCsvHandler();
}

function exportFreeeHandler() {
  JournalLib.exportFreeeHandler();
}

function setGeminiApiKey() {
  JournalLib.setGeminiApiKey();
}

// === freee連携メニューからの呼び出し用 ===
function showFreeeSidebar() {
  JournalLib.showFreeeSidebar();
}

function showCompanySelectorSidebar() {
  JournalLib.showCompanySelectorSidebar();
}

function fetchFreeeMasters() {
  JournalLib.fetchFreeeMasters();
}

function toggleAutoRefreshToken() {
  JournalLib.toggleAutoRefreshToken();
}

function runDealLifecycleSample() {
  JournalLib.runDealLifecycleSample();
}

function reset() {
  JournalLib.reset();
}

// === イベントトリガー ===
function onEdit(e) {
  JournalLib.onEdit(e);
}
// === UIコールバック用（google.script.run から呼ばれる）===
function getPopupData(rowIndex) {
  return JournalLib.getPopupData(rowIndex);
}
function updateAndProcessNext(rowIndex, updateData, action) {
  return JournalLib.updateAndProcessNext(rowIndex, updateData, action);
}
function deleteFreeeDetailRow(rowIndex, detailIndex) {
  return JournalLib.deleteFreeeDetailRow(rowIndex, detailIndex);
}
function insertFreeeDetailRow(rowIndex, detailIndex) {
  return JournalLib.insertFreeeDetailRow(rowIndex, detailIndex);
}
function getActiveFileId() {
  return JournalLib.getActiveFileId();
}
function getSidebarState() {
  return JournalLib.getSidebarState();
}
function checkAuth() {
  return JournalLib.checkAuth();
}
function getCompanySelectorInfo() {
  return JournalLib.getCompanySelectorInfo();
}
function saveCompanySelection(companyId) {
  return JournalLib.saveCompanySelection(companyId);
}
function executeExportProcess(statuses) {
  return JournalLib.executeExportProcess(statuses);
}
function executeFreeeExportProcess(statuses) {
  return JournalLib.executeFreeeExportProcess(statuses);
}
