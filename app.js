/**
 * SheetView - Excel Interactive Viewer
 * Application Logic
 */

document.addEventListener('DOMContentLoaded', () => {
    // --- Application State ---
    let state = {
        workbook: null,
        fileName: '',
        fileSize: '',
        sheetNames: [],
        activeSheetName: '',
        headers: [],       // Array of header names
        sheetData: [],     // 2D Array of all rows (excluding header row)
        filteredData: [],  // Rows after search filter
        currentPage: 1,
        pageSize: 25,
        sortColumnIndex: null,
        sortDirection: 'asc',
        viewMode: 'list',
        groupFilter: 'all',
        expandedGroups: new Set(),
        theme: 'dark'
    };

    let qrExportState = null;

    const CONTACT_COLUMN_FALLBACK = 7; // Cột H nếu không tìm thấy tên "LIÊN HỆ"
    const NAME_COLUMN_FALLBACK = 4; // Cột E nếu không tìm thấy tên "HỌ VÀ TÊN"
    const CCCD_COLUMN_FALLBACK = 6; // Cột G nếu không tìm thấy tên "CCCD"
    const EMPTY_CONTACT_LABEL = '(Chưa có liên hệ)';
    const PDF_WORKER_SRC = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    const PDF_MARKER_KEYS = ['2.hotenthisinh:', '2.hovatenthisinh:'];
    const PDF_MARKER_RE = /2\.\s*(?:Họ|Ho|HỌ)(?:\s*(?:và|va))?\s*(?:tên|ten|TÊN)\s*(?:thí\s*sinh|thi\s*sinh|THÍ\s*SINH)?\s*:/i;

    // --- DOM Elements ---
    const themeToggleBtn = document.getElementById('themeToggleBtn');
    const themeIconSun = document.getElementById('themeIconSun');
    const themeIconMoon = document.getElementById('themeIconMoon');
    const uploadSection = document.getElementById('uploadSection');
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const workspaceSection = document.getElementById('workspaceSection');
    
    const fileNameDisplay = document.getElementById('fileNameDisplay');
    const fileSizeDisplay = document.getElementById('fileSizeDisplay');
    const createQrBtn = document.getElementById('createQrBtn');
    const pdfInput = document.getElementById('pdfInput');
    const qrStatsModal = document.getElementById('qrStatsModal');
    const qrStatsCloseBtn = document.getElementById('qrStatsCloseBtn');
    const qrStatsSubtitle = document.getElementById('qrStatsSubtitle');
    const qrStatsAlert = document.getElementById('qrStatsAlert');
    const qrStatsSummary = document.getElementById('qrStatsSummary');
    const qrFoundCount = document.getElementById('qrFoundCount');
    const qrMissingPdfCount = document.getElementById('qrMissingPdfCount');
    const qrMissingExcelCount = document.getElementById('qrMissingExcelCount');
    const qrFoundTable = document.getElementById('qrFoundTable');
    const qrMissingPdfTable = document.getElementById('qrMissingPdfTable');
    const qrMissingExcelTable = document.getElementById('qrMissingExcelTable');
    const searchInput = document.getElementById('searchInput');
    const clearSearchBtn = document.getElementById('clearSearchBtn');
    
    const viewModeSelect = document.getElementById('viewModeSelect');
    const groupFilterWrap = document.getElementById('groupFilterWrap');
    const groupFilterSelect = document.getElementById('groupFilterSelect');
    const pageSizeWrap = document.getElementById('pageSizeWrap');
    const pageSizeSelect = document.getElementById('pageSizeSelect');
    const rowCounter = document.getElementById('rowCounter');
    
    const tableHead = document.getElementById('tableHead');
    const tableBody = document.getElementById('tableBody');
    
    const paginationContainer = document.getElementById('paginationContainer');
    const pageNumbers = document.getElementById('pageNumbers');
    const prevPageBtn = document.getElementById('prevPageBtn');
    const nextPageBtn = document.getElementById('nextPageBtn');
    
    const loadingOverlay = document.getElementById('loadingOverlay');
    const loadingText = document.getElementById('loadingText');
    const notificationContainer = document.getElementById('notificationContainer');

    // --- Initialize Lucide Icons (không được chặn app nếu lỗi) ---
    function initLucideIcons() {
        if (typeof lucide !== 'undefined' && typeof lucide.createIcons === 'function') {
            lucide.createIcons();
        }
    }

    // --- Kiểm tra thư viện bắt buộc ---
    function checkLibraries() {
        const missing = [];
        if (typeof XLSX === 'undefined') missing.push('SheetJS (xlsx)');
        if (missing.length > 0) {
            showToast('Thiếu thư viện: ' + missing.join(', ') + '. Kiểm tra mạng và tải lại trang (Ctrl+F5).', 'error');
            return false;
        }
        return true;
    }

    // --- Theme Controller ---
    function initTheme() {
        const savedTheme = localStorage.getItem('sheetview-theme') || 'dark';
        setTheme(savedTheme);
    }

    function setTheme(theme) {
        state.theme = theme;
        document.documentElement.setAttribute('data-theme', theme);
        localStorage.setItem('sheetview-theme', theme);
        
        if (theme === 'dark') {
            themeIconSun.classList.remove('hidden');
            themeIconMoon.classList.add('hidden');
        } else {
            themeIconSun.classList.add('hidden');
            themeIconMoon.classList.remove('hidden');
        }
    }

    themeToggleBtn.addEventListener('click', () => {
        setTheme(state.theme === 'dark' ? 'light' : 'dark');
    });

    // --- Notifications ---
    function showToast(message, type = 'success') {
        const toast = document.createElement('div');
        toast.className = `toast toast-${type}`;
        
        let iconName = 'check-circle';
        if (type === 'error') iconName = 'alert-triangle';
        if (type === 'warning') iconName = 'info';
        
        toast.innerHTML = `
            <i data-lucide="${iconName}"></i>
            <span>${message}</span>
        `;
        
        notificationContainer.appendChild(toast);
        if (typeof lucide !== 'undefined' && typeof lucide.createIcons === 'function') {
            lucide.createIcons({ attrs: { class: 'toast-icon' } });
        }

        // Auto remove
        setTimeout(() => {
            toast.style.animation = 'slideIn 0.3s reverse forwards';
            toast.addEventListener('animationend', () => {
                toast.remove();
            });
        }, 3500);
    }

    // --- Loading Utility ---
    function showLoading(text) {
        loadingText.textContent = text || 'Processing...';
        loadingOverlay.classList.remove('hidden');
    }

    function hideLoading() {
        loadingOverlay.classList.add('hidden');
    }

    // --- File Drag & Drop & Selection ---
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => {
            dropZone.classList.remove('drag-over');
        });
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            handleFile(files[0]);
        }
    });

    fileInput.addEventListener('change', (e) => {
        const files = e.target.files;
        if (files.length > 0) {
            handleFile(files[0]);
        }
        e.target.value = '';
    });

    // --- File Parser (SheetJS) ---
    function handleFile(file) {
        if (!file) return;
        if (!checkLibraries()) return;
        
        const validExtensions = ['.xlsx', '.xls', '.csv'];
        const fileExt = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
        if (!validExtensions.includes(fileExt)) {
            showToast('Định dạng không hợp lệ. Vui lòng tải file .xlsx, .xls hoặc .csv.', 'error');
            return;
        }

        state.fileName = file.name;
        state.fileSize = formatBytes(file.size);

        showLoading('Đang đọc file Excel...');
        
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { 
                    type: 'array',
                    cellDates: true
                });
                
                processWorkbook(workbook);
                showToast('Đã tải workbook thành công!', 'success');
            } catch (err) {
                console.error(err);
                showToast('Không thể đọc file Excel. File có thể bị hỏng.', 'error');
                hideLoading();
            }
        };
        reader.onerror = () => {
            showToast('Lỗi khi đọc file.', 'error');
            hideLoading();
        };
        reader.readAsArrayBuffer(file);
    }

    function formatBytes(bytes, decimals = 1) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const dm = decimals < 0 ? 0 : decimals;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
    }

    // Lấy danh sách sheet không bị ẩn (Hidden: 1 = ẩn, 2 = ẩn hoàn toàn)
    function getVisibleSheetNames(workbook) {
        const allNames = workbook.SheetNames || [];
        const sheetsMeta = workbook.Workbook && workbook.Workbook.Sheets;

        if (!sheetsMeta || !Array.isArray(sheetsMeta)) {
            return allNames;
        }

        const hiddenByName = new Map();
        sheetsMeta.forEach(meta => {
            if (meta && meta.name != null) {
                hiddenByName.set(meta.name, meta.Hidden || 0);
            }
        });

        return allNames.filter(name => {
            const hidden = hiddenByName.get(name) ?? 0;
            return hidden !== 1 && hidden !== 2;
        });
    }

    // --- Process Workbook Data ---
    function processWorkbook(workbook) {
        try {
            state.workbook = workbook;

            const visibleSheets = getVisibleSheetNames(workbook);
            if (visibleSheets.length === 0) {
                showToast('Không có sheet hiển thị nào trong file.', 'warning');
                hideLoading();
                return;
            }

            const firstSheet = visibleSheets[0];
            state.sheetNames = [firstSheet];

            uploadSection.classList.add('hidden');
            workspaceSection.classList.remove('hidden');
            fileNameDisplay.textContent = state.fileName;
            fileSizeDisplay.textContent = `${state.fileSize} • Sheet: ${firstSheet}`;

            loadSheet(firstSheet);

            setTimeout(() => {
                workspaceSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }, 150);
        } catch (err) {
            console.error(err);
            showToast('Lỗi hiển thị dữ liệu: ' + err.message, 'error');
        } finally {
            hideLoading();
        }
    }

    // Kiểm tra ô có giá trị "STT" / "stt"
    function cellMatchesStt(value) {
        if (value === null || value === undefined) return false;
        return String(value).trim().replace(/\u00a0/g, ' ').toLowerCase() === 'stt';
    }

    // Tìm dòng tiêu đề chứa cột STT
    function findHeaderRowIndex(rawRows) {
        for (let r = 0; r < rawRows.length; r++) {
            const row = rawRows[r];
            if (!row || !row.length) continue;
            if (row.some(cell => cellMatchesStt(cell))) {
                return r;
            }
        }
        return -1;
    }

    function getSttColumnIndex(rawHeaders, maxCols) {
        for (let c = 0; c < maxCols; c++) {
            if (cellMatchesStt(rawHeaders[c])) return c;
        }
        return 0;
    }

    function cellValueIsSttOne(value) {
        if (value === null || value === undefined || value === '') return false;
        const num = Number(value);
        if (!isNaN(num) && num === 1) return true;
        return String(value).trim() === '1';
    }

    // Tìm dòng dữ liệu đầu tiên có STT = 1
    function findFirstDataRowIndex(rawRows, headerRowIndex, sttColIndex) {
        for (let r = headerRowIndex + 1; r < rawRows.length; r++) {
            const row = rawRows[r];
            if (!row) continue;
            if (cellValueIsSttOne(row[sttColIndex])) return r;
        }
        return -1;
    }

    function isCellEmpty(value) {
        if (value === null || value === undefined) return true;
        if (value instanceof Date) return false;
        if (typeof value === 'number' && !isNaN(value)) return false;
        return String(value).trim() === '';
    }

    function isRowFullyEmpty(rowData, colCount) {
        for (let c = 0; c < colCount; c++) {
            if (!isCellEmpty(rowData[c])) return false;
        }
        return true;
    }

    // --- Load Specific Sheet ---
    function loadSheet(sheetName) {
        state.activeSheetName = sheetName;

        const sheet = state.workbook.Sheets[sheetName];
        
        const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
        
        if (rawRows.length === 0) {
            state.headers = [];
            state.sheetData = [];
            state.filteredData = [];
            renderEmptyGrid();
            showToast('Sheet này trống.', 'warning');
            return;
        }

        const headerRowIndex = findHeaderRowIndex(rawRows);
        if (headerRowIndex === -1) {
            state.headers = [];
            state.sheetData = [];
            state.filteredData = [];
            renderEmptyGrid();
            showToast('Không tìm thấy dòng tiêu đề chứa "STT" hoặc "stt".', 'warning');
            return;
        }

        const MAX_COLUMNS = 8;
        const colCounts = rawRows.map(r => (r && r.length) ? r.length : 0);
        const totalCols = colCounts.length ? Math.max(...colCounts) : 0;
        const maxCols = Math.min(totalCols, MAX_COLUMNS);

        if (maxCols === 0) {
            state.headers = [];
            state.sheetData = [];
            state.filteredData = [];
            renderEmptyGrid();
            showToast('Sheet không có cột dữ liệu.', 'warning');
            return;
        }
        
        const rawHeaders = rawRows[headerRowIndex];
        state.headers = [];
        
        for (let c = 0; c < maxCols; c++) {
            let label = rawHeaders[c];
            if (label === undefined || label === null || label === '') {
                label = `Cột ${c + 1}`;
            }
            state.headers.push(String(label).trim());
        }

        const sttColIndex = getSttColumnIndex(rawHeaders, maxCols);
        const firstDataRowIndex = findFirstDataRowIndex(rawRows, headerRowIndex, sttColIndex);

        if (firstDataRowIndex === -1) {
            state.headers = [];
            state.sheetData = [];
            state.filteredData = [];
            renderEmptyGrid();
            showToast('Không tìm thấy dòng có STT = 1.', 'warning');
            return;
        }

        state.sheetData = [];
        for (let r = firstDataRowIndex; r < rawRows.length; r++) {
            const rawRow = rawRows[r];
            const rowData = [];
            for (let c = 0; c < maxCols; c++) {
                rowData.push(rawRow[c] !== undefined ? rawRow[c] : '');
            }
            if (isRowFullyEmpty(rowData, maxCols)) continue;

            rowData._originalIndex = r + 1;
            state.sheetData.push(rowData);
        }

        // Reset states
        state.filteredData = [...state.sheetData];
        state.currentPage = 1;
        state.sortColumnIndex = null;
        state.sortDirection = 'asc';
        state.groupFilter = 'all';
        state.expandedGroups = new Set();
        searchInput.value = '';
        clearSearchBtn.classList.add('hidden');

        populateGroupFilter();

        try {
            renderTable();
        } catch (err) {
            console.error(err);
            renderEmptyGrid();
            showToast('Lỗi render bảng: ' + err.message, 'error');
        }
    }

    function renderEmptyGrid() {
        tableHead.innerHTML = '<tr><th>#</th><th>Sheet trống</th></tr>';
        tableBody.innerHTML = '<tr><td colspan="2" style="text-align: center; padding: 40px; color: var(--text-muted);">Không có dữ liệu trong sheet này.</td></tr>';
        rowCounter.textContent = 'Hiển thị 0-0 / 0 dòng';
        paginationContainer.classList.add('hidden');
        groupFilterWrap.classList.add('hidden');
        createQrBtn.classList.add('hidden');
    }

    const BULK_LIMITS = [250, 350, 550];
    const DEFAULT_MAX_ROWS = 550;
    let searchDebounceTimer = null;

    function isBulkPageSize() {
        return BULK_LIMITS.includes(parseInt(state.pageSize));
    }

    function getMaxRowsCap() {
        const size = parseInt(state.pageSize);
        if (BULK_LIMITS.includes(size)) return size;
        return DEFAULT_MAX_ROWS;
    }

    function normalizeHeaderName(value) {
        return String(value).trim().replace(/\u00a0/g, ' ').toLowerCase();
    }

    function getContactColumnIndex() {
        const idx = state.headers.findIndex(h =>
            normalizeHeaderName(h) === 'liên hệ'
        );
        return idx >= 0 ? idx : CONTACT_COLUMN_FALLBACK;
    }

    function getNameColumnIndex() {
        const idx = state.headers.findIndex(h => {
            const n = normalizeHeaderName(h);
            return n === 'họ và tên' || n.includes('họ và tên') || n === 'họ tên' || n.includes('họ tên');
        });
        return idx >= 0 ? idx : NAME_COLUMN_FALLBACK;
    }

    function getCccdColumnIndex() {
        const idx = state.headers.findIndex(h => {
            const n = normalizeHeaderName(h);
            return n === 'cccd' || n === 'cmnd' || n.includes('cccd') || n.includes('cmnd');
        });
        return idx >= 0 ? idx : CCCD_COLUMN_FALLBACK;
    }

    function foldVietnameseChars(str) {
        return String(str)
            .normalize('NFC')
            .replace(/đ/g, 'd')
            .replace(/Đ/g, 'D')
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '');
    }

    function sanitizeName(raw) {
        if (raw == null) return '';

        return String(raw)
            .trim()
            .replace(/\u00a0/g, ' ')
            .replace(/\s+/g, ' ')
            .replace(/\s*[-–—].*$/, '')
            .trim();
    }

    function sanitizeCccd(raw) {
        if (raw == null) return '';

        let text = formatCellValue(raw).trim().replace(/\u00a0/g, ' ');
        if (!text) return '';

        const digitsOnly = text.replace(/\D/g, '');
        if (digitsOnly.length >= 9 && digitsOnly.length <= 12) {
            return digitsOnly;
        }

        const match = text.match(/\d{9,12}/);
        return match ? match[0] : digitsOnly;
    }

    function createCccdMatchKey(cccd) {
        const sanitized = sanitizeCccd(cccd);
        return sanitized.length >= 9 ? sanitized : '';
    }

    function getStudentNameValue(row) {
        return sanitizeName(formatCellValue(row[getNameColumnIndex()]));
    }

    function getStudentCccdValue(row) {
        return sanitizeCccd(row[getCccdColumnIndex()]);
    }

    function initPdfJs() {
        if (typeof pdfjsLib === 'undefined') return false;
        if (!pdfjsLib.GlobalWorkerOptions.workerSrc) {
            pdfjsLib.GlobalWorkerOptions.workerSrc = PDF_WORKER_SRC;
        }
        return true;
    }

    async function extractTextFromPdfPage(page) {
        const textContent = await page.getTextContent();
        const items = textContent.items
            .filter(item => item.str && item.str.trim())
            .map(item => ({
                str: item.str,
                x: item.transform[4],
                y: item.transform[5]
            }));

        items.sort((a, b) => {
            const yDiff = b.y - a.y;
            if (Math.abs(yDiff) > 4) return yDiff;
            return a.x - b.x;
        });

        const lines = [];
        let currentLine = [];
        let currentY = null;

        items.forEach(item => {
            if (currentY === null || Math.abs(item.y - currentY) <= 4) {
                currentLine.push(item);
                if (currentY === null) currentY = item.y;
            } else {
                if (currentLine.length) {
                    lines.push(
                        currentLine
                            .sort((a, b) => a.x - b.x)
                            .map(part => part.str)
                            .join(' ')
                    );
                }
                currentLine = [item];
                currentY = item.y;
            }
        });

        if (currentLine.length) {
            lines.push(
                currentLine
                    .sort((a, b) => a.x - b.x)
                    .map(part => part.str)
                    .join(' ')
            );
        }

        return lines.join('\n');
    }

    function buildPdfCandidate(pageNumber, rawNamePart, cccd) {
        const displayName = sanitizeName(rawNamePart);
        const matchKey = createCccdMatchKey(cccd);
        if (!matchKey) return null;

        return { pageNumber, displayName, cccd: matchKey, matchKey };
    }

    function extractPdfCandidateFromCompactText(text) {
        const compact = text.replace(/\s+/g, '');
        const foldedCompact = foldVietnameseChars(compact).toLowerCase();
        let nameStart = -1;

        for (const key of PDF_MARKER_KEYS) {
            const idx = foldedCompact.indexOf(key);
            if (idx >= 0) {
                nameStart = idx + key.length;
                break;
            }
        }

        if (nameStart < 0) return null;

        const tail = compact.slice(nameStart);
        const foldedTail = foldedCompact.slice(nameStart);
        const cccdMatch = foldedTail.match(/cccd(\d{9,12})/i);
        if (!cccdMatch) return null;

        const cccdIdx = foldedTail.indexOf(cccdMatch[0]);
        const displayName = tail.slice(0, cccdIdx).replace(/[-–—]+$/, '').trim();

        return {
            displayName,
            cccd: cccdMatch[1]
        };
    }

    function extractPdfCandidateFromSpacedLine(line) {
        const normalized = line.replace(/\s+/g, ' ').trim();
        const markerMatch = normalized.match(PDF_MARKER_RE);
        if (!markerMatch) return null;

        const afterMarker = normalized.slice(markerMatch.index + markerMatch[0].length);
        const cccdMatch = afterMarker.match(/^(.+?)\s*[-–—]\s*CCCD\s*(\d{9,12})/i);
        if (!cccdMatch) return null;

        return {
            displayName: cccdMatch[1],
            cccd: cccdMatch[2]
        };
    }

    function parseCandidateFromPageText(text, pageNumber) {
        const compactCandidate = extractPdfCandidateFromCompactText(text);
        if (compactCandidate) {
            const candidate = buildPdfCandidate(
                pageNumber,
                compactCandidate.displayName,
                compactCandidate.cccd
            );
            if (candidate) return candidate;
        }

        for (const line of text.split('\n')) {
            const spacedCandidate = extractPdfCandidateFromSpacedLine(line);
            if (spacedCandidate) {
                const candidate = buildPdfCandidate(
                    pageNumber,
                    spacedCandidate.displayName,
                    spacedCandidate.cccd
                );
                if (candidate) return candidate;
            }
        }

        return null;
    }

    function indexPdfCandidate(matchIndex, candidate) {
        if (!matchIndex.has(candidate.matchKey)) {
            matchIndex.set(candidate.matchKey, []);
        }
        matchIndex.get(candidate.matchKey).push(candidate);
    }

    function findPdfMatches(pdfResult, matchKey) {
        if (!matchKey) return null;
        return pdfResult.matchIndex.get(matchKey) || null;
    }

    async function parsePdfCandidates(file, buffer) {
        if (!initPdfJs()) {
            throw new Error('Thiếu thư viện PDF.js. Kiểm tra mạng và tải lại trang.');
        }

        const pdfData = buffer || await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: pdfData }).promise;
        const candidates = [];
        const matchIndex = new Map();
        let parsedPages = 0;

        for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
            if (pageNum === 1 || pageNum % 25 === 0 || pageNum === pdf.numPages) {
                loadingText.textContent = `Đang đọc PDF: trang ${pageNum}/${pdf.numPages}...`;
            }

            const page = await pdf.getPage(pageNum);
            const text = await extractTextFromPdfPage(page);
            const candidate = parseCandidateFromPageText(text, pageNum);

            if (candidate) {
                parsedPages++;
                candidates.push(candidate);
                indexPdfCandidate(matchIndex, candidate);
            }
        }

        return {
            fileName: file.name,
            totalPages: pdf.numPages,
            parsedPages,
            candidates,
            matchIndex
        };
    }

    function buildDuplicateNameGroups(entries, getKey) {
        const groups = new Map();

        entries.forEach(entry => {
            const key = getKey(entry);
            if (!key) return;
            if (!groups.has(key)) groups.set(key, []);
            groups.get(key).push(entry);
        });

        const duplicates = new Map();
        groups.forEach((items, key) => {
            if (items.length > 1) duplicates.set(key, items);
        });

        return duplicates;
    }

    function buildExcelDuplicateMap() {
        return buildDuplicateNameGroups(
            state.sheetData.map(row => ({
                stt: row._originalIndex,
                name: getStudentNameValue(row),
                cccd: getStudentCccdValue(row),
                key: createCccdMatchKey(getStudentCccdValue(row))
            })),
            entry => entry.key
        );
    }

    function buildPdfDuplicateMap(candidates) {
        return buildDuplicateNameGroups(candidates, candidate => candidate.matchKey);
    }

    function getExcelDuplicateWarning(cccd, excelDuplicateMap) {
        const entries = excelDuplicateMap.get(createCccdMatchKey(cccd));
        if (!entries || entries.length <= 1) return '';

        const stts = entries.map(e => e.stt).join(', ');
        return `Trùng CCCD Excel (STT ${stts})`;
    }

    function getPdfDuplicateWarning(matchKey, pdfDuplicateMap) {
        const entries = pdfDuplicateMap.get(matchKey);
        if (!entries || entries.length <= 1) return '';

        const pages = [...new Set(entries.map(e => e.pageNumber))].sort((a, b) => a - b).join(', ');
        return `Trùng CCCD PDF (tr. ${pages})`;
    }

    function isDuplicateMatchKey(matchKey, excelDuplicateMap, pdfDuplicateMap) {
        if (!matchKey) return false;
        return excelDuplicateMap.has(matchKey) || pdfDuplicateMap.has(matchKey);
    }

    function buildWarningCell(...parts) {
        const text = parts.filter(Boolean).join(' • ');
        if (!text) return '—';
        return `<span class="qr-warn-badge" title="${escapeHTML(text)}">${escapeHTML(text)}</span>`;
    }

    function matchPdfWithExcel(pdfResult) {
        const found = [];
        const missing = [];
        const matchedCandidates = new Set();
        const excelDuplicateMap = buildExcelDuplicateMap();
        const pdfDuplicateMap = buildPdfDuplicateMap(pdfResult.candidates);

        state.sheetData.forEach(row => {
            const name = getStudentNameValue(row);
            const cccd = getStudentCccdValue(row);
            const matchKey = createCccdMatchKey(cccd);
            const pdfMatches = findPdfMatches(pdfResult, matchKey);

            if (pdfMatches && pdfMatches.length > 0) {
                if (isDuplicateMatchKey(matchKey, excelDuplicateMap, pdfDuplicateMap)) {
                    return;
                }

                pdfMatches.forEach(c => {
                    matchedCandidates.add(`${c.pageNumber}:${c.matchKey}`);
                });
                found.push({
                    stt: row._originalIndex,
                    name,
                    cccd,
                    contact: getContactValue(row),
                    pages: [...new Set(pdfMatches.map(c => c.pageNumber))].sort((a, b) => a - b)
                });
            } else {
                missing.push({
                    stt: row._originalIndex,
                    name,
                    cccd
                });
            }
        });

        const unmatchedPdf = pdfResult.candidates.filter(
            c => !matchedCandidates.has(`${c.pageNumber}:${c.matchKey}`)
        );

        return { found, missing, unmatchedPdf, excelDuplicateMap, pdfDuplicateMap };
    }

    function sanitizePathSegment(name) {
        return String(name)
            .replace(/[\\/:*?"<>|]/g, '_')
            .replace(/\s+/g, ' ')
            .trim() || 'Khong-ten';
    }

    function buildMatchedPdfFileName(item) {
        const namePart = sanitizePathSegment(item.name || 'Khong-ten');
        return `${namePart} ${item.stt}.pdf`;
    }

    function clonePdfBytes(buffer) {
        if (!buffer) return null;
        if (buffer instanceof Uint8Array) {
            return new Uint8Array(buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength));
        }
        return new Uint8Array(buffer.slice(0));
    }

    function toPdfLibBytes(buffer) {
        const bytes = clonePdfBytes(buffer);
        if (!bytes || bytes.byteLength === 0) {
            throw new Error('Dữ liệu PDF rỗng hoặc đã bị hủy. Chạy lại đối chiếu QR.');
        }
        return bytes;
    }

    function triggerFileDownload(blob, filename) {
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        link.remove();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
    }

    function escapeCsvCell(value) {
        const text = String(value ?? '');
        if (/[",\n\r]/.test(text)) {
            return `"${text.replace(/"/g, '""')}"`;
        }
        return text;
    }

    function buildCsvContent(headers, rows) {
        const lines = [headers.map(escapeCsvCell).join(',')];
        rows.forEach(row => lines.push(row.map(escapeCsvCell).join(',')));
        return '\uFEFF' + lines.join('\r\n');
    }

    function findRowByStt(stt) {
        return state.sheetData.find(row => row._originalIndex === stt) || null;
    }

    function buildSummaryCard({ count, label, color = '', downloadType = null }) {
        const downloadBtn = downloadType
            ? `<button type="button" class="qr-summary-download btn btn-outline btn-sm" data-qr-download="${downloadType}" title="Tải xuống" aria-label="Tải xuống">
                <i data-lucide="download"></i>
               </button>`
            : '';

        return `
            <div class="qr-summary-item">
                <strong${color ? ` style="color: ${color}"` : ''}>${count}</strong>
                <span>${escapeHTML(label)}</span>
                ${downloadBtn}
            </div>
        `;
    }

    function buildMatchedContactGroups(found) {
        const groups = new Map();

        found.forEach(item => {
            const key = item.contact || EMPTY_CONTACT_LABEL;
            if (!groups.has(key)) groups.set(key, []);
            groups.get(key).push(item);
        });

        return [...groups.entries()]
            .sort((a, b) => sortContactNames(a[0], b[0]))
            .map(([contact, items]) => ({ contact, count: items.length, items }));
    }

    function buildMatchedSummaryCard(count, found) {
        if (count === 0) {
            return buildSummaryCard({ count, label: 'Đã khớp', color: 'var(--success)' });
        }

        const contactGroups = buildMatchedContactGroups(found);
        const options = [
            `<option value="all">Tất cả (${count})</option>`,
            ...contactGroups.map((group, index) =>
                `<option value="${index}">${escapeHTML(group.contact)} (${group.count})</option>`
            )
        ].join('');

        return `
            <div class="qr-summary-item qr-summary-item--matched">
                <strong style="color: var(--success)">${count}</strong>
                <span>Đã khớp</span>
                <div class="qr-matched-download">
                    <select class="qr-matched-select" id="qrMatchedContactSelect" aria-label="Chọn nhóm Liên hệ">
                        ${options}
                    </select>
                    <button type="button" class="qr-summary-download btn btn-outline btn-sm" data-qr-download="matched-zip" title="Tải xuống" aria-label="Tải xuống">
                        <i data-lucide="download"></i>
                    </button>
                </div>
            </div>
        `;
    }

    async function extractSinglePdfPage(sourcePdf, pageNumber) {
        const newPdf = await PDFLib.PDFDocument.create();
        const [copiedPage] = await newPdf.copyPages(sourcePdf, [pageNumber - 1]);
        newPdf.addPage(copiedPage);
        return new Uint8Array(await newPdf.save());
    }

    async function downloadMatchedPdfZip(contactFilter = null) {
        if (!qrExportState?.pdfBytes || !qrExportState?.matchResult) {
            showToast('Không có dữ liệu PDF để tải. Chạy lại đối chiếu QR.', 'error');
            return;
        }

        if (typeof PDFLib === 'undefined' || typeof JSZip === 'undefined') {
            showToast('Thiếu thư viện PDF/ZIP. Tải lại trang (Ctrl+F5).', 'error');
            return;
        }

        let found = qrExportState.matchResult.found;
        if (found.length === 0) {
            showToast('Không có học viên đã khớp để tải.', 'warning');
            return;
        }

        const downloadAll = contactFilter == null;
        if (!downloadAll) {
            found = found.filter(item => (item.contact || EMPTY_CONTACT_LABEL) === contactFilter);
            if (found.length === 0) {
                showToast('Nhóm Liên hệ đã chọn không có học viên khớp.', 'warning');
                return;
            }
        }

        const groupCounts = new Map();
        found.forEach(item => {
            const key = item.contact || EMPTY_CONTACT_LABEL;
            groupCounts.set(key, (groupCounts.get(key) || 0) + 1);
        });

        const loadingLabel = downloadAll
            ? 'Đang tách PDF và đóng gói ZIP (tất cả)...'
            : `Đang tách PDF nhóm "${contactFilter}"...`;
        showLoading(loadingLabel);

        try {
            const sourcePdf = await PDFLib.PDFDocument.load(toPdfLibBytes(qrExportState.pdfBytes), {
                ignoreEncryption: true
            });
            const zip = new JSZip();
            const pageCache = new Map();
            let processed = 0;
            let added = 0;
            let failed = 0;

            for (const item of found) {
                processed++;
                if (processed % 20 === 0 || processed === found.length) {
                    loadingText.textContent = `Đang tách PDF: ${processed}/${found.length}...`;
                    await new Promise(resolve => setTimeout(resolve, 0));
                }

                const pageNumber = item.pages[0];
                if (!pageNumber) continue;

                try {
                    let pageBytes = pageCache.get(pageNumber);
                    if (!pageBytes) {
                        pageBytes = await extractSinglePdfPage(sourcePdf, pageNumber);
                        pageCache.set(pageNumber, pageBytes);
                    }

                    const fileName = buildMatchedPdfFileName(item);
                    if (downloadAll) {
                        const contact = item.contact || EMPTY_CONTACT_LABEL;
                        const folderName = `${sanitizePathSegment(contact)} ${groupCounts.get(contact)}`;
                        zip.folder(folderName).file(fileName, pageBytes);
                    } else {
                        zip.file(fileName, pageBytes);
                    }
                    added++;
                } catch (pageErr) {
                    failed++;
                    console.error(`Tách trang ${pageNumber} (STT ${item.stt}) thất bại:`, pageErr);
                }
            }

            if (added === 0) {
                throw new Error('Không tách được trang PDF nào. Chạy lại đối chiếu QR.');
            }

            loadingText.textContent = 'Đang nén file ZIP...';
            const zipBlob = await zip.generateAsync({
                type: 'blob',
                compression: 'DEFLATE',
                compressionOptions: { level: 1 }
            });
            const pdfBaseName = sanitizePathSegment(
                (qrExportState.pdfFileName || 'qr-da-khop').replace(/\.pdf$/i, '')
            );
            const zipName = downloadAll
                ? `${pdfBaseName}-da-khop.zip`
                : `${pdfBaseName}-${sanitizePathSegment(contactFilter)}-${found.length}-da-khop.zip`;
            triggerFileDownload(zipBlob, zipName);

            const failMsg = failed > 0 ? ` (${failed} trang lỗi)` : '';
            const successMsg = downloadAll
                ? `Đã tải ${added} file PDF (tất cả nhóm Liên hệ)${failMsg}.`
                : `Đã tải ${added} file PDF nhóm "${contactFilter}"${failMsg}.`;
            showToast(successMsg, failed > 0 ? 'warning' : 'success');
        } catch (err) {
            console.error(err);
            showToast('Lỗi tạo ZIP: ' + err.message, 'error');
        } finally {
            hideLoading();
        }
    }

    function downloadExcelDuplicatesCsv() {
        if (!qrExportState?.matchResult) {
            showToast('Không có dữ liệu. Chạy lại đối chiếu QR.', 'error');
            return;
        }

        const { excelDuplicateMap } = qrExportState.matchResult;
        if (excelDuplicateMap.size === 0) {
            showToast('Không có CCCD trùng trong Excel.', 'warning');
            return;
        }

        const headers = ['STT', 'Họ và tên', 'CCCD', 'Liên hệ', 'Số người trùng', 'Danh sách STT trùng'];
        const rows = [];

        excelDuplicateMap.forEach(entries => {
            const stts = entries.map(e => e.stt).sort((a, b) => a - b);
            const sttList = stts.join(', ');

            entries.forEach(entry => {
                const row = findRowByStt(entry.stt);
                rows.push([
                    entry.stt,
                    entry.name,
                    entry.cccd,
                    row ? getContactValue(row) : '',
                    entries.length,
                    sttList
                ]);
            });
        });

        rows.sort((a, b) => Number(a[0]) - Number(b[0]));
        const csv = buildCsvContent(headers, rows);
        const baseName = (state.fileName || 'danh-sach').replace(/\.[^.]+$/, '');
        triggerFileDownload(new Blob([csv], { type: 'text/csv;charset=utf-8' }), `${baseName}-cccd-trung-excel.csv`);
        showToast(`Đã tải ${rows.length} dòng CCCD trùng Excel.`, 'success');
    }

    async function downloadPdfDuplicatesZip() {
        if (!qrExportState?.pdfBytes || !qrExportState?.matchResult) {
            showToast('Không có dữ liệu PDF. Chạy lại đối chiếu QR.', 'error');
            return;
        }

        if (typeof PDFLib === 'undefined' || typeof JSZip === 'undefined') {
            showToast('Thiếu thư viện PDF/ZIP. Tải lại trang (Ctrl+F5).', 'error');
            return;
        }

        const { pdfDuplicateMap } = qrExportState.matchResult;
        if (pdfDuplicateMap.size === 0) {
            showToast('Không có CCCD trùng trong PDF.', 'warning');
            return;
        }

        showLoading('Đang tách các trang PDF trùng CCCD...');

        try {
            const sourcePdf = await PDFLib.PDFDocument.load(toPdfLibBytes(qrExportState.pdfBytes), {
                ignoreEncryption: true
            });
            const zip = new JSZip();
            const pageCache = new Map();
            let added = 0;
            let failed = 0;
            let groupIndex = 0;
            const groupTotal = pdfDuplicateMap.size;

            for (const entries of pdfDuplicateMap.values()) {
                groupIndex++;
                const displayName = entries[0]?.displayName || 'Khong-ten';
                const cccd = entries[0]?.cccd || 'Khong-cccd';
                const folderName = `${sanitizePathSegment(cccd)} ${entries.length}`;
                const pageNumbers = [...new Set(entries.map(e => e.pageNumber))].sort((a, b) => a - b);

                loadingText.textContent = `Đang tách PDF trùng CCCD: ${groupIndex}/${groupTotal}...`;
                await new Promise(resolve => setTimeout(resolve, 0));

                for (const pageNumber of pageNumbers) {
                    try {
                        let pageBytes = pageCache.get(pageNumber);
                        if (!pageBytes) {
                            pageBytes = await extractSinglePdfPage(sourcePdf, pageNumber);
                            pageCache.set(pageNumber, pageBytes);
                        }

                        const fileName = `${sanitizePathSegment(displayName)} ${sanitizePathSegment(cccd)} ${pageNumber}.pdf`;
                        zip.folder(folderName).file(fileName, pageBytes);
                        added++;
                    } catch (pageErr) {
                        failed++;
                        console.error(`Tách trang ${pageNumber} (${displayName}) thất bại:`, pageErr);
                    }
                }
            }

            if (added === 0) {
                throw new Error('Không tách được trang PDF trùng CCCD nào.');
            }

            loadingText.textContent = 'Đang nén file ZIP...';
            const zipBlob = await zip.generateAsync({
                type: 'blob',
                compression: 'DEFLATE',
                compressionOptions: { level: 1 }
            });
            const baseName = sanitizePathSegment(
                (qrExportState.pdfFileName || 'qr-trung-cccd').replace(/\.pdf$/i, '')
            );
            triggerFileDownload(zipBlob, `${baseName}-cccd-trung-pdf.zip`);

            const failMsg = failed > 0 ? ` (${failed} trang lỗi)` : '';
            showToast(`Đã tải ${added} trang PDF trùng CCCD${failMsg}.`, failed > 0 ? 'warning' : 'success');
        } catch (err) {
            console.error(err);
            showToast('Lỗi tải PDF trùng CCCD: ' + err.message, 'error');
        } finally {
            hideLoading();
        }
    }

    async function downloadUnmatchedPdfZip() {
        if (!qrExportState?.pdfBytes || !qrExportState?.matchResult) {
            showToast('Không có dữ liệu PDF. Chạy lại đối chiếu QR.', 'error');
            return;
        }

        if (typeof PDFLib === 'undefined' || typeof JSZip === 'undefined') {
            showToast('Thiếu thư viện PDF/ZIP. Tải lại trang (Ctrl+F5).', 'error');
            return;
        }

        const unmatched = qrExportState.matchResult.unmatchedPdf;
        if (unmatched.length === 0) {
            showToast('Không có trang PDF ngoài Excel.', 'warning');
            return;
        }

        showLoading('Đang tách các trang không có trong Excel...');

        try {
            const sourcePdf = await PDFLib.PDFDocument.load(toPdfLibBytes(qrExportState.pdfBytes), {
                ignoreEncryption: true
            });
            const zip = new JSZip();
            const pageCache = new Map();
            let processed = 0;
            let added = 0;
            let failed = 0;

            for (const item of unmatched) {
                processed++;
                if (processed % 10 === 0 || processed === unmatched.length) {
                    loadingText.textContent = `Đang tách PDF: ${processed}/${unmatched.length}...`;
                    await new Promise(resolve => setTimeout(resolve, 0));
                }

                const pageNumber = item.pageNumber;
                if (!pageNumber) continue;

                try {
                    let pageBytes = pageCache.get(pageNumber);
                    if (!pageBytes) {
                        pageBytes = await extractSinglePdfPage(sourcePdf, pageNumber);
                        pageCache.set(pageNumber, pageBytes);
                    }

                    const fileName = `${sanitizePathSegment(item.displayName || 'Khong-ten')} ${sanitizePathSegment(item.cccd || 'Khong-cccd')} ${pageNumber}.pdf`;
                    zip.file(fileName, pageBytes);
                    added++;
                } catch (pageErr) {
                    failed++;
                    console.error(`Tách trang ${pageNumber} (${item.displayName}) thất bại:`, pageErr);
                }
            }

            if (added === 0) {
                throw new Error('Không tách được trang PDF nào.');
            }

            loadingText.textContent = 'Đang nén file ZIP...';
            const zipBlob = await zip.generateAsync({
                type: 'blob',
                compression: 'DEFLATE',
                compressionOptions: { level: 1 }
            });
            const baseName = sanitizePathSegment(
                (qrExportState.pdfFileName || 'qr-khong-co-excel').replace(/\.pdf$/i, '')
            );
            triggerFileDownload(zipBlob, `${baseName}-khong-co-trong-excel.zip`);

            const failMsg = failed > 0 ? ` (${failed} trang lỗi)` : '';
            showToast(`Đã tải ${added} trang PDF không có trong Excel${failMsg}.`, failed > 0 ? 'warning' : 'success');
        } catch (err) {
            console.error(err);
            showToast('Lỗi tải PDF: ' + err.message, 'error');
        } finally {
            hideLoading();
        }
    }

    function handleQrSummaryDownload(downloadType) {
        if (downloadType === 'matched-zip') {
            const select = document.getElementById('qrMatchedContactSelect');
            const selected = select?.value ?? 'all';

            if (selected === 'all') {
                downloadMatchedPdfZip();
                return;
            }

            const groups = buildMatchedContactGroups(qrExportState.matchResult.found);
            const group = groups[Number(selected)];
            if (!group) {
                showToast('Không tìm thấy nhóm Liên hệ.', 'error');
                return;
            }
            downloadMatchedPdfZip(group.contact);
        } else if (downloadType === 'excel-dup-csv') {
            downloadExcelDuplicatesCsv();
        } else if (downloadType === 'pdf-dup-zip') {
            downloadPdfDuplicatesZip();
        } else if (downloadType === 'unmatched-excel-zip') {
            downloadUnmatchedPdfZip();
        }
    }

    function renderQrStatsTableBody(tbody, rows, columns) {
        tbody.innerHTML = '';

        if (rows.length === 0) {
            const tr = document.createElement('tr');
            tr.className = 'empty-row';
            const td = document.createElement('td');
            td.colSpan = columns;
            td.textContent = 'Không có dữ liệu';
            tr.appendChild(td);
            tbody.appendChild(tr);
            return;
        }

        rows.forEach(row => {
            const tr = document.createElement('tr');
            tr.innerHTML = row.map(cell => {
                if (cell && typeof cell === 'object' && cell.html != null) {
                    return `<td>${cell.html}</td>`;
                }
                return `<td>${escapeHTML(String(cell ?? ''))}</td>`;
            }).join('');
            tbody.appendChild(tr);
        });
    }

    function showQrStatsModal(pdfResult, matchResult) {
        const cccdColHeader = state.headers[getCccdColumnIndex()] || 'CCCD';
        const totalExcel = state.sheetData.length;
        const totalFound = matchResult.found.length;
        const totalMissingPdf = matchResult.missing.length;
        const totalMissingExcel = matchResult.unmatchedPdf.length;
        const excelDupCount = matchResult.excelDuplicateMap.size;
        const pdfDupCount = matchResult.pdfDuplicateMap.size;
        const hasMismatch = totalMissingPdf > 0 || totalMissingExcel > 0;
        const hasDuplicates = excelDupCount > 0 || pdfDupCount > 0;

        const pdfTooSmall = pdfResult.totalPages < totalExcel;

        qrStatsSubtitle.textContent =
            `File PDF: ${pdfResult.fileName} • ${pdfResult.parsedPages}/${pdfResult.totalPages} trang có CCCD • ` +
            `Cột đối chiếu: ${cccdColHeader}` +
            (hasMismatch ? ' • Có chênh lệch cần kiểm tra' : ' • Khớp hoàn toàn') +
            (hasDuplicates ? ` • ${excelDupCount} CCCD trùng Excel, ${pdfDupCount} CCCD trùng PDF` : '');

        if (pdfTooSmall) {
            qrStatsAlert.className = 'qr-stats-alert qr-stats-alert--error';
            qrStatsAlert.innerHTML = `
                <i data-lucide="alert-triangle"></i>
                <div>
                    <strong>PDF quá ít trang so với Excel</strong>
                    <p>PDF có <b>${pdfResult.totalPages}</b> trang nhưng Excel có <b>${totalExcel}</b> học viên.
                    Bạn có thể đang chọn <b>file mẫu 1 trang</b> (tên có "34" hoặc file nhỏ ~15KB).
                    Hãy chọn file PDF <b>đầy đủ ~${totalExcel} trang</b> (khoảng 5–6 MB), ví dụ:
                    <b>MA QR THANH TOAN LE PHI SH 23.06.2026 (1).pdf</b></p>
                </div>
            `;
            qrStatsAlert.classList.remove('hidden');
        } else if (pdfResult.parsedPages < pdfResult.totalPages) {
            qrStatsAlert.className = 'qr-stats-alert qr-stats-alert--warning';
            qrStatsAlert.innerHTML = `
                <i data-lucide="alert-circle"></i>
                <div>
                    <strong>Một số trang PDF không đọc được CCCD</strong>
                    <p>Đọc được ${pdfResult.parsedPages}/${pdfResult.totalPages} trang. Kiểm tra lại file PDF.</p>
                </div>
            `;
            qrStatsAlert.classList.remove('hidden');
        } else {
            qrStatsAlert.classList.add('hidden');
            qrStatsAlert.innerHTML = '';
        }

        qrStatsSummary.innerHTML = [
            buildSummaryCard({ count: totalExcel, label: 'Học viên trong Excel' }),
            buildSummaryCard({ count: pdfResult.candidates.length, label: 'Học viên trong PDF' }),
            buildMatchedSummaryCard(totalFound, matchResult.found),
            buildSummaryCard({ count: totalMissingPdf, label: 'Không có trong PDF', color: 'var(--danger)' }),
            buildSummaryCard({ count: totalMissingExcel, label: 'Không có trong Excel', color: 'var(--warning)', downloadType: 'unmatched-excel-zip' }),
            buildSummaryCard({ count: excelDupCount, label: 'CCCD trùng trong Excel', color: 'var(--warning)', downloadType: 'excel-dup-csv' }),
            buildSummaryCard({ count: pdfDupCount, label: 'CCCD trùng trong PDF', color: 'var(--warning)', downloadType: 'pdf-dup-zip' })
        ].join('');

        qrFoundCount.textContent = String(totalFound);
        qrMissingPdfCount.textContent = String(totalMissingPdf);
        qrMissingExcelCount.textContent = String(totalMissingExcel);

        const { excelDuplicateMap, pdfDuplicateMap } = matchResult;

        renderQrStatsTableBody(
            qrFoundTable.querySelector('tbody'),
            matchResult.found.map(item => {
                const matchKey = createCccdMatchKey(item.cccd);
                return [
                    item.stt,
                    item.name,
                    item.cccd,
                    item.pages.join(', '),
                    {
                        html: buildWarningCell(
                            getExcelDuplicateWarning(item.cccd, excelDuplicateMap),
                            getPdfDuplicateWarning(matchKey, pdfDuplicateMap)
                        )
                    }
                ];
            }),
            5
        );

        renderQrStatsTableBody(
            qrMissingPdfTable.querySelector('tbody'),
            matchResult.missing.map(item => [
                item.stt,
                item.name,
                item.cccd,
                {
                    html: buildWarningCell(
                        getExcelDuplicateWarning(item.cccd, excelDuplicateMap)
                    )
                }
            ]),
            4
        );

        renderQrStatsTableBody(
            qrMissingExcelTable.querySelector('tbody'),
            matchResult.unmatchedPdf.map(item => [
                item.pageNumber,
                item.displayName,
                item.cccd,
                {
                    html: buildWarningCell(
                        getPdfDuplicateWarning(item.matchKey, pdfDuplicateMap)
                    )
                }
            ]),
            4
        );

        qrStatsModal.classList.remove('hidden');
        initLucideIcons();
    }

    function closeQrStatsModal() {
        qrStatsModal.classList.add('hidden');
    }

    async function handlePdfFileSelected(e) {
        const file = e.target.files?.[0];
        pdfInput.value = '';

        if (!file) return;

        if (!file.name.toLowerCase().endsWith('.pdf')) {
            showToast('Vui lòng chọn file PDF.', 'error');
            return;
        }

        showLoading('Đang đọc và đối chiếu file PDF...');

        try {
            const rawBuffer = await file.arrayBuffer();
            const pdfResult = await parsePdfCandidates(file, rawBuffer.slice(0));
            const matchResult = matchPdfWithExcel(pdfResult);

            qrExportState = {
                pdfBytes: clonePdfBytes(rawBuffer),
                pdfFileName: file.name,
                pdfResult,
                matchResult
            };

            showQrStatsModal(pdfResult, matchResult);

            if (pdfResult.totalPages < state.sheetData.length) {
                showToast(
                    `PDF chỉ có ${pdfResult.totalPages} trang, Excel có ${state.sheetData.length} học viên. ` +
                    'Hãy chọn file PDF đầy đủ (~5 MB, ~376 trang), không phải file mẫu 1 trang.',
                    'error'
                );
            } else if (pdfResult.candidates.length === 0) {
                showToast('Không tìm thấy CCCD trong PDF. Tất cả học viên Excel thiếu QR.', 'warning');
            } else if (
                matchResult.missing.length > 0 ||
                matchResult.unmatchedPdf.length > 0 ||
                matchResult.excelDuplicateMap.size > 0 ||
                matchResult.pdfDuplicateMap.size > 0
            ) {
                const dupMsg = (matchResult.excelDuplicateMap.size || matchResult.pdfDuplicateMap.size)
                    ? ` Trùng CCCD: ${matchResult.excelDuplicateMap.size} Excel, ${matchResult.pdfDuplicateMap.size} PDF.`
                    : '';
                showToast(
                    `Chênh lệch: ${matchResult.missing.length} không có trong PDF, ` +
                    `${matchResult.unmatchedPdf.length} không có trong Excel.${dupMsg}`,
                    'warning'
                );
            } else {
                showToast(`Đối chiếu hoàn tất: ${matchResult.found.length} học viên khớp.`, 'success');
            }
        } catch (err) {
            console.error(err);
            showToast('Lỗi đọc PDF: ' + err.message, 'error');
        } finally {
            hideLoading();
        }
    }

    function getContactValue(row) {
        const val = row[getContactColumnIndex()];
        const text = formatCellValue(val).trim();
        return text || EMPTY_CONTACT_LABEL;
    }

    function sortContactNames(a, b) {
        if (a === EMPTY_CONTACT_LABEL) return 1;
        if (b === EMPTY_CONTACT_LABEL) return -1;
        return a.localeCompare(b, 'vi', { sensitivity: 'base', numeric: true });
    }

    function buildContactGroups(rows) {
        const groups = new Map();

        rows.forEach(row => {
            const key = getContactValue(row);
            if (!groups.has(key)) groups.set(key, []);
            groups.get(key).push(row);
        });

        return Array.from(groups.entries()).sort((a, b) =>
            sortContactNames(a[0], b[0])
        );
    }

    function getContactCounts(rows) {
        const counts = new Map();
        rows.forEach(row => {
            const key = getContactValue(row);
            counts.set(key, (counts.get(key) || 0) + 1);
        });
        return Array.from(counts.entries()).sort((a, b) =>
            sortContactNames(a[0], b[0])
        );
    }

    function getDisplayData() {
        if (state.groupFilter === 'all') return state.filteredData;
        return state.filteredData.filter(row => getContactValue(row) === state.groupFilter);
    }

    function populateGroupFilter() {
        const contactCounts = getContactCounts(state.sheetData);
        const contactNames = contactCounts.map(([name]) => name);

        groupFilterSelect.innerHTML = '<option value="all">— Tất cả nhóm —</option>';
        contactCounts.forEach(([contact, count]) => {
            const option = document.createElement('option');
            option.value = contact;
            option.textContent = `${contact} (${count})`;
            groupFilterSelect.appendChild(option);
        });

        if (state.groupFilter !== 'all' && !contactNames.includes(state.groupFilter)) {
            state.groupFilter = 'all';
        }
        groupFilterSelect.value = state.groupFilter;
    }

    function updateViewModeUI() {
        const hasData = state.headers.length > 0;
        const isGroup = state.viewMode === 'group';

        groupFilterWrap.classList.toggle('hidden', !hasData);
        pageSizeWrap.classList.toggle('hidden', isGroup);
        if (isGroup) {
            paginationContainer.classList.add('hidden');
        }
    }

    function renderTableHeader() {
        const headerTr = document.createElement('tr');
        const indexTh = document.createElement('th');
        indexTh.textContent = '#';
        headerTr.appendChild(indexTh);

        state.headers.forEach((header, index) => {
            const th = document.createElement('th');
            th.className = 'sortable';

            const colLetter = getColumnLetter(index + 1);
            let arrow = '';
            if (state.sortColumnIndex === index) {
                arrow = state.sortDirection === 'asc' ? ' ↑' : ' ↓';
                th.style.color = 'var(--primary)';
            }

            th.innerHTML = `
                <div style="font-size: 10px; color: var(--text-muted); font-weight: 500; font-family: var(--font-mono);">${colLetter}</div>
                <div>${escapeHTML(header)}${arrow}</div>
            `;

            th.addEventListener('click', () => sortColumn(index));
            headerTr.appendChild(th);
        });

        tableHead.appendChild(headerTr);
    }

    function updateCreateQrButton() {
        createQrBtn.classList.toggle('hidden', state.sheetData.length === 0);
    }

    function findFirstEmptyCell(rows) {
        for (let r = 0; r < rows.length; r++) {
            const row = rows[r];
            for (let c = 0; c < state.headers.length; c++) {
                if (isCellEmpty(row[c])) {
                    return {
                        originalIndex: row._originalIndex,
                        colIndex: c,
                        header: state.headers[c]
                    };
                }
            }
        }
        return null;
    }

    function ensureRowVisible(originalIndex) {
        const row = state.sheetData.find(r => r._originalIndex === originalIndex);
        if (!row) return false;

        let changed = false;

        if (!state.filteredData.some(r => r._originalIndex === originalIndex)) {
            searchInput.value = '';
            clearSearchBtn.classList.add('hidden');
            state.filteredData = [...state.sheetData];
            changed = true;
        }

        const contact = getContactValue(row);
        if (state.groupFilter !== 'all' && contact !== state.groupFilter) {
            state.groupFilter = 'all';
            groupFilterSelect.value = 'all';
            changed = true;
        }

        if (state.viewMode === 'group' && isGroupCollapsed(contact)) {
            state.expandedGroups.add(contact);
            changed = true;
        }

        return changed;
    }

    function focusEmptyCellElement(originalIndex, colIndex) {
        const rowEl = tableBody.querySelector(`tr[data-original-index="${originalIndex}"]`);
        const cellEl = rowEl?.querySelector(`td[data-col-index="${colIndex}"]`);

        if (!cellEl) return false;

        document.querySelectorAll('.cell-focus-flash').forEach(el => {
            el.classList.remove('cell-focus-flash');
        });

        cellEl.classList.add('cell-focus-flash');
        cellEl.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'nearest' });

        setTimeout(() => cellEl.classList.remove('cell-focus-flash'), 3000);
        return true;
    }

    function scrollToEmptyCell(emptyCell) {
        const { originalIndex, colIndex } = emptyCell;
        const viewChanged = ensureRowVisible(originalIndex);

        const displayData = getDisplayData();
        const dataIdx = displayData.findIndex(r => r._originalIndex === originalIndex);
        if (dataIdx === -1) return;

        let needsRender = viewChanged;

        if (state.viewMode === 'list' && !isBulkPageSize()) {
            const targetPage = Math.floor(dataIdx / parseInt(state.pageSize)) + 1;
            if (state.currentPage !== targetPage) {
                state.currentPage = targetPage;
                needsRender = true;
            }
        }

        if (needsRender) {
            renderTable();
            setTimeout(() => focusEmptyCellElement(originalIndex, colIndex), 80);
        } else {
            focusEmptyCellElement(originalIndex, colIndex);
        }
    }

    function handleCreateQrClick() {
        if (state.sheetData.length === 0) {
            showToast('Không có dữ liệu để tạo QR.', 'warning');
            return;
        }

        const emptyCell = findFirstEmptyCell(state.sheetData);
        if (emptyCell) {
            const colLetter = getColumnLetter(emptyCell.colIndex + 1);
            showToast(
                `Cảnh báo: Dòng ${emptyCell.originalIndex}, cột ${colLetter} (${emptyCell.header}) đang trống.`,
                'error'
            );
            scrollToEmptyCell(emptyCell);
            return;
        }

        pdfInput.click();
    }

    function encodeGroupKey(contact) {
        return encodeURIComponent(contact);
    }

    function isGroupCollapsed(contact) {
        return !state.expandedGroups.has(contact);
    }

    function toggleGroupCollapse(contact) {
        if (state.expandedGroups.has(contact)) {
            state.expandedGroups.delete(contact);
        } else {
            state.expandedGroups.add(contact);
        }

        const key = encodeGroupKey(contact);
        const collapsed = isGroupCollapsed(contact);

        tableBody.querySelectorAll(`tr.group-data-row[data-group-key="${key}"]`).forEach(row => {
            row.classList.toggle('hidden', collapsed);
        });

        const headerRow = tableBody.querySelector(`tr.group-header-row[data-group-key="${key}"]`);
        if (headerRow) {
            headerRow.classList.toggle('is-collapsed', collapsed);
            const chevron = headerRow.querySelector('.group-chevron');
            if (chevron) {
                chevron.setAttribute('data-lucide', collapsed ? 'chevron-right' : 'chevron-down');
                initLucideIcons();
            }
        }
    }

    function rowHasEmptyCells(row) {
        for (let c = 0; c < state.headers.length; c++) {
            if (isCellEmpty(row[c])) return true;
        }
        return false;
    }

    function renderDataRow(row, rowIndex, query, fragment, groupKey) {
        const tr = document.createElement('tr');
        const hasEmpty = rowHasEmptyCells(row);

        if (groupKey) tr.classList.add('group-data-row');
        if (hasEmpty) {
            tr.classList.add('row-warning');
            tr.title = 'Cảnh báo: dòng có ô trống';
        }

        if (groupKey) {
            tr.dataset.groupKey = groupKey;
            if (isGroupCollapsed(decodeURIComponent(groupKey))) {
                tr.classList.add('hidden');
            }
        }

        tr.dataset.originalIndex = row._originalIndex || rowIndex;

        const rowNumTd = document.createElement('td');
        rowNumTd.textContent = row._originalIndex || rowIndex;
        tr.appendChild(rowNumTd);

        for (let c = 0; c < state.headers.length; c++) {
            const td = document.createElement('td');
            const cellValue = row[c];
            const displayVal = c === getCccdColumnIndex()
                ? getStudentCccdValue(row)
                : formatCellValue(cellValue);
            const empty = c === getCccdColumnIndex()
                ? !displayVal
                : isCellEmpty(cellValue);

            td.dataset.colIndex = c;

            if (isNumeric(cellValue) || c === getCccdColumnIndex()) td.classList.add('numeric');
            if (empty) td.classList.add('cell-empty');
            if (c === getContactColumnIndex() && state.viewMode === 'group') {
                td.classList.add('contact-cell');
            }

            if (query) {
                td.innerHTML = highlightMatch(displayVal, query);
            } else {
                td.textContent = displayVal;
            }
            tr.appendChild(td);
        }

        fragment.appendChild(tr);
    }

    function renderGroupHeaderRow(contact, count, fragment) {
        const tr = document.createElement('tr');
        const groupKey = encodeGroupKey(contact);
        const collapsed = isGroupCollapsed(contact);

        tr.className = 'group-header-row' + (collapsed ? ' is-collapsed' : '');
        tr.dataset.groupKey = groupKey;

        const td = document.createElement('td');
        td.colSpan = state.headers.length + 1;
        td.innerHTML = `
            <button type="button" class="group-header group-header-toggle" aria-expanded="${!collapsed}">
                <i data-lucide="${collapsed ? 'chevron-right' : 'chevron-down'}" class="group-chevron"></i>
                <i data-lucide="users" class="group-header-icon"></i>
                <span class="group-header-title">${escapeHTML(contact)}</span>
                <span class="group-header-count">${count} học viên</span>
            </button>
        `;
        tr.appendChild(td);
        fragment.appendChild(tr);
    }

    // --- Table Rendering (with Pagination & Search Highlights) ---
    function renderTable() {
        if (state.headers.length === 0) {
            renderEmptyGrid();
            return;
        }

        updateViewModeUI();

        tableHead.innerHTML = '';
        tableBody.innerHTML = '';
        renderTableHeader();

        if (state.viewMode === 'group') {
            renderGroupedTable();
        } else {
            renderListTable();
        }

        updateCreateQrButton();
        initLucideIcons();
    }

    function renderListTable() {
        const displayData = getDisplayData();

        // 2. Pagination Calculations (list mode)
        const totalRows = displayData.length;
        let startIdx = 0;
        let endIdx = totalRows;

        let cappedBulk = false;
        const maxCap = getMaxRowsCap();

        if (isBulkPageSize()) {
            endIdx = Math.min(totalRows, maxCap);
            cappedBulk = totalRows > maxCap;
        } else {
            const size = parseInt(state.pageSize);
            startIdx = (state.currentPage - 1) * size;
            endIdx = Math.min(startIdx + size, totalRows);
        }

        if (isBulkPageSize() || totalRows <= parseInt(state.pageSize)) {
            paginationContainer.classList.add('hidden');
        } else {
            paginationContainer.classList.remove('hidden');
            renderPaginationControls(totalRows);
        }

        if (totalRows === 0) {
            rowCounter.textContent = 'Hiển thị 0-0 / 0 dòng';
        } else if (cappedBulk) {
            rowCounter.textContent = `Hiển thị ${startIdx + 1}-${endIdx} / ${totalRows} dòng (giới hạn ${maxCap})`;
        } else {
            rowCounter.textContent = `Hiển thị ${startIdx + 1}-${endIdx} / ${totalRows} dòng`;
        }

        const query = searchInput.value.trim().toLowerCase();
        const fragment = document.createDocumentFragment();

        for (let i = startIdx; i < endIdx; i++) {
            renderDataRow(displayData[i], displayData[i]._originalIndex || (i + 2), query, fragment);
        }

        tableBody.appendChild(fragment);
        renderEmptyStateIfNeeded(totalRows);
    }

    function renderGroupedTable() {
        const groups = buildContactGroups(getDisplayData());

        const query = searchInput.value.trim().toLowerCase();
        const fragment = document.createDocumentFragment();
        let totalRows = 0;
        let renderedRows = 0;
        let groupCount = 0;

        groups.forEach(([contact, rows]) => {
            totalRows += rows.length;
        });

        const maxCap = getMaxRowsCap();

        for (const [contact, rows] of groups) {
            if (renderedRows >= maxCap) break;

            groupCount++;
            const rowsToShow = rows.slice(0, maxCap - renderedRows);
            const groupKey = encodeGroupKey(contact);
            renderGroupHeaderRow(contact, rows.length, fragment);

            rowsToShow.forEach(row => {
                renderDataRow(row, row._originalIndex, query, fragment, groupKey);
                renderedRows++;
            });
        }

        tableBody.appendChild(fragment);

        if (totalRows === 0) {
            rowCounter.textContent = 'Hiển thị 0-0 / 0 dòng';
            renderEmptyStateIfNeeded(0);
            return;
        }

        const capped = totalRows > maxCap;
        if (capped) {
            rowCounter.textContent = `${groupCount} nhóm • ${renderedRows}/${totalRows} học viên (giới hạn ${maxCap})`;
        } else {
            rowCounter.textContent = `${groupCount} nhóm • ${totalRows} học viên`;
        }
    }

    function renderEmptyStateIfNeeded(totalRows) {
        if (totalRows === 0) {
            const emptyTr = document.createElement('tr');
            const emptyTd = document.createElement('td');
            emptyTd.colSpan = state.headers.length + 1;
            emptyTd.style.textAlign = 'center';
            emptyTd.style.padding = '32px';
            emptyTd.style.color = 'var(--text-muted)';
            emptyTd.textContent = 'Không tìm thấy dòng phù hợp.';
            emptyTr.appendChild(emptyTd);
            tableBody.appendChild(emptyTr);
        }
    }

    // Helper: Excel Column letters (1 = A, 27 = AA)
    function getColumnLetter(colNum) {
        let letter = '';
        let temp = colNum;
        while (temp > 0) {
            let modulo = (temp - 1) % 26;
            letter = String.fromCharCode(65 + modulo) + letter;
            temp = Math.floor((temp - modulo) / 26);
        }
        return letter;
    }

    function isNumeric(val) {
        if (typeof val === 'number') return true;
        if (typeof val !== 'string') return false;
        return !isNaN(val) && !isNaN(parseFloat(val));
    }

    function formatCellValue(val) {
        if (val === null || val === undefined) return '';
        if (val instanceof Date) {
            // Short date format
            return val.toLocaleDateString();
        }
        if (typeof val === 'number') {
            // Format floats slightly to avoid long JS floating point decimals
            if (!Number.isInteger(val)) {
                return Number(val.toFixed(4)).toString(); // Max 4 decimal digits
            }
            return val.toString();
        }
        return String(val);
    }

    function escapeHTML(str) {
        return str
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;');
    }

    function highlightMatch(text, query) {
        if (!text) return '';
        const index = text.toLowerCase().indexOf(query);
        if (index === -1) return escapeHTML(text);
        
        // Highlight logic
        const before = text.substring(0, index);
        const match = text.substring(index, index + query.length);
        const after = text.substring(index + query.length);
        
        return escapeHTML(before) + `<mark class="highlight">${escapeHTML(match)}</mark>` + highlightMatch(after, query);
    }

    // --- Sorting ---
    function sortColumn(index) {
        showLoading('Đang sắp xếp...');
        setTimeout(() => {
            if (state.sortColumnIndex === index) {
                state.sortDirection = state.sortDirection === 'asc' ? 'desc' : 'asc';
            } else {
                state.sortColumnIndex = index;
                state.sortDirection = 'asc';
            }

            state.filteredData.sort((a, b) => {
                let valA = a[index];
                let valB = b[index];

                // Standardize comparison values
                let numA = parseFloat(valA);
                let numB = parseFloat(valB);

                let isNumA = !isNaN(numA) && isFinite(valA) && valA !== '';
                let isNumB = !isNaN(numB) && isFinite(valB) && valB !== '';

                if (isNumA && isNumB) {
                    return state.sortDirection === 'asc' ? numA - numB : numB - numA;
                }

                // Treat empty values as low priority
                if (valA === '' || valA === null) return 1;
                if (valB === '' || valB === null) return -1;

                // String comparison
                let strA = String(valA).toLowerCase();
                let strB = String(valB).toLowerCase();

                if (strA < strB) return state.sortDirection === 'asc' ? -1 : 1;
                if (strA > strB) return state.sortDirection === 'asc' ? 1 : -1;
                return 0;
            });

            state.currentPage = 1;
            renderTable();
            hideLoading();
            showToast(`Đã sắp xếp theo "${state.headers[index]}" (${state.sortDirection === 'asc' ? 'Tăng dần' : 'Giảm dần'})`);
        }, 30);
    }

    // --- Search Filtering ---
    function applySearch(query) {
        if (query) {
            clearSearchBtn.classList.remove('hidden');
        } else {
            clearSearchBtn.classList.add('hidden');
        }

        state.filteredData = state.sheetData.filter(row => {
            return row.some(cellValue => {
                if (cellValue === null || cellValue === undefined) return false;
                const formatted = formatCellValue(cellValue).toLowerCase();
                return formatted.includes(query);
            });
        });

        state.currentPage = 1;
        renderTable();
    }

    searchInput.addEventListener('input', (e) => {
        const query = e.target.value.trim().toLowerCase();
        clearTimeout(searchDebounceTimer);
        searchDebounceTimer = setTimeout(() => applySearch(query), 200);
    });

    clearSearchBtn.addEventListener('click', () => {
        searchInput.value = '';
        clearTimeout(searchDebounceTimer);
        applySearch('');
    });

    tableBody.addEventListener('click', (e) => {
        const toggleBtn = e.target.closest('.group-header-toggle');
        if (!toggleBtn) return;
        const headerRow = toggleBtn.closest('tr.group-header-row');
        if (!headerRow || !headerRow.dataset.groupKey) return;
        toggleGroupCollapse(decodeURIComponent(headerRow.dataset.groupKey));
    });

    // --- View Mode & Group Filter ---
    viewModeSelect.addEventListener('change', (e) => {
        state.viewMode = e.target.value;
        state.currentPage = 1;
        renderTable();
    });

    groupFilterSelect.addEventListener('change', (e) => {
        state.groupFilter = e.target.value;
        state.currentPage = 1;
        renderTable();
    });

    // --- Page Size Handler ---
    pageSizeSelect.addEventListener('change', (e) => {
        state.pageSize = e.target.value;
        state.currentPage = 1;
        renderTable();
    });

    // --- Pagination Buttons ---
    function renderPaginationControls(totalRows) {
        pageNumbers.innerHTML = '';
        const size = parseInt(state.pageSize);
        const totalPages = Math.ceil(totalRows / size);
        
        prevPageBtn.disabled = state.currentPage === 1;
        nextPageBtn.disabled = state.currentPage === totalPages;

        // Smart pagination logic (limits visible numbers)
        const maxPagesToShow = 5;
        let startPage = Math.max(1, state.currentPage - 2);
        let endPage = Math.min(totalPages, startPage + maxPagesToShow - 1);
        
        if (endPage - startPage + 1 < maxPagesToShow) {
            startPage = Math.max(1, endPage - maxPagesToShow + 1);
        }

        // First Page link
        if (startPage > 1) {
            addPageButton(1);
            if (startPage > 2) {
                const ellipsis = document.createElement('span');
                ellipsis.className = 'page-num-ellipsis';
                ellipsis.textContent = '...';
                pageNumbers.appendChild(ellipsis);
            }
        }

        // Page Numbers
        for (let p = startPage; p <= endPage; p++) {
            addPageButton(p);
        }

        // Last Page link
        if (endPage < totalPages) {
            if (endPage < totalPages - 1) {
                const ellipsis = document.createElement('span');
                ellipsis.className = 'page-num-ellipsis';
                ellipsis.textContent = '...';
                pageNumbers.appendChild(ellipsis);
            }
            addPageButton(totalPages);
        }
    }

    function addPageButton(pageNum) {
        const btn = document.createElement('button');
        btn.className = 'page-num-btn';
        if (pageNum === state.currentPage) {
            btn.classList.add('active');
        }
        btn.textContent = pageNum;
        btn.addEventListener('click', () => {
            state.currentPage = pageNum;
            renderTable();
            // Scroll to toolbar level so they see the table head
            document.querySelector('.toolbar').scrollIntoView({ behavior: 'smooth' });
        });
        pageNumbers.appendChild(btn);
    }

    prevPageBtn.addEventListener('click', () => {
        if (state.currentPage > 1) {
            state.currentPage--;
            renderTable();
        }
    });

    nextPageBtn.addEventListener('click', () => {
        const size = parseInt(state.pageSize);
        const totalPages = Math.ceil(state.filteredData.length / size);
        if (state.currentPage < totalPages) {
            state.currentPage++;
            renderTable();
        }
    });

    createQrBtn.addEventListener('click', handleCreateQrClick);
    qrStatsSummary.addEventListener('click', (e) => {
        const btn = e.target.closest('[data-qr-download]');
        if (!btn) return;
        handleQrSummaryDownload(btn.dataset.qrDownload);
    });
    pdfInput.addEventListener('change', handlePdfFileSelected);
    qrStatsCloseBtn.addEventListener('click', closeQrStatsModal);
    qrStatsModal.addEventListener('click', (e) => {
        if (e.target === qrStatsModal) closeQrStatsModal();
    });
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && !qrStatsModal.classList.contains('hidden')) {
            closeQrStatsModal();
        }
    });

    // --- Page Boot ---
    initTheme();
    initLucideIcons();
    checkLibraries();
});
