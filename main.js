const SOURCE_WORKBOOK_PATH = "./2물류통합재고목록.xlsm";

const elements = {
	reloadBtn: document.getElementById("reloadBtn"),
	excelFileInput: document.getElementById("excelFileInput"),
	fileInfo: document.getElementById("fileInfo"),
	stats: document.getElementById("stats"),
	sheetTabs: document.getElementById("sheetTabs"),
	previewCount: document.getElementById("previewCount"),
	excelWrap: document.getElementById("excelWrap"),
	rowEditor: document.getElementById("rowEditor"),
	addRowBtn: document.getElementById("addRowBtn"),
	downloadBtn: document.getElementById("downloadBtn")
};

let workbook = null;
let activeSheetName = "";

function formatNumber(value) {
	return new Intl.NumberFormat("ko-KR").format(value);
}

function decodeRange(ref) {
	if (!ref) {
		return { rows: 0, cols: 0 };
	}
	const range = XLSX.utils.decode_range(ref);
	return {
		rows: range.e.r - range.s.r + 1,
		cols: range.e.c - range.s.c + 1
	};
}

function renderStats() {
	if (!workbook) {
		elements.stats.innerHTML = "";
		return;
	}

	const sheetNames = workbook.SheetNames || [];
	let totalRows = 0;
	let totalCols = 0;
	let mergedCount = 0;

	sheetNames.forEach((name) => {
		const ws = workbook.Sheets[name];
		const ref = ws["!ref"];
		const size = decodeRange(ref);
		totalRows += size.rows;
		totalCols = Math.max(totalCols, size.cols);
		mergedCount += (ws["!merges"] || []).length;
	});

	const cards = [
		{ label: "시트 수", value: formatNumber(sheetNames.length) },
		{ label: "총 행 범위", value: formatNumber(totalRows) },
		{ label: "최대 열 범위", value: formatNumber(totalCols) },
		{ label: "병합셀 수", value: formatNumber(mergedCount) }
	];

	elements.stats.innerHTML = cards
		.map(
			(item) =>
				`<article class="stat-card"><div class="stat-label">${item.label}</div><div class="stat-value">${item.value}</div></article>`
		)
		.join("");
}

function renderActiveSheet() {
	if (!workbook || !activeSheetName) {
		elements.previewCount.textContent = "표시할 시트가 없습니다.";
		elements.excelWrap.innerHTML = "";
		renderRowEditor();
		return;
	}

	const ws = workbook.Sheets[activeSheetName];
	if (!ws || !ws["!ref"]) {
		elements.previewCount.textContent = `${activeSheetName} 시트 데이터 없음`;
		elements.excelWrap.innerHTML = "";
		renderRowEditor();
		return;
	}

	const rangeInfo = decodeRange(ws["!ref"]);
	elements.previewCount.textContent = `${activeSheetName} / ${formatNumber(rangeInfo.rows)}행 x ${formatNumber(
		rangeInfo.cols
	)}열`;

	const html = XLSX.utils.sheet_to_html(ws, {
		id: "excel-sheet",
		editable: false
	});
	elements.excelWrap.innerHTML = html;
	renderRowEditor();
}

function getSheetColumns(ws) {
	if (!ws || !ws["!ref"]) {
		return [];
	}
	const rows = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false, defval: "" });
	const headerRow = rows[0] || [];
	const colCount = decodeRange(ws["!ref"]).cols;

	return Array.from({ length: colCount }, (_, idx) => {
		const header = String(headerRow[idx] ?? "").trim();
		return header || `열 ${idx + 1}`;
	});
}

function renderRowEditor() {
	if (!workbook || !activeSheetName) {
		elements.rowEditor.innerHTML = '<p class="subtle">시트를 선택하면 입력 칸이 표시됩니다.</p>';
		elements.addRowBtn.disabled = true;
		elements.downloadBtn.disabled = true;
		return;
	}

	const ws = workbook.Sheets[activeSheetName];
	const columns = getSheetColumns(ws);
	if (!columns.length) {
		elements.rowEditor.innerHTML = '<p class="subtle">현재 시트에 사용할 열 정보가 없습니다.</p>';
		elements.addRowBtn.disabled = true;
		elements.downloadBtn.disabled = false;
		return;
	}

	elements.rowEditor.innerHTML = columns
		.map(
			(label, idx) =>
				`<label class="field"><span>${label}</span><input type="text" data-col-index="${idx}" placeholder="${label} 값 입력" /></label>`
		)
		.join("");
	elements.addRowBtn.disabled = false;
	elements.downloadBtn.disabled = false;
}

function addRowToActiveSheet() {
	if (!workbook || !activeSheetName) {
		return;
	}

	const ws = workbook.Sheets[activeSheetName];
	const inputs = Array.from(elements.rowEditor.querySelectorAll("input[data-col-index]"));
	if (!inputs.length) {
		return;
	}

	const values = inputs.map((input) => input.value.trim());
	const hasValue = values.some((value) => value !== "");
	if (!hasValue) {
		elements.fileInfo.textContent = "추가할 데이터를 한 칸 이상 입력해 주세요.";
		return;
	}

	XLSX.utils.sheet_add_aoa(ws, [values], { origin: -1 });
	inputs.forEach((input) => {
		input.value = "";
	});

	elements.fileInfo.textContent = `${activeSheetName} 시트에 새 행이 추가되었습니다.`;
	renderStats();
	renderActiveSheet();
}

function downloadWorkbook() {
	if (!workbook) {
		return;
	}
	const dateTag = new Date().toISOString().slice(0, 10).replace(/-/g, "");
	const fileName = `물류통합재고목록_편집본_${dateTag}.xlsx`;
	XLSX.writeFile(workbook, fileName, { bookType: "xlsx", compression: true });
	elements.fileInfo.textContent = `${fileName} 파일로 저장을 시작했습니다.`;
}

async function loadWorkbookFromFile(file) {
	if (!file) {
		return;
	}

	try {
		elements.fileInfo.textContent = `선택 파일 로딩 중: ${file.name}`;
		const buffer = await file.arrayBuffer();
		workbook = XLSX.read(buffer, {
			type: "array",
			cellStyles: true,
			cellFormula: true,
			cellNF: true,
			cellDates: true
		});

		elements.fileInfo.textContent = `${file.name} 분석 완료 (${workbook.SheetNames.length}개 시트)`;
		renderStats();
		renderSheetTabs();
		setActiveSheet(workbook.SheetNames[0] || "");
	} catch (error) {
		resetView(error.message || "선택한 엑셀 파일 파싱 중 오류가 발생했습니다.");
	}
}

function setActiveSheet(sheetName) {
	activeSheetName = sheetName;
	Array.from(elements.sheetTabs.querySelectorAll(".sheet-tab")).forEach((button) => {
		const isActive = button.dataset.sheet === sheetName;
		button.classList.toggle("active", isActive);
	});
	renderActiveSheet();
}

function renderSheetTabs() {
	if (!workbook) {
		elements.sheetTabs.innerHTML = "";
		return;
	}

	elements.sheetTabs.innerHTML = workbook.SheetNames.map(
		(name) => `<button class="sheet-tab" data-sheet="${name}">${name}</button>`
	).join("");

	Array.from(elements.sheetTabs.querySelectorAll(".sheet-tab")).forEach((button) => {
		button.addEventListener("click", () => setActiveSheet(button.dataset.sheet));
	});
}

function resetView(message) {
	workbook = null;
	activeSheetName = "";
	elements.fileInfo.textContent = message;
	elements.stats.innerHTML = "";
	elements.sheetTabs.innerHTML = "";
	elements.previewCount.textContent = "표시할 시트가 없습니다.";
	elements.excelWrap.innerHTML = "";
	renderRowEditor();
}

async function loadWorkbook() {
	try {
		elements.fileInfo.textContent = `원본 파일 로딩 중: ${SOURCE_WORKBOOK_PATH}`;
		const response = await fetch(encodeURI(SOURCE_WORKBOOK_PATH), { cache: "no-store" });
		if (!response.ok) {
			throw new Error(
				`원본 파일을 찾을 수 없습니다. ${SOURCE_WORKBOOK_PATH} 파일을 프로젝트 루트에 배치하세요.`
			);
		}

		const buffer = await response.arrayBuffer();
		workbook = XLSX.read(buffer, {
			type: "array",
			cellStyles: true,
			cellFormula: true,
			cellNF: true,
			cellDates: true
		});

		elements.fileInfo.textContent = `${SOURCE_WORKBOOK_PATH} 분석 완료 (${workbook.SheetNames.length}개 시트)`;
		renderStats();
		renderSheetTabs();
		setActiveSheet(workbook.SheetNames[0] || "");
	} catch (error) {
		resetView(error.message || "엑셀 파일 파싱 중 오류가 발생했습니다.");
	}
}

function bindEvents() {
	elements.reloadBtn.addEventListener("click", loadWorkbook);
	elements.excelFileInput.addEventListener("change", (event) => {
		const [file] = event.target.files || [];
		loadWorkbookFromFile(file);
	});
	elements.addRowBtn.addEventListener("click", addRowToActiveSheet);
	elements.downloadBtn.addEventListener("click", downloadWorkbook);
}

function init() {
	bindEvents();
	loadWorkbook();
}

init();
