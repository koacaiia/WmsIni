import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js";
import { get, getDatabase, push, ref, remove, set } from "https://www.gstatic.com/firebasejs/10.12.5/firebase-database.js";

const FIREBASE_DB_ROOT_PATH = "wms/consignee";
const CONSIGNEE_ROW_NODE_KEYS = ["재고목록", "반입일_BL_품목"];

const firebaseConfig = {
	apiKey: "AIzaSyDLzmZyt5nZwCk98iZ6wi01y7Jxio1ppZQ",
	authDomain: "fine-bondedwarehouse.firebaseapp.com",
	databaseURL: "https://fine-bondedwarehouse-default-rtdb.asia-southeast1.firebasedatabase.app",
	projectId: "fine-bondedwarehouse",
	storageBucket: "fine-bondedwarehouse.appspot.com",
	messagingSenderId: "415417723331",
	appId: "1:415417723331:web:15212f190062886281b576",
	measurementId: "G-SWBR4359JQ"
};

const headers = [
	"재고이관일",
	"반입일",
	"BL",
	"내국관리번호",
	"팔렛트 유형",
	"품목",
	"유통기한",
	"팔렛트 보관료",
	"사용 팔렛트 입고수량",
	"반입수량(EA)",
	"반입수량(PLT)",
	"부족수량(SHORTAGE)/EA",
	"파손수량(EA)",
	"파손수량(PLT)",
	"파손출고수량(EA)",
	"출고가능수량(EA)",
	"출고가능수량(PLT)",
	"총재고수량(EA)",
	"총재고수량(PLT)",
	"비고",
	"적재수량(PLT당)",
	"보관료발생일",
	"프리타임",
	"보관요율(팔렛트당)"
];

const defaultVisibleHeaders = new Set([
	"반입일",
	"BL",
	"품목",
	"총재고수량(EA)",
	"총재고수량(PLT)",
	"비고",
	"적재수량(PLT당)"
]);

const app = document.getElementById("app");
const state = {
	database: null,
	sheetNames: [],
	sheetRowsByName: {},
	activeSheetName: "",
	showAllColumns: false,
	selectedRowIndices: new Set(),
	selectedCell: null,
	sheetLoadError: "",
	activeView: "schedule",
	scheduleYear: "",
	scheduleYears: [],
	scheduleTemplateHtml: "",
	expandedScheduleGroups: new Set(),
	scheduleMonths: [],
	scheduleRecordsByMonth: {},
	scheduleLoadedStart: -1,
	scheduleLoadedEnd: -1,
	scheduleBusy: false,
	scheduleRequestId: 0,
	scheduleListMode: null,
	scheduleSelection: new Map(),
	scheduleRowRecords: new Map(),
	summaryTabOpen: false,
	summary: { mode: "week", date: "", startDate: "", endDate: "", requestId: 0 }
};

const INCARGO_ROOT_PATH = "DeptName/WareHouseDept2/InCargo";
const RECYCLEBIN_PATH = "DeptName/WareHouseDept2/recyclebin";
const LIST_MODE_LABELS = { edit: "목록수정", delete: "목록삭제" };
const SUMMARY_TAB_LABEL = "입고목록 요약";
const SUMMARY_MODE_LABELS = { week: "주간", month: "월간", year: "연간", range: "기간검색" };
const SUMMARY_WEEKDAY_LABELS = ["일", "월", "화", "수", "목", "금", "토"];

const VIEW_TABS = [
	{ id: "schedule", label: "스케줄" },
	{ id: "stock", label: "재고" }
];

function sanitizeFirebaseFieldKey(fieldName) {
	const normalized = String(fieldName ?? "").trim().replace(/[.$#[\]/]/g, "_");
	return normalized || "field";
}

function escapeHtml(value) {
	return String(value ?? "")
		.replace(/&/g, "&amp;")
		.replace(/</g, "&lt;")
		.replace(/>/g, "&gt;")
		.replace(/"/g, "&quot;")
		.replace(/'/g, "&#39;");
}

function initFirebaseDatabase() {
	const firebaseApp = initializeApp(firebaseConfig);
	state.database = getDatabase(firebaseApp);
}

function getVisibleColumnIndexes() {
	return headers.reduce((indexes, headerLabel, index) => {
		if (state.showAllColumns || defaultVisibleHeaders.has(headerLabel)) {
			indexes.push(index);
		}
		return indexes;
	}, []);
}

function parseRowsFromSheetNode(sheetNode) {
	if (!sheetNode || typeof sheetNode !== "object") {
		return [];
	}

	if (Array.isArray(sheetNode)) {
		return sheetNode
			.map((row) => (Array.isArray(row) ? row.map((cell) => String(cell ?? "").trim()) : []))
			.filter((row) => row.some((cell) => cell !== ""));
	}

	const keyedContainer = CONSIGNEE_ROW_NODE_KEYS.find(
		(nodeKey) => sheetNode[nodeKey] && typeof sheetNode[nodeKey] === "object"
	);
	const rowContainer = keyedContainer ? sheetNode[keyedContainer] : sheetNode;

	const rowEntries = Object.entries(rowContainer).filter(([key, value]) => {
		if (key === "_meta") {
			return false;
		}
		return value && typeof value === "object" && !Array.isArray(value);
	});

	return rowEntries.map(([, rowObject]) => {
		const fieldLookup = rowObject;
		return headers.map((headerLabel) => {
			if (headerLabel in fieldLookup) {
				return String(fieldLookup[headerLabel] ?? "").trim();
			}

			const safeKey = sanitizeFirebaseFieldKey(headerLabel);
			return String(fieldLookup[safeKey] ?? "").trim();
		});
	});
}

function extractSheetRows(sheetName) {
	const rows = state.sheetRowsByName[sheetName];
	if (!Array.isArray(rows)) {
		return [];
	}
	return rows;
}

function openRowEditPopup(rowIndex, focusColIndex) {
	closeRowEditPopup();

	const row = state.sheetRowsByName[state.activeSheetName]?.[rowIndex];
	if (!row) {
		return;
	}

	const fieldsHtml = getVisibleColumnIndexes()
		.map(
			(index) => `
				<label class="row-edit-field">
					<span class="row-edit-label">${escapeHtml(headers[index])}</span>
					<input type="text" class="row-edit-input" data-col-index="${index}" value="${escapeHtml(row[index] ?? "")}">
				</label>`
		)
		.join("");

	const overlay = document.createElement("div");
	overlay.className = "row-edit-overlay";
	overlay.innerHTML = `
		<div class="row-edit-popup" role="dialog" aria-modal="true" aria-label="행 편집">
			<header class="row-edit-header">행 편집</header>
			<div class="row-edit-body">${fieldsHtml}</div>
			<footer class="row-edit-footer">
				<button type="button" class="row-edit-cancel">취소</button>
				<button type="button" class="row-edit-save">저장</button>
			</footer>
		</div>
	`;
	document.body.append(overlay);

	overlay.addEventListener("click", (event) => {
		if (event.target === overlay) {
			closeRowEditPopup();
		} else if (event.target.closest(".row-edit-cancel")) {
			closeRowEditPopup();
			setSelectedCell();
		} else if (event.target.closest(".row-edit-save")) {
			saveRowEditPopup(rowIndex);
		}
	});
	overlay.addEventListener("keydown", (event) => {
		if (event.key === "Escape") {
			closeRowEditPopup();
		} else if (event.key === "Enter" && event.target.classList.contains("row-edit-input")) {
			event.preventDefault();
			saveRowEditPopup(rowIndex);
		}
	});

	const focusInput = overlay.querySelector(`.row-edit-input[data-col-index="${focusColIndex}"]`);
	(focusInput ?? overlay.querySelector(".row-edit-input"))?.select();
}

function saveRowEditPopup(rowIndex) {
	const row = state.sheetRowsByName[state.activeSheetName]?.[rowIndex];
	if (row) {
		document.querySelectorAll(".row-edit-input").forEach((input) => {
			row[Number(input.dataset.colIndex)] = input.value.trim();
		});
	}
	closeRowEditPopup();
	renderTable();
}

function closeRowEditPopup() {
	document.querySelector(".row-edit-overlay")?.remove();
}

const SPEC_GROUPS = [
	["20FT", "40FT", "LCL"],
	["Pallet", "Bulk"],
];

function buildSpecFieldHtml(label, value) {
	const selected = String(value ?? "").split(",").map((part) => part.trim());
	const groupsHtml = SPEC_GROUPS.map(
		(options, index) => `<div class="spec-option-group">${options
			.map(
				(option) =>
					`<label><input type="radio" name="spec-group${index + 1}" value="${escapeHtml(option)}"${selected[index] === option ? " checked" : ""}> ${escapeHtml(option)}</label>`
			)
			.join("")}</div>`
	).join("");
	return `
				<div class="row-edit-field spec-field" data-key="spec">
					<span class="row-edit-label">${escapeHtml(label)}</span>
					${groupsHtml}
				</div>`;
}

function readSpecField(root) {
	const parts = SPEC_GROUPS.map((_, index) => root.querySelector(`input[name="spec-group${index + 1}"]:checked`)?.value ?? "");
	return parts.some(Boolean) ? parts.join(",") : "";
}

function collectFormValues(overlay) {
	const values = Object.fromEntries(
		Array.from(overlay.querySelectorAll(".row-edit-input[data-key]")).map((input) => [input.dataset.key, input.value.trim()])
	);
	if (overlay.querySelector(".spec-field")) {
		values.spec = readSpecField(overlay);
	}
	return values;
}

function openNewRecordPopup() {
	const fieldHeaders = Array.from(app.querySelectorAll(".schedule-table thead th[data-key]"));
	if (!fieldHeaders.length || document.querySelector(".new-record-overlay")) {
		return;
	}
	const today = getTodayDateString();
	const fieldsHtml = fieldHeaders
		.map((th) => {
			const key = th.dataset.key;
			if (key === "spec") {
				return buildSpecFieldHtml(th.textContent.trim(), "");
			}
			const value = key === "date" ? today : "";
			return `
				<label class="row-edit-field">
					<span class="row-edit-label">${escapeHtml(th.textContent.trim())}</span>
					<input type="${key === "date" ? "date" : "text"}" class="row-edit-input" data-key="${escapeHtml(key)}" value="${escapeHtml(value)}">
				</label>`;
		})
		.join("");

	const overlay = document.createElement("div");
	overlay.className = "row-edit-overlay new-record-overlay";
	overlay.innerHTML = `
		<div class="row-edit-popup new-record-popup" role="dialog" aria-modal="true" aria-label="신규자료">
			<header class="row-edit-header">신규자료</header>
			<div class="row-edit-body">
				${fieldsHtml}
				<div class="new-list-builder" hidden>
					<label class="row-edit-field">
						<span class="row-edit-label">수량</span>
						<input type="number" min="1" step="1" class="row-edit-input new-list-count">
					</label>
					<div class="new-list-type">
						<label><input type="radio" name="new-list-type" value="container"> Container</label>
						<label><input type="radio" name="new-list-type" value="description"> 품명</label>
					</div>
					<p class="new-list-preview"></p>
				</div>
				<ol class="new-list"></ol>
			</div>
			<footer class="row-edit-footer">
				<span class="new-record-message"></span>
				<button type="button" class="new-record-cancel">취소</button>
				<button type="button" class="row-edit-save new-record-save">저장</button>
				<button type="button" class="new-record-add-list">목록추가</button>
			</footer>
		</div>
	`;
	document.body.append(overlay);

	const builder = overlay.querySelector(".new-list-builder");
	const countInput = overlay.querySelector(".new-list-count");
	const preview = overlay.querySelector(".new-list-preview");
	const listEl = overlay.querySelector(".new-list");
	const messageEl = overlay.querySelector(".new-record-message");
	const items = [];

	const collectForm = () => collectFormValues(overlay);
	const getBatch = () => {
		const field = overlay.querySelector('input[name="new-list-type"]:checked')?.value;
		const count = Number(countInput.value);
		if (!field || !Number.isInteger(count) || count < 1) {
			return null;
		}
		const base = collectForm()[field] ?? "";
		return { field, texts: Array.from({ length: count }, (_, index) => `${base}+${index + 1}`) };
	};
	const updatePreview = () => {
		const batch = getBatch();
		preview.textContent = batch ? `미리보기: ${batch.texts.join(", ")}` : "";
	};
	const renderItems = () => {
		listEl.innerHTML = items
			.map((item) => `<li>${escapeHtml(item.field === "container" ? "Container" : "품명")}: ${escapeHtml(item.text)}</li>`)
			.join("");
	};
	const showMessage = (text) => {
		messageEl.textContent = text;
	};
	const addToList = () => {
		if (builder.hidden) {
			builder.hidden = false;
			countInput.focus();
			return;
		}
		const batch = getBatch();
		if (!batch) {
			showMessage("수량과 Container/품명을 선택해 주세요.");
			return;
		}
		const form = collectForm();
		batch.texts.forEach((text) => {
			items.push({ field: batch.field, text, record: { ...form, [batch.field]: text } });
		});
		countInput.value = "";
		overlay.querySelectorAll('input[name="new-list-type"]').forEach((radio) => {
			radio.checked = false;
		});
		builder.hidden = true;
		updatePreview();
		showMessage("");
		renderItems();
	};
	const saveRecords = async () => {
		const records = items.length ? items.map((item) => item.record) : [collectForm()];
		const error = records.map(getScheduleRecordError).find(Boolean);
		if (error) {
			showMessage(error);
			return;
		}

		const saveButton = overlay.querySelector(".new-record-save");
		saveButton.disabled = true;
		try {
			await Promise.all(
				records.map((record) =>
					set(
						ref(state.database, buildScheduleRecordPath(record.date, record.consignee, getScheduleNodeKey(record))),
						buildScheduleNodeValue(record)
					)
				)
			);
		} catch (error) {
			saveButton.disabled = false;
			showMessage(error.message || "저장 중 오류가 발생했습니다.");
			return;
		}

		overlay.remove();
		state.scheduleYears = [];
		state.scheduleYear = records[0].date.slice(0, 4);
		await loadScheduleView();
	};

	overlay.addEventListener("input", updatePreview);
	overlay.addEventListener("change", updatePreview);
	overlay.addEventListener("click", (event) => {
		if (event.target === overlay || event.target.closest(".new-record-cancel")) {
			overlay.remove();
		} else if (event.target.closest(".new-record-save")) {
			saveRecords();
		} else if (event.target.closest(".new-record-add-list")) {
			addToList();
		}
	});
	overlay.addEventListener("keydown", (event) => {
		if (event.key === "Escape") {
			overlay.remove();
		}
	});
}


async function loadSheetNamesFromDatabase() {
	if (!state.database) {
		throw new Error("Firebase Database 초기화 실패");
	}

	const rootUrl = `${firebaseConfig.databaseURL}/${FIREBASE_DB_ROOT_PATH}.json?shallow=true`;
	const response = await fetch(rootUrl, { cache: "no-store" });
	if (!response.ok) {
		throw new Error("시트 목록 조회에 실패했습니다.");
	}

	const rootKeys = await response.json();
	if (!rootKeys || typeof rootKeys !== "object") {
		throw new Error("Realtime Database 시트 목록 형식이 올바르지 않습니다.");
	}

	const sheetNames = Object.keys(rootKeys).filter((key) => key !== "_meta");
	if (!sheetNames.length) {
		throw new Error("시트 노드가 없습니다. wms/consignee 하위 데이터를 확인하세요.");
	}

	state.sheetNames = sheetNames;
}

async function loadSheetRowsFromDatabase(sheetName, forceReload = false) {
	if (!state.database || !sheetName) {
		return;
	}

	if (!forceReload && Array.isArray(state.sheetRowsByName[sheetName])) {
		return;
	}

	const sheetSnapshot = await get(ref(state.database, `${FIREBASE_DB_ROOT_PATH}/${sheetName}`));
	if (!sheetSnapshot.exists()) {
		state.sheetRowsByName[sheetName] = [];
		return;
	}

	state.sheetRowsByName[sheetName] = parseRowsFromSheetNode(sheetSnapshot.val());
}

function renderTable() {
	if (!app) {
		return;
	}

	const rows = extractSheetRows(state.activeSheetName);
	const visibleColumnIndexes = getVisibleColumnIndexes();
	const visibleHeaders = visibleColumnIndexes.map((index) => headers[index]);
	const blColumnIndex = headers.indexOf("BL");
	const blGroupColors = new Map();
	let nextBlGroupIndex = 0;
	const headHtml = `<th class="select-header"></th>` + visibleHeaders
		.map(
			(label) =>
				`<th class="toggle-header" title="클릭하여 ${state.showAllColumns ? "핵심 컬럼만" : "전체 컬럼"} 보기">${label}</th>`
		)
		.join("");
	const bodyHtml = rows.length
		? rows
				.map((row, rowIndex) => {
					const currentBlValue = blColumnIndex >= 0 ? String(row[blColumnIndex] ?? "").trim() : "";
					const groupKey = currentBlValue || "__EMPTY__";
					const rowBackgroundColor = groupKey === "__EMPTY__"
						? "#ffffff"
						: (() => {
								if (!blGroupColors.has(groupKey)) {
									blGroupColors.set(groupKey, nextBlGroupIndex % 2 === 0 ? "#e9edf2" : "#ffffff");
									nextBlGroupIndex += 1;
								}
								return blGroupColors.get(groupKey);
							})();
					const rowStyle = ` style="background-color:${rowBackgroundColor};"`;
					return `<tr data-row-index="${rowIndex}"${rowStyle}><td class="select-cell"><input type="radio" aria-label="행 선택"${state.selectedRowIndices.has(rowIndex) ? " checked" : ""}></td>${visibleColumnIndexes
						.map((index) => {
							const value = row[index] ?? "";
							const isRemark = headers[index] === "비고";
							const isSelectedCell = rowIndex === state.selectedCell?.rowIndex && index === state.selectedCell?.colIndex;
							const classNames = [isRemark ? "remark-cell" : "", isSelectedCell ? "cell-inverted" : ""].filter(Boolean);
							const className = classNames.length ? ` class="${classNames.join(" ")}"` : "";
							const title = isRemark && value ? ` title="${value.replace(/"/g, "&quot;")}"` : "";
							return `<td data-col-index="${index}"${className}${title}${rowStyle}>${value}</td>`;
						})
						.join("")}</tr>`;
				})
				.join("")
		: `<tr><td class="empty" colspan="${visibleHeaders.length + 1}">${escapeHtml(state.sheetLoadError || "데이터가 없습니다.")}</td></tr>`;

	const tabs = state.sheetNames
		.map(
			(name) =>
				`<button class="sheet-tab ${name === state.activeSheetName ? "active" : ""}" data-sheet="${name}">${name}</button>`
		)
		.join("") + `<button class="sheet-tab sheet-add-tab" type="button" title="업체 추가(+)">업체추가(+)</button>`;
	const viewTabsHtml = VIEW_TABS
		.map(
			(view) =>
				`<button class="sheet-tab view-tab ${view.id === state.activeView ? "active" : ""}" type="button" data-view="${view.id}">${view.label}</button>`
		)
		.join("");
	const contentHtml = state.activeView === "stock"
		? `
		<section class="app-stock">
			<section class="table-section">
				<div class="sheet-tabs">${tabs}</div>
				<div class="table-wrap">
					<table>
						<thead>
							<tr>${headHtml}</tr>
						</thead>
						<tbody>${bodyHtml}</tbody>
					</table>
				</div>
			</section>
		</section>`
		: state.activeView === "summary"
			? `<section class="app-summary"></section>`
			: `<section class="app-schedule"></section>`;
	const summaryTabHtml = state.summaryTabOpen
		? `<button class="sheet-tab view-tab ${state.activeView === "summary" ? "active" : ""}" type="button" data-view="summary">${SUMMARY_TAB_LABEL}</button>`
		: "";
	const yearPickerHtml = state.activeView === "schedule"
		? `<label class="year-picker">년도 <select class="year-select" aria-label="년도 선택">${buildYearOptionsHtml()}</select></label><button type="button" class="new-record-btn">신규자료+</button><button type="button" class="list-mode-btn" data-mode="edit">목록수정</button><button type="button" class="list-mode-btn" data-mode="delete">목록삭제</button><button type="button" class="summary-open-btn">입고요약</button>`
		: "";
	app.innerHTML = `
		<div class="view-tabs">${viewTabsHtml}${summaryTabHtml}${yearPickerHtml}</div>
		${contentHtml}
	`;

	if (state.activeView === "schedule") {
		loadScheduleView();
		app.querySelector(".new-record-btn")?.addEventListener("click", openNewRecordPopup);
		app.querySelectorAll(".list-mode-btn").forEach((button) => {
			button.addEventListener("click", () => setListMode(button.dataset.mode));
		});
		updateListModeControls();
	}

	app.querySelector(".summary-open-btn")?.addEventListener("click", openSummaryTab);
	if (state.activeView === "summary") {
		loadSummaryView();
	}

	Array.from(app.querySelectorAll(".view-tab[data-view]")).forEach((button) => {
		button.addEventListener("click", () => {
			state.activeView = button.dataset.view;
			renderTable();
		});
	});

	app.querySelector(".year-select")?.addEventListener("change", (event) => {
		state.scheduleYear = event.target.value;
		state.scheduleSelection = new Map();
		loadScheduleView();
	});

	Array.from(app.querySelectorAll(".sheet-tab[data-sheet]")).forEach((button) => {
		button.addEventListener("click", async () => {
			state.selectedRowIndices = new Set();
			state.selectedCell = null;
			state.activeSheetName = button.dataset.sheet || "";
			await loadSheetRowsFromDatabase(state.activeSheetName, true);
			renderTable();
		});
	});

	Array.from(app.querySelectorAll(".sheet-add-tab")).forEach((button) => {
		button.addEventListener("click", () => {
			const inputName = window.prompt("추가할 업체명을 입력하세요.");
			const sheetName = String(inputName ?? "").trim();
			if (!sheetName) {
				return;
			}

			if (!state.sheetNames.includes(sheetName)) {
				state.sheetNames = [...state.sheetNames, sheetName];
			}
			state.sheetRowsByName[sheetName] = state.sheetRowsByName[sheetName] ?? [];
			state.activeSheetName = sheetName;
			state.selectedRowIndices = new Set();
			state.selectedCell = null;
			renderTable();
		});
	});

	Array.from(app.querySelectorAll("th.toggle-header")).forEach((headerCell) => {
		headerCell.addEventListener("click", () => {
			state.showAllColumns = !state.showAllColumns;
			renderTable();
		});
	});

	app.querySelector("tbody")?.addEventListener("click", (event) => {
		const selectCell = event.target.closest("td.select-cell");
		if (selectCell) {
			const radio = selectCell.querySelector('input[type="radio"]');
			const rowIndex = Number(selectCell.parentElement.dataset.rowIndex);
			const selected = !state.selectedRowIndices.has(rowIndex);
			radio.checked = selected;
			if (selected) {
				state.selectedRowIndices.add(rowIndex);
			} else {
				state.selectedRowIndices.delete(rowIndex);
				if (state.selectedCell?.rowIndex === rowIndex) {
					setSelectedCell();
				}
			}
			return;
		}

		const cell = event.target.closest("td[data-col-index]");
		if (!cell) {
			return;
		}
		const rowIndex = Number(cell.parentElement.dataset.rowIndex);
		if (!state.selectedRowIndices.has(rowIndex)) {
			return;
		}
		const colIndex = Number(cell.dataset.colIndex);
		setSelectedCell(rowIndex, colIndex);
		openRowEditPopup(rowIndex, colIndex);
	});
}

function setSelectedCell(rowIndex = null, colIndex = null) {
	state.selectedCell = rowIndex === null ? null : { rowIndex, colIndex };
	app.querySelectorAll("td.cell-inverted").forEach((cell) => cell.classList.remove("cell-inverted"));
	if (rowIndex === null) {
		return;
	}
	const cell = app.querySelector(`tbody tr[data-row-index="${rowIndex}"] td[data-col-index="${colIndex}"]`);
	cell?.classList.add("cell-inverted");
}

function buildYearOptionsHtml() {
	if (!state.scheduleYears.length) {
		return `<option value="">년도 없음</option>`;
	}
	return state.scheduleYears
		.map((year) => `<option value="${year}"${year === state.scheduleYear ? " selected" : ""}>${year}</option>`)
		.join("");
}

async function loadScheduleYears() {
	const response = await fetch(`${firebaseConfig.databaseURL}/${INCARGO_ROOT_PATH}.json?shallow=true`, { cache: "no-store" });
	if (!response.ok) {
		throw new Error("년도 목록 조회에 실패했습니다.");
	}
	const yearKeys = await response.json();
	state.scheduleYears = Object.keys(yearKeys ?? {})
		.filter((key) => /^\d{4}$/.test(key))
		.sort((a, b) => Number(b) - Number(a));
}

function collectScheduleRecords(node, path, records = []) {
	if (!node || typeof node !== "object") {
		return records;
	}
	if ("keyValue" in node || "refValue" in node) {
		records.push({ ...node, _path: path });
		return records;
	}
	Object.entries(node).forEach(([key, child]) => collectScheduleRecords(child, `${path}/${key}`, records));
	return records;
}

async function loadScheduleMonths(year) {
	const response = await fetch(`${firebaseConfig.databaseURL}/${INCARGO_ROOT_PATH}/${year}.json?shallow=true`, { cache: "no-store" });
	if (!response.ok) {
		throw new Error("월 목록 조회에 실패했습니다.");
	}
	const monthKeys = await response.json();
	return Object.keys(monthKeys ?? {})
		.filter((key) => /^\d{2}$/.test(key))
		.sort();
}

async function loadScheduleMonthRecords(year, month) {
	const monthPath = `${INCARGO_ROOT_PATH}/${year}/${month}`;
	const snapshot = await get(ref(state.database, monthPath));
	return collectScheduleRecords(snapshot.exists() ? snapshot.val() : null, monthPath);
}

function getInitialScheduleMonth(months) {
	const today = getTodayDateString();
	const currentMonth = today.slice(5, 7);
	if (state.scheduleYear === today.slice(0, 4) && months.includes(currentMonth)) {
		return currentMonth;
	}
	return months[months.length - 1] || "";
}

function renderLoadedScheduleRows(scrollToToday) {
	const { scheduleMonths, scheduleLoadedStart, scheduleLoadedEnd } = state;
	const records = [];
	for (let index = scheduleLoadedStart; index <= scheduleLoadedEnd && index >= 0; index += 1) {
		records.push(...(state.scheduleRecordsByMonth[scheduleMonths[index]] ?? []));
	}
	renderScheduleRows(records, scrollToToday);
}

async function extendScheduleMonths(direction) {
	const generation = state.scheduleRequestId;
	if (state.scheduleBusy || state.scheduleLoadedStart < 0) {
		return;
	}
	const nextIndex = direction > 0 ? state.scheduleLoadedEnd + 1 : state.scheduleLoadedStart - 1;
	const month = state.scheduleMonths[nextIndex];
	if (month === undefined) {
		return;
	}

	state.scheduleBusy = true;
	try {
		if (!state.scheduleRecordsByMonth[month]) {
			const records = await loadScheduleMonthRecords(state.scheduleYear, month);
			if (generation !== state.scheduleRequestId) {
				return;
			}
			state.scheduleRecordsByMonth[month] = records;
		}
		if (generation !== state.scheduleRequestId) {
			return;
		}
		if (direction > 0) {
			state.scheduleLoadedEnd = nextIndex;
		} else {
			state.scheduleLoadedStart = nextIndex;
		}
		const section = app.querySelector(".app-schedule");
		const previousHeight = section?.scrollHeight ?? 0;
		renderLoadedScheduleRows(false);
		if (section && direction < 0) {
			section.scrollTop += section.scrollHeight - previousHeight;
		}
	} catch (error) {
		console.error(error);
	} finally {
		if (generation === state.scheduleRequestId) {
			state.scheduleBusy = false;
		}
	}
}

function getScheduleRecordDate(record) {
	const dateMatch = [record.date, record.keyValue]
		.map((value) => String(value ?? "").match(/\d{4}-\d{2}-\d{2}/))
		.find(Boolean);
	return dateMatch ? dateMatch[0] : "";
}

function formatDateString(date) {
	const month = String(date.getMonth() + 1).padStart(2, "0");
	const day = String(date.getDate()).padStart(2, "0");
	return `${date.getFullYear()}-${month}-${day}`;
}

function getTodayDateString() {
	return formatDateString(new Date());
}

function renderScheduleRows(records, scrollToToday) {
	const tbody = app.querySelector("#schedule-tbody");
	if (!tbody) {
		return;
	}
	const keys = Array.from(app.querySelectorAll(".schedule-table thead th")).map((th) => th.dataset.key || "");
	const sortedRecords = records
		.map((record) => ({ record, date: getScheduleRecordDate(record) }))
		.sort((a, b) => (a.date || "\uffff").localeCompare(b.date || "\uffff"));
	const containerFields = ["consignee", "count", "container"];
	const blFields = ["consignee", "bl", "description"];
	const groupRuns = (fields) => {
		const starts = [];
		const sizes = new Map();
		sortedRecords.forEach(({ record }, index) => {
			const previous = sortedRecords[index - 1]?.record;
			const same = Boolean(previous) && fields.every((key) => String(record[key] ?? "") === String(previous[key] ?? ""));
			starts[index] = same ? starts[index - 1] : index;
			sizes.set(starts[index], (sizes.get(starts[index]) ?? 0) + 1);
		});
		return { starts, sizes };
	};
	const containerRuns = groupRuns(containerFields);
	const blRuns = groupRuns(blFields);
	const runKey = (prefix, { record, date }, fields) =>
		[prefix, date, ...fields.map((key) => record[key] ?? "")].join("|");
	const specIndex = keys.indexOf("spec");
	state.scheduleRowRecords = new Map(sortedRecords.map(({ record }) => [record._path, record]));

	tbody.innerHTML = sortedRecords.length
		? sortedRecords
				.map((item, index) => {
					const { record, date } = item;
					const aStart = containerRuns.starts[index];
					const bStart = blRuns.starts[index];
					const aKey = containerRuns.sizes.get(aStart) > 1 ? runKey("a", sortedRecords[aStart], containerFields) : "";
					const bKey = blRuns.sizes.get(bStart) > 1 ? runKey("b", sortedRecords[bStart], blFields) : "";
					const aIsChild = aStart !== index;
					const bIsChild = bStart !== index;
					const aHasChildren = Boolean(aKey) && !aIsChild;
					const bHasChildren = Boolean(bKey) && !bIsChild;
					const aExpanded = state.expandedScheduleGroups.has(aKey);
					const bExpanded = state.expandedScheduleGroups.has(bKey);
					const bCount = blRuns.sizes.get(bStart);
					const selectHtml = state.scheduleListMode
						? `<input type="radio" class="schedule-select" data-path="${escapeHtml(record._path)}"${state.scheduleSelection.has(record._path) ? " checked" : ""}>`
						: "";

					const classes = [];
					if (aKey) {
						classes.push("same-group", aIsChild ? "group-child" : "group-parent");
					}
					if (bKey) {
						classes.push("bl-group");
					}
					if (index === 0 || sortedRecords[index - 1].date !== date) {
						classes.push("date-first");
					}
					if (index === sortedRecords.length - 1 || sortedRecords[index + 1].date !== date) {
						classes.push("date-last");
					}
					classes.push("date-block");
					const attrs = [`data-date="${date}"`];
					if (aIsChild) {
						attrs.push(`data-a-key="${escapeHtml(aKey)}"`);
					}
					if (bIsChild) {
						attrs.push(`data-b-key="${escapeHtml(bKey)}"`);
					}
					if (bHasChildren) {
						attrs.push(`data-count="${bCount}"`);
					}
					if ((aIsChild && !aExpanded) || (bIsChild && !bExpanded)) {
						attrs.push("hidden");
					}
					if (classes.length) {
						attrs.push(`class="${classes.join(" ")}"`);
					}

					const toggleButton = (key, expanded) =>
						`<button type="button" class="group-toggle" data-group-key="${escapeHtml(key)}" aria-expanded="${expanded}">${expanded ? "-" : "+"}</button>`;
					const cells = keys
						.map((key, cellIndex) => {
							const toggles = cellIndex === 0
								? (aHasChildren ? toggleButton(aKey, aExpanded) : "") + (bHasChildren ? toggleButton(bKey, bExpanded) : "")
								: "";
							if (cellIndex === specIndex && bHasChildren) {
								const original = escapeHtml(record[key] ?? "");
								const shown = bExpanded ? original : escapeHtml(`${record[key] ?? ""}*${bCount}대`);
								return `<td data-original="${original}">${shown}</td>`;
							}
							return `<td>${cellIndex === 0 ? selectHtml : ""}${toggles}${key ? escapeHtml(record[key] ?? "") : ""}</td>`;
						})
						.join("");
					return `<tr ${attrs.join(" ")}>${cells}</tr>`;
				})
				.join("")
		: `<tr><td class="empty" colspan="${keys.length}">데이터가 없습니다.</td></tr>`;

	tbody.onclick = (event) => {
		const select = event.target.closest(".schedule-select");
		if (select) {
			handleScheduleSelect(select);
			return;
		}
		const toggle = event.target.closest(".group-toggle");
		if (!toggle) {
			return;
		}
		const groupKey = toggle.dataset.groupKey;
		const expanded = !state.expandedScheduleGroups.has(groupKey);
		if (expanded) {
			state.expandedScheduleGroups.add(groupKey);
		} else {
			state.expandedScheduleGroups.delete(groupKey);
		}
		toggle.setAttribute("aria-expanded", String(expanded));
		toggle.textContent = expanded ? "-" : "+";
		tbody.querySelectorAll("tr[data-a-key], tr[data-b-key]").forEach((row) => {
			row.hidden = Boolean(
				(row.dataset.aKey && !state.expandedScheduleGroups.has(row.dataset.aKey)) ||
				(row.dataset.bKey && !state.expandedScheduleGroups.has(row.dataset.bKey))
			);
		});
		if (groupKey.startsWith("b|") && specIndex >= 0) {
			const parentRow = toggle.closest("tr");
			const specCell = parentRow.cells[specIndex];
			specCell.textContent = expanded ? specCell.dataset.original : `${specCell.dataset.original}*${parentRow.dataset.count}대`;
		}
	};

	const scrollContainer = app.querySelector(".app-schedule");
	const todayRow = Array.from(tbody.querySelectorAll("tr[data-date]")).find(
		(row) => !row.hidden && row.dataset.date >= getTodayDateString()
	);
	if (scrollToToday && scrollContainer && todayRow) {
		const headerHeight = app.querySelector(".schedule-table thead")?.getBoundingClientRect().height ?? 0;
		const offset = todayRow.getBoundingClientRect().top - scrollContainer.getBoundingClientRect().top;
		scrollContainer.scrollTop += offset - headerHeight;
	}
}


function getScheduleRecordData(record) {
	const { _path, ...data } = record;
	return data;
}

function getScheduleRecordError(record) {
	if (!/^\d{4}-\d{2}-\d{2}$/.test(record.date ?? "") || !record.consignee) {
		return "Date(YYYY-MM-DD)와 화주를 확인해 주세요.";
	}
	if (/[.#$[\]/]/.test(record.consignee)) {
		return "화주명에는 . # $ [ ] / 를 사용할 수 없습니다.";
	}
	if (/[.#$[\]/]/.test(getScheduleNodeKey(record))) {
		return "BL·품명·수량에는 . # $ [ ] / 를 사용할 수 없습니다.";
	}
	return "";
}

function buildScheduleRecordPath(date, consignee, recordKey) {
	return `${INCARGO_ROOT_PATH}/${date.slice(0, 4)}/${date.slice(5, 7)}/${date.slice(8, 10)}/${consignee}/${recordKey}`;
}

function getScheduleNodeKey(record) {
	return `${record.date}_${record.bl ?? ""}_${record.description ?? ""}_${record.count ?? ""}_Container`;
}

function buildScheduleNodeValue(record) {
	const { _path, keyValue, refValue, ...data } = record;
	const fullPath = buildScheduleRecordPath(record.date, record.consignee, getScheduleNodeKey(record));
	return { ...data, refValue: fullPath };
}

function getSelectedScheduleRecords() {
	return Array.from(state.scheduleSelection.values()).sort(
		(a, b) => getScheduleRecordDate(a).localeCompare(getScheduleRecordDate(b)) || a._path.localeCompare(b._path)
	);
}

function updateListModeControls() {
	const mode = state.scheduleListMode;
	app.querySelectorAll(".list-mode-btn").forEach((button) => {
		button.classList.toggle("active", button.dataset.mode === mode);
	});
	const countText = `${LIST_MODE_LABELS[mode] ?? ""} 선택 ${state.scheduleSelection.size}건`;
	const panel = ensureListActionPanel();
	panel.hidden = !mode;
	panel.querySelector(".list-action-count").textContent = countText;
}

function setListMode(mode) {
	state.scheduleListMode = state.scheduleListMode === mode ? null : mode;
	state.scheduleSelection = new Map();
	updateListModeControls();
	renderLoadedScheduleRows(false);
}

function handleScheduleSelect(radio) {
	const path = radio.dataset.path;
	if (state.scheduleSelection.has(path)) {
		state.scheduleSelection.delete(path);
		radio.checked = false;
		updateListModeControls();
		return;
	}
	const record = state.scheduleRowRecords.get(path);
	if (!record) {
		radio.checked = false;
		return;
	}
	state.scheduleSelection.set(path, record);
	radio.checked = true;
	updateListModeControls();
}

function ensureListActionPanel() {
	let panel = document.querySelector(".list-action-panel");
	if (!panel) {
		panel = document.createElement("div");
		panel.className = "list-action-panel";
		panel.hidden = true;
		panel.innerHTML = `<span class="list-action-count"></span><button type="button" class="list-action-cancel">취소</button><button type="button" class="list-action-complete">선택 완료</button>`;
		panel.querySelector(".list-action-cancel").addEventListener("click", cancelListMode);
		panel.querySelector(".list-action-complete").addEventListener("click", completeListAction);
		document.body.append(panel);
	}
	return panel;
}

function cancelListMode() {
	setListMode(null);
}

async function moveScheduleRecordToRecycleBin(record) {
	await push(ref(state.database, RECYCLEBIN_PATH), { ...getScheduleRecordData(record), deletedFrom: record._path });
	await remove(ref(state.database, record._path));
}

function openScheduleEditPopup(record, sequence) {
	const data = getScheduleRecordData(record);
	const title = [data.date, data.consignee, data.container, data.bl, data.spec, data.description]
		.map((value) => value ?? "")
		.join("_") + `+(${sequence})`;
	const fieldsHtml = Array.from(app.querySelectorAll(".schedule-table thead th[data-key]"))
		.map((th) => {
			const key = th.dataset.key;
			if (key === "spec") {
				return buildSpecFieldHtml(th.textContent.trim(), data.spec);
			}
			return `
				<label class="row-edit-field">
					<span class="row-edit-label">${escapeHtml(th.textContent.trim())}</span>
					<input type="${key === "date" ? "date" : "text"}" class="row-edit-input" data-key="${escapeHtml(key)}" value="${escapeHtml(data[key] ?? "")}">
				</label>`;
		})
		.join("");

	const overlay = document.createElement("div");
	overlay.className = "row-edit-overlay edit-record-overlay";
	overlay.innerHTML = `
		<div class="row-edit-popup" role="dialog" aria-modal="true" aria-label="${escapeHtml(title)}">
			<header class="row-edit-header">${escapeHtml(title)}</header>
			<div class="row-edit-body">${fieldsHtml}</div>
			<footer class="row-edit-footer">
				<span class="new-record-message"></span>
				<button type="button" class="edit-record-cancel">취소</button>
				<button type="button" class="row-edit-save edit-record-save">수정</button>
			</footer>
		</div>
	`;
	document.body.append(overlay);
	overlay.querySelector(".row-edit-input")?.focus();

	return new Promise((resolve) => {
		const close = (result) => {
			overlay.remove();
			resolve(result);
		};
		overlay.addEventListener("click", (event) => {
			if (event.target === overlay || event.target.closest(".edit-record-cancel")) {
				close(null);
			} else if (event.target.closest(".edit-record-save")) {
				const form = collectFormValues(overlay);
				const values = { ...data, ...form };
				const error = getScheduleRecordError(values);
				if (error) {
					overlay.querySelector(".new-record-message").textContent = error;
					return;
				}
				close(values);
			}
		});
		overlay.addEventListener("keydown", (event) => {
			if (event.key === "Escape") {
				close(null);
			}
		});
	});
}

async function completeListAction() {
	const mode = state.scheduleListMode;
	const records = getSelectedScheduleRecords();
	if (!mode) {
		return;
	}
	if (!records.length) {
		window.alert("선택된 목록이 없습니다.");
		return;
	}

	try {
		if (mode === "delete") {
			if (records.length > 1 && !window.confirm(`선택한 ${records.length}건을 삭제하시겠습니까?`)) {
				return;
			}
			for (const record of records) {
				await moveScheduleRecordToRecycleBin(record);
			}
		} else {
			for (const [index, record] of records.entries()) {
				const values = await openScheduleEditPopup(record, index + 1);
				if (!values) {
					continue;
				}
				await moveScheduleRecordToRecycleBin(record);
				await set(
					ref(state.database, buildScheduleRecordPath(values.date, values.consignee, getScheduleNodeKey(values))),
					buildScheduleNodeValue(values)
				);
			}
		}
	} catch (error) {
		window.alert(error.message || "처리 중 오류가 발생했습니다.");
	}

	state.scheduleListMode = null;
	state.scheduleSelection = new Map();
	updateListModeControls();
	await loadScheduleView();
}

async function loadScheduleView() {
	const requestId = ++state.scheduleRequestId;
	const isStale = () => requestId !== state.scheduleRequestId;
	const showMessage = (message) => {
		const section = app.querySelector(".app-schedule");
		if (section) {
			section.innerHTML = `<p class="message">${escapeHtml(message)}</p>`;
		}
	};

	try {
		if (!state.scheduleTemplateHtml) {
			const response = await fetch("./schedule.html", { cache: "no-store" });
			if (!response.ok) {
				throw new Error(`스케줄 화면을 불러오지 못했습니다. (${response.status})`);
			}
			state.scheduleTemplateHtml = await response.text();
		}

		if (!state.scheduleYears.length) {
			await loadScheduleYears();
		}
		if (!state.scheduleYear || !state.scheduleYears.includes(state.scheduleYear)) {
			state.scheduleYear = state.scheduleYears[0] || "";
		}

		const months = state.scheduleYear ? await loadScheduleMonths(state.scheduleYear) : [];
		if (isStale()) {
			return;
		}
		state.scheduleBusy = false;
		state.scheduleMonths = months;
		state.scheduleRecordsByMonth = {};
		const initialMonth = getInitialScheduleMonth(months);
		const initialIndex = months.indexOf(initialMonth);
		state.scheduleLoadedStart = initialIndex;
		state.scheduleLoadedEnd = initialIndex;
		if (initialIndex >= 0) {
			state.scheduleRecordsByMonth[initialMonth] = await loadScheduleMonthRecords(state.scheduleYear, initialMonth);
		}
		if (isStale()) {
			return;
		}

		let section = app.querySelector(".app-schedule");
		if (!section) {
			return;
		}
		if (!section.querySelector(".schedule-table")) {
			section.outerHTML = state.scheduleTemplateHtml;
			section = app.querySelector(".app-schedule");
		}
		if (!section) {
			return;
		}
		const yearSelect = app.querySelector(".year-select");
		if (yearSelect) {
			yearSelect.innerHTML = buildYearOptionsHtml();
		}
		section.onscroll = () => {
			const nearBottom = section.scrollTop + section.clientHeight >= section.scrollHeight - 40;
			const nearTop = section.scrollTop <= 40;
			if (nearBottom) {
				extendScheduleMonths(1);
			} else if (nearTop) {
				extendScheduleMonths(-1);
			}
		};
		renderLoadedScheduleRows(true);
	} catch (error) {
		if (!isStale()) {
			showMessage(error.message || "스케줄 데이터 로딩 중 오류가 발생했습니다.");
		}
	}
}

function parseDateString(dateString) {
	const [year, month, day] = dateString.split("-").map(Number);
	return new Date(year, month - 1, day);
}

function shiftDateString(dateString, days) {
	const date = parseDateString(dateString);
	date.setDate(date.getDate() + days);
	return formatDateString(date);
}

function getSummaryRange() {
	const { mode, date, startDate, endDate } = state.summary;
	if (mode === "range") {
		return { start: startDate, end: endDate, label: `${startDate} ~ ${endDate}` };
	}
	const base = parseDateString(date);
	if (mode === "week") {
		const start = shiftDateString(date, -((base.getDay() + 6) % 7));
		const end = shiftDateString(start, 5);
		return { start, end, label: `${start} ~ ${end}` };
	}
	if (mode === "month") {
		const year = base.getFullYear();
		const month = base.getMonth();
		return {
			start: formatDateString(new Date(year, month, 1)),
			end: formatDateString(new Date(year, month + 1, 0)),
			label: `${year}-${String(month + 1).padStart(2, "0")}`
		};
	}
	return { start: `${base.getFullYear()}-01-01`, end: `${base.getFullYear()}-12-31`, label: `${base.getFullYear()}년` };
}

function shiftSummaryDate(delta) {
	const { mode, date } = state.summary;
	const base = parseDateString(date);
	if (mode === "week") {
		state.summary.date = shiftDateString(date, delta * 7);
	} else if (mode === "month") {
		state.summary.date = formatDateString(new Date(base.getFullYear(), base.getMonth() + delta, 1));
	} else if (mode === "year") {
		state.summary.date = formatDateString(new Date(base.getFullYear() + delta, 0, 1));
	}
}

async function loadSummaryRecords(startDate, endDate) {
	const monthCursor = parseDateString(startDate);
	monthCursor.setDate(1);
	const lastDate = parseDateString(endDate);
	const requests = [];
	while (monthCursor <= lastDate) {
		const year = String(monthCursor.getFullYear());
		const month = String(monthCursor.getMonth() + 1).padStart(2, "0");
		requests.push(loadScheduleMonthRecords(year, month));
		monthCursor.setMonth(monthCursor.getMonth() + 1);
	}
	const records = (await Promise.all(requests)).flat();
	return records
		.filter((record) => {
			const date = getScheduleRecordDate(record);
			return date >= startDate && date <= endDate;
		})
		.sort(
			(a, b) =>
				getScheduleRecordDate(a).localeCompare(getScheduleRecordDate(b)) ||
				String(a.consignee ?? "").localeCompare(String(b.consignee ?? ""), "ko")
		);
}

function groupRecordsByDate(records) {
	const groups = new Map();
	records.forEach((record) => {
		const date = getScheduleRecordDate(record);
		groups.set(date, [...(groups.get(date) ?? []), record]);
	});
	return groups;
}

function isBlankOrZero(value) {
	const text = String(value ?? "").trim();
	return text === "" || text === "0";
}

function hasNonZeroContainer(record) {
	return !isBlankOrZero(record.container20) || !isBlankOrZero(record.container40);
}

function getSummarySpec(record) {
	if (String(record.container20 ?? "").trim() === "0" && String(record.container40 ?? "").trim() === "0" && isBlankOrZero(record.spec)) {
		return `${record.Pqty ?? ""}P(${record.incargo ?? ""}EA)`;
	}
	return record.spec ?? "";
}

function buildSummaryTableHtml(records, showDate = true) {
	if (!records.length) {
		return `<p class="message">데이터가 없습니다.</p>`;
	}
	const groups = new Map();
	records.forEach((record, index) => {
		const container = String(record.container ?? "").trim();
		const key = container ? `${getScheduleRecordDate(record)}|${container}` : `row-${index}`;
		groups.set(key, [...(groups.get(key) ?? []), record]);
	});
	const rows = [...groups.values()]
		.map((members, groupIndex) => {
			const groupId = `g${groupIndex}`;
			const groupSpec =
				members.length > 1
					? getSummarySpec(members.find(hasNonZeroContainer) ?? members[0])
					: getSummarySpec(members[0]);
			return members
				.map((record, memberIndex) => {
					const isSub = memberIndex > 0;
					const toggleHtml =
						memberIndex === 0 && members.length > 1
							? `<button type="button" class="summary-toggle" data-summary-toggle="${groupId}" aria-expanded="false" title="같은 컨테이너 ${members.length}건">+</button>`
							: "";
					return `<tr class="${isSub ? "summary-sub-row" : ""}" data-summary-group="${groupId}">
				${showDate ? `<td>${toggleHtml}${escapeHtml(getScheduleRecordDate(record))}</td>` : ""}
				<td>${showDate ? "" : toggleHtml}${escapeHtml(record.consignee ?? "")}</td>
				<td>${escapeHtml(record.description ?? "")}</td>
				<td>${escapeHtml(members.length > 1 ? groupSpec : getSummarySpec(record))}</td>
			</tr>`;
				})
				.join("");
		})
		.join("");
	return `<div class="table-wrap"><table class="summary-table">
		<thead><tr>${showDate ? "<th>날짜</th>" : ""}<th>화주</th><th>품명</th><th>Spec</th></tr></thead>
		<tbody>${rows}</tbody>
	</table></div>`;
}

function buildSummaryWeekHtml(records, startDate) {
	const recordsByDate = groupRecordsByDate(records);
	const cards = Array.from({ length: 6 }, (_, index) => {
		const date = shiftDateString(startDate, index);
		const weekday = SUMMARY_WEEKDAY_LABELS[parseDateString(date).getDay()];
		return `<section class="summary-day-card">
			<header class="summary-day-head">${date} (${weekday})</header>
			${buildSummaryTableHtml(recordsByDate.get(date) ?? [], false)}
		</section>`;
	});
	return `<div class="summary-week-grid">${cards.join("")}</div>`;
}

function buildSummaryCalendarHtml(records, startDate) {
	const recordsByDate = groupRecordsByDate(records);
	const first = parseDateString(startDate);
	const year = first.getFullYear();
	const month = first.getMonth();
	const daysInMonth = new Date(year, month + 1, 0).getDate();
	const cells = Array.from({ length: first.getDay() }, () => `<td class="summary-cal-empty"></td>`);
	for (let day = 1; day <= daysInMonth; day += 1) {
		const date = formatDateString(new Date(year, month, day));
		const items = (recordsByDate.get(date) ?? [])
			.map(
				(record) =>
					`<div class="summary-cal-item">${escapeHtml(record.consignee ?? "")} · ${escapeHtml(record.description ?? "")}</div>`
			)
			.join("");
		cells.push(`<td><div class="summary-cal-date">${day}</div>${items}</td>`);
	}
	while (cells.length % 7) {
		cells.push(`<td class="summary-cal-empty"></td>`);
	}
	const rows = [];
	for (let index = 0; index < cells.length; index += 7) {
		rows.push(`<tr>${cells.slice(index, index + 7).join("")}</tr>`);
	}
	const headHtml = SUMMARY_WEEKDAY_LABELS.map((label) => `<th>${label}</th>`).join("");
	return `<div class="table-wrap"><table class="summary-calendar">
		<thead><tr>${headHtml}</tr></thead>
		<tbody>${rows.join("")}</tbody>
	</table></div>`;
}

function buildSummaryFilterHtml(range) {
	const { mode, date, startDate, endDate } = state.summary;
	const modeButtonsHtml = Object.entries(SUMMARY_MODE_LABELS)
		.map(
			([id, label]) =>
				`<button type="button" class="${mode === id ? "active" : ""}" data-summary-mode="${id}">${label}</button>`
		)
		.join("");
	const controlsHtml = mode === "range"
		? `<input type="date" class="summary-date" data-summary-field="startDate" value="${startDate}" aria-label="검색 시작일">
			~ <input type="date" class="summary-date" data-summary-field="endDate" value="${endDate}" aria-label="검색 종료일">
			<button type="button" data-summary-search>조회</button>`
		: `<button type="button" data-summary-nav="-1" aria-label="이전">◀</button>
			<input type="date" class="summary-date" data-summary-field="date" value="${date}" aria-label="기준일">
			<button type="button" data-summary-nav="1" aria-label="다음">▶</button>`;
	return `<div class="summary-filter">${modeButtonsHtml}${controlsHtml}<span class="summary-filter-label">${range.label}</span></div>`;
}

function renderSummarySection(bodyHtml) {
	const section = app.querySelector(".app-summary");
	if (!section) {
		return;
	}
	section.innerHTML = `${buildSummaryFilterHtml(getSummaryRange())}<div class="summary-body">${bodyHtml}</div>`;
	section.onclick = handleSummaryClick;
	section.onchange = handleSummaryChange;
}

function handleSummaryClick(event) {
	const toggleButton = event.target.closest("[data-summary-toggle]");
	if (toggleButton) {
		const isOpen = toggleButton.classList.toggle("open");
		toggleButton.textContent = isOpen ? "-" : "+";
		toggleButton.setAttribute("aria-expanded", String(isOpen));
		event.currentTarget
			.querySelectorAll(`tr.summary-sub-row[data-summary-group="${toggleButton.dataset.summaryToggle}"]`)
			.forEach((row) => row.classList.toggle("open", isOpen));
		return;
	}
	const modeButton = event.target.closest("[data-summary-mode]");
	if (modeButton) {
		state.summary.mode = modeButton.dataset.summaryMode;
		loadSummaryView();
		return;
	}
	const navButton = event.target.closest("[data-summary-nav]");
	if (navButton) {
		shiftSummaryDate(Number(navButton.dataset.summaryNav));
		loadSummaryView();
		return;
	}
	if (event.target.closest("[data-summary-search]")) {
		const { startDate, endDate } = state.summary;
		if (!startDate || !endDate) {
			window.alert("검색 기간을 입력해 주세요.");
		} else if (startDate > endDate) {
			window.alert("시작일이 종료일보다 늦을 수 없습니다.");
		} else {
			loadSummaryView();
		}
	}
}

function handleSummaryChange(event) {
	const field = event.target.dataset.summaryField;
	if (!field || !event.target.value) {
		return;
	}
	state.summary[field] = event.target.value;
	if (field === "date") {
		loadSummaryView();
	}
}

function openSummaryTab() {
	if (!state.summaryTabOpen) {
		const today = getTodayDateString();
		state.summary = { ...state.summary, mode: "week", date: today, startDate: today, endDate: today };
	}
	state.summaryTabOpen = true;
	state.activeView = "summary";
	renderTable();
}

async function loadSummaryView() {
	const requestId = ++state.summary.requestId;
	const range = getSummaryRange();
	renderSummarySection(`<p class="message">불러오는 중...</p>`);
	try {
		const records = await loadSummaryRecords(range.start, range.end);
		if (requestId !== state.summary.requestId) {
			return;
		}
		const { mode } = state.summary;
		let bodyHtml;
		if (mode === "week") {
			bodyHtml = buildSummaryWeekHtml(records, range.start);
		} else if (mode === "month") {
			bodyHtml = buildSummaryCalendarHtml(records, range.start);
		} else {
			bodyHtml = buildSummaryTableHtml(records);
		}
		renderSummarySection(bodyHtml);
	} catch (error) {
		if (requestId === state.summary.requestId) {
			renderSummarySection(`<p class="message">${escapeHtml(error.message || "요약 데이터 로딩 중 오류가 발생했습니다.")}</p>`);
		}
	}
}

function renderMessage(message) {
	if (!app) {
		return;
	}
	app.innerHTML = `<section class="table-section"><p class="message">${message}</p></section>`;
}

async function loadWorkbook() {
	try {
		initFirebaseDatabase();
	} catch (error) {
		renderMessage(error.message || "Firebase 초기화 중 오류가 발생했습니다.");
		return;
	}

	try {
		await loadSheetNamesFromDatabase();
		state.activeSheetName = state.sheetNames[0] || "";
		await loadSheetRowsFromDatabase(state.activeSheetName);
	} catch (error) {
		state.sheetLoadError = error.message || "Firebase 데이터 로딩 중 오류가 발생했습니다.";
	}

	renderTable();
}

if (app) {
	loadWorkbook();
}
