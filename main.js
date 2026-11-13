import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js";
import { get, getDatabase, ref } from "https://www.gstatic.com/firebasejs/10.12.5/firebase-database.js";

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
	showAllColumns: false
};

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

function getRowValues(targetRow) {
	return Array.from(targetRow.querySelectorAll("td")).map((cell) => String(cell.textContent ?? "").trim());
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
	if (!app || !state.sheetNames.length || !state.activeSheetName) {
		return;
	}

	const rows = extractSheetRows(state.activeSheetName);
	const visibleColumnIndexes = getVisibleColumnIndexes();
	const visibleHeaders = visibleColumnIndexes.map((index) => headers[index]);
	const blColumnIndex = headers.indexOf("BL");
	const blGroupColors = new Map();
	let nextBlGroupIndex = 0;
	const headHtml = visibleHeaders
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
					return `<tr data-row-index="${rowIndex}"${rowStyle}>${visibleColumnIndexes
						.map((index) => {
							const value = row[index] ?? "";
							const isRemark = headers[index] === "비고";
							const className = isRemark ? ' class="remark-cell"' : "";
							const title = isRemark && value ? ` title="${value.replace(/"/g, "&quot;")}"` : "";
							return `<td${className}${title}${rowStyle}>${value}</td>`;
						})
						.join("")}</tr>`;
				})
				.join("")
		: `<tr><td class="empty" colspan="${visibleHeaders.length}">데이터가 없습니다.</td></tr>`;

	const tabs = state.sheetNames
		.map(
			(name) =>
				`<button class="sheet-tab ${name === state.activeSheetName ? "active" : ""}" data-sheet="${name}">${name}</button>`
		)
		.join("") + `<button class="sheet-tab sheet-add-tab" type="button" title="업체 추가(+)">업체추가(+)</button>`;
	app.innerHTML = `
		<section class="app-main">
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
		</section>
	`;

	Array.from(app.querySelectorAll(".sheet-tab[data-sheet]")).forEach((button) => {
		button.addEventListener("click", async () => {
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
			renderTable();
		});
	});

	Array.from(app.querySelectorAll("th.toggle-header")).forEach((headerCell) => {
		headerCell.addEventListener("click", () => {
			state.showAllColumns = !state.showAllColumns;
			renderTable();
		});
	});
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
		await loadSheetNamesFromDatabase();
		state.activeSheetName = state.sheetNames[0] || "";
		if (!state.activeSheetName) {
			renderMessage("표시할 시트 데이터가 없습니다.");
			return;
		}

		await loadSheetRowsFromDatabase(state.activeSheetName);

		renderTable();
	} catch (error) {
		renderMessage(error.message || "Firebase 데이터 로딩 중 오류가 발생했습니다.");
	}
}

if (app) {
	loadWorkbook();
}
