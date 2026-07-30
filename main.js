const SOURCE_WORKBOOK_PATH = "./2물류통합재고목록.xlsm";

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
	"전월까지총출고수량",
	"당월출고수량",
	"총출고수량",
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
	workbook: null,
	activeSheetName: "",
	showAllColumns: false
};

function getVisibleColumnIndexes() {
	return headers.reduce((indexes, headerLabel, index) => {
		if (state.showAllColumns || defaultVisibleHeaders.has(headerLabel)) {
			indexes.push(index);
		}
		return indexes;
	}, []);
}

function normalizeText(value) {
	return String(value ?? "")
		.replace(/["']/g, "")
		.replace(/\s+/g, "")
		.trim();
}

function findHeaderRowIndex(rows) {
	const targetSet = new Set(headers.map(normalizeText));
	let bestIndex = -1;
	let bestScore = 0;

	rows.forEach((row, rowIndex) => {
		const score = new Set(
			row
				.map((cell) => normalizeText(cell))
				.filter((normalized) => normalized && targetSet.has(normalized))
		).size;

		if (score > bestScore) {
			bestScore = score;
			bestIndex = rowIndex;
		}
	});

	return bestScore > 0 ? bestIndex : -1;
}

function buildColumnMap(headerRow) {
	return headers.map((headerLabel) => {
		const normalizedLabel = normalizeText(headerLabel);
		return headerRow.findIndex((cell) => normalizeText(cell) === normalizedLabel);
	});
}

function extractSheetRows(sheetName) {
	const ws = state.workbook?.Sheets?.[sheetName];
	if (!ws) {
		return [];
	}

	const rows = XLSX.utils.sheet_to_json(ws, {
		header: 1,
		defval: "",
		blankrows: false,
		raw: false
	});
	if (!rows.length) {
		return [];
	}

	const headerRowIndex = findHeaderRowIndex(rows);
	if (headerRowIndex < 0) {
		return [];
	}

	const columnMap = buildColumnMap(rows[headerRowIndex]);
	const dataRows = rows.slice(headerRowIndex + 1);

	return dataRows
		.map((row) =>
			columnMap.map((columnIndex) => {
				if (columnIndex < 0) {
					return "";
				}
				return String(row[columnIndex] ?? "").trim();
			})
		)
		.filter((row) => row.some((cell) => cell !== ""));
}

function renderTable() {
	if (!app || !state.workbook || !state.activeSheetName) {
		return;
	}

	const rows = extractSheetRows(state.activeSheetName);
	const visibleColumnIndexes = getVisibleColumnIndexes();
	const visibleHeaders = visibleColumnIndexes.map((index) => headers[index]);
	const headHtml = visibleHeaders
		.map(
			(label) =>
				`<th class="toggle-header" title="클릭하여 ${state.showAllColumns ? "핵심 컬럼만" : "전체 컬럼"} 보기">${label}</th>`
		)
		.join("");
	const bodyHtml = rows.length
		? rows
				.map(
					(row) =>
						`<tr>${visibleColumnIndexes
							.map((index) => {
								const value = row[index] ?? "";
								const isRemark = headers[index] === "비고";
								const className = isRemark ? ' class="remark-cell"' : "";
								const title = isRemark && value ? ` title="${value.replace(/"/g, "&quot;")}"` : "";
								return `<td${className}${title}>${value}</td>`;
							})
							.join("")}</tr>`
				)
				.join("")
		: `<tr><td class="empty" colspan="${visibleHeaders.length}">데이터가 없습니다.</td></tr>`;

	const tabs = state.workbook.SheetNames.map(
		(name) =>
			`<button class="sheet-tab ${name === state.activeSheetName ? "active" : ""}" data-sheet="${name}">${name}</button>`
	).join("");

	app.innerHTML = `
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
	`;

	Array.from(app.querySelectorAll(".sheet-tab")).forEach((button) => {
		button.addEventListener("click", () => {
			state.activeSheetName = button.dataset.sheet || "";
			renderTable();
		});
	});

	Array.from(app.querySelectorAll("th.toggle-header")).forEach((headerCell) => {
		headerCell.addEventListener("click", () => {
			state.showAllColumns = !state.showAllColumns;
			renderTable();
		});
	});

	const tableBody = app.querySelector("tbody");
	if (!tableBody) {
		return;
	}

	tableBody.addEventListener("click", (event) => {
		const targetCell = event.target.closest("td");
		if (!targetCell) {
			return;
		}

		const targetRow = targetCell.closest("tr");
		if (!targetRow) {
			return;
		}

		const isEmptyRow = targetRow.querySelector("td.empty");
		if (isEmptyRow) {
			return;
		}

		Array.from(tableBody.querySelectorAll("tr.editing")).forEach((row) => {
			row.classList.remove("editing");
			Array.from(row.querySelectorAll("td")).forEach((cell) => {
				cell.removeAttribute("contenteditable");
			});
		});

		targetRow.classList.add("editing");
		Array.from(targetRow.querySelectorAll("td")).forEach((cell) => {
			cell.setAttribute("contenteditable", "true");
		});

		targetCell.focus();
		const selection = window.getSelection();
		const range = document.createRange();
		range.selectNodeContents(targetCell);
		range.collapse(false);
		selection.removeAllRanges();
		selection.addRange(range);
	});
}

function renderMessage(message) {
	if (!app) {
		return;
	}
	app.innerHTML = `<section class="table-section"><p class="message">${message}</p></section>`;
}

async function loadWorkbook() {
	if (!window.XLSX) {
		renderMessage("XLSX 라이브러리를 불러오지 못했습니다.");
		return;
	}

	try {
		const response = await fetch(encodeURI(SOURCE_WORKBOOK_PATH), { cache: "no-store" });
		if (!response.ok) {
			throw new Error("루트 경로에서 엑셀 파일을 찾을 수 없습니다.");
		}

		const buffer = await response.arrayBuffer();
		state.workbook = XLSX.read(buffer, {
			type: "array",
			cellDates: true,
			raw: false
		});

		state.activeSheetName = state.workbook.SheetNames[0] || "";
		if (!state.activeSheetName) {
			renderMessage("워크북에 시트가 없습니다.");
			return;
		}

		renderTable();
	} catch (error) {
		renderMessage(error.message || "엑셀 파일 로딩 중 오류가 발생했습니다.");
	}
}

if (app) {
	loadWorkbook();
}
