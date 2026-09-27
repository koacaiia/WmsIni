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
	showAllColumns: false,
	selectedRowData: null,
	viewMode: "edit",
	attachments: []
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

function escapeAttribute(value) {
	return String(value ?? "")
		.replace(/&/g, "&amp;")
		.replace(/"/g, "&quot;");
}

function formatFileSize(bytes) {
	if (!Number.isFinite(bytes) || bytes <= 0) {
		return "0 B";
	}
	const units = ["B", "KB", "MB", "GB"];
	let size = bytes;
	let unitIndex = 0;
	while (size >= 1024 && unitIndex < units.length - 1) {
		size /= 1024;
		unitIndex += 1;
	}
	return `${size.toFixed(size >= 10 || unitIndex === 0 ? 0 : 1)} ${units[unitIndex]}`;
}

function initFirebaseDatabase() {
	const firebaseApp = initializeApp(firebaseConfig);
	state.database = getDatabase(firebaseApp);
}

function getAttachmentPreviewMarkup(fileInfo) {
	if (fileInfo.previewKind === "image" && fileInfo.previewUrl) {
		return `<div class="attachment-preview-image" data-preview-name="${escapeAttribute(fileInfo.name)}" data-preview-type="image" data-preview-url="${escapeAttribute(fileInfo.previewUrl)}"><img src="${escapeAttribute(fileInfo.previewUrl)}" alt="${escapeAttribute(fileInfo.name)}" /></div>`;
	}
	if (fileInfo.previewText) {
		return `<div class="attachment-preview-text" data-preview-name="${escapeAttribute(fileInfo.name)}" data-preview-type="text" data-preview-text="${escapeAttribute(fileInfo.previewText)}">${escapeHtml(fileInfo.previewText)}</div>`;
	}
	return `<div class="attachment-preview-text">미리보기 불가</div>`;
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

function parseNumericCellValue(value) {
	const raw = String(value ?? "").trim();
	if (!raw) {
		return 0;
	}

	const numericTokens = raw
		.replace(/,/g, "")
		.match(/-?\d+(?:\.\d+)?/g);

	if (!numericTokens || !numericTokens.length) {
		return 0;
	}

	return numericTokens.reduce((sum, token) => {
		const parsed = Number(token);
		return Number.isFinite(parsed) ? sum + parsed : sum;
	}, 0);
}

function buildSideSummaryStats(rows) {
	const blIndex = headers.indexOf("BL");
	const eaIndex = headers.indexOf("총재고수량(EA)");
	const pltIndex = headers.indexOf("총재고수량(PLT)");
	const summaryByBl = new Map();

	rows.forEach((row) => {
		if (!Array.isArray(row)) {
			return;
		}

		const blValue = String(row[blIndex] ?? "").trim();
		if (!blValue) {
			return;
		}

		const eaValue = parseNumericCellValue(row[eaIndex]);
		const pltValue = parseNumericCellValue(row[pltIndex]);
		const existing = summaryByBl.get(blValue) || { bl: blValue, ea: 0, plt: 0 };
		existing.ea += eaValue;
		existing.plt += pltValue;
		summaryByBl.set(blValue, existing);
	});

	const summaries = Array.from(summaryByBl.values());
	const totals = summaries.reduce(
		(acc, item) => {
			acc.totalEa += item.ea;
			acc.totalPlt += item.plt;
			return acc;
		},
		{ totalEa: 0, totalPlt: 0 }
	);

	return {
		summaries,
		blCount: summaryByBl.size,
		totalEa: totals.totalEa,
		totalPlt: totals.totalPlt
	};
}

function renderSideSummaryHtml() {
	const rows = extractSheetRows(state.activeSheetName);
	const stats = buildSideSummaryStats(rows);
	const summaryRowsHtml = stats.summaries.length
		? stats.summaries
			.map(
				(item) => `
					<tr>
						<td>${escapeHtml(item.bl)}</td>
						<td>${item.ea.toLocaleString()}</td>
						<td>${item.plt.toLocaleString()}</td>
					</tr>
				`
			)
			.join("")
		: `
				<tr>
					<td colspan="3">데이터가 없습니다.</td>
				</tr>
		  `;

	return `
		<div class="side-summary-layout">
			<div class="side-summary-table">
				<table class="side-table">
					<thead>
						<tr>
							<th>BL</th>
							<th>EA</th>
							<th>PLT</th>
						</tr>
					</thead>
					<tbody>
						${summaryRowsHtml}
					</tbody>
				</table>
			</div>
			<div class="side-summary-total">
				<div class="side-summary-meta">
					<p>총 BL 건수: ${stats.blCount.toLocaleString()}</p>
					<p>total EA 수량: ${stats.totalEa.toLocaleString()}</p>
					<p>total PLT 수량: ${stats.totalPlt.toLocaleString()}</p>
				</div>
			</div>
		</div>
	`;
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

const shipViewHeaders = [
	"반입일",
	"BL",
	"품목",
	"총재고수량(EA)",
	"총재고수량(PLT)",
	"출고가능수량(EA)",
	"출고가능수량(PLT)",
	"적재수량(PLT당)",
	"비고"
];

function renderSidePanelHtml() {
	if (!state.selectedRowData) {
		return `
			<div class="side-title-wrap">
				<h3 class="side-title">재고 요약</h3>
			</div>
			<div class="side-table-wrap" id="side-summary">
				${renderSideSummaryHtml()}
			</div>
		`;
	}

	if (state.viewMode === "ship") {
		const { headers: selectedHeaders, values: selectedValues } = state.selectedRowData;
		const headerIndexMap = new Map(selectedHeaders.map((header, index) => [header, index]));
		const getValue = (header) => {
			const aliases = header === "부족수량(EA)" ? [header, "부족수량(SHORTAGE)/EA"] : [header];
			for (const candidate of aliases) {
				const index = headerIndexMap.get(candidate);
				if (index !== undefined) {
					return String(selectedValues[index] ?? "").trim();
				}
			}
			return "";
		};
		const hasWarning = [getValue("부족수량(EA)"), getValue("부족수량(SHORTAGE)/EA"), getValue("파손수량(EA)")]
			.some((value) => value !== "" && value.toUpperCase() !== "EA");
		const shipRowsHtml = shipViewHeaders
			.map((header) => {
				const index = headerIndexMap.get(header);
				const value = index !== undefined ? selectedValues[index] ?? "" : "";
				const isWarningRow = hasWarning && (header === "출고가능수량(EA)" || header === "출고가능수량(PLT)");
				return `
					<tr class="${isWarningRow ? "ship-warning-row" : ""}">
						<th>${escapeHtml(header)}</th>
						<td>${escapeHtml(value)}</td>
					</tr>
				`;
			})
			.join("");

		const attachmentRowsHtml = state.attachments.length
			? state.attachments
				.map((file) => `
					<tr>
						<td>${escapeHtml(file.name)}</td>
						<td>${escapeHtml(formatFileSize(file.size))}</td>
						<td>${getAttachmentPreviewMarkup(file)}</td>
					</tr>
				`)
				.join("")
			: `
					<tr>
						<td colspan="3">첨부파일이 없습니다.</td>
					</tr>
				`;

		return `
			<div class="side-title-wrap">
				<h3 class="side-title">출고 등록 정보</h3>
				<div class="side-actions">
					<button class="side-action-btn" data-action="register">수정내용등록</button>
					<button class="side-action-btn" data-action="ship">출고등록</button>
				</div>
			</div>
			<div class="side-table-wrap">
				<table class="side-table">
					<tbody>${shipRowsHtml}</tbody>
				</table>
			</div>
			<div class="ship-footer-grid">
				<div class="ship-footer-section ship-footer-header">
					<div class="ship-footer-title">첨부파일</div>
					<div class="ship-footer-actions">
						<label class="side-action-btn ship-attachment-label">
							첨부파일 등록
							<input class="ship-attachment-input" type="file" multiple hidden />
						</label>
					</div>
				</div>
				<div class="ship-footer-section ship-footer-preview">
					<table class="ship-attachment-table">
						<thead>
							<tr>
								<th>파일명</th>
								<th>크기</th>
								<th>미리보기</th>
							</tr>
						</thead>
						<tbody>${attachmentRowsHtml}</tbody>
					</table>
				</div>
			</div>
		`;
	}

	const { headers: selectedHeaders, values: selectedValues } = state.selectedRowData;
	const rowsHtml = selectedHeaders
		.map((header, index) => {
			const value = selectedValues[index] ?? "";
			return `
				<tr>
					<th>${escapeHtml(header)}</th>
					<td><input class="side-input" data-index="${index}" value="${escapeHtml(value)}" /></td>
				</tr>
			`;
		})
		.join("");

	return `
		<div class="side-title-wrap">
			<h3 class="side-title">선택된 행</h3>
			<div class="side-actions">
				<button class="side-action-btn" data-action="register">수정내용등록</button>
				<button class="side-action-btn" data-action="ship">출고등록</button>
			</div>
		</div>
		<div class="side-table-wrap">
			<table class="side-table">
				<tbody>${rowsHtml}</tbody>
			</table>
		</div>
		<div class="side-click-spacer" aria-hidden="true"></div>
	`;
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
					const isSelectedRow = state.selectedRowData?.rowIndex === rowIndex;
					const currentBlValue = blColumnIndex >= 0 ? String(row[blColumnIndex] ?? "").trim() : "";
					const groupKey = currentBlValue || "__EMPTY__";
					const rowBackgroundColor = isSelectedRow
						? "#d9dde2"
						: groupKey === "__EMPTY__"
							? "#ffffff"
							: (() => {
									if (!blGroupColors.has(groupKey)) {
										blGroupColors.set(groupKey, nextBlGroupIndex % 2 === 0 ? "#e9edf2" : "#ffffff");
										nextBlGroupIndex += 1;
									}
									return blGroupColors.get(groupKey);
								})();
					const rowClass = isSelectedRow ? ' class="selected-row"' : "";
					const rowStyle = ` style="background-color:${rowBackgroundColor};"`;
					return `<tr data-row-index="${rowIndex}"${rowClass}${rowStyle}>${visibleColumnIndexes
						.map((index) => {
							const value = row[index] ?? "";
							const isRemark = headers[index] === "비고";
							const className = isRemark ? ' class="remark-cell"' : "";
							const title = isRemark && value ? ` title="${value.replace(/"/g, "&quot;")}"` : "";
							const cellStyle = isSelectedRow ? ' style="background-color:#d9dde2;"' : ` style="background-color:${rowBackgroundColor};"`;
							return `<td${className}${title}${cellStyle}>${value}</td>`;
						})
						.join("")}</tr>`;
				})
				.join("")
		: `<tr><td class="empty" colspan="${visibleHeaders.length}">데이터가 없습니다.</td></tr>`;

	const tabs = state.sheetNames.map(
		(name) =>
			`<button class="sheet-tab ${name === state.activeSheetName ? "active" : ""}" data-sheet="${name}">${name}</button>`
	).join("");
	const sidePanelStateClass = !state.selectedRowData
		? "app-side-summary"
		: state.viewMode === "ship"
			? "app-side-ship"
			: "app-side-selected";

	app.innerHTML = `
		<div class="app-layout">
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
			<aside class="app-side ${sidePanelStateClass}">
				${renderSidePanelHtml()}
			</aside>
		</div>
	`;

	Array.from(app.querySelectorAll(".sheet-tab")).forEach((button) => {
		button.addEventListener("click", async () => {
			state.activeSheetName = button.dataset.sheet || "";
			await loadSheetRowsFromDatabase(state.activeSheetName, true);
			renderTable();
		});
	});

	Array.from(app.querySelectorAll("th.toggle-header")).forEach((headerCell) => {
		headerCell.addEventListener("click", () => {
			state.showAllColumns = !state.showAllColumns;
			renderTable();
		});
	});

	const tableBody = app.querySelector(".app-main table tbody");
	if (!tableBody) {
		return;
	}

	tableBody.querySelectorAll("tr.selected-row").forEach((row) => row.classList.remove("selected-row"));

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

		const rowIndex = Number(targetRow.dataset.rowIndex);
		if (Number.isNaN(rowIndex) || !rows[rowIndex]) {
			return;
		}

		Array.from(tableBody.querySelectorAll("tr.selected-row")).forEach((row) => row.classList.remove("selected-row"));
		targetRow.classList.add("selected-row");

		const fullRowValues = rows[rowIndex] ? [...rows[rowIndex]] : [];
		state.selectedRowData = {
			rowIndex,
			headers,
			values: fullRowValues
		};
		window.lastClickedRowValues = fullRowValues;
		renderTable();
	});

	Array.from(app.querySelectorAll(".side-action-btn")).forEach((button) => {
		button.addEventListener("click", () => {
			const action = button.dataset.action;
			if (action === "register") {
				state.viewMode = "edit";
				renderTable();
			} else if (action === "ship") {
				state.viewMode = "ship";
				renderTable();

				if (!state.selectedRowData) {
					window.alert("선택된 행이 없습니다.");
					return;
				}

				const { headers: selectedHeaders, values: selectedValues } = state.selectedRowData;
				const headerIndexMap = new Map(selectedHeaders.map((header, index) => [header, index]));
				const getValue = (header) => {
					const aliases = header === "부족수량(EA)" ? [header, "부족수량(SHORTAGE)/EA"] : [header];
					for (const candidate of aliases) {
						const index = headerIndexMap.get(candidate);
						if (index !== undefined) {
							return String(selectedValues[index] ?? "").trim();
						}
					}
					return "";
				};
				const parseNumericValue = (header) => {
					const rawValue = getValue(header);
					if (!rawValue) {
						return 0;
					}
					const normalized = rawValue.replace(/,/g, "").replace(/\s+/g, "");
					const parsed = Number(normalized);
					return Number.isNaN(parsed) ? 0 : parsed;
				};

				const totalEa = parseNumericValue("총재고수량(EA)");
				const availableEa = parseNumericValue("출고가능수량(EA)");
				const shortageRaw = getValue("부족수량(EA)");
				const damageEaRaw = getValue("파손수량(EA)");
				const shortage = parseNumericValue("부족수량(EA)");
				const damageEa = parseNumericValue("파손수량(EA)");
				const remarks = getValue("비고");

				const warnings = [];
				const hasShortage = shortageRaw !== "" && shortageRaw.toUpperCase() !== "EA";
				const hasDamage = damageEaRaw !== "" && damageEaRaw.toUpperCase() !== "EA";
				if (hasShortage) {
					warnings.push(`출고 시 부족수량(EA): ${shortageRaw}`);
				}
				if (hasDamage) {
					warnings.push(`출고 시 파손수량(EA): ${damageEaRaw}`);
				}
				if (remarks && (hasShortage || hasDamage)) {
					warnings.push(`출고 관련 비고: ${remarks}`);
				}

				if (hasShortage || hasDamage) {
					window.alert(`출고 등록 경고:\n${warnings.join("\n")}`);
				}
			}
		});
	});

	const attachmentInput = app.querySelector(".ship-attachment-input");
	if (attachmentInput) {
		attachmentInput.addEventListener("change", async (event) => {
			const files = Array.from(event.target.files || []);
			const attachmentInfos = await Promise.all(files.map(async (file) => {
				const attachmentInfo = {
					name: file.name,
					size: file.size,
					type: file.type || "기타",
					previewKind: "text",
					previewText: "",
					previewUrl: ""
				};

				if (file.type.startsWith("image/")) {
					attachmentInfo.previewKind = "image";
					attachmentInfo.previewUrl = URL.createObjectURL(file);
					return attachmentInfo;
				}

				const isTextLike = file.type.startsWith("text/") || /\.(txt|csv|json|md|log|js|ts|html|css|xml)$/i.test(file.name);
				if (isTextLike) {
					try {
						const text = await file.text();
						attachmentInfo.previewText = text.slice(0, 240).replace(/\s+/g, " ").trim();
					} catch (error) {
						attachmentInfo.previewText = "파일 내용을 읽을 수 없습니다.";
					}
				}

				return attachmentInfo;
			}));

			state.attachments = [...state.attachments, ...attachmentInfos];
			renderTable();
		});
	}

	app.querySelectorAll(".attachment-preview-text, .attachment-preview-image").forEach((preview) => {
		preview.addEventListener("click", (event) => {
			event.stopPropagation();
			const previewName = preview.dataset.previewName || "첨부파일";
			const previewType = preview.dataset.previewType || "text";
			const previewText = preview.dataset.previewText || "";
			const previewUrl = preview.dataset.previewUrl || "";
			const modalHtml = `
				<div class="attachment-modal-backdrop" role="presentation">
					<div class="attachment-modal" role="dialog" aria-modal="true" aria-label="첨부파일 미리보기">
						<div class="attachment-modal-header">
							<strong>${escapeHtml(previewName)}</strong>
							<button class="attachment-modal-close" type="button">닫기</button>
						</div>
						<div class="attachment-modal-body">
							${previewType === "image" && previewUrl
								? `<img src="${escapeAttribute(previewUrl)}" alt="${escapeAttribute(previewName)}" />`
								: `<pre>${escapeHtml(previewText || "미리보기 내용이 없습니다.")}</pre>`}
						</div>
					</div>
				</div>
			`;
			const existingModal = document.querySelector(".attachment-modal-backdrop");
			if (existingModal) {
				existingModal.remove();
			}
			document.body.insertAdjacentHTML("beforeend", modalHtml);
			const closeButton = document.querySelector(".attachment-modal-close");
			if (closeButton) {
				closeButton.addEventListener("click", () => {
					const modal = document.querySelector(".attachment-modal-backdrop");
					if (modal) {
						modal.remove();
					}
				});
			}
			document.querySelector(".attachment-modal-backdrop")?.addEventListener("click", (modalEvent) => {
				if (modalEvent.target.classList.contains("attachment-modal-backdrop")) {
					modalEvent.currentTarget.remove();
				}
			});
		});
	});

	Array.from(app.querySelectorAll(".side-input")).forEach((input) => {
		input.addEventListener("input", (event) => {
			const targetInput = event.target;
			const inputIndex = Number(targetInput.dataset.index);
			if (!state.selectedRowData || Number.isNaN(inputIndex)) {
				return;
			}

			const rowIndex = state.selectedRowData.rowIndex;
			if (Number.isInteger(rowIndex) && rows[rowIndex]) {
				rows[rowIndex][inputIndex] = targetInput.value;
			}

			state.selectedRowData.values[inputIndex] = targetInput.value;

			const mainTableRow = app.querySelector(`.app-main table tbody tr[data-row-index="${rowIndex}"]`);
			if (mainTableRow) {
				const visibleIndex = visibleColumnIndexes.indexOf(inputIndex);
				if (visibleIndex >= 0) {
					const cell = mainTableRow.querySelectorAll("td")[visibleIndex];
					if (cell) {
						cell.textContent = targetInput.value;
					}
				}
			}
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
