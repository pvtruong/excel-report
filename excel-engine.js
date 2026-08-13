const ExcelJS = require('exceljs');
const ejs = require('ejs');

// ==========================================
// PHẦN 1: CÁC HÀM HỖ TRỢ (HELPERS)
// ==========================================

function getFunctionSignature(fn) {
    try {
        let fnStr = fn.toString();
        let match = fnStr.match(/^[^(]*\(\s*([^)]*)\)/);
        let args = match ? match[1].trim() : "...";
        return `Function(${args})`;
    } catch (e) {
        return "Function()";
    }
}

function getPlainText(cell) {
    if (!cell || cell.value === null || cell.value === undefined) return '';
    if (typeof cell.value === 'string') return cell.value;
    if (cell.value.richText) return cell.value.richText.map(rt => rt.text || '').join('');
    if (cell.value.result !== undefined) return String(cell.value.result);
    return String(cell.value);
}

function cloneValue(val) {
    if (val === null || val === undefined) return val;
    if (val instanceof Date) return new Date(val);
    if (typeof val === 'object') {
        try { return JSON.parse(JSON.stringify(val)); } catch(e) { return val; }
    }
    return val;
}

function decodeRange(rangeStr) {
    let parts = rangeStr.split(':');
    let start = parts[0], end = parts[1] || parts[0];
    return {
        startCol: start.replace(/[0-9]/g, ''), startRow: parseInt(start.replace(/[^0-9]/g, '')),
        endCol: end.replace(/[0-9]/g, ''), endRow: parseInt(end.replace(/[^0-9]/g, ''))
    };
}

/**
 * [ĐÃ FIX LỖI NỀN ĐỎ] Sử dụng trực tiếp cell.style của ExcelJS.
 * Không dùng JSON.stringify để không làm hỏng cấu trúc Fill.
 */
function extractRowTemplate(worksheet, rowIdx) {
    let row = worksheet.getRow(rowIdx);
    let cells = [];
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        // Lấy trực tiếp object style do ExcelJS cấp phát
        cells.push({ col: colNumber, value: cloneValue(cell.value), style: cell.style });
    });

    let merges = [];
    if (worksheet.model && worksheet.model.merges) {
        worksheet.model.merges.forEach(rangeStr => {
            let m = decodeRange(rangeStr);
            if (m.startRow === rowIdx && m.endRow === rowIdx) {
                merges.push({ startCol: m.startCol, endCol: m.endCol });
            }
        });
    }
    return { height: row.height, hidden: row.hidden, cells: cells, merges: merges };
}

function applyRowTemplate(worksheet, rowIdx, template) {
    let row = worksheet.getRow(rowIdx);
    row.height = template.height;
    row.hidden = template.hidden;
    
    template.cells.forEach(c => {
        let cell = row.getCell(c.col);
        cell.value = cloneValue(c.value); 
        // Gán trực tiếp, ExcelJS tự động xử lý ánh xạ ID Style
        cell.style = c.style; 
    });

    template.merges.forEach(m => {
        try { worksheet.mergeCells(`${m.startCol}${rowIdx}:${m.endCol}${rowIdx}`); } catch(e) {}
    });
}

function renderCellWithEJS(cell, context) {
    let cellText = getPlainText(cell);
    if (cellText.includes('<%')) {
        try {
            let renderedValue = ejs.render(cellText, context).trim();
            if (renderedValue !== '' && !isNaN(renderedValue)) {
                cell.value = Number(renderedValue);
            } else {
                cell.value = renderedValue;
            }
        } catch (err) {
            console.error(`[LỖI EJS] Dòng ${cell.row}, Cột ${cell.col}:`, err.message);
            // Ô lỗi sẽ in chữ bình thường, không bôi đỏ
            cell.value = `LỖI TÊN BIẾN`; 
        }
    }
}


// ==========================================
// PHẦN 2: HÀM LÕI (CORE RENDER)
// ==========================================

async function renderExcelTemplate(templateBuffer, data) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(templateBuffer); 

    // ==========================================
    // TẠO SHEET "TÀI LIỆU SỐNG" (HD_DU_LIEU_DEBUG)
    // ==========================================
    const helpSheetName = "HD_DU_LIEU_DEBUG";
    let helpSheet = workbook.getWorksheet(helpSheetName);
    if (!helpSheet) helpSheet = workbook.addWorksheet(helpSheetName);
    else helpSheet.spliceRows(1, helpSheet.rowCount);

    helpSheet.columns = [
        { header: '', key: 'col1', width: 25 },
        { header: '', key: 'col2', width: 35 },
        { header: '', key: 'col3', width: 50 }
    ];

    const setHeaderStyle = (cell, isTitle = false) => {
        cell.font = { 
            name: 'Arial', bold: true, size: isTitle ? 14 : 11, 
            color: { argb: isTitle ? 'FF0070C0' : 'FF000000' } // Cả 2 title đều dùng màu Xanh Blue
        };
        if (!isTitle) {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } };
        }
    };

    // --- PHẦN HƯỚNG DẪN CÚ PHÁP ---
    helpSheet.addRow(['--- HƯỚNG DẪN CÚ PHÁP (CHEAT SHEET) ---']);
    helpSheet.mergeCells('A1:C1');
    setHeaderStyle(helpSheet.getCell('A1'), true);

    let rowHeader1 = helpSheet.rowCount + 1;
    helpSheet.addRow(['Mục đích', 'Cú pháp mẫu', 'Ghi chú / Cách dùng']);
    ['A', 'B', 'C'].forEach(col => setHeaderStyle(helpSheet.getCell(`${col}${rowHeader1}`)));

    helpSheet.addRow(['In giá trị', '<%= ten_bien %>', 'Dùng cho văn bản, tên khách hàng, số...']);
    helpSheet.addRow(['Vòng lặp cơ bản', '[[FOR:ten_mang]]', 'Đặt ở Cột A. Tự gọi biến bằng `<%= ten_mang.thuoc_tinh %>`']);
    helpSheet.addRow(['Dòng Mẹ (Master)', '[[MASTER:ten_mang]]', 'Đặt ở Cột A. Tự gọi biến `<%= ten_mang.thuoc_tinh %>`']);
    helpSheet.addRow(['Dòng Con (Detail)', '[[DETAIL:mang_con]]', 'Đặt ngay dưới dòng Mẹ. Tự gọi biến `<%= mang_con.thuoc_tinh %>`']);
    helpSheet.addRow(['In STT dòng', '<%= ten_mang_index + 1 %>', 'Hệ thống tự tạo ra biến đuôi _index bắt đầu từ số 0']);
    helpSheet.addRow([]); 

    // --- PHẦN DANH SÁCH BIẾN DỮ LIỆU ---
    let rowTitle2 = helpSheet.rowCount + 1;
    helpSheet.addRow(['--- DANH SÁCH BIẾN DỮ LIỆU ĐƯỢC CẤP TỪ HỆ THỐNG ---']);
    helpSheet.mergeCells(`A${rowTitle2}:C${rowTitle2}`);
    setHeaderStyle(helpSheet.getCell(`A${rowTitle2}`), true);

    let rowHeader2 = helpSheet.rowCount + 1;
    helpSheet.addRow(['Tên Biến (Key)', 'Kiểu Dữ Liệu', 'Dữ liệu thô / Cấu trúc']);
    ['A', 'B', 'C'].forEach(col => setHeaderStyle(helpSheet.getCell(`${col}${rowHeader2}`)));

    const jsonSafeData = {};

    Object.keys(data).forEach(key => {
        let val = data[key];
        let isObject = typeof val === 'object' && val !== null;
        let isArray = Array.isArray(val);
        let isFunction = typeof val === 'function';

        let type = isArray ? 'Mảng (Array)' : typeof val;
        let displayVal = '';

        if (isFunction) {
            displayVal = getFunctionSignature(val); 
            type = 'Hàm Xử Lý (Function)';
        } else {
            jsonSafeData[key] = val; 
            displayVal = isObject ? '(Xem cấu trúc JSON bên dưới)' : String(val);
        }

        helpSheet.addRow([key, type, displayVal]);
    });

    // --- PHẦN RAW JSON ---
    helpSheet.addRow([]);
    let rowJson = helpSheet.rowCount + 1;
    helpSheet.mergeCells(`A${rowJson}:C${rowJson + 30}`); 
    let cellJson = helpSheet.getCell(`A${rowJson}`);
    cellJson.value = JSON.stringify(jsonSafeData, null, 2);
    cellJson.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };


    // ==========================================
    // LOGIC RENDER EJS CHÍNH (XỬ LÝ DỮ LIỆU)
    // ==========================================
    const globalContext = { ...data };

    workbook.eachSheet((worksheet) => {
        if (worksheet.name === helpSheetName) return; 

        let blocks = [];
        let r = 1;
        
        while (r <= worksheet.rowCount) {
            let cell = worksheet.getRow(r).getCell(1);
            let val = getPlainText(cell);
            
            if (val.includes('[[MASTER:')) {
                let masterMatch = val.match(/\[\[MASTER:([a-zA-Z0-9_]+)\]\]/);
                if (masterMatch) {
                    let nextCell = worksheet.getRow(r+1).getCell(1);
                    let nextVal = getPlainText(nextCell);
                    let detailMatch = nextVal.match(/\[\[DETAIL:([a-zA-Z0-9_]+)\]\]/);

                    if (detailMatch) {
                        blocks.push({
                            type: 'MASTER_DETAIL', masterRowIdx: r, detailRowIdx: r+1,
                            masterVar: masterMatch[1], detailVar: detailMatch[1]
                        });
                        r += 2; continue;
                    }
                }
            } else if (val.includes('[[FOR:')) {
                let match = val.match(/\[\[FOR:([a-zA-Z0-9_]+)\]\]/);
                if (match) {
                    blocks.push({ type: 'FOR', rowIdx: r, arrayVar: match[1] });
                }
            }
            r++;
        }

        for (let b = blocks.length - 1; b >= 0; b--) {
            let block = blocks[b];

            if (block.type === 'MASTER_DETAIL') {
                let mTemplate = extractRowTemplate(worksheet, block.masterRowIdx);
                let dTemplate = extractRowTemplate(worksheet, block.detailRowIdx);
                worksheet.spliceRows(block.masterRowIdx, 2); 

                let mDataArray = data[block.masterVar] || [];
                let totalRows = 0;
                mDataArray.forEach(m => { totalRows += 1 + (m[block.detailVar] || []).length; });

                if (totalRows > 0) {
                    worksheet.spliceRows(block.masterRowIdx, 0, ...new Array(totalRows).fill([]));
                    let currIdx = block.masterRowIdx;
                    
                    mDataArray.forEach((mItem, mIndex) => {
                        applyRowTemplate(worksheet, currIdx, mTemplate);
                        let cellM = worksheet.getRow(currIdx).getCell(1);
                        let textM = getPlainText(cellM);
                        if (textM.includes('[[MASTER:')) {
                            cellM.value = textM.replace(/\[\[MASTER:[a-zA-Z0-9_]+\]\]/, '').trim();
                        }
                        
                        let rowContext = { ...globalContext };
                        rowContext[block.masterVar] = mItem;
                        rowContext[`${block.masterVar}_index`] = mIndex;

                        worksheet.getRow(currIdx).eachCell((c) => renderCellWithEJS(c, rowContext));
                        currIdx++;

                        let dDataArray = mItem[block.detailVar] || [];
                        dDataArray.forEach((dItem, dIndex) => {
                            applyRowTemplate(worksheet, currIdx, dTemplate);
                            let cellD = worksheet.getRow(currIdx).getCell(1);
                            let textD = getPlainText(cellD);
                            if (textD.includes('[[DETAIL:')) {
                                cellD.value = textD.replace(/\[\[DETAIL:[a-zA-Z0-9_]+\]\]/, '').trim();
                            }
                            
                            let detailContext = { ...rowContext };
                            detailContext[block.detailVar] = dItem;
                            detailContext[`${block.detailVar}_index`] = dIndex;

                            worksheet.getRow(currIdx).eachCell((c) => renderCellWithEJS(c, detailContext));
                            currIdx++;
                        });
                    });
                }
            } 
            else if (block.type === 'FOR') {
                let template = extractRowTemplate(worksheet, block.rowIdx);
                worksheet.spliceRows(block.rowIdx, 1);

                let dataArray = data[block.arrayVar] || [];
                if (dataArray.length > 0) {
                    worksheet.spliceRows(block.rowIdx, 0, ...new Array(dataArray.length).fill([]));
                    let currIdx = block.rowIdx;
                    
                    dataArray.forEach((item, i) => {
                        applyRowTemplate(worksheet, currIdx, template);
                        let cell = worksheet.getRow(currIdx).getCell(1);
                        let textFor = getPlainText(cell);
                        if (textFor.includes('[[FOR:')) {
                            cell.value = textFor.replace(/\[\[FOR:[a-zA-Z0-9_]+\]\]/, '').trim();
                        }
                        
                        let forContext = { ...globalContext };
                        forContext[block.arrayVar] = item;
                        forContext[`${block.arrayVar}_index`] = i;

                        worksheet.getRow(currIdx).eachCell((c) => renderCellWithEJS(c, forContext));
                        currIdx++;
                    });
                }
            }
        }

        // 3. Render các ô tĩnh không nằm trong vòng lặp
        worksheet.eachRow((row) => {
            row.eachCell((cell) => renderCellWithEJS(cell, globalContext));
        });
    });

    return await workbook.xlsx.writeBuffer();
}

module.exports = renderExcelTemplate;