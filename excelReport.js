const fs = require("fs");
const util = require("util");
const nodezip = require("node-zip");
const eletree = require("elementtree");
const numeral = require('numeral');
const Moment = require('moment-timezone');
const readFileAsync = util.promisify(fs.readFile);

// Import Engine mới (Thay thế ejsexcel)
const renderExcelTemplate = require("./excel-engine");

// Import code cũ (Legacy)
const legacyExporter = require("./excel-legacy");

/**
 * Hàm kiểm tra loại Template (Giữ nguyên của bạn)
 */
function detectTemplateType(buffer) {
    try {
        const zip = new nodezip(buffer, { base64: false, checkCRC32: true });
        if (zip.files["xl/sharedStrings.xml"]) {
            const text = zip.files["xl/sharedStrings.xml"].asText();
            if (/{{tb:/.test(text) || /{{.*?}}/.test(text)) return 'LEGACY';
        }
        if (zip.files["xl/worksheets/sheet1.xml"]) {
            const sheet = zip.files["xl/worksheets/sheet1.xml"].asText();
            if (/{{tb:/.test(sheet) || /{{.*?}}/.test(sheet)) return 'LEGACY';
        }
        return 'NEW_ENGINE';
    } catch (e) {
        return 'LEGACY';
    }
}

/**
 * CORE LOGIC
 */
async function processExport(templatePath, data, options) {
    if (!fs.existsSync(templatePath)) {
        throw new Error("Template file not exists: " + templatePath);
    }

    // 1. Đọc file
    const templateBuf = await readFileAsync(templatePath);

    // 2. Detect Engine
    const engineType = detectTemplateType(templateBuf);
    
    // 3. Routing
    if (engineType === 'LEGACY') {
        // --- LEGACY MODE (Code cũ) ---
        return new Promise((resolve, reject) => {
            legacyExporter(templatePath, data, (err, resultBuffer) => {
                if (err) return reject(err);
                if (Buffer.isBuffer(resultBuffer)) {
                    resolve(resultBuffer);
                } else {
                    resolve(Buffer.from(resultBuffer, 'binary'));
                }
            }, options);
        });

    } else {
        // --- NEW ENGINE MODE (ExcelJS + EJS) ---
        
        // B1: Trích xuất Data
        // Engine mới không cần mảng cho từng sheet như ejsexcel. Nó dùng 1 cục data chung.
        // Nếu dự án của bạn đang truyền mảng (do thói quen cũ), ta lấy object đầu tiên làm data.
        let finalData = data;
        if (Array.isArray(data)) {
            finalData = data[0] || {};
        }

        // B2: Nhúng helper moment (và numeral nếu file excel vẫn xài text)
        const momentFn = (time) => {
            return time ? Moment.tz(time, options.timezone) : Moment().tz(options.timezone);
        };
        
        const dataWithHelpers = {
            ...finalData,
            numeral: numeral, // Giữ lại cho tương thích ngược nếu có lỡ dùng <%~ numeral(...) %>
            moment: momentFn
        };

        // B3: Gọi Engine mới truyền thẳng Buffer (RAM) thay vì file path
        // Engine sẽ trả về 1 resultBuf (Node Buffer)
        const resultBuf = await renderExcelTemplate(templateBuf, dataWithHelpers);
        
        return resultBuf;
    }
}

/**
 * MAIN EXPORT FUNCTION
 */
module.exports = function (templatePath, data, arg3, arg4) {
    let callback = null;
    let options = { timezone: 'Asia/Ho_Chi_Minh' };

    if (typeof arg3 === 'function') {
        callback = arg3;
        if (typeof arg4 === 'object') options = { ...options, ...arg4 };
    } else if (typeof arg3 === 'object') {
        options = { ...options, ...arg3 };
        if (typeof arg4 === 'function') callback = arg4;
    }

    if (!callback) {
        return processExport(templatePath, data, options);
    }

    processExport(templatePath, data, options)
        .then(buffer => callback(null, buffer))
        .catch(err => {
            console.error("[excelReport Error]", err);
            callback(err);
        });
};