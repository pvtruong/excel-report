const JSZip = require("jszip");
const eletree = require("elementtree");
const fs = require("fs");
const Moment = require('moment-timezone');

Moment.locale("vi");
const moment = (time) => {
    return Moment.tz(time, Moment().tz());
};

// Hàm format số chuẩn Việt Nam thay thế cho numeral
const formatNumberVN = (num) => {
    return new Intl.NumberFormat('vi-VN').format(num);
};

function removeXMLInvalidChars(str, removeDiscouragedChars) {
    let regex = /((?:[\0-\x08\x0B\f\x0E-\x1F\uFFFD\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?:[^\uD800-\uDBFF]|^)[\uDC00-\uDFFF]))/g;
    str = String(str || '').replace(regex, '');

    if (removeDiscouragedChars) {
        regex = new RegExp(
            '([\\x7F-\\x84]|[\\x86-\\x9F]|[\\uFDD0-\\uFDEF]|(?:\\uD83F[\\uDFFE\\uDFFF])|(?:\\uD87F[\\uDF' +
            'FE\\uDFFF])|(?:\\uD8BF[\\uDFFE\\uDFFF])|(?:\\uD8FF[\\uDFFE\\uDFFF])|(?:\\uD93F[\\uDFFE\\uD' +
            'FFF])|(?:\\uD97F[\\uDFFE\\uDFFF])|(?:\\uD9BF[\\uDFFE\\uDFFF])|(?:\\uD9FF[\\uDFFE\\uDFFF])' +
            '|(?:\\uDA3F[\\uDFFE\\uDFFF])|(?:\\uDA7F[\\uDFFE\\uDFFF])|(?:\\uDABF[\\uDFFE\\uDFFF])|(?:\\' +
            'uDAFF[\\uDFFE\\uDFFF])|(?:\\uDB3F[\\uDFFE\\uDFFF])|(?:\\uDB7F[\\uDFFE\\uDFFF])|(?:\\uDBBF' +
            '[\\uDFFE\\uDFFF])|(?:\\uDBFF[\\uDFFE\\uDFFF])(?:[\\0-\\t\\x0B\\f\\x0E-\\u2027\\u202A-\\uD7FF\\' +
            'uE000-\\uFFFF]|[\\uD800-\\uDBFF][\\uDC00-\\uDFFF]|[\\uD800-\\uDBFF](?![\\uDC00-\\uDFFF])|' +
            '(?:[^\\uD800-\\uDBFF]|^)[\\uDC00-\\uDFFF]))', 'g');
        str = str.replace(regex, '');
    }
    return str;
}

function escapeRegExp(string) {
    return string.replace(/([.*+?^=!:${}()|\[\]\/\\])/g, "\\$1");
}

function replaceAll(string, find, replace) {
    if (!string || !find) return "";
    if (Array.isArray(replace) || (typeof replace === 'object' && replace !== null && !(replace instanceof Date))) {
        return string.replace(new RegExp(escapeRegExp(find), 'g'), "");
    }
    let replace_removed_special = replace;
    if (replace_removed_special) {
        replace_removed_special = removeXMLInvalidChars(replace_removed_special, true);
    }
    return string.replace(new RegExp(escapeRegExp(find), 'g'), replace_removed_special);
}

function isDate(val) {
    return val instanceof Date && !isNaN(val);
}

function isMatch(object, attrs) {
    const keys = Object.keys(attrs);
    const length = keys.length;
    if (object == null) return !length;
    for (let i = 0; i < length; i++) {
        const key = keys[i];
        if (attrs[key] !== object[key] || !(key in object)) return false;
    }
    return true;
}

async function fillData(zip, data, begin_row, stt_sharedString) {
    let sharedStrings = [];
    let calcChain = [];
    let refs = {};

    let addSharedStrings = function (v) {
        if (v) v = removeXMLInvalidChars(v, true);
        sharedStrings.push(v);
        return sharedStrings.length + stt_sharedString - 1;
    }

    let addCalcChain = function (calc) {
        calcChain.push(calc);
        return calcChain.length - 1;
    }

    // Obtains sharedStrings (Logic y hệt bản gốc, parse từ file có sẵn)
    const ssText = await zip.file("xl/sharedStrings.xml").async("string");
    eletree.parse(ssText).getroot().findall("si/t").forEach(function (t) {
        let ts = t.text;
        addSharedStrings(ts);
    });

    // Parse sheet1
    let sheet1 = await zip.file("xl/worksheets/sheet1.xml").async("string");
    let sheet1Doc = eletree.parse(sheet1);
    
    // Merge cells
    let mergeCells = {};
    sheet1Doc.findall("./mergeCells/mergeCell").forEach(function (mergeCell) {
        let ref = mergeCell.attrib.ref;
        if (ref) {
            let d_ref = ref.replace(":", "_") + begin_row.toString();
            ref.split(":").forEach(function (r) {
                mergeCells[r] = d_ref;
            });
            refs[d_ref] = ref;
        }
    });

    // Create table
    let table = [];
    let t_i = 0;
    let pre_row_i = 0;
    let now_row_i = 0;

    const rows = sheet1Doc.findall("./sheetData/row");

    for (let row of rows) {
        let first_row = false;
        let table_name;
        let filter;
        now_row_i = Number(row.attrib.r) + begin_row;

        let cells = row.findall("c");
        for (let c = 0; c < cells.length; c++) {
            let cell = cells[c];
            let t = cell.attrib.t;
            let v_cell = cell.find("v");
            if (v_cell) {
                let i = v_cell.text;
                let s;
                if (i) {
                    i = Number(i);
                    s = sharedStrings[i];
                }
                if (t == "s" && s) {
                    if (/{{tb:(.*)\.(.*)}}/.test(s)) {
                        first_row = true;
                        table_name = /{{tb:(.*)\.(.*)}}/.exec(s)[1];
                        let ff = /{{tb:(.*)\.(.*)}}/.exec(s)[2].split("|");
                        if (ff.length > 1) {
                            filter = "{" + ff[1] + "}";
                            try {
                                filter = eval("(" + filter + ")");
                            } catch (e) {
                                filter = { zzz: 'zzz' };
                            }
                        }
                    } else {
                        let ms = s.match(/{{([a-zA-Z0-9_]+)}}/gi);
                        if (ms && ms.length > 0) {
                            let fields = [];
                            ms.forEach(function (m) {
                                let exec = /{{([a-zA-Z0-9_]+)}}/gi.exec(m);
                                fields.push(exec[1]);
                            });

                            if (ms[0] !== s) {
                                let str = s;
                                fields.forEach(function (field) {
                                    if (data[field]) {
                                        if (typeof data[field] === 'number') {
                                            str = replaceAll(str, "{{" + field + "}}", formatNumberVN(data[field]));
                                        } else {
                                            if (isDate(data[field])) {
                                                str = replaceAll(str, "{{" + field + "}}", moment(data[field]).format('L'));
                                            } else {
                                                str = replaceAll(str, "{{" + field + "}}", data[field]);
                                            }
                                        }
                                    }
                                });
                                sharedStrings[i] = str;
                                v_cell.text = (i + stt_sharedString).toString();
                            } else {
                                let field = fields[0];
                                if (typeof data[field] === 'number') {
                                    v_cell.text = data[field].toString();
                                    delete cell.attrib["t"];
                                } else {
                                    if (isDate(data[field])) {
                                        let originDate = new Date(Date.UTC(1899, 11, 30));
                                        let v = data[field];
                                        v = new Date(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate()));
                                        v_cell.text = ((v - originDate) / (24 * 60 * 60 * 1000)).toString();
                                        delete cell.attrib["t"];
                                    } else {
                                        if (data[field] === 0) {
                                            v_cell.text = "0";
                                            delete cell.attrib["t"];
                                        } else {
                                            if (data[field]) {
                                                sharedStrings[i] = replaceAll(s, "{{" + field + "}}", data[field]);
                                                v_cell.text = (i + stt_sharedString).toString();
                                                cell.set('t', 's');
                                            }
                                        }
                                    }
                                }
                            }
                        } else {
                            v_cell.text = (i + stt_sharedString).toString();
                            cell.set('t', 's');
                        }
                    }
                }
            }
        }

        if (first_row && table_name && data[table_name]) {
            let i_r = t_i + (now_row_i - pre_row_i);
            let rows_data;
            if (!filter) {
                rows_data = data[table_name];
            } else {
                rows_data = data[table_name].filter(function (r) {
                    return isMatch(r, filter);
                });
            }

            let stt = 0;
            for (let d of rows_data) {
                stt = stt + 1;
                let current_i_r = i_r;
                t_i = i_r;
                i_r = i_r + 1;

                let rtable = new eletree.Element("row");
                rtable.set("r", current_i_r.toString());
                if (row.attrib.spans) {
                    rtable.set("spans", row.attrib.spans);
                }

                row._children.forEach(function (cell) {
                    let ctable = new eletree.Element("c");
                    ctable.set("r", cell.attrib.r.match(/[a-zA-Z]+/g)[0] + current_i_r);

                    if (cell.attrib.s) {
                        ctable.set("s", cell.attrib.s);
                    }
                    rtable.append(ctable);

                    if (cell._children.length > 0) {
                        let vtable = new eletree.Element("v");
                        ctable.append(vtable);
                        let s = cell.find("v").text;
                        vtable.text = s;

                        let t = cell.attrib.t;
                        if (t == 's' && s) {
                            s = Number(s);
                            s = sharedStrings[s];
                            if (s == '{{stt}}') {
                                vtable.text = stt.toString();
                            } else {
                                if (/{{tb:(.*)\.(.*)}}/.test(s)) {
                                    let ff = /{{tb:(.*)\.(.*)}}/.exec(s)[2].split("|");
                                    let field = ff[0];
                                    let v = d[field];

                                    if (v !== undefined && v !== null && v !== "") {
                                        if (isDate(v)) {
                                            let originDate = new Date(Date.UTC(1899, 11, 30));
                                            v = new Date(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate()));
                                            vtable.text = ((v - originDate) / (24 * 60 * 60 * 1000)).toString();
                                            delete ctable.attrib["t"];
                                        } else {
                                            if (typeof v === 'number') {
                                                vtable.text = v.toString();
                                                delete ctable.attrib["t"];
                                            } else {
                                                vtable.text = addSharedStrings(v).toString();
                                                ctable.set("t", "s");
                                            }
                                        }
                                    } else {
                                        if (v === 0) {
                                            vtable.text = "0";
                                            delete ctable.attrib["t"];
                                        } else {
                                            vtable.text = addSharedStrings("").toString();
                                            ctable.set("t", "s");
                                        }
                                    }
                                } else {
                                    if (cell.attrib.t) {
                                        ctable.set("t", cell.attrib.t);
                                    }
                                }
                            }
                        }
                    }
                });
                table.push(rtable);
            }
        } else {
            t_i = t_i + (now_row_i - pre_row_i);
            row.set("r", t_i.toString());
            row._children.forEach(function (c) {
                let oldCell = c.attrib.r;
                let newCell = oldCell.match(/[a-zA-Z]+/g)[0] + t_i;
                c.set("r", newCell);

                let r_merge = mergeCells[oldCell];
                if (r_merge) {
                    let ref = refs[r_merge];
                    ref = ref.replace(oldCell, newCell);
                    refs[r_merge] = ref;
                }

                let f = c.find("f");
                if (f) {
                    addCalcChain(c.attrib.r);
                }
            });
            table.push(row);
        }
        pre_row_i = now_row_i;
    }

    let end_row = Number(table[table.length - 1].attrib.r) + 1;

    let fn_sharedStrings = [];
    for (let str of sharedStrings) {
        str = str.toString();
        for (let key of Object.keys(data)) {
            let v = data[key];
            if (Array.isArray(v) || (typeof v === 'object' && v !== null && !isDate(v))) {
                str = replaceAll(str, "{{" + key + "}}", "");
                continue;
            }
            if (v == undefined || v == null) {
                v = "";
            }
            if (typeof v === 'number') {
                v = formatNumberVN(v);
            }
            if (isDate(v)) {
                v = moment(v).format('L');
            }
            str = replaceAll(str, "{{" + key + "}}", v);
        }

        str = str.replace(/({{[a-zA-Z0-9_]+}})/gi, '');
        let regex = /c\((.*)\)/gi;
        let exec = regex.exec(str);
        if (exec && exec[1]) {
            try {
                str = eval("(" + exec[1] + ")");
            } catch (e) { }
        }
        
        let si = new eletree.Element("si"),
            t = new eletree.Element("t");
        t.text = str.toString();
        si.append(t);
        fn_sharedStrings.push(si);
    }

    return {
        table: table,
        sharedStrings: fn_sharedStrings,
        calcChain: calcChain,
        refs: refs,
        end_row: end_row
    };
}

module.exports = async function (file_template, datas, callback, options = { timezone: 'Asia/Ho_Chi_Minh' }) {
    try {
        if (!fs.existsSync(file_template)) {
            throw new Error("Template file not exists");
        }
        Moment.tz.setDefault(options.timezone || 'Asia/Ho_Chi_Minh');

        const dataTmp = fs.readFileSync(file_template);
        // Tương đương new nodezip, nhưng sử dụng JSZip hiện đại
        const zip = await JSZip.loadAsync(dataTmp); 

        let sharedStrings = [];
        let calcChain = [];
        let refs = {};
        let table = [];

        if (!Array.isArray(datas)) {
            datas = [datas];
        }
        let begin_row = 0;

        for (let data of datas) {
            const rs = await fillData(zip, data, begin_row, sharedStrings.length);
            sharedStrings = sharedStrings.concat(rs.sharedStrings);
            calcChain = calcChain.concat(rs.calcChain);
            table = table.concat(rs.table);
            for (let k in rs.refs) {
                refs[k] = rs.refs[k];
            }
            begin_row = rs.end_row + 1;
        }

        // Save sharedStrings (Tôn trọng root gốc của Excel, không tự tạo mới)
        let ssTextRaw = await zip.file("xl/sharedStrings.xml").async("string");
        let root = eletree.parse(ssTextRaw).getroot();
        let children = root.getchildren();
        root.delSlice(0, children.length);
        
        sharedStrings.forEach(function (si) {
            root.append(si);
        });

        root.attrib.count = sharedStrings.length.toString();
        root.attrib.uniqueCount = sharedStrings.length.toString();
        zip.file("xl/sharedStrings.xml", eletree.tostring(root));

        // Save calcChain
        if (zip.file("xl/calcChain.xml")) {
            let calcTextRaw = await zip.file("xl/calcChain.xml").async("string");
            root = eletree.parse(calcTextRaw).getroot();
            children = root.getchildren();
            root.delSlice(0, children.length);
            calcChain.forEach(function (str) {
                let c = new eletree.Element("c");
                c.attrib.i = "1";
                c.attrib.r = str;
                root.append(c);
            });
            zip.file("xl/calcChain.xml", eletree.tostring(root));
        }

        // Save table
        let sheet1 = await zip.file("xl/worksheets/sheet1.xml").async("string");
        root = eletree.parse(sheet1).getroot();
        
        // SheetData
        let sheetData = root.find("sheetData");
        let sheetData_children = sheetData.getchildren();
        sheetData.delSlice(0, sheetData_children.length);
        table.forEach(function (r) {
            sheetData.append(r);
        });

        // MergeCells
        let megeCellsR = root.find("mergeCells");
        if (megeCellsR) {
            let megeCellsR_children = megeCellsR.getchildren();
            megeCellsR.delSlice(0, megeCellsR_children.length);
            for (let key in refs) {
                let r = new eletree.Element("mergeCell"); 
                r.attrib.ref = refs[key];
                megeCellsR.append(r);
            }
        }

        // Save sheet1
        zip.file("xl/worksheets/sheet1.xml", eletree.tostring(root));

        // Get binary data (Tương đương với buffer cũ của bạn)
        const buffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
        
        if (callback) return callback(null, buffer);
        return buffer;

    } catch (error) {
        if (callback) return callback(error);
        throw error;
    }
}