const nodezip = require("node-zip");
const eletree = require("elementtree");
const fs = require("fs");
const underscore = require("underscore");
const numeral = require('numeral');
const async = require("async");
// load a locale
numeral.register('locale', 'vn', {
    delimiters: {
        thousands: '.',
        decimal: ','
    },
    abbreviations: {
        thousand: 'ngàn',
        million: 'triệu',
        billion: 'tỷ',
        trillion: 'ngàn tỷ'
    },
    ordinal : function (number) {
        return number === 1 ? 'er' : 'ème';
    },
    currency: {
        symbol: 'VND'
    }
});

// switch between locales
numeral.locale('vn');
//
const Moment = require('moment-timezone');
Moment.locale("vi");
const moment = (time)=>{
	return Moment.tz(time,Moment().tz());
}
//
/**
 * Removes invalid XML characters from a string
 * @param {string} str - a string containing potentially invalid XML characters (non-UTF8 characters, STX, EOX etc)
 * @param {boolean} removeDiscouragedChars - should it remove discouraged but valid XML characters
 * @return {string} a sanitized string stripped of invalid XML characters
 */
 function removeXMLInvalidChars(str, removeDiscouragedChars) {
    // remove everything forbidden by XML 1.0 specifications, plus the unicode replacement character U+FFFD
    let regex = /((?:[\0-\x08\x0B\f\x0E-\x1F\uFFFD\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?:[^\uD800-\uDBFF]|^)[\uDC00-\uDFFF]))/g;
    // ensure we have a string
    str = String(str || '').replace(regex, '');

    if (removeDiscouragedChars) {

        // remove everything discouraged by XML 1.0 specifications
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
function escapeRegExp(string) {
	return string.replace(/([.*+?^=!:${}()|\[\]\/\\])/g, "\\$1");
}
function replaceAll(string, find, replace) {

  if(!string || !find) return "";
  if(underscore.isArray(replace) || underscore.isObject(replace)){
	  return string.replace(new RegExp(escapeRegExp(find), 'g'), "");
  }

  let replace_removed_special = replace;
  if(replace_removed_special){
	replace_removed_special = removeXMLInvalidChars(replace_removed_special,true)
  }
  return string.replace(new RegExp(escapeRegExp(find), 'g'), replace_removed_special);
}
function fillData(zip,data,begin_row,stt_sharedString,callback){
	let sharedStrings = [];
	let calcChain = [];
	let refs={};
	let addSharedStrings = function(v){
		if(v) v = removeXMLInvalidChars(v,true);
		sharedStrings.push(v);
		return sharedStrings.length + stt_sharedString -1;
	}
	let addCalcChain = function(calc){
		calcChain.push(calc);
		return calcChain.length-1;
	}
	//obtains sharedStrings
	eletree.parse(zip.files["xl/sharedStrings.xml"].asText()).getroot().findall("si/t").forEach(function(t){
		let ts = t.text;
		addSharedStrings(ts);
	});
	//parse sheet1
	let sheet1 = zip.files["xl/worksheets/sheet1.xml"].asText();
	//merge cells
	let mergeCells ={};

	eletree.parse(sheet1).findall("./mergeCells/mergeCell").forEach(function(mergeCell){
		let ref = mergeCell.attrib.ref
		if(ref){
			let d_ref =ref.replace(":","_") + begin_row.toString();
			ref.split(":").forEach(function(r){
				mergeCells[r] = d_ref;
			})
			refs[d_ref] = ref;
		}
	});
	//create table
	let table =[];
	let t_i =0;
	let pre_row_i=0;
	let now_row_i=0;
	async.eachSeries(eletree.parse(sheet1).findall("./sheetData/row")
	,function(row,callback){
		let first_row=false;
		let table_name;
		let filter;
		now_row_i = Number(row.attrib.r) + begin_row;
		//identify first row and name of table
		let cells = row.findall("c");
		for(let c =0;c<cells.length;c++){
			let cell = cells[c];
			let t = cell.attrib.t;
			let v_cell = cell.find("v");
			if(v_cell){
				let i = v_cell.text;
				let s;
				if(i){
					i = Number(i);
					s = sharedStrings[i];
				}
				if(t=="s" && s){
					if(/{{tb:(.*)\.(.*)}}/.test(s)){
						//first row
						first_row =true;
						table_name = /{{tb:(.*)\.(.*)}}/.exec(s)[1];
						let ff = /{{tb:(.*)\.(.*)}}/.exec(s)[2].split("|");
						if(ff.length>1){
							filter = "{" + ff[1] + "}";
							try{
								filter = eval("(" + filter + ")");
							}catch(e){
								console.log("can't parse JSON: ",filter);
								filter ={zzz:'zzz'} ;
							}

						}
						//break;
					}else{
						let ms = s.match(/{{([a-zA-Z0-9_]+)}}/gi)
						if(ms && ms.length>0){
							let fields =[];
							ms.forEach(function(m){
								let exec =  /{{([a-zA-Z0-9_]+)}}/gi.exec(m);
								fields.push(exec[1]);
							})
							//string
							if(ms[0]!==s){
								let str =s;
								fields.forEach(function(field){

									if(data[field]){
										if(underscore.isNumber(data[field])){
											str = replaceAll(str,"{{" + field + "}}",numeral(data[field]).format());
										}else{
											if(underscore.isDate(data[field])){
												str = replaceAll(str,"{{" + field + "}}",moment(data[field]).format('L'));
											}else{
												str = replaceAll(str,"{{" + field + "}}",data[field]);
											}

										}

									}

								})
								sharedStrings[i] = str;
								v_cell.text = (i+stt_sharedString).toString();

							}else{
								field = fields[0];
								//number
								if(underscore.isNumber(data[field])){
									v_cell.text = data[field].toString();
									delete cell.attrib["t"];
								}else{
									//date
									if(underscore.isDate(data[field])){
										let originDate = new Date(Date.UTC(1899,11,30));
										let v = data[field];
										v = new Date(Date.UTC(v.getFullYear(),v.getMonth(),v.getDate()));

										v_cell.text = (v - originDate) / (24 * 60 * 60 * 1000);
										//cell.set('t','');
										delete cell.attrib["t"];
									}else{
										if(data[field]===0){
											v_cell.text = "0";
											//cell.set('t','');
											delete cell.attrib["t"];
										}else{
											if(data[field]){
												sharedStrings[i] = replaceAll(s,"{{" + field + "}}",data[field]);
												v_cell.text = (i+stt_sharedString).toString();
												cell.set('t','s');
											}
										}

									}
								}
							}

						}else{
							v_cell.text = (i+stt_sharedString).toString();
							cell.set('t','s');
						}
					}
				}
			}
		};
		async.series({
			create_rows:(callback1)=>{
				if(first_row && table_name && data[table_name]){
					let i_r = t_i + (now_row_i - pre_row_i);
					let rows_data;
					if(!filter){
						rows_data = data[table_name];
					}else{
						rows_data = underscore.filter(data[table_name],function(r){
							return underscore.isMatch(r,filter);
						})
					}
					let stt =0

					for(let r of rows_data){
						stt = stt+1;
						r.stt = stt;
						r.i_r = i_r;
						//
						t_i =i_r;
						i_r=i_r + 1;

					}
					//console.log("begin fill rows",new Date());
					async.mapSeries(rows_data,(d,cb)=>{
						setImmediate(()=>{
							let stt = d.stt;
							let i_r = d.i_r;
							let rtable =new eletree.Element("row");
							rtable.set("r",i_r)
							if(row.attrib.spans){
								rtable.set("spans",row.attrib.spans);
							}
							row._children.forEach(function(cell){
								let ctable =  new eletree.Element("c");
								//ctable.set("r",cell.attrib.r.substring(0,1) + i_r);

								ctable.set("r",cell.attrib.r.match(/[a-zA-Z]+/g)[0] + i_r);
								

								if(cell.attrib.s){
									ctable.set("s",cell.attrib.s);
								}
								rtable.append(ctable);
								//value
								if(cell._children.length>0){
									let vtable = new eletree.Element("v");
									ctable.append(vtable);
									let s = cell.find("v").text;

									vtable.text = s;
									let t =  cell.attrib.t;
									if(t=='s' && s){
										s = Number(s);
										s = sharedStrings[s];
										if(s=='{{stt}}'){
											vtable.text = stt;
										}else{
											if(/{{tb:(.*)\.(.*)}}/.test(s)){
												let ff = /{{tb:(.*)\.(.*)}}/.exec(s)[2].split("|");
												let field = ff[0]
												let v = d[field];

												if(v){
													if(v && underscore.isDate(v)){
														let originDate = new Date(Date.UTC(1899,11,30));
														v = new Date(Date.UTC(v.getFullYear(),v.getMonth(),v.getDate()));
														vtable.text = (v - originDate) / (24 * 60 * 60 * 1000);
														delete ctable.attrib["t"];
													}else{
														if(underscore.isNumber(v)){
															vtable.text = v.toString();
															//ctable.set('t','');
															delete ctable.attrib["t"];
														}else{
															vtable.text = addSharedStrings(v);
															ctable.set("t","s");
														}
													}
												}else{
													if(v==0){
														vtable.text = "0";
														//ctable.set('t','');
														delete ctable.attrib["t"];
													}else{
														vtable.text = addSharedStrings("");
														ctable.set("t","s");
													}

												}

											}else{
												if(cell.attrib.t){
													ctable.set("t",cell.attrib.t);
												}
											}
										}

									}
								}
							});
							table.push(rtable);
							cb();
						})
					},()=>{
						callback1();
					})
				}else{
					callback1();
				}
			},
			create_others:(callback)=>{
				setImmediate(()=>{
					if(!(first_row && table_name && data[table_name])){
						t_i = t_i + (now_row_i - pre_row_i);
						row.set("r",t_i);
						row._children.forEach(function(c){
							let oldCell = c.attrib.r;
							//let newCell = oldCell.substring(0,1) + t_i;
							let newCell = oldCell.match(/[a-zA-Z]+/g)[0] + t_i;
							c.set("r",newCell);
							//merge
							let r_merge = mergeCells[oldCell];
							if(r_merge){
								let ref = refs[r_merge];
								ref = ref.replace(oldCell,newCell);
								refs[r_merge] = ref;
							}
							//function
							let f = c.find("f");
							if(f){
								addCalcChain(c.attrib.r)
							}
						});
						table.push(row);
					}
					callback();
				})
			}
		},function(e,rs){
			pre_row_i = now_row_i;
			callback()
		});
	},function(e,rs){
		let end_row = Number(table[table.length-1].attrib.r) + 1;
		//create sharedStrings
		//console.log("begin create sharestrings",new Date());
		let fn_sharedStrings =[];
		async.mapSeries(sharedStrings,function(str,cb){
			str = str.toString();
			async.mapSeries(underscore.keys(data),function(key,callback){
				let v = data[key];
				if(underscore.isArray(v) || underscore.isObject(v)){
					str = replaceAll(str,"{{" + key + "}}","");
					return callback();
				}
				if(v==undefined || v==null){
					v ="";
				}
				if(underscore.isNumber(v)){
					v = numeral(v).format();
				}
				if(underscore.isDate(v)){
					v = moment(v).format('L');
				}
				str = replaceAll(str,"{{" + key + "}}",v);

				callback();
			},function(e,rs){
				str = str.replace(/({{[a-zA-Z0-9_]+}})/gi,'');
				let regex =/c\((.*)\)/gi
				let exec = regex.exec(str)
				if(exec && exec[1]){
					try{
						str = eval("(" + exec[1] +  ")");
					}catch(e){

					}
				}
				let si = new eletree.Element("si"),
					t  = new eletree.Element("t");
				t.text = str;
				si.append(t);
				fn_sharedStrings.push(si);
				cb();
			});
		},function(e,rs){
			//console.log("end create sharedStrings",new Date());
			//result
			callback(null,{
				table:table,
				sharedStrings:fn_sharedStrings,
				calcChain:calcChain,
				refs:refs,
				end_row:end_row

			})
		});
	})
}
module.exports = function(file_template,datas,callback,options={timezone:'Asia/Ho_Chi_Minh'}){
	//check exists of template file
	if(!fs.existsSync(file_template)){
		return callback(new Error("Template file not exists"));
	}
	Moment.tz.setDefault(options.timezone || 'Asia/Ho_Chi_Minh');
	//read template file
	fs.readFile(file_template,function(error,dataTmp){
		if(error) return callback(error);
		//unzip template. a xlsx file is a zip file
		const zip = new nodezip(dataTmp, {base64: false, checkCRC32: true});
		let sharedStrings = [];
		let calcChain = [];
		let refs={};
		let table =[];
		// fill data for each data
		if(!underscore.isArray(datas)){
			//console.log("create array");
			datas =[datas];
		}
		let begin_row =0;
		//console.log("begin fill data",new Date())
		async.mapSeries(datas,(data,callback)=>{

			fillData(zip,data,begin_row,sharedStrings.length,function(e,rs){
				if(e) return callback(e);
				sharedStrings = sharedStrings.concat(rs.sharedStrings);
				calcChain = calcChain.concat(rs.calcChain);
				table = table.concat(rs.table);
				for(let k in rs.refs){
					refs[k] = rs.refs[k];
				}
				begin_row = rs.end_row + 1;
				callback();
			})
		},(e,rs)=>{
			//console.log("begin save data",new Date())
			//save sharedStrings
			let root = eletree.parse(zip.files["xl/sharedStrings.xml"].asText()).getroot(),
				children = root.getchildren();
			root.delSlice(0, children.length);
			sharedStrings.forEach(function(si) {
				root.append(si);
			});

			root.attrib.count = sharedStrings.length;
			root.attrib.uniqueCount = sharedStrings.length;
			zip.file("xl/sharedStrings.xml",eletree.tostring(root))
			//save calcChain
			if(zip.files["xl/calcChain.xml"]){
				root = eletree.parse(zip.files["xl/calcChain.xml"].asText()).getroot(),
				children = root.getchildren();
				root.delSlice(0, children.length);
				calcChain.forEach(function(str) {
					let c = new eletree.Element("c");
					c.attrib.i =1;
					c.attrib.r = str;
					root.append(c);
				});
				zip.file("xl/calcChain.xml",eletree.tostring(root))
			}
			//save table
			let sheet1 = zip.files["xl/worksheets/sheet1.xml"].asText();
			root = eletree.parse(sheet1).getroot();
			//sheetData
			let sheetData = root.find("sheetData");
			let sheetData_children = sheetData.getchildren();
			sheetData.delSlice(0, sheetData_children.length);
			table.forEach(function(r){
				sheetData.append(r);
			});
			//mergeCells
			let megeCellsR = root.find("mergeCells");
			if(megeCellsR){
				let megeCellsR_children = megeCellsR.getchildren();
				megeCellsR.delSlice(0, megeCellsR_children.length);
				for(let key in refs){
					let r =new eletree.Element("mergeCell ");
					r.attrib.ref =refs[key] ;
					megeCellsR.append(r);
				}
			}
			//console.log("begin zip data",new Date())
			//save sheet1
			zip.file("xl/worksheets/sheet1.xml",eletree.tostring(root));
			// Get binary data
			const result = zip.generate({base64: false, checkCRC32: true});
			//console.log("create binary data",new Date())
			callback(null,result);
		});



	});


}
