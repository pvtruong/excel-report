var nodezip = require("node-zip");
var eletree = require("elementtree");
var fs = require("fs");
var underscore = require("underscore");
function escapeRegExp(string) {
	return string.replace(/([.*+?^=!:${}()|\[\]\/\\])/g, "\\$1");
}
function replaceAll(string, find, replace) {
  if(!string || !find) return ""
  return string.replace(new RegExp(escapeRegExp(find), 'g'), replace);
}
module.exports = function(file_template,data,callback){
	var sharedStrings = []
	var calcChain = []
	var addSharedStrings = function(v){
		sharedStrings.push(v);
		return sharedStrings.length-1;
	}
	var addCalcChain = function(calc){
		calcChain.push(calc);
		return calcChain.length-1;
	}
	if(!fs.existsSync(file_template)){
		return callback(new Error("Template file not exists"));
	}
	//read file
	fs.readFile(file_template,function(error,dataTmp){
		if(error) return callback(error);
		// Create a template
		var zip = new nodezip(dataTmp, {base64: false, checkCRC32: true});
		//obtains sharedStrings
		eletree.parse(zip.files["xl/sharedStrings.xml"].asText()).getroot().findall("si/t").forEach(function(t){
			var ts = t.text;
			addSharedStrings(ts);
		});
		//parse sheet1
		var sheet1 = zip.files["xl/worksheets/sheet1.xml"].asText();
		//merge cells
		var mergeCells ={}
		var refs={}
		eletree.parse(sheet1).findall("./mergeCells/mergeCell").forEach(function(mergeCell){
			var ref = mergeCell.attrib.ref
			if(ref){
				var d_ref =ref.replace(":","_");
				ref.split(":").forEach(function(r){
					mergeCells[r] = d_ref;
				})
				refs[d_ref] = ref
			}
		});
		//create table
		var table =[]
		var t_i =0;
		var pre_row_i=0;
		var now_row_i=0;
		eletree.parse(sheet1).findall("./sheetData/row").forEach(function(row){
			var first_row=false;
			var table_name;
			var filter;
			now_row_i = Number(row.attrib.r);
			//identify first row and name of table
			var cells = row.findall("c");
			for(var c =0;c<cells.length;c++){
				var cell = cells[c];
				var t = cell.attrib.t;
				var v = cell.find("v");
				if(v){
					var i = v.text;
					if(t=="s" && i){
						i = Number(i);
						var s = sharedStrings[i];
						if(/{{tb:(.*)\.(.*)}}/.test(s)){
							//first row
							first_row =true;
							table_name = /{{tb:(.*)\.(.*)}}/.exec(s)[1];
							var ff = /{{tb:(.*)\.(.*)}}/.exec(s)[2].split("|");
							if(ff.length>1){
								filter = "{" + ff[1] + "}"
								try{
									filter = eval("(" + filter + ")")
								}catch(e){
									console.log("can't parse JSON: ",filter)
									filter ={zzz:'zzz'} //
								}
								
							}
							//break;
						}else{
							var ms = s.match(/{{([a-zA-Z0-9_]+)}}/gi)
							if(ms && ms.length>0){
								var fields =[]
								ms.forEach(function(m){
									var exec =  /{{([a-zA-Z0-9_]+)}}/gi.exec(m);
									fields.push(exec[1])
								})
								//string
								if(ms[0]!==s){
									var str =s
									fields.forEach(function(field){
										if(data[field]){
											str = replaceAll(str,"{{" + field + "}}",data[field]);
										}
										
									})
									sharedStrings[i] = str
								}else{
									field = fields[0]
									//number
									if(underscore.isNumber(data[field])){
										v.text = data[field];
										cell.set('t','');
									}else{
										//date
										if(underscore.isDate(data[field])){
											var originDate = new Date(Date.UTC(1899,11,30));
											v.text = (data[field] - originDate) / (24 * 60 * 60 * 1000);
											cell.set('t','');
										}else{
											if(data[field]){
												sharedStrings[i] = replaceAll(s,"{{" + field + "}}",data[field]);
											}
										}
									}
								}
								
							}
						}
					}
				}
			};
			//create rows
			if(first_row && table_name && data[table_name]){
				var i_r = t_i + (now_row_i - pre_row_i)
				var rows_data;
				if(!filter){
					rows_data = data[table_name]
				}else{
					rows_data = underscore.filter(data[table_name],function(r){
						return underscore.isMatch(r,filter);
					})
				}
				var stt =0
				rows_data.forEach(function(d){
					stt = stt+1;
					var rtable =new eletree.Element("row");
					rtable.set("r",i_r)
					if(row.attrib.spans){
						rtable.set("spans",row.attrib.spans);
					}
					
					row._children.forEach(function(cell){
						
						var ctable =  new eletree.Element("c");
						ctable.set("r",cell.attrib.r.substring(0,1) + i_r);
						
						if(cell.attrib.s){
							ctable.set("s",cell.attrib.s);
						}
						rtable.append(ctable);
						//value
						if(cell._children.length>0){
							var vtable = new eletree.Element("v");
							ctable.append(vtable);
							var s = cell.find("v").text;
							vtable.text = s;
							var t =  cell.attrib.t;
							if(t=='s' && s){
								s = Number(s);
								s = sharedStrings[s];
								if(s=='{{stt}}'){
									vtable.text = stt;
								}else{
									if(/{{tb:(.*)\.(.*)}}/.test(s)){
										var ff = /{{tb:(.*)\.(.*)}}/.exec(s)[2].split("|");
										var field = ff[0]
										var v = d[field];
										if(v){
											if(underscore.isDate(v)){
												var originDate = new Date(Date.UTC(1899,11,30));
												vtable.text = (v - originDate) / (24 * 60 * 60 * 1000);
											}else{
												if(underscore.isNumber(v)){
													vtable.text = v;
												}else{
													vtable.text = addSharedStrings(v);
													ctable.set("t","s");
												}
											}
										}else{
											vtable.text = addSharedStrings("");
											ctable.set("t","s");
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
					table.push(rtable)
					t_i =i_r;
					i_r=i_r + 1;
				});
				
			}else{
				t_i = t_i + (now_row_i - pre_row_i)
				row.set("r",t_i);
				row._children.forEach(function(c){
					var oldCell = c.attrib.r;
					var newCell = oldCell.substring(0,1) + t_i;
					c.set("r",newCell);
					//merge
					var r_merge = mergeCells[oldCell]; 
					if(r_merge){
						var ref = refs[r_merge]
						ref = ref.replace(oldCell,newCell);
						refs[r_merge] = ref;
					}
					//function
					var f = c.find("f");
					if(f){
						addCalcChain(c.attrib.r)
					}
				});
				table.push(row);

			}
			pre_row_i = now_row_i;
			
		});
		//save sharedStrings
		var root = eletree.parse(zip.files["xl/sharedStrings.xml"].asText()).getroot(),
			children = root.getchildren();
		
		root.delSlice(0, children.length);
		sharedStrings.forEach(function(str) {
			str = str.toString()
			for(var key in data){
				var v = data[key];
				if(v==undefined || v==null){
					v ="";
				}
				str = replaceAll(str,"{{" + key + "}}",v);
				
			}
			//replace by '' if don't have value
			str = str.replace(/({{[a-zA-Z0-9_]+}})/gi,'');
			var regex =/c\((.*)\)/gi 
			var exec = regex.exec(str)
			if(exec && exec[1]){
				try{
					str = eval("(" + exec[1] +  ")")
				}catch(e){
					
				}
			}
			
			
			var si = new eletree.Element("si"),
				t  = new eletree.Element("t");
			t.text = str;
			si.append(t);
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
				var c = new eletree.Element("c")
				c.attrib.i =1
				c.attrib.r = str;
				root.append(c);
			});
			zip.file("xl/calcChain.xml",eletree.tostring(root))
		}
		//save table
		root = eletree.parse(sheet1).getroot()
		//sheetData
		var sheetData = root.find("sheetData");
		var sheetData_children = sheetData.getchildren();
		sheetData.delSlice(0, sheetData_children.length);		
		table.forEach(function(r){
			sheetData.append(r);
		});
		//mergeCells
		var megeCellsR = root.find("mergeCells");
		if(megeCellsR){
			var megeCellsR_children = megeCellsR.getchildren();
			megeCellsR.delSlice(0, megeCellsR_children.length);	
			for(var key in refs){
				var r =new eletree.Element("mergeCell ");
				r.attrib.ref =refs[key] 
				megeCellsR.append(r);
			}
		}
		//save sheet1
		zip.file("xl/worksheets/sheet1.xml",eletree.tostring(root));
		// Get binary data
		var result = zip.generate({base64: false, checkCRC32: true});
		callback(null,result);
	})
	
}					
								
								