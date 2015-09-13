var excelReport = require("../excelReport");
var http = require('http');
http.createServer(function (req, res) {
  var data ={title:'Voucher List',company:'STP software',address:'56, 13C Street, Binh Tri Dong B ward, Binh Tan district, Ho Chi Minh City',user_created:'TRUONGPV'}
  data.table1 =[{date:new Date(Date.UTC(2015,0,13)),number:1,description:'description 1',qty:10}
				,{date:new Date(Date.UTC(2015,0,14)),number:2,description:'description 2',qty:20}
				,{date:new Date(Date.UTC(2015,0,14)),number:5,description:'description 3',qty:30}
				,{date:new Date(Date.UTC(2015,0,26)),number:0,description:'description 4',qty:0}
			]

  var template_file ='template.xlsx'

  excelReport(template_file,data,function(error,binary){
		if(error){
			res.writeHead(400, {'Content-Type': 'text/plain'});
			res.end(error);
			return
		}
		res.setHeader('Content-Type', 'application/vnd.openxmlformats');
		res.setHeader("Content-Disposition", "attachment; filename=report.xlsx");
		res.end(binary, 'binary');
  })
  
}).listen(3000, '127.0.0.1');
console.log("server running at http://127.0.0.1:3000")