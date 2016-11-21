var excelReport = require("../excelReport");
var http = require('http');
http.createServer(function (req, res) {
  
  
  var data1 ={title:'Voucher List',company:'STP software',address:'56, 13C Street, Binh Tri Dong B ward, Binh Tan district, Ho Chi Minh City',user_created:'TRUONGPV',t_qty:10,price:100}
  data1.table1 =[{date:new Date(Date.UTC(2015,0,13)),number:1,description:'description 1',qty:10,price:1000}
				,{date:new Date(Date.UTC(2015,0,14)),number:2,qty:30,price:1000}
				,{date:new Date(Date.UTC(2015,0,14)),number:5,description:'description 3',qty:30,price:1000}
				,{date:new Date(Date.UTC(2015,0,14)),number:5,description:'description 3',qty:30,price:1000}
				,{date:new Date(Date.UTC(2015,0,14)),number:5,description:'description 3',qty:30,price:1000}
				,{date:new Date(Date.UTC(2015,0,14)),number:5,qty:30,price:1000}
				,{date:new Date(Date.UTC(2015,0,14)),number:5,description:'description 3',qty:30,price:1000}
				,{date:new Date(Date.UTC(2015,0,14)),number:5,description:'description 3',qty:30,price:1000}
				,{date:new Date(Date.UTC(2015,0,14)),number:5,description:'description 3',qty:30,price:1000}
				,{date:new Date(Date.UTC(2015,0,26)),number:0,description:'description 4',qty:0,price:1000}
			]
  
   var data2 ={title:'Voucher List',company:'STP software',address:'519, Hoang Sa Street, 8 ward, 3 district, Ho Chi Minh City',user_created:'LONGPH',t_qty:20,price:100}
  data2.table1 =[{date:new Date(Date.UTC(2015,0,13)),number:7,description:'description 41',qty:30,price:1000}
				,{date:new Date(Date.UTC(2015,0,18)),number:12,description:'description 42',qty:50,price:1000}
				,{date:new Date(Date.UTC(2015,0,14)),number:21,description:'description 43',qty:90,price:1000}
				,{date:new Date(Date.UTC(2015,0,18)),number:12,description:'description 42',qty:50,price:1000}
				,{date:new Date(Date.UTC(2015,0,14)),number:21,description:'description 43',qty:90,price:1000}
				,{date:new Date(Date.UTC(2015,0,18)),number:12,description:'description 42',qty:50,price:1000}
				,{date:new Date(Date.UTC(2015,0,14)),number:21,description:'description 43',qty:90,price:1000}
				,{date:new Date(Date.UTC(2015,0,18)),number:12,description:'description 42',qty:50,price:1000}
				,{date:new Date(Date.UTC(2015,0,14)),number:21,description:'description 43',qty:90,price:1000}
				,{date:new Date(Date.UTC(2015,0,18)),number:12,description:'description 42',qty:50,price:1000}
				,{date:new Date(Date.UTC(2015,0,14)),number:21,description:'description 43',qty:90,price:1000}
				,{date:new Date(Date.UTC(2015,0,18)),number:12,description:'description 42',qty:50,price:1000}
				,{date:new Date(Date.UTC(2015,0,14)),number:21,description:'description 43',qty:90,price:1000}
			]
			
			
 var data3 ={title:'Voucher List',company:'STP software',address:'519, Hoang Sa Street, 8 ward, 3 district, Ho Chi Minh City',user_created:'VIETPQ',t_qty:50,price:100}
  data3.table1 =[{date:new Date(Date.UTC(2015,0,13)),number:9,description:'description 61',qty:30,price:1000}
				,{date:new Date(Date.UTC(2015,0,14)),number:12,description:'description 62',qty:50,price:1000}
				,{date:new Date(Date.UTC(2015,0,14)),number:21,description:'description 63',qty:90,price:1000}
				,{date:new Date(Date.UTC(2015,0,28)),number:32,description:'description 64',qty:100,price:1000}
				,{date:new Date(Date.UTC(2015,0,14)),number:21,description:'description 43',qty:90,price:1000}
				,{date:new Date(Date.UTC(2015,0,14)),number:21,description:'description 43',qty:90,price:1000}
				,{date:new Date(Date.UTC(2015,0,14)),number:21,description:'description 43',qty:90,price:1000}
			]
			
			
 var datas =[data1,data2,data3];
  var template_file ='template.xlsx'

  excelReport(template_file,datas,function(error,binary){
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