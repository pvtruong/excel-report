# excel-report

create report from excel template. Only supports xlsx for now

```
npm install excel-report
```

## Usage

``` js
var excelReport = require('excel-report')

var data ={title:'Voucher List',company:'STP software',user_created:'TRUONGPV'}
data.table1 =[{date:new Date(),number:1,description:'description 1',qty:10},{date:new Date(),number:2,description:'description 2',qty:20}]

var template_file ='template.xlsx'

excelReport(template_file,data,function(error,binary){
	//check error
	//use binary
})
```
please view example
```
##Logs
Version 0.0.5:
- Support merge cells
- Support filter: {{tb:detail.qty|number:1}}
Version 0.1.1
- Fixed value =0
## License

MIT
