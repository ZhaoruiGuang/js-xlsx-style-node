var fs = require('fs');
var path = require('path');
const XLSX = require('../index.js').default;

let ws = XLSX.parseToHtml(path.resolve(__dirname,'read.xlsx'),{
	cellDates:true,		
	// header:'',
	// footer:'',
	tableStyle:{
		"border-collapse":'collapse',
	},
	tdStyle:{
		"border":"1px solid #000"
	}
});
console.log(ws[0])
