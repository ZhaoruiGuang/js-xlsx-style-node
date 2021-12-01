var fs = require('fs');
var path = require('path');
const XLSX = require('../index.js').default;

let ws = XLSX.parseToHtml(path.resolve(__dirname,'read.xlsx'),{
	cellDates:true,			
	// header:'',
	// footer:'',
});

console.log(ws)
