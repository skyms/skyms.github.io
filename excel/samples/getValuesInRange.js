var ctx = new Excel.ExcelClientContext();
var range = ctx.workbook.worksheets.getItem("Sheet1").getRange("A1:C3");
ctx.load(range);
ctx.executeAsync().then(function() {
	for (var i = 0; i < range.values.length; i++) {
		for (var j = 0; j < range.values[i].length; j++) {
			console.log(range.values[i][j]);
		}
	}
	console.log("done");
});