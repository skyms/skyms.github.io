var ctx = new Excel.ExcelClientContext();
var title = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").title;	
ctx.load(title);
ctx.executeAsync().then(function () {
		console.log(title.text);
		console.log("done");
});