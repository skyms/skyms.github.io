
Excel.run(function (ctx) {
	var range = ctx.workbook.worksheets.getItem("Sheet1").getRange("Sheet1!A1:D5");
	ctx.workbook.worksheets.getItem("Sheet1").charts.add("ColumnClustered", range , Excel.ChartSeriesBy.auto);
	return ctx.sync();
}).catch(function (error) {
	console.log(error);
});
