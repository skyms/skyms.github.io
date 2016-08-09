
Excel.run(function (ctx) {
	var range = ctx.workbook.worksheets.getActiveWorksheet().getUsedRange();
	ctx.workbook.worksheets.getActiveWorksheet().charts.add("ColumnClustered", range , Excel.ChartSeriesBy.auto);
	return ctx.sync();
}).catch(function (error) {
	console.log(JSON.stringify(error));
});
