
Excel.run(function (ctx) {
	var range = ctx.workbook.worksheets.getItem("Sheet1").getRange("A1:C3");
	range.
	range.values = [[1, 2, 3], [4, 5, 6], [7, 8, 9]];
	var table = ctx.workbook.tables.add("Sheet1!A1:C3", false);

	var column = table.columns.getItemAt(0);
	column.filter.applyDynamicFilter(Excel.DynamicFilterCriteria.aboveAverage);

	return ctx.sync();
}).catch(function (error) {
	console.log(error);
});