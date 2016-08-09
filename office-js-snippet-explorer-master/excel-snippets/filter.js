
Excel.run(function (ctx) {
	var filterColumn = ctx.workbook.tables.getItemAt(0).columns.getItemAt(2);
	filterColumn.filter.applyDynamicFilter(Excel.DynamicFilterCriteria.aboveAverage);

	return ctx.sync();
}).catch(function (error) {
	console.log(JSON.stringify(error));
});