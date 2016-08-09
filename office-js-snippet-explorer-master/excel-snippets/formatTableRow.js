
Excel.run(function (ctx) {
	var range = ctx.workbook.tables.getItem('Table1').rows.getItemAt(0).getRange();
	range.format.fill.color = "Green";
	return ctx.sync();
}).catch(function (error) {
	console.log(JSON.stringify(error));
});