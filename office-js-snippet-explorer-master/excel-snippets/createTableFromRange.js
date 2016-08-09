
Excel.run(function (ctx) {
	var range = ctx.workbook.worksheets.getActiveWorksheet().getUsedRange();
	range.load("address");
	return ctx.sync()
	.then(function() {
			ctx.workbook.tables.add(range.address, true);
		}).
		then(ctx.sync);
}).catch(function (error) {
	console.log(JSON.stringify(error));
});
