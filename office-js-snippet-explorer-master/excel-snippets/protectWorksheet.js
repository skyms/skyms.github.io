
Excel.run(function (ctx) {
	var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
	worksheet.protection.protect();

	return ctx.sync();	
}).catch(function (error) {
	console.log(error);
});