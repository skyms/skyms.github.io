
Excel.run(function (ctx) {
	var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:C4");
	range.values = [['Month','Quantity','Sales'],['Jan', 20, 300], ['Feb', 15, 260], ['Mar', 28, 480]];
	return ctx.sync();
}).catch(function (error) {
	console.log(JSON.stringify(error));
});