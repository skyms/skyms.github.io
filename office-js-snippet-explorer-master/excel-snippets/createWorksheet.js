
Excel.run(function (ctx) {
	var ws = ctx.workbook.worksheets.add("Sheet" + Math.floor(Math.random()*100000).toString());
	ws.activate();
	return ctx.sync();	
}).catch(function (error) {
	console.log(JSON.stringify(error));
});