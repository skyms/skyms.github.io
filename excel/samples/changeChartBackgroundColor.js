var ctx = new Excel.RequestContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(0);	
chart.format.fill.setSolidColor("#FF0000");
ctx.executeAsync().then();