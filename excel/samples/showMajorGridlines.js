var ctx = new Excel.RequestContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(0);	
chart.axes.valueAxis.majorGridlines.visible = true;
ctx.executeAsync().then();
