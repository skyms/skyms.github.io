var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(0);	
chart.setData("A1:B4", Excel.ChartSeriesBy.columns);
ctx.executeAsync().then();