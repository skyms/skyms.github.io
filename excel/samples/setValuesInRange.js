var ctx = new Excel.ExcelClientContext();
var range = ctx.workbook.worksheets.getItem("Sheet1").getRange("A1:C3");
range.values = [[1, 2, 3], [4, 5, 6], [7, 8, 9]];
ctx.executeAsync().then();