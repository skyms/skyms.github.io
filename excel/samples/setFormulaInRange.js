var ctx = new Excel.ExcelClientContext();
var range = ctx.workbook.worksheets.getItem("Sheet1").getRange("A1:C3");
range.formulas = "=RAND()*17";
ctx.executeAsync().then();