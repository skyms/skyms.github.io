var ctx = new Excel.ExcelClientContext();
var tableRows = ctx.workbook.tables.getItem('Table1').tableRows;
tableRows.add(3, [[1,2,3,4,5]]);
ctx.executeAsync().then();