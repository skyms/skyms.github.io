/// <reference pashowAlert() th="myscript.js" />

Office.initialize = function (reason) {
    insideOffice = true;
  
};

function getMyDietData() {
    // create a new Request Context
    var ctx = new Excel.RequestContext();

    // get the Diet Worksheet
    var dietSheet = ctx.workbook.worksheets.getActiveWorksheet();

    //clear the sheet before inserting data
    dietSheet.getUsedRange().clear();
    dietSheet.name = "Diet";

    var values = [["DATE", "CARBS", "SUGARS", "FIBER", "FAT", "Total", "TIME", "DESCRIPTION", "NOTES"],
            ["41334", "10", "0", "0", "10", "20", "0.291667", "Coffee", "Morning coffee"],
            ["41334", "10", "2", "10", "10", "32", "0.333333", "Bagel", "Light breakfast"],
            ["41334", "35", "12", "45", "350", "442", "0.5", "Lunch", "Turkey sandwich"],
            ["41334", "75", "32", "95", "575", "777", "0.791667", "Dinner", "Tater tot casserole"],
            ["41335", "10", "0", "0", "10", "20", "0.291667", "Coffee", "Morning coffee"],
            ["41335", "10", "2", "10", "10", "32", "0.333333", "Toast", "Light breakfast"],
            ["41335", "40", "15", "55", "325", "435", "0.5", "Lunch", "Sandwich"],
            ["41335", "45", "45", "45", "445", "580", "0.791667", "Dinner", "Dinner"],
            ["41336", "10", "0", "0", "10", "20", "0.291667", "Coffee", "Morning coffee"],
            ["41336", "10", "2", "10", "10", "32", "0.333333", "Bagel", "Light breakfast"],
            ["41336", "10", "2", "2", "50", "64", "0.5", "Lunch", "Salad"],
            ["41336", "64", "32", "22", "456", "574", "0.791667", "Dinner", "Dinner"],
            ["41337", "10", "5", "0", "10", "25", "0.291667", "Coffee", "Coffee"],
            ["41337", "10", "5", "0", "10", "25", "0.416667", "Coffee", "Coffee"],
            ["41337", "15", "0", "35", "125", "175", "0.510417", "Lunch", "Salad"]];
    ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:I16").values = values;
    ctx.executeAsync()
        .then(function () {
            console.log("My Diet Data Inserted!");
        })
        .catch(function (error) {
            console.log(JSON.stringify(error));
        });
}


// Format Diet Worksheet, change the data format, insert title & subtitle and create a table.
function formatMyDietData() {
    // create a new Request Context
    var ctx = new Excel.RequestContext();

    // get the Diet Worksheet
    var dietSheet = ctx.workbook.worksheets.getItem("Diet");

    var dietRange;
    // insert a range to make room for titles
    dietSheet.getRange("A1:I4").insert("Down");
    ctx.executeAsync()
        .then(function () {
            //insert tiltes and subtitles, together with formatting
            dietSheet.getRange("A1:A1").values = "Diet";
            dietSheet.getRange("A1:A1").format.font.name = "Arial Black";
            dietSheet.getRange("A1:A1").format.font.size = 24;
            dietSheet.getRange("A1:I2").format.borders.getItem("InsideHorizontal").style = "Continuous";
            dietSheet.getRange("A2:A2").values = "DIET & EXERCISE JOURNAL";
            dietSheet.getRange("A2:A2").format.font.size = 12;
            dietSheet.getRange("A2:A2").format.font.name = "Arial";
        })
        //Remember to call execute Async
        .then(ctx.executeAsync)
        .then(function () {
            //Get the Range
            var range = ctx.workbook.worksheets.getItem("Diet").getRange("A6:A20");

            //setting data format
            range.numberFormat = "[$-en-US]mmmm d, yyyy;@";
        })
        .then(function () {
            // Get a Range
            dietRange = ctx.workbook.worksheets.getItem("Diet").getUsedRange();
            //load rowCount in order to know the last row of the table
            ctx.load(dietRange, "rowCount");
        })
        .then(ctx.executeAsync)
        .then(function () {
            // add the table
            var dietTable = ctx.workbook.tables.add("Diet!A5:I" + dietRange.rowCount, true);
            // insert the formula to get the maximum 
            ctx.workbook.worksheets.getItem("Diet").getRange("A4:A4").formulas = "=MAX(F6:F" + dietRange.rowCount + ")";

        })
        .then(ctx.executeAsync)
        .then(function () {
            alert("Diet Worksheet Format Updated!");
        })
        .catch(function (error) {
            console.log(JSON.stringify(error));
        });
}

// Highlight the meal with most calories in red on the Diet worksheet
function highlightMax() {
    // create new Request Context
    var ctx = new Excel.RequestContext();

    // find the maximum value
    var maxMeal = ctx.workbook.worksheets.getItem("Diet").getRange("A4:A4");
    // load the values
    ctx.load(maxMeal, "values");

    //load the diet range 
    var dietRange = ctx.workbook.worksheets.getItem("Diet").getUsedRange();
    ctx.load(dietRange);
    ctx.executeAsync()
        .then(function () {

            // loop through and find the row with the maximum value
            for (var i = 6; i < dietRange.rowCount; i++) {
                if (dietRange.values[i][5] == maxMeal.values[0][0]) {
                    var j = i + 1;
                    var rangeAddr = "A" + j + ":I" + j;
                    //hightlight and select
                    ctx.workbook.worksheets.getItem("Diet").getRange(rangeAddr).format.fill.color = "red";
                    ctx.workbook.worksheets.getItem("Diet").getRange(rangeAddr).select();
                }
            }
        })
        .then(ctx.executeAsync)
        .then(function () {
            console.log("Meal wtih Most Calories Highlighted!");
        })
        .catch(function (error) {
            console.log(JSON.stringify(error));
        });
}

function createDashboardWorksheet() {
    var ctx = new Excel.RequestContext();
    var worksheets = ctx.workbook.worksheets;
    //load items to know the total number of the worksheets
    ctx.load(worksheets, "items, name");
    ctx.executeAsync()
        .then(function () {
            var i;
            // loop through worksheets collection to find out whether Goals already exsits
            for (i = 0; i < worksheets.items.length; i++) {
                if (worksheets.items[i].name == "Goals") {
                    break;
                }
            }
            // if it doesn't exist, add a new worksheet.
            if (worksheets.items.length == i) {
                worksheets.add("Goals");
                return ctx.executeAsync();
            }
        })
        .then(function () {
            // activate Worksheet
            var worksheets = ctx.workbook.worksheets.getItem("Goals").activate();
        })
        .then(ctx.executeAsync)
        .then(function () {
            console.log("Worksheet Goals is Created and Activated.");
        })
        .catch(function (error) {
            console.log(JSON.stringify(error));
        });
}

// Format Worksheet Goals
function formatGoals() {
    var ctx = new Excel.RequestContext();
    var goalsSheet = ctx.workbook.worksheets.getItem("Goals");
    // formatting the range
    for (var i = 1; i <= 14; i++) {
        var rangeAddr = "A" + i + ":A" + i;
        if (i % 2 == 0) {
            goalsSheet.getRange(rangeAddr).format.font.name = "Arial";
            goalsSheet.getRange(rangeAddr).format.font.size = 11;
        }
        else {
            goalsSheet.getRange(rangeAddr).format.font.name = "Arial Black";
            goalsSheet.getRange(rangeAddr).format.font.size = 18;
        }

    }
    //set values
    goalsSheet.getRange("A2:A2").values = "Start Date";
    goalsSheet.getRange("A4:A4").values = "End Date";
    goalsSheet.getRange("A6:A6").values = "Start Weight";
    goalsSheet.getRange("A8:A8").values = "End Weight";
    goalsSheet.getRange("A10:A10").values = "Goal Loss";
    goalsSheet.getRange("A12:A12").values = "Days to Go";
    goalsSheet.getRange("A14:A14").values = "Loss Per Day";

    //set formulas
    goalsSheet.getRange("A9:A9").formulas = "= A5-A7";
    goalsSheet.getRange("A11:A11").formulas = "=DAYS(A3,A1)";
    goalsSheet.getRange("A13:A13").formulas = "=A9/A11";
    goalsSheet.getRange("A13:A13").numberFormat = "0.0";

    //set format
    goalsSheet.getRange("A1:A14").format.font.color = "white";
    goalsSheet.getRange("A1:A4").format.fill.color = "32CD32";
    goalsSheet.getRange("A5:A8").format.fill.color = "1C9C85";
    goalsSheet.getRange("A9:A14").format.fill.color = "1E8496";

    goalsSheet.getRange("B1:B1").values = "Goals";
    goalsSheet.getRange("B1:B1").format.font.name = "Arial";
    goalsSheet.getRange("B1:B1").format.font.size = 24;
    goalsSheet.getRange("B1:H2").format.borders.getItem("InsideHorizontal").style = "Continuous";
    goalsSheet.getRange("B2:B2").values = "DIET & EXERCISE JOURNAL";
    goalsSheet.getRange("B2:B2").format.font.size = 12;
    goalsSheet.getRange("B2:B2").format.font.name = "Arial";

    goalsSheet.getRange("B5:B5").values = "DIETARY ANALYSIS";
    goalsSheet.getRange("B5:B5").format.font.name = "Arial Black";
    goalsSheet.getRange("B5:B5").format.font.size = 14;
    goalsSheet.getRange("B5:B5").format.font.color = "white";
    goalsSheet.getRange("B5:K5").format.fill.color = "1E8496";

    ctx.executeAsync()
        .then(function () {
            console.log("Worksheet Goals Updated!");
        })
        .catch(function (error) {
            console.log(JSON.stringify(error));
        });
}

// Insert sample Goals
function insertGoals() {
    var ctx = new Excel.RequestContext();
    var goalsSheet = ctx.workbook.worksheets.getItem("Goals");

    //set values
    goalsSheet.getRange("A1:A1").values = "07/04/2015";
    goalsSheet.getRange("A3:A3").values = "12/04/2015";
    goalsSheet.getRange("A5:A5").values = 200;
    goalsSheet.getRange("A7:A7").values = 150;

    ctx.executeAsync()
        .then(function () {
            console.log("Sample Goals Defined");
        })
        .catch(function (error) {
            console.log(JSON.stringify(error));
        });
}

function createChart() {
    var ctx = new Excel.RequestContext();
    var goalsSheet = ctx.workbook.worksheets.getItem("Goals");
    var dietRange = ctx.workbook.worksheets.getItem("Diet").getUsedRange();
    ctx.load(dietRange, "rowCount,values");

    ctx.executeAsync()
        .then(function () {
            var chartSource = "Diet!A5:E" + dietRange.rowCount;

            //create a chart
            var goalChart = goalsSheet.charts.add("ColumnStacked100", chartSource, "auto");

            //settings
            goalChart.legend.position = "right";
            goalChart.top = 162;
            goalChart.left = 55;
            goalChart.title.visible = false;
            goalChart.axes.valueAxis.majorGridlines.visible = true;
            goalChart.axes.valueAxis.majorUnit = 0.25;

            goalChart.series.getItemAt(0).format.fill.setSolidColor("1E8496");
            goalChart.series.getItemAt(1).format.fill.setSolidColor("1C9C85");
            goalChart.series.getItemAt(2).format.fill.setSolidColor("708090");
            goalChart.series.getItemAt(3).format.fill.setSolidColor("FFC000");


        })
        .then(ctx.executeAsync)
        .then(function () {
            console.log("Chart Created!");
        })
        .catch(function (error) {
            console.log(JSON.stringify(error));
        });
}
