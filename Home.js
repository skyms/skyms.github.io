/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            //link button with functions
            $('#set-numberformat').click(setNumberFormat);
            $('#insert-title').click(insertTitle);
            $('#get-data').click(setNumberFormat);
            $('#create-table').click(createTable);
            $('#highlight-max').click(highlightMax);
            $('#create-worksheet').click(createDashboardWorksheet);
            $('#select-worksheet').click(selectWorksheet);
            $('#insert-goals').click(insertGoals);
            $('#create-chart').click(createChart);
            $('#get-weather-forecast').click(getWeatherForecast);


            
        });
    };

    // Insert Title in the Diet Sheet
    function insertTitle() {
        // create a new Request Context
        var ctx = new Excel.RequestContext();

        // get the Diet Worksheet
        var dietSheet = ctx.workbook.worksheets.getItem("Diet");
        
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
                app.showNotification("Title is added for Diet!");
            })
            .catch(function (error) {
                app.showNotification("Error", JSON.stringify(error));
            });
    }

    // change the date format
    function setNumberFormat() {
        // create a new Request Context 
        var ctx = new Excel.RequestContext();

        //Get the Range
        var range = ctx.workbook.worksheets.getItem("Diet").getRange("A6:A20");

        //setting data format
        range.numberFormat = "[$-en-US]mmmm d, yyyy;@";

        ctx.executeAsync()
            .then(function () {
                app.showNotification("NumberFormat is Set Succesfully!");
            })
            .catch(function (error) {
                app.showNotification("Error", JSON.stringify(error));
            });
    }

    // create a table to better present the data
    function createTable() {
        // create a new Request Context
        var ctx = new Excel.RequestContext();

        // Get a Range
        var dietRange = ctx.workbook.worksheets.getItem("Diet").getUsedRange();

        //load rowCount in order to know the last row of the table
        ctx.load(dietRange, "rowCount");
        ctx.executeAsync()
            .then(function () {
                // add the table
                var dietTable = ctx.workbook.tables.add("Diet!A5:I" + dietRange.rowCount, true);
                // insert the formula to get the maximum 
                ctx.workbook.worksheets.getItem("Diet").getRange("A4:A4").formulas = "=MAX(F6:F" + dietRange.rowCount+")";
            })
            .then(ctx.executeAsync)
            .then(function () {
                app.showNotification("Table Added");
            })
            .catch(function (error) {
                app.showNotification("Error", JSON.stringify(error));
            });
    }

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
                for (var i = 6; i<dietRange.rowCount; i++){
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
                app.showNotification("Meal wtih Most Calories Highlighted!");
            })
            .catch(function (error) {
                app.showNotification("Error", JSON.stringify(error));
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
                for(i = 0; i < worksheets.items.length; i++) {
                    if (worksheets.items[i].name == "Goals"){
                        break;
                    }
                }
                // if it doesn't exist, add a new worksheet.
                if (worksheets.items.length == i){
                    worksheets.add("Goals");
                    return ctx.executeAsync();
                }
            })
            .then(function () {
                app.showNotification("Worksheet Goals is Created.");
            })
            .catch(function (error) {
                app.showNotification("Error", JSON.stringify(error));
            });
    }

    function selectWorksheet() {
        var ctx = new Excel.RequestContext();
        // activate Worksheet
        var worksheets = ctx.workbook.worksheets.getItem("Goals").activate();
       
        ctx.executeAsync()
            .then(function () {
                app.showNotification("Worksheet Goals is activited!");
            })
            .catch(function (error) {
                app.showNotification("Error", JSON.stringify(error));
            });
    }


    function insertGoals() {
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
                app.showNotification("Goals Defined");
            })
            .catch(function (error) {
                app.showNotification("Error", JSON.stringify(error));
            });
    }


    function createChart() {
        var ctx = new Excel.RequestContext();
        var goalsSheet = ctx.workbook.worksheets.getItem("Goals");
        goalsSheet.getRange("B5:B5").values = "DIETARY ANALYSIS";
        goalsSheet.getRange("B5:B5").format.font.name = "Arial Black";
        goalsSheet.getRange("B5:B5").format.font.size = 14;
        goalsSheet.getRange("B5:B5").format.font.color = "white";
        goalsSheet.getRange("B5:K5").format.fill.color = "1E8496";
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
                app.showNotification("Chart Created!");
            })
            .catch(function (error) {
                app.showNotification("Error", JSON.stringify(error));
            });
    }

    function getWeatherForecast() {
        var ctx = new Excel.RequestContext();
        var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:B1");
        range.load("values");

        ctx.executeAsync().then(function () {
            var state = range.values[0][0];
            var city = range.values[0][1];
            var weatherUrl = "https://api.wunderground.com/api/533b7efdc6786c66/forecast10day/q/" + state + "/" + city + ".json"
            console.log(weatherUrl);

            $.ajax({
                url: weatherUrl,
                dataType: "jsonp",
                success: function (parsed_json) {
                    //parsed_json['location']['city'];
                    //parsed_json['current_observation']['temp_f'];
                    console.log(parsed_json.stringify)
                    writeWeatherRange(parsed_json)
                },
                failure: function (error) {
                    var err = "Error in getWeather: " + error;
                    app.showNotification(JSON.stringify(err));
                }
            });

        });
    }


    function writeWeatherRange(weather) {

        var rangevalues = [["Date", "Weekday", "Celsius-High", "Fahrenheit-High", "Conditions"]];

        for (var i = 0; i < weather.forecast.simpleforecast.forecastday.length; i++) {

            var date = weather.forecast.simpleforecast.forecastday[i].date.pretty
            var weekday = weather.forecast.simpleforecast.forecastday[i].date.weekday
            var celsius = weather.forecast.simpleforecast.forecastday[i].high.celsius
            var farh = weather.forecast.simpleforecast.forecastday[i].high.fahrenheit
            var conditions = weather.forecast.simpleforecast.forecastday[i].conditions
            var array = [date, weekday, celsius, farh, conditions];
            rangevalues.push(array);

        }

        var rangeAddress = "A2:E12"
        var ctx = new Excel.RequestContext();
        ctx.workbook.worksheets.getActiveWorksheet().getRange("A2:E12").delete();
        var range = ctx.workbook.worksheets.getActiveWorksheet().getRange(rangeAddress);

        range.values = rangevalues;

        ctx.executeAsync().then(function () {
            app.showNotification("Write to Range" + rangeAddress + "is Successful!");
        }, function (error) {
            app.showNotification("Error", JSON.stringify(error));
        });
    }


})();
