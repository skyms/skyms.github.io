/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $("#tabs").tabs();

            $(".click-button").button();

            $('#add-expense').click(addExpense);
            $('#add-income').click(addIncome);

            createBudgetAnalyzer();
        });
    };

    function createBudgetAnalyzer() {
        // create a new Request Context
        var ctx = new Excel.RequestContext();

        // get the active worksheet
        var dashboardSheet = ctx.workbook.worksheets.getActiveWorksheet();

        //clear the sheet before inserting data
        dashboardSheet.getUsedRange().clear();
        dashboardSheet.name = "Dashboard";

        var title = "College Budget Analysis";
        dashboardSheet.getRange("A1:A1").values = title;
        dashboardSheet.getRange("A1:A1").format.font.name = "Rockwell";
        dashboardSheet.getRange("A1:A1").format.font.size = 22.5;



        ctx.workbook.tables.add('Dashboard!A117:C117', true);
        ctx.workbook.tables.getItem('Table1').getHeaderRowRange().values = [["Description", "Cost", "Category"]];
        var tableRows = ctx.workbook.tables.getItem('Table1').rows;
        tableRows.add(null, [["Rent", "$600", "Housing"]]);
        tableRows.add(null, [["Food", "$450", "Food"]]);
        tableRows.add(null, [["Car", "$150", "Transportation"]]);
        tableRows.add(null, [["Tuition", "$800", "School costs"]]);
        tableRows.add(null, [["Books", "$150", "School costs"]]);
        tableRows.add(null, [["Gift", "$100", "Other"]]);
        tableRows.add(null, [["Loan", "$250", "Loans/Payments"]]);


        ctx.workbook.tables.add('Dashboard!F117:H117', true);
        ctx.workbook.tables.getItem('Table2').getHeaderRowRange().values = [["Description", "Amount", "Category"]];
        var tableRows = ctx.workbook.tables.getItem('Table2').rows;

        tableRows.add(null, [["Wages", "$1500", "Wages"]]);
        tableRows.add(null, [["Parents", "$700", "Assistance from parents"]]);
        tableRows.add(null, [["Gift", "$100", "Other"]]);
        tableRows.add(null, [["Bank interest", "$250", "From savings"]]);
        tableRows.add(null, [["Scholarship", "$500", "Financial aid"]]);


        var expenseTableTitle = "Monthly Expenses";
        dashboardSheet.getRange("A116:A116").values = expenseTableTitle;
        dashboardSheet.getRange("A116:A116").format.font.name = "Rockwell";
        dashboardSheet.getRange("A116:A116").format.font.size = 18;

        var incomeTableTitle = "Monthly Income";
        dashboardSheet.getRange("F116:F116").values = incomeTableTitle;
        dashboardSheet.getRange("F116:F116").format.font.name = "Rockwell";
        dashboardSheet.getRange("F116:F116").format.font.size = 18;
     

        var summaryValues = [["Percentage of income spent", "=D4/D3"],
                              ["Income", '=SUM(G117:G122)'],
                              ["Expenses", '=SUM(B117:B124)'],
                              ["Balance", "=D3-D4"]];

        dashboardSheet.getRange("C2:D5").values = summaryValues;

        dashboardSheet.getRange("D2:D2").numberFormat = "0.00%";
        dashboardSheet.getRange("D3:D5").numberFormat = "$#";

        dashboardSheet.getRange("C2:D2").format.font.size = 18;
        dashboardSheet.getRange("C2:D2").format.font.color = "red";
        dashboardSheet.getRange("C2:D5").format.font.name = "Rockwell";
        dashboardSheet.getRange("C3:D5").format.font.size = 10;
        dashboardSheet.getRange("C2:D5").format.borders.getItem("InsideHorizontal").style = "Continuous";
        //dashboardSheet.getRange("C5:F10").format.borders.getItem('InsideHorizontal').style = 'Continuous';
        //dashboardSheet.getRange("C5:F10").format.borders.getItem('InsideVertical').style = 'Continuous';
        dashboardSheet.getRange("C2:D5").format.borders.getItem('EdgeBottom').style = 'Continuous';
        //dashboardSheet.getRange("C5:F10").format.borders.getItem('EdgeLeft').style = 'Continuous';
        //dashboardSheet.getRange("C5:F10").format.borders.getItem('EdgeRight').style = 'Continuous';
        dashboardSheet.getRange("C2:D5").format.borders.getItem('EdgeTop').style = 'Continuous';
        dashboardSheet.getRange("C5:D5").format.font.size = 13;
        dashboardSheet.getRange("C5:D5").format.font.name = "Rockwell";


        var moneyInValues = [["Money coming in", ""],
                             ["Category", "Amount"],
                             ["Wages", '=IFERROR(SUMIFS(G117:G122,H117:H122,C10),"")'],
                             ["Financial aid", '=IFERROR(SUMIFS(G117:G122,H117:H122,C11),"")'],
                             ["From savings", '=IFERROR(SUMIFS(G117:G122,H117:H122,C12),"")'],
                             ["Assistance from parents", '=IFERROR(SUMIFS(G117:G122,H117:H122,C13),"")'],
                             ["Other", '=IFERROR(SUMIFS(G117:G122,H117:H122,C14),"")'],
                             ["Total", "=sum(D10:D14)"]];

        dashboardSheet.getRange("C8:D15").values = moneyInValues;

        dashboardSheet.getRange("D10:D15").numberFormat = "$#";

        dashboardSheet.getRange("C8:D8").format.font.size = 18;
        dashboardSheet.getRange("C8:D8").format.font.color = "red";
        dashboardSheet.getRange("C8:D15").format.font.name = "Rockwell";
        dashboardSheet.getRange("C9:D9").format.font.size = 13;;
        dashboardSheet.getRange("C10:D14").format.font.size = 10;
        dashboardSheet.getRange("C8:D15").format.borders.getItem("InsideHorizontal").style = "Continuous";
        //dashboardSheet.getRange("C5:F10").format.borders.getItem('InsideHorizontal').style = 'Continuous';
        //dashboardSheet.getRange("C5:F10").format.borders.getItem('InsideVertical').style = 'Continuous';
        dashboardSheet.getRange("C8:D15").format.borders.getItem('EdgeBottom').style = 'Continuous';
        //dashboardSheet.getRange("C5:F10").format.borders.getItem('EdgeLeft').style = 'Continuous';
        //dashboardSheet.getRange("C5:F10").format.borders.getItem('EdgeRight').style = 'Continuous';
        dashboardSheet.getRange("C8:D15").format.borders.getItem('EdgeTop').style = 'Continuous';
        dashboardSheet.getRange("C15:D15").format.font.size = 13;
        dashboardSheet.getRange("C15:D15").format.font.name = "Rockwell";

        var moneyOutValues = [["Money going out", ""],
                             ["Category", "Cost"],
                             ["School costs", '=IFERROR(SUMIFS(B117:B124,C117:C124,C20),"")'],
                             ["Food", '=IFERROR(SUMIFS(B117:B124,C117:C124,C21),"")'],
                             ["Housing", '=IFERROR(SUMIFS(B117:B124,C117:C124,C22),"")'],
                             ["Transportation", '=IFERROR(SUMIFS(B117:B124,C117:C124,C23),"")'],
                             ["Loans/Payments", '=IFERROR(SUMIFS(B117:B124,C117:C124,C24),"")'],
                            ["Other", '=IFERROR(SUMIFS(B117:B124,C117:C124,C25),"")'],
                             ["Total", "=sum(D20:D25)"]];

        dashboardSheet.getRange("C18:D26").values = moneyOutValues;

        dashboardSheet.getRange("D19:D26").numberFormat = "$#";

        dashboardSheet.getRange("C18:D18").format.font.size = 18;
        dashboardSheet.getRange("C18:D18").format.font.color = "red";
        dashboardSheet.getRange("C18:D26").format.font.name = "Rockwell";
        dashboardSheet.getRange("C19:D19").format.font.size = 13;;
        dashboardSheet.getRange("C20:D25").format.font.size = 10;
        dashboardSheet.getRange("C18:D26").format.borders.getItem("InsideHorizontal").style = "Continuous";
        //dashboardSheet.getRange("C5:F10").format.borders.getItem('InsideHorizontal').style = 'Continuous';
        //dashboardSheet.getRange("C5:F10").format.borders.getItem('InsideVertical').style = 'Continuous';
        dashboardSheet.getRange("C18:D26").format.borders.getItem('EdgeBottom').style = 'Continuous';
        //dashboardSheet.getRange("C5:F10").format.borders.getItem('EdgeLeft').style = 'Continuous';
        //dashboardSheet.getRange("C5:F10").format.borders.getItem('EdgeRight').style = 'Continuous';
        dashboardSheet.getRange("C18:D26").format.borders.getItem('EdgeTop').style = 'Continuous';
        dashboardSheet.getRange("C26:D26").format.font.size = 13;
        dashboardSheet.getRange("C26:D26").format.font.name = "Rockwell";
    





        var incomeChartDataRange = dashboardSheet.getRange("C10:D14");

        
            
        var chart = dashboardSheet.charts.add("doughnut", incomeChartDataRange, "auto");

        chart.setPosition("A3", "A13");

        chart.title.text = "Income";
        chart.title.format.font.size = 15;
        chart.title.format.font.color = "red";

        chart.legend.position = "left";
        chart.legend.format.font.name = "Trebuchet MS (Body)";
        chart.legend.format.font.size = 8;

        chart.dataLabels.showPercentage = true;
        chart.dataLabels.format.font.size = 8;
        chart.dataLabels.format.font.color = "white";

        var points = chart.series.getItemAt(0).points;
        points.getItemAt(0).format.fill.setSolidColor("#ff3300");
        points.getItemAt(1).format.fill.setSolidColor("#00cccc");
        points.getItemAt(2).format.fill.setSolidColor("#bf6514");
        points.getItemAt(3).format.fill.setSolidColor("#2be6c2");
        points.getItemAt(4).format.fill.setSolidColor("#993cf3");





        var expenseChartDataRange = dashboardSheet.getRange("C20:D25");
       


        var expenseChart = dashboardSheet.charts.add("doughnut", expenseChartDataRange, "auto");

        expenseChart.setPosition("A16", "A26");

        expenseChart.title.text = "Expenses";
        expenseChart.title.format.font.size = 15;
        expenseChart.title.format.font.color = "red";

        expenseChart.legend.position = "left";
        expenseChart.legend.format.font.name = "Trebuchet MS (Body)";
        expenseChart.legend.format.font.size = 8;

        expenseChart.dataLabels.showPercentage = true;
        expenseChart.dataLabels.format.font.size = 8;
        expenseChart.dataLabels.format.font.color = "white";

        var points = expenseChart.series.getItemAt(0).points;
        points.getItemAt(0).format.fill.setSolidColor("#ff3300");
        points.getItemAt(1).format.fill.setSolidColor("#00cccc");
        points.getItemAt(2).format.fill.setSolidColor("#bf6514");
        points.getItemAt(3).format.fill.setSolidColor("#2be6c2");
        points.getItemAt(4).format.fill.setSolidColor("#993cf3");



        ctx.executeAsync()
            .then(function () {
                //app.showNotification("Success");
            })
            .catch(function (error) {
                app.showNotification("Error", JSON.stringify(error));
            });

    }
    
    function addExpense() {

        var ctx = new Excel.RequestContext();
 
        var tableRows = ctx.workbook.tables.getItem('Table1').rows;

        tableRows.add(null, [[$("#expense-description").val(), $("#expense-cost").val(), $("#expense-category").val()]]);


        ctx.executeAsync()
            .then(function () {
                //app.showNotification("Course added");
            })
            .catch(function (error) {
                //app.showNotification("Error", JSON.stringify(error));
            });

    }


    function addIncome() {
        var ctx = new Excel.RequestContext();

        var tableRows = ctx.workbook.tables.getItem('Table2').rows;

        tableRows.add(null, [[$("#income-description").val(), $("#income-amount").val(), $("#income-category").val()]]);


        ctx.executeAsync()
            .then(function () {
                //app.showNotification("Course added");
            })
            .catch(function (error) {
                //app.showNotification("Error", JSON.stringify(error));
            });


    }

})();