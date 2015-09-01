/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            createCollegeCreditTracker();

            //$('#generate-template').button;
            //$('#generate-template').click(createCollegeCreditTracker);
        });
    };

   
    function createCollegeCreditTracker() {
        // create a new Request Context
        var ctx = new Excel.RequestContext();

        // get the active worksheet
        var dashboardSheet = ctx.workbook.worksheets.getActiveWorksheet();

        //clear the sheet before inserting data
        dashboardSheet.getUsedRange().clear();
        dashboardSheet.name = "Dashboard";

        //var title;
        //if ($("#fullname").val() != null) {
        //    title = "College Credit Planner for " + $("#fullname").val();
        //}
        //else {
        //    title = "College Credit Planner";
        //}
        var title = "College Credit Planner";
        dashboardSheet.getRange("A1:A1").values = title;
        dashboardSheet.getRange("A1:A1").format.font.name = "Century";
        dashboardSheet.getRange("A1:A1").format.font.size = 26;
        dashboardSheet.getRange("A1:K1").format.fill.color = "1E8FEB";
        dashboardSheet.getRange("A1:A1").format.font.color = "white";

        var degreeName = "Bachelor of Arts in Music History";
        dashboardSheet.getRange("C1:C1").values = degreeName;
        dashboardSheet.getRange("C1:C1").format.font.name = "Century";
        dashboardSheet.getRange("C1:C1").format.font.size = 14;
        dashboardSheet.getRange("C1:C1").format.font.color = "white";

        var creditreqvalues = [["CREDIT REQUIREMENTS", "TOTAL", "EARNED", "NEEDED"],
                       ["Academic Major", 90, '=IFERROR(SUMIFS(D22:D122,C22:C122,C4,E22:E122,"=Yes"),"")', "=D4-E4"],
                       ["Academic Minor", 0, '=IFERROR(SUMIFS(D22:D122,C22:C122,C5,E22:E122,"=Yes"),"")', "=D5-E5"],
                       ["Elective Course", 4, '=IFERROR(SUMIFS(D22:D122,C22:C122,C6,E22:E122,"=Yes"),"")', "=D6-E6"],
                       ["General Study", 66, '=IFERROR(SUMIFS(D22:D122,C22:C122,C7,E22:E122,"=Yes"),"")', "=D7-E7"],
                       ["Totals", "=SUM(D4:D7)", "=SUM(E4:E7)", "=SUM(F4:F7)"]];

        dashboardSheet.getRange("C3:F8").values = creditreqvalues;

        dashboardSheet.getRange("C3:F3").format.font.size = 11;
        dashboardSheet.getRange("C3:F3").format.font.name = "Franklin Gothic Medium";
        dashboardSheet.getRange("C4:F8").format.font.name = "Bookman Old Style";
        dashboardSheet.getRange("C4:F7").format.font.size = 10;
        dashboardSheet.getRange("C3:F8").format.borders.getItem("InsideHorizontal").style = "Continuous";
        //dashboardSheet.getRange("C5:F10").format.borders.getItem('InsideHorizontal').style = 'Continuous';
        //dashboardSheet.getRange("C5:F10").format.borders.getItem('InsideVertical').style = 'Continuous';
        dashboardSheet.getRange("C3:F8").format.borders.getItem('EdgeBottom').style = 'Continuous';
        //dashboardSheet.getRange("C5:F10").format.borders.getItem('EdgeLeft').style = 'Continuous';
        //dashboardSheet.getRange("C5:F10").format.borders.getItem('EdgeRight').style = 'Continuous';
        dashboardSheet.getRange("C3:F8").format.borders.getItem('EdgeTop').style = 'Continuous';
        dashboardSheet.getRange("C8:F8").format.font.size = 11;
        dashboardSheet.getRange("C8:F8").format.font.name = "Franklin Gothic Medium";
        //dashboardSheet.getRange("F3:F8").format.font.color = "red";




        //courses
        var title = "College Courses";
        dashboardSheet.getRange("A23:A23").values = title;
        dashboardSheet.getRange("A23:A23").format.font.name = "Century";
        dashboardSheet.getRange("A23:A23").format.font.size = 26;
        dashboardSheet.getRange("A23:K23").format.fill.color = "1E8FEB";
        dashboardSheet.getRange("A23:A23").format.font.color = "white";

        var coursevalues = [["COURSE TITLE", "COURSE#", "DEGREE REQUIREMENT", "CREDITS", "COMPLETED?", "SEMESTER"],
                            ["Anthropology", "GEN 108", "General Study", 20, "Yes", "Semester 1"],
                            ["Applied Music", "MUS 215", "Academic Major", 3, , "Semester 3"],
                            ["Art History", "ART 101", "General Study", 2, "Yes", "Semester 1"],
                            ["Art History", "ART 201", "General Study", 2, "Yes", "Semester 2"],
                            ["Aural Skills I", "MUS 113", "Academic Major", 2, "Yes", "Semester 1"],
                            ["Aural Skills II", "MUS 213", "Academic Major", 2, "Yes", "Semester 2"],
                            ["Aural Skills III", "MUS 313", "Academic Major", 2, , "Semester 3"],
                            ["Aural Skills IV", "MUS 413", "Academic Major", 2, , "Semester 4"],
                            ["Conducting I", "MUS 114", "Academic Major", 2, "Yes", "Semester 1"],
                            ["English Writing", "Eng 101", "Generic Study", 3, "Yes", "Semester 1"],
                            ["English Writing", "Eng 201", "Generic Study", 3, "Yes", "Semester 2"],
                            ["Form and Analysis", "MUS 214", "Academic Major", 2, "Yes", "Semester 2"],
                            ["Intro to Anthropology", "GEN 208", "General Study", 3, "Yes", "Semester 2"],
                            ["Mathematics 101", "MAT 101", "General Study", 3, "Yes", "Semester 1"],
                            ["Music History in Western Civilization", "MUS 101", "Academic Major", 2, "Yes", "Semester 1"],
                            ["Music History in Western Civilization", "MUS 201", "Academic Major", 2, "Yes", "Semester 1"],
                            ["Music Theory I", "MUS 110", "Academic Major", 2, "Yes", "Semester 2"],
                            ["Music Theory II", "MUS 210", "Academic Major", 2, "Yes", "Semester 3"],
                            ["Music Theory III", "MUS 310", "Academic Major", 2, , "Semester 4"],
                            ["Music Theory IV", "MUS 410", "Academic Major", 2, , "Semester 5"],
                            ["Piano Class", "MUS 109", "Academic Major", 2, "Yes", "Semester 1"],
                            ["Social Sciences 101", "SOC 101", "General Study", 3, "Yes", "Semester 1"],
                            ["Social 101", "SOC 201", "General Study", 3, "Yes", "Semester 1"],
                            ["World of Jazz", "MUS 105", "Elective Course", 4, "Yes", "Semester 2"],
                            ["World of Music I", "MUS 112", "Academic Major", 6, "Yes", "Semester 3"],
                            ["World of Music II", "MUS 212", "Academic Major", 6, "Yes", "Semester 4"],
                            ["World of Music III", "MUS 312", "Academic Major", 6, "Yes", "Semester 5"],
                            ["World of Music IV", "MUS 412", "Academic Major", 6, "Yes", "Semester 6"],
                            ["World of Music V", "MUS 512", "Academic Major", 6, "Yes", "Semester 7"],
                            ["World of Music VI", "MUS 612", "Academic Major", 6, "Yes", "Semester 8"]

        ];

        dashboardSheet.getRange("A24:F54").values = coursevalues;

        dashboardSheet.getRange("A24:I24").format.font.size = 11;
        dashboardSheet.getRange("A24:I24").format.font.name = "Franklin Gothic Medium";
        dashboardSheet.getRange("A24:I24").format.font.color = "white";
        dashboardSheet.getRange("A25:I125").format.font.name = "Bookman Old Style";
        dashboardSheet.getRange("A25:I125").format.font.size = 10;
        dashboardSheet.getRange("A24:I24").format.fill.color = "2A4C69";

        //dashboardSheet.getRange("A16:I117").format.borders.getItem("InsideHorizontal").style = "Continuous";
        //dashboardSheet.getRange("C5:F10").format.borders.getItem('InsideHorizontal').style = 'Continuous';
        //dashboardSheet.getRange("C5:F10").format.borders.getItem('InsideVertical').style = 'Continuous';
        dashboardSheet.getRange("A24:I125").format.borders.getItem('EdgeBottom').style = 'Continuous';
        //dashboardSheet.getRange("C5:F10").format.borders.getItem('EdgeLeft').style = 'Continuous';
        //dashboardSheet.getRange("C5:F10").format.borders.getItem('EdgeRight').style = 'Continuous';
        dashboardSheet.getRange("A24:I125").format.borders.getItem('EdgeTop').style = 'Continuous';



        var charttitle = "SEMESTER SUMMARY";
        dashboardSheet.getRange("A3:A3").values = charttitle;
        dashboardSheet.getRange("A3:A3").format.font.name = "Franklin Gothic Medium";
        dashboardSheet.getRange("A3:A3").format.font.size = 11;


        var semestersummarytitle = "Semester Summary Data";
        dashboardSheet.getRange("C10:C10").values = semestersummarytitle;
        dashboardSheet.getRange("C10:C10").format.font.name = "Century";
        dashboardSheet.getRange("C10:E10").format.font.size = 11;
        dashboardSheet.getRange("C10:E10").format.fill.color = "1E8FEB";
        dashboardSheet.getRange("C10:E10").format.font.color = "white";
        //dashboardSheet.getRange("C12:E12").format.font.color = "white";



        var semestersummarydata = [["SEMESTER", "CREDITS", "CLASSES"],
                       ["Semester 1", '=IFERROR(SUMIFS(D22:D122,F22:F122,C12),"")', '=IFERROR(COUNTIFS(F22:F122,C12),"")'],
                       ["Semester 2", '=IFERROR(SUMIFS(D22:D122,F22:F122,C13),"")', '=IFERROR(COUNTIFS(F22:F122,C13),"")'],
                       ["Semester 3", '=IFERROR(SUMIFS(D22:D122,F22:F122,C14),"")', '=IFERROR(COUNTIFS(F22:F122,C14),"")'],
                       ["Semester 4", '=IFERROR(SUMIFS(D22:D122,F22:F122,C15),"")', '=IFERROR(COUNTIFS(F22:F122,C15),"")'],
                       ["Semester 5", '=IFERROR(SUMIFS(D22:D122,F22:F122,C16),"")', '=IFERROR(COUNTIFS(F22:F122,C16),"")'],
                       ["Semester 6", '=IFERROR(SUMIFS(D22:D122,F22:F122,C17),"")', '=IFERROR(COUNTIFS(F22:F122,C17),"")'],
                       ["Semester 7", '=IFERROR(SUMIFS(D22:D122,F22:F122,C18),"")', '=IFERROR(COUNTIFS(F22:F122,C18),"")'],
                       ["Semester 8", '=IFERROR(SUMIFS(D22:D122,F22:F122,C19),"")', '=IFERROR(COUNTIFS(F22:F122,C19),"")'],
                       ["Total", "=sum(D12:D19)", "=sum(E12:E19)"]];

        dashboardSheet.getRange("C11:E20").values = semestersummarydata;

        dashboardSheet.getRange("C11:E11").format.font.size = 11;
        dashboardSheet.getRange("C11:E11").format.font.name = "Franklin Gothic Medium";
       // dashboardSheet.getRange("A15:C15").format.fill.color = "2A4C69";
        dashboardSheet.getRange("C12:E19").format.font.name = "Bookman Old Style";
        dashboardSheet.getRange("C12:E19").format.font.size = 10;
        dashboardSheet.getRange("C11:E20").format.borders.getItem("InsideHorizontal").style = "Continuous";
        //dashboardSheet.getRange("C5:F10").format.borders.getItem('InsideHorizontal').style = 'Continuous';
        //dashboardSheet.getRange("C5:F10").format.borders.getItem('InsideVertical').style = 'Continuous';
        dashboardSheet.getRange("C11:E20").format.borders.getItem('EdgeBottom').style = 'Continuous';
        //dashboardSheet.getRange("C5:F10").format.borders.getItem('EdgeLeft').style = 'Continuous';
        //dashboardSheet.getRange("C5:F10").format.borders.getItem('EdgeRight').style = 'Continuous';
        dashboardSheet.getRange("C11:E20").format.borders.getItem('EdgeTop').style = 'Continuous';
        dashboardSheet.getRange("C20:E20").format.font.size = 11;
        dashboardSheet.getRange("C20:E20").format.font.name = "Franklin Gothic Medium";
        dashboardSheet.getRange("C20:E20").format.font.color = "white";
        dashboardSheet.getRange("C20:E20").format.fill.color = "2A4C69";


        //create a chart
        var chartSource = "Dashboard!C11:E19";
        var semestersummarychart = dashboardSheet.charts.add("BarClustered", chartSource, "auto");

        //settings
        //semestersummarychart.width = 500;
        //semestersummarychart.height = 300;
        semestersummarychart.setPosition("A4", "A19");
        semestersummarychart.legend.position = "right";
        //semestersummarychart.top = 162;
        //semestersummarychart.left = 55;
        semestersummarychart.title.visible = false;
        //semestersummarychart.axes.valueAxis.majorGridlines.visible = true;
        //semestersummarychart.axes.valueAxis.majorUnit = 0.25;

        semestersummarychart.dataLabels.showValue = true;

        semestersummarychart.series.getItemAt(0).format.fill.setSolidColor("green");
        semestersummarychart.series.getItemAt(1).format.fill.setSolidColor("dark green");

        ctx.executeAsync()
            .then(function () {
                //app.showNotification("Title inserted!");
            })
            .catch(function (error) {
                //app.showNotification("Error", JSON.stringify(error));
            });
    }
})();