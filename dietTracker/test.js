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
                app.showNotification("My Diet Data Inserted!");
            })
            .catch(function (error) {
                app.showNotification("Error", JSON.stringify(error));
            });
    }
