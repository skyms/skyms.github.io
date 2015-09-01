/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            // create a new Request Context
            var ctx = new Excel.RequestContext();

            // get the worksheets collection 
            var worksheets = ctx.workbook.worksheets;
            var items;

            // load the items collection of worksheets so you 
            // can read its length property
            worksheets.load("items");

            //execute the above load statement
            ctx.executeAsync()
                .then(function () {
                    //add some worksheets to the workbook
                    if (worksheets.items.length == 1) {
                        for (var i = 2; i <=20; i++) {
                            worksheets.add("Sheet"+i);
                        }
                    }
                })
                //execute the above worksheets add statements
                .then(ctx.executeAsync)
                .then(function () {
                    //load the worksheets collection again
                    ctx.load(worksheets);
                    })
                .then(ctx.executeAsync)
                .then(function () {
                    //create a button for each sheet in the taskpane
                    for (var i = 0; i < worksheets.items.length; i++) {
                            var buttonName = worksheets.items[i].name;
                            var $input = $('<input type="button" class="button" value=' + buttonName + '>');
                            $input.appendTo($("#all-sheets"));
                            (function (buttonName) {
                                $input.click(function (e) {
                                    makeActiveSheet(buttonName);
                                });
                            })(buttonName);
                        }
                })
             .catch(function (error) {
                app.showNotification("Error", error);
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.RuntimeError) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
                console.log("Error: " + JSON.stringify(error.debugInfo));
            });
        });
    }

    function makeActiveSheet(buttonName) {
        var ctx = new Excel.RequestContext();
        var worksheets = ctx.workbook.worksheets;
        var clickedSheet = worksheets.getItem(buttonName);

        //insert the sheet name into a cell for better readability
        clickedSheet.getCell(0, 0).values = buttonName;

        //activate the sheet
        clickedSheet.activate();

        //execute the above the statements
        ctx.executeAsync().then(function () {
        });

    }
 
})();