Excel.run(function (ctx) { 
    var table = ctx.workbook.tables.getItemAt(0);
    table.sort.apply([ 
            {
                key: 2,
                ascending: false
            },
        ], true);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});