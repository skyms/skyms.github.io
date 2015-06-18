var ctx = new Word.WordClientContext();
ctx.customData = OfficeExtension.Constants.iterativeExecutor;

var paras = ctx.document.body.paragraphs;
ctx.load(paras);

ctx.executeAsync().then(
    function () {
        for (var i = 0; i < paras.count; i++) {
            paras.items[i].insertContentControl();
        }
        ctx.executeAsync().then(
            function () {
                console.log("Success");
            }
        );
    },
    function (result) {
        console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
        console.log(result.traceMessages);
    }
);