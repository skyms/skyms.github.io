// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to get the current selection and then 
    // create a proxy range object with the results.
    var range = context.document.getSelection();
    
    // Queue a commmand to insert the paragraph at the end of the range.
    range.insertText('New text inserted into the range.', Word.InsertLocation.end);
    
    // Synchronize the document state by executing the queued-up commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Text added to the end of the range.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});