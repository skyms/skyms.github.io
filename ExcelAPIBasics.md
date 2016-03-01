## The basics

Let's begin with a brief introduction to the key concepts that are fundamental to using the APIs, such as RequestContext, JavaScript proxy objects, sync(), Excel.run(), and load(). The example code at the end of the section shows these concepts in use.


#### RequestContext

The RequestContext object facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, request context is required to get access to Excel and related objects such as worksheets, tables, etc. from the add-in. A request context is created as shown below.

```js
var ctx = new Excel.RequestContext();
```

#### Proxy Objects 

The Excel JavaScript objects declared and used in an add-in are proxy objects for the real objects in an Excel document. All actions taken on proxy objects are not realized in Excel, and the state of the Excel document is not realized in the proxy objects until the document state has been synchronized. The document state is synchronized when context.sync() is run (see below). 

For example, the local JavaScript object `selectedRange` is declared to reference the selected range. This can be used to queue the setting of its properties and invoking methods. The actions on such objects are not realized until the sync() method is run. 

```js
var selectedRange = ctx.workbook.getSelectedRange();
```    

#### sync()

The sync() method available on the request context synchronizes the state between JavaScript proxy objects and real objects in Excel by executing instructions queued on the context and retrieving properties of loaded Office objects for use in your code.  This method returns a promise, which is resolved when  synchronization is complete.

#### Excel.run(function(context) { batch })

Excel.run() executes a batch script that performs actions on the Excel object model. The batch commands include definitions of local JavaScript proxy objects and sync() methods that synchronize the state between local and Excel objects and promise resolution. The advantage of batching requests in Excel.run() is that when the promise is resolved, any tracked range objects that were allocated during the execution will be automatically released. 

The run method takes in RequestContext and returns a promise (typically, just the result of ctx.sync()). It is possible to run the batch operation outside of the Excel.run(). However, in such a scenario, any range object references needs to be manually tracked and managed. 

#### load()

The load() method is used to fill in the proxy objects created in the add-in JavaScript layer. When trying to retrieve an object, a worksheet for example, a local proxy object is created first in the JavaScript layer. Such an object can be used to queue the setting of its properties and invoking methods. However, for reading object properties or relations, the load() and sync() methods need to be invoked first. The load() method takes in the properties and relations that need to be loaded when the sync() method is called. 

_Syntax:_

```js
object.load(string: properties);
//or 
object.load(array: properties);
//or
object.load({loadOption});
```
Where, 

* `properties` is the list of properties and/or relationship names to be loaded specified as comma-delimited strings or array of names. See .load() methods under each object for details.
* `loadOption` specifies an object that describes the selection, expansion, top, and skip options. See object load [options](resources/loadoption.md) for details.

##### Example

The following example puts the above concepts together. The sample code shows writing of values from an array to a range object. 

The Excel.run() contains a batch of instructions. As part of this batch, a proxy object is created that references a range (address A1:B2) on the active worksheet. The value of this proxy range object is set locally. In order to read the values back, the `text` property of the range is instructed to be loaded onto the proxy object. All these commands are queued and run when ctx.sync() is called. The sync() method returns a promise that can be used to chain it with other operations.

```js
// Run a batch operation against the Excel object model. Use the context argument to get access to the Excel document.
Excel.run(function (ctx) { 

	// Create a proxy object for the sheet
	var sheet = ctx.workbook.worksheets.getActiveWorksheet();
	// Values to be updated
	var values = [
				 ["Type", "Estimate"],
				 ["Transportation", 1670]
				 ];
	// Create a proxy object for the range
	var range = sheet.getRange("A1:B2");

	// Assign array value to the proxy object's values property.
	range.values = values;
	
	// Queue a command to load the text property for the proxy range object.	
	range.load('text');

	// Synchronizes the state between JavaScript proxy objects and real objects in Excel by executing instructions queued on the context 
	return ctx.sync().then(function() {
			console.log("Done");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

##### Example

The following example shows how to copy the values from Range A1:A2 to B1:B2 of the active worksheet by using load() method on the range object. 

```js
// Run a batch operation against the Excel object model. Use the context argument to get access to the Excel document.
Excel.run(function (ctx) { 

	// Create a proxy object for the range
	var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:A2");

	// Queue a command to load the following properties on the proxy range object.	
	range.load ("address, values, range/format"); 

	// Synchronizes the state between JavaScript proxy objects and real objects in Excel by executing instructions queued on the context 
	return ctx.sync().then(function() {
		// Assign the previously loaded values to the new range proxy object. The values will be updated once the following .then() function is invoked. 
		ctx.workbook.worksheets. getActiveWorksheet().getRange("B1:B2").values= range.values;
	});
}).then(function() {
	  console.log("done");
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

