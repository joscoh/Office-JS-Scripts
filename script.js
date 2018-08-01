$("#run").click(() => tryCatch(run));

function run() {
    return Word.run(function (context) {
		var bindingName = "" + Math.random();
		var range = context.document.getSelection();
            var myContentControl = range.insertContentControl();
            myContentControl.title = bindingName;
            return context.sync().then(function () {
                console.log('Wrapped a content control around the selected text with title ' + bindingName);
                Office.context.document.bindings.addFromNamedItemAsync(bindingName,
                Office.BindingType.Text, { id: bindingName },
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        console.log('Control bound. Binding.id: '
                            + result.value.id + ' Binding.type: ' + result.value.type);
							//add content to binding
							Office.select("bindings#" + bindingName).setDataAsync('<h1>Title</h1><p>Added content to binding.</p>', { coercionType: 'html' }, function (asyncResult) {
							console.log("Added content to " + bindingName + "binding");
							});
                    } else {
                        console.log('Error:', result.error.message);
                    }
                });
            });
        )
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
    }
}

/** Default helper for invoking an action and handling errors. */
function tryCatch(callback) {
    Promise.resolve()
        .then(callback)
        .catch(function (error) {
            OfficeHelpers.UI.notify(error);
            OfficeHelpers.Utilities.log(error);
        });
}