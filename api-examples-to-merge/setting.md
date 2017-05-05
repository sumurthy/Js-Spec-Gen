### delete()


```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue commands add a setting.
    var settings = context.document.settings;
    var startMonth = settings.add('startMonth', { month: 'March', year: 1998 });

    // Queue a command to get the count of settings.
    var count = settings.getCount();

    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log(count.value);

        // Queue a command to delete the setting.
        startMonth.delete();

        // Queue a command to get the new count of settings.
        count = settings.getCount();
    })

    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    .then(context.sync)
    .then(function () {
        console.log(count.value);
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```
