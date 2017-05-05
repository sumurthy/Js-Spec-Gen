### add(key: string, value: object)

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue commands add a setting.
    var settings = context.document.settings;
    settings.add('startMonth', { month: 'March', year: 1998 });

    // Queue a command to load the settings object.
    context.load(settings, 'value');

    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        for (var i = 0; i < settings.items.length; i++) {
            console.log(settings.items[i].value);
        }
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### deleteAll()

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue commands add a setting.
    var settings = context.document.settings;
    settings.add('startMonth', { month: 'March', year: 1998 });

    // Queue a command to get the count of settings.
    var count = settings.getCount();

    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log(count.value);

        // Queue a command to delete all settings.
        settings.deleteAll();

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

### getCount()

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue commands add a setting.
    var settings = context.document.settings;
    settings.add('startMonth', { month: 'March', year: 1998 });

    // Queue a command to get the count of settings.
    var count = settings.getCount();

    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log(count.value);

        // Queue a command to delete all settings.
        settings.deleteAll();

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


### getItem(key: string)

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue commands add a setting.
    var settings = context.document.settings;
    settings.add('startMonth', { month: 'March', year: 1998 });

    // Queue a command to retrieve a setting.
    var startMonth = settings.getItem('startMonth');

    // Queue a command to load the setting.
    context.load(startMonth);

    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log(JSON.stringify(startMonth.value));
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### getItem(key: string)

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue commands add a setting.
    var settings = context.document.settings;
    settings.add('startMonth', { month: 'March', year: 1998 });
    
    // Queue commands to retrieve settings.
    var startMonth = settings.getItemOrNullObject('startMonth');
    var endMonth = settings.getItemOrNullObject('endMonth');

    // Queue commands to load settings.
    context.load(startMonth);
    context.load(endMonth);

    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
       return context.sync().then(function () {
           if (startMonth.isNullObject) {
               console.log("No such setting.");
           }
           else {
               console.log(JSON.stringify(startMonth.value));
           }
            if (endMonth.isNullObject) {
               console.log("No such setting.");
           }
           else {
               console.log(JSON.stringify(endMonth.value));
           }
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

