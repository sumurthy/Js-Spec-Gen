### Getter
**id**
```js
OneNote.run(function (context) {
        
    // Get the current section.
    var section = context.application.activeSection;
            
    // Queue a command to load the section. 
    // For best performance, request specific properties.           
    section.load("id");
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            console.log("Section ID: " + section.id);
        });
    })
    .catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
```

**name and notebook**
```js
OneNote.run(function (context) {
        
    // Get the current section.
    var section = context.application.activeSection;
            
    // Queue a command to load the section with the specified properties. 
    section.load("name,notebook/name");
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            console.log("Section name: " + section.name);
            console.log("Parent notebook name: " + section.notebook.name);
        });
    })
    .catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
```


### addPage(title: string)
```js
OneNote.run(function (context) {
            
    // Queue a command to add a page to the current section.
    var page = context.application.activeSection.addPage("Wish list");
            
    // Queue a command to load the id and title of the new page. 
    // This example loads the new page so it can read its properties later.           
    page.load('id,title');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
             
            // Display the properties.       
            console.log("Page name: " + page.title);
            console.log("Page ID: " + page.id);

        });
    })
    .catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
```

### getPages()
```js
OneNote.run(function (context) {
            
    // Get the pages in the current section.
    var pages = context.application.activeSection.getPages();
            
    // Queue a command to load the id and title for each page.            
    pages.load('id,title');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Display the properties.         
            $.each(pages.items, function(index, page) {
                console.log("Page name: " + page.title);
                console.log("Page ID: " + page.id);
            });
        });
    })
    .catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
```

### insertSectionAsSibling(location: string, title: string)
```js
OneNote.run(function (context) {
            
    // Queue a command to insert a section after the current section.
    var section = context.application.activeSection.insertSectionAsSibling("After", "New section");
            
    // Queue a command to load the id and name of the new section. 
    // This example loads the new section so it can read its properties later.           
    section.load('id,name');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
             
            // Display the properties.       
            console.log("Section name: " + section.name);
            console.log("Section ID: " + section.id);
        });
    })
    .catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
```
