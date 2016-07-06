# Document Object (JavaScript API for Visio)

_Visio 2016, Visio for iPad, Visio for Mac_

Dispatch Ids

## Properties

None

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|activePage|[Page](page.md)|Represents the active Page in the Document. ReadWrite|1.1||
|lastRefreshTime|[DateTime](datetime.md)|Last DateTime value of Diagram RefreshReloadOpen. Mohan: Put this in page level. Also what will be the time-zone ?. Potential misuse can happen. Read-only.|1.1||
|pages|[Pages](pages.md)|Represents the Pages in the Document. Read-only.|1.1||
|path|[String](string.md)|Represents the path of the document. Read-only.|1.1||
|zoom|[Double](double.md)|GetSet Document's Zoom level. Readwrite Mohan: move this to page level|1.1||

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[clearHandlers()](#clearhandlers)|void|Removes the handlers of the document. Mohan: Remove for a single one too.|1.1|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|
|[open(FileNameFullPath: string, PageIndex: number, FileMode: visFileMode)](#openfilenamefullpath-string-pageindex-number-filemode-visfilemode)|void|Open's a Document|1.1|
|[refreshData()](#refreshdata)|void|Refresh's the Data for the Diagram in all pages. Mohan: Park, potential blocking perf impact|1.1|
|[reload(ShouldDataRefresh: bool)](#reloadshoulddatarefresh-bool)|void|Reloads the Diagram with latest version and refreshed data. Mohan: provide an optional|1.1|

## Method Details


### clearHandlers()
Removes the handlers of the document. Mohan: Remove for a single one too.

#### Syntax
```js
documentObject.clearHandlers();
```

#### Parameters
None

#### Returns
void

### load(param: object)
Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document.
    var thisDocument = context.document;
    
    // Queue a command to load content control properties.
    context.load(thisDocument, 'contentControls/id, contentControls/text, contentControls/tag');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (thisDocument.contentControls.items.length !== 0) {
            for (var i = 0; i < thisDocument.contentControls.items.length; i++) {
                console.log(thisDocument.contentControls.items[i].id);
                console.log(thisDocument.contentControls.items[i].text);
                console.log(thisDocument.contentControls.items[i].tag);
            }
        } else {
            console.log('No content controls in this document.');
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


### open(FileNameFullPath: string, PageIndex: number, FileMode: visFileMode)
Open's a Document

#### Syntax
```js
documentObject.open(FileNameFullPath, PageIndex, FileMode);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|FileNameFullPath|string|File Name Full Path|
|PageIndex|number|Optional. Page Index to open|
|FileMode|visFileMode|Optional. File Mode. Default = Read|

#### Returns
void

### refreshData()
Refresh's the Data for the Diagram in all pages. Mohan: Park, potential blocking perf impact

#### Syntax
```js
documentObject.refreshData();
```

#### Parameters
None

#### Returns
void

### reload(ShouldDataRefresh: bool)
Reloads the Diagram with latest version and refreshed data. Mohan: provide an optional

#### Syntax
```js
documentObject.reload(ShouldDataRefresh);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|ShouldDataRefresh|bool|Optional. Should Data connected be refreshed. Default = false|

#### Returns
void
