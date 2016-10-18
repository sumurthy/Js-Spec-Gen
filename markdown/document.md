# Document Object (JavaScript API for Visio)

_Visio Online_

Represents the Document class.

## Properties

None

## Relationships
| Relationship | Type	|Description| Req. Set| Feedback|
|:---------------|:--------|:----------|:----|:---|
|application|[Application](application.md)|Represents a Visio application instance that contains this document. Read-only.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-document-application)|
|pages|[PageCollection](pagecollection.md)|Represents a collection of pages associated with the document. Read-only.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-document-pages)|
|view|[DocumentView](documentview.md)|Returns the DocumentView object. Read-only.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-document-view)|

## Methods

| Method		   | Return Type	|Description| Req. Set| Feedback|
|:---------------|:--------|:----------|:----|:---|
|[getActivePage()](#getactivepage)|[Page](page.md)|Returns the Active Page of the document.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-document-getActivePage)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-document-load)|
|[setActivePage(PageName: string)](#setactivepagepagename-string)|void|Set the Active Page of the document.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-document-setActivePage)|
|[startDataRefresh()](#startdatarefresh)|void|Triggers the refresh of the data in the Diagram, for all pages.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-document-startDataRefresh)|

## Method Details


### getActivePage()
Returns the Active Page of the document.

#### Syntax
```js
documentObject.getActivePage();
```

#### Parameters
None

#### Returns
[Page](page.md)

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

### setActivePage(PageName: string)
Set the Active Page of the document.

#### Syntax
```js
documentObject.setActivePage(PageName);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|PageName|string|Name of the page|

#### Returns
void

### startDataRefresh()
Triggers the refresh of the data in the Diagram, for all pages.

#### Syntax
```js
documentObject.startDataRefresh();
```

#### Parameters
None

#### Returns
void
