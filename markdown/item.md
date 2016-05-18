# item resource type

Provides methods and properties for accessing a message or appointment in an Outlook add-in.

The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType]{@link Office.context.mailbox.item#itemType} property. 	 
// The initialize function is required for all apps. 	 
Office.initialize = function () { 	 
// Checks for the DOM to load using the jQuery ready function. 	 
$(document).ready(function () { 	 
// After the DOM is loaded, app-specific code can run. 	 
var item = Office.context.mailbox.item; 	 
var subject = item.subject; 	 
// Continue with processing the subject of the current item, 	 
// which can be a message or appointment. 	 
}); 	 
} 	 
##### Example 
 	 

```js 	 
The following JavaScript code example shows how to access the `subject` property of the current item in Outlook. 	 
// The initialize function is required for all apps. 	 
Office.initialize = function () { 	 
// Checks for the DOM to load using the jQuery ready function. 	 
$(document).ready(function () { 	 
// After the DOM is loaded, app-specific code can run. 	 
var item = Office.context.mailbox.item; 	 
var subject = item.subject; 	 
// Continue with processing the subject of the current item, 	 
// which can be a message or appointment. 	 
}); 	 
} 	 
```



### Properties

| Property	   | Type	| Description| Requirements|
|:-------------|:-------|:-----------|:------------|
>|%propname%      | %proptype% | %propdescription% | %propreq% |

%propertygetset%
%propertynotes%

### Properties

| Option	   | Type	| Description| Requirements|
|:-------------|:-------|:-----------|:------------|
>|%enumname%      | %enumtype% | %enumdescription% | %enumreq% |

%propertygetset%
%propertynotes%


### Relationships
| Relationship | Type	| Description| Requirements|
|:-------------|:-------|:-----------|:------------|
>|%name%      | [%type%](%link%) | %description% | %req% |

%relationshipnotes%


## Methods

| Method	   | Return Type    | Description | Requirements|
|:-------------|:---------------|:------------|:----|
>| [%name%](%link%)     | %dtype% | %description% | %req%|

%methodnotes%

