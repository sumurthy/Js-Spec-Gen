# Body resource type

##### Namespace: *Body*

Provides methods for adding and updating the content of an item in an Outlook add-in.

The `body` object provides methods for adding and updating the content of the message or appointment. It is returned in the `body` property of the selected item.

|Requirement| Value|
|:----------|:-----|
|Minimum requirement set/version: | *1.1*|
|Minimum permission level |*ReadItem* |
|Modes supported | *Read, Compose*|



## Methods

| Method	   | Return Type    | Description | Requirements|
|:-------------|:---------------|:------------|:----|
| [getAsync](getasync)     |  | Returns the current body in a specified format.  | 1.3|  
| [getTypeAsync](gettypeasync)     |  | Gets a value that indicates whether the content is in HTML or text format.  | 1.1|  
| [prependAsync](prependasync)     |  | Adds the specified content to the beginning of the item body.  | 1.1|  
| [setAsync](setasync)     |  | Replaces the entire body with the specified text.  | 1.3|  
| [setSelectedDataAsync](setselecteddataasync)     |  | Replaces the selection in the body with the specified text.  | 1.1|  

