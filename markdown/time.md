# Time resource type

##### Namespace: *Time*

Provides methods to get and set the start or end time of a meeting in an Outlook add-in.

The `Time` object is returned as the [start]{@linkcode Office.context.mailbox.item#start} or [end]{@linkcode Office.context.mailbox.item#end} property of an appointment in compose mode.

|Requirement| Value|
|:----------|:-----|
|Minimum requirement set/version: | *1.1*|
|Minimum permission level |*ReadItem* |
|Modes supported | *Compose*|



## Methods

| Method	   | Return Type    | Description | Requirements|
|:-------------|:---------------|:------------|:----|
| [getAsync](getasync)     |  | Gets the start or end time of an appointment.  | 1.1|  
| [setAsync](setasync)     |  | Sets the start or end time of an appointment.  | 1.1|  

