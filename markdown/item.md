# item resource type

*Namespace: Office.context.mailbox*

*Minimum requirement set/version: 1.0*

*Minimum permission level: Restricted*

*Modes supported: Read, Compose*


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
|attachments      | AttachmentDetails[] | Gets an array of attachments for the item. | 1.0 |  
|bcc      | Recipients | Gets or sets the recipients on the Bcc (blind carbon copy) line of a message. | 1.1 |  
|body      | Body | Gets an object that provides methods for manipulating the body of an item. | 1.1 |  
|cc      | EmailAddressDetails[], Recipients | Gets or sets the Cc (carbon copy) recipients of a message. | 1.0 |  
|conversationId      | String | Gets an identifier for the email conversation that contains a particular message. | 1.0 |  
|dateTimeCreated      | Date | Gets the date and time that an item was created. | 1.0 |  
|dateTimeModified      | Date | Gets the date and time that an item was last modified. | 1.0 |  
|end      | Date, Time | Gets or sets the date and time that the appointment is to end. | 1.0 |  
|from      | EmailAddressDetails | Gets the email address of the sender of a message. | 1.0 |  
|internetMessageId      | String | Gets the Internet message identifier for an email message. | 1.0 |  
|itemClass      | String | Gets the Exchange Web Services item class of the selected item. | 1.0 |  
|itemId      | String | Gets the Exchange Web Services item identifier for the current item. | 1.0 |  
|itemType      | Office.MailboxEnums.ItemType | Gets the type of item that an instance represents. | 1.0 |  
|location      | String, Location | Gets or sets the location of an appointment. | 1.0 |  
|normalizedSubject      | String | Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). | 1.0 |  
|notificationMessages      | NotificationMessages | Gets the notification messages for an item. | 1.3 |  
|optionalAttendees      | EmailAddressDetails[], Recipients | Gets or sets a list of email addresses for optional attendees. | 1.0 |  
|organizer      | EmailAddressDetails | Gets the email address of the meeting organizer for a specified meeting. | 1.0 |  
|requiredAttendees      | EmailAddressDetails[], Recipients | Gets or sets a list of email addresses for required attendees. | 1.0 |  
|resources      | EmailAddressDetails | Gets the resources required for an appointment. | 1.0 |  
|sender      | EmailAddressDetails | Gets the email address of the sender of an email message. | 1.0 |  
|start      | Date, Time | Gets or sets the date and time that the appointment is to begin. | 1.0 |  
|subject      | String, Subject | Gets or sets the description that appears in the subject field of an item. | 1.0 |  
|to      | EmailAddressDetails[], Recipients | Gets or sets the recipients of an email message. | 1.0 |  
>|%name%      | %type% | %description% | %req% |

%propertygetset%
%propertynotes%

### Enumerations

| Option	   | Type	| Description| Requirements|
|:-------------|:-------|:-----------|:------------|
>|%name%      | %type% | %description% | %enumreq% |

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

