# item resource type

##### Namespace: *Office.context.mailbox*

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


|Requirement| Value|
|:----------|:-----|
|Minimum requirement set/version: | *1.0*|
|Minimum permission level |*Restricted* |
|Modes supported | *Read, Compose*|


### Properties

| Property	   | Type	| Description| Requirements|
|:-------------|:-------|:-----------|:------------|
|attachments      | [AttachmentDetails[]](attachmentdetails.md) | Gets an array of attachments for the item. | 1.0  %readonly%|  
|bcc      | [Recipients](recipients.md) | Gets or sets the recipients on the Bcc (blind carbon copy) line of a message. | 1.1  %readonly%|  
|body      | [Body](body.md) | Gets an object that provides methods for manipulating the body of an item. | 1.1  %readonly%|  
|cc      | [EmailAddressDetails[]](emailaddressdetails.md) or [Recipients](recipients.md) | Gets or sets the Cc (carbon copy) recipients of a message. | 1.0  %readonly%|  
|conversationId      | String | Gets an identifier for the email conversation that contains a particular message. | 1.0  %readonly%|  
|dateTimeCreated      | Date | Gets the date and time that an item was created. | 1.0  %readonly%|  
|dateTimeModified      | Date | Gets the date and time that an item was last modified. | 1.0  %readonly%|  
|end      | Date or [Time](time.md) | Gets or sets the date and time that the appointment is to end. | 1.0  %readonly%|  
|from      | [EmailAddressDetails](emailaddressdetails.md) | Gets the email address of the sender of a message. | 1.0  %readonly%|  
|internetMessageId      | String | Gets the Internet message identifier for an email message. | 1.0  %readonly%|  
|itemClass      | String | Gets the Exchange Web Services item class of the selected item. | 1.0  %readonly%|  
|itemId      | String | Gets the Exchange Web Services item identifier for the current item. | 1.0  %readonly%|  
|itemType      | Office.MailboxEnums.ItemType | Gets the type of item that an instance represents. | 1.0  %readonly%|  
|location      | String or [Location](location.md) | Gets or sets the location of an appointment. | 1.0  %readonly%|  
|normalizedSubject      | String | Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). | 1.0  %readonly%|  
|notificationMessages      | [NotificationMessages](notificationmessages.md) | Gets the notification messages for an item. | 1.3  %readonly%|  
|optionalAttendees      | [EmailAddressDetails[]](emailaddressdetails.md) or [Recipients](recipients.md) | Gets or sets a list of email addresses for optional attendees. | 1.0  %readonly%|  
|organizer      | [EmailAddressDetails](emailaddressdetails.md) | Gets the email address of the meeting organizer for a specified meeting. | 1.0  %readonly%|  
|requiredAttendees      | [EmailAddressDetails[]](emailaddressdetails.md) or [Recipients](recipients.md) | Gets or sets a list of email addresses for required attendees. | 1.0  %readonly%|  
|resources      | [EmailAddressDetails](emailaddressdetails.md) | Gets the resources required for an appointment. | 1.0  %readonly%|  
|sender      | [EmailAddressDetails](emailaddressdetails.md) | Gets the email address of the sender of an email message. | 1.0  %readonly%|  
|start      | Date or [Time](time.md) | Gets or sets the date and time that the appointment is to begin. | 1.0  %readonly%|  
|subject      | String or [Subject](subject.md) | Gets or sets the description that appears in the subject field of an item. | 1.0  %readonly%|  
|to      | [EmailAddressDetails[]](emailaddressdetails.md) or [Recipients](recipients.md) | Gets or sets the recipients of an email message. | 1.0  %readonly%|  



## Methods

| Method	   | Return Type    | Description | Requirements|
|:-------------|:---------------|:------------|:----|
| [addFileAttachmentAsync](addfileattachmentasync)     |  | Adds a file to a message or appointment as an attachment.  | 1.1|  
| [addItemAttachmentAsync](additemattachmentasync)     |  | Adds an Exchange item, such as a message, as an attachment to the message or appointment.  | 1.1|  
| [close](close)     |  | Closes the current item that is being composed.  | 1.3|  
| [displayReplyAllForm](displayreplyallform)     |  | Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.  | 1.0|  
| [displayReplyForm](displayreplyform)     |  | Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.  | 1.0|  
| [getEntities](getentities)     | [Entities](entities.md) |   | 1.0|  
| [getEntitiesByType](getentitiesbytype)     | String[] or [Contact[]](contact.md) or [MeetingSuggestion[]](meetingsuggestion.md) or [PhoneNumber[]](phonenumber.md) or [TaskSuggestion[]](tasksuggestion.md) | Gets an array of all the entities of the specified entity type found in the selected item.  This method can return null. | 1.0|  
| [getFilteredEntitiesByName](getfilteredentitiesbyname)     | String[] or [Contact[]](contact.md) or [MeetingSuggestion[]](meetingsuggestion.md) or [PhoneNumber[]](phonenumber.md) or [TaskSuggestion[]](tasksuggestion.md) | Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.  This method can return null. | 1.0|  
| [getRegExMatches](getregexmatches)     | Object | Returns string values in the selected item that match the regular expressions defined in the manifest XML file.  | 1.0|  
| [getRegExMatchesByName](getregexmatchesbyname)     | String[] | Returns string values in the selected item that match the named regular expression defined in the manifest XML file.  This method can return null. | 1.0|  
| [getSelectedDataAsync](getselecteddataasync)     | String | Asynchronously returns selected data from the subject or body of a message.  | 1.0|  
| [loadCustomPropertiesAsync](loadcustompropertiesasync)     |  | Asynchronously loads custom properties for this add-in on the selected item.  | 1.0|  
| [removeAttachmentAsync](removeattachmentasync)     |  | Removes an attachment from a message or appointment.  | 1.1|  
| [saveAsync](saveasync)     |  | Asynchronously saves an item.  | 1.3|  
| [setSelectedDataAsync](setselecteddataasync)     |  | Asynchronously inserts data into the body or subject of a message.  | 1.2|  

