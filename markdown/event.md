# Event resource type

##### Namespace: *Event*

Passed as a parameter to UI-less command functions in an Outlook add-in. Used to signal completion of processing.

The `event` object is passed as a parameter to add-in functions invoked by UI-less command buttons. The object allows the add-in to identify which button was clicked and to signal the host that it has completed its processing. 	 
 	 
For example, consider a button defined in an add-in manifest as follows: 	 
 	 
<Control xsi:type="Button" id="eventTestButton"> 	 
<Label resid="eventButtonLabel" /> 	 
<Tooltip resid="eventButtonTooltip" /> 	 
<Supertip> 	 
<Title resid="eventSuperTipTitle" /> 	 
<Description resid="eventSuperTipDescription" /> 	 
</Supertip> 	 
<Icon> 	 
<bt:Image size="16" resid="blue-icon-16" /> 	 
<bt:Image size="32" resid="blue-icon-32" /> 	 
<bt:Image size="80" resid="blue-icon-80" /> 	 
</Icon> 	 
<Action xsi:type="ExecuteFunction"> 	 
<FunctionName>testEventObject</FunctionName> 	 
</Action> 	 
</Control> 	 
 	 
The button has an `id` attribute set to `eventTestButton`, and will invoke the `testEventObject` function defined in the add-in. That function looks like this: 	 
 	 
function testEventObject(event) { 	 
// The event object implements the Event interface 	 
 	 
// This value will be "eventTestButton" 	 
var buttonId = event.source.id; 	 
 	 
// Signal to the host app that processing is complete. 	 
event.completed(); 	 
}

|Requirement| Value|
|:----------|:-----|
|Minimum requirement set/version: | *1.3*|
|Minimum permission level |*Restricted* |
|Modes supported | *Read, Compose*|


### Properties

| Property	   | Type	| Description| Requirements|
|:-------------|:-------|:-----------|:------------|
|source      | Object | Gets the identifier of the add-in command button that invoked the method. | 1.3  %readonly%|  



## Methods

| Method	   | Return Type    | Description | Requirements|
|:-------------|:---------------|:------------|:----|
| [completed](completed)     |  | Indicates that the add-in has completed processing that was triggered by an add-in command button.  | 1.3|  

