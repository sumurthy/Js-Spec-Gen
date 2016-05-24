# diagnostics resource type

##### Namespace: *Office.context.mailbox*

Provides diagnostic information to an Outlook add-in.

Provides diagnostic information to an Outlook add-in.

|Requirement| Value|
|:----------|:-----|
|Minimum requirement set/version: | *1.0*|
|Minimum permission level |*ReadItem* |
|Modes supported | *Read, Compose*|


### Properties

| Property	   | Type	| Description| Requirements|
|:-------------|:-------|:-----------|:------------|
|OWAView      | String | Gets a string that represents the current view of Outlook Web App. | 1.0  %readonly%|  
|hostName      | String | Gets a string that represents the name of the host application. | 1.0  %readonly%|  
|hostVersion      | String | Gets a string that represents the version of either the host application or the Exchange Server. | 1.0  %readonly%|  


