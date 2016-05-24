# userProfile resource type

##### Namespace: *Office.context.mailbox*

Provides information about the user in an Outlook add-in.



|Requirement| Value|
|:----------|:-----|
|Minimum requirement set/version: | *1.0*|
|Minimum permission level |*ReadItem* |
|Modes supported | *Read, Compose*|


### Properties

| Property	   | Type	| Description| Requirements|
|:-------------|:-------|:-----------|:------------|
|displayName      | String | Gets the user's display name. | 1.0  %readonly%|  
|emailAddress      | String | Gets the user's SMTP email address. | 1.0  %readonly%|  
|timeZone      | String | Gets the user's default time zone. | 1.0  %readonly%|  


