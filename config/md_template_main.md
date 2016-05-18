<resource>
# %resourcename% resource type

*Namespace: %resourcenamespace%*

*Minimum requirement set/version: %minreqset%*

*Minimum permission level: %minpermission%*

*Modes supported: %modes%*


%resourcedescription%

%longobjectdescription%

</resource>

<properties>
### Properties

| Property	   | Type	| Description| Requirements|
|:-------------|:-------|:-----------|:------------|
>|%name%      | %type% | %description% | %req% |

%propertygetset%
%propertynotes%
</properties>

<enums>
### Enumerations

| Option	   | Type	| Description| Requirements|
|:-------------|:-------|:-----------|:------------|
>|%name%      | %type% | %description% | %enumreq% |

%propertygetset%
%propertynotes%

</enums>

<relationships>
### Relationships
| Relationship | Type	| Description| Requirements|
|:-------------|:-------|:-----------|:------------|
>|%name%      | [%type%](%link%) | %description% | %req% |

%relationshipnotes%
</relationships>

<methods>

## Methods

| Method	   | Return Type    | Description | Requirements|
|:-------------|:---------------|:------------|:----|
>| [%name%](%link%)     | %dtype% | %description% | %req%|

%methodnotes%

