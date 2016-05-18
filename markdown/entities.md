# Entities resource type

Represents a collection of entities found in an email message or appointment.

The `Entities` object is a container for the entity arrays returned by the `getEntities` and `getEntitiesByType` methods when the item (either an email message or an appointment) contains one or more entities that have been found by the server. You can use these entities in your code to provide additional context information to the viewer, such as a map to an address found in the item, or to open a dialer for a phone number found in the item.

*	Namespace: *Entities*
*	Minimum requirement set/version: *1.0*
*	Minimum permission level: *ReadItem*
*	Modes supported: *Read*


