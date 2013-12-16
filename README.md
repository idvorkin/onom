Onenote Object Model (onom)
====

The OneNote Object Model exposed via C#. This project contains the following sub directories:

* OneNoteObjectModel - An assembly exposing the ObjectModel
* ListNotebooks - An example showing listing the notebooks. 

Many OneNote APIs, like GetHierarchy, are not conveniently typed, so ONOM implements typed helpers, like GetNotebooks, which are exposed on the OneNoteApp class. 


TODO List:

* Create a NuGet project for easy access from other projects and LinqPad
* Complete typed helpers on OneNoteApp
* Add a unit test project

