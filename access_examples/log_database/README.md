# Access Examples

## Introduction

This is a collection of Microsoft Access database used for specific purposes, like build a Bill of Materials or manage test logs.

### HWLL_Log_Database_Empty

This is a simple Access database to track and manage Test Logs, it allow to insert, view and print logs including attachments.

**Attachments**

It use a standard Microsoft Access feature to handle attachments, that are stored (and compressed) directly into the database file. As result the database file increas as increas the number of attachments. **Is not required to store the attachments into a dedicated folder.**

**Export Macro**

A set of export macro allow to export logs, attachments and relevant index quickly. Most of macro are *_byDate* and export only data that match the date range typed into *Select Date*.

**Multi-user**

Use the Microsoft Access *split* function to create a *Back End* database that can be saved into a network folder for shared access. Only the *Front End* database shall be shared (it contains forms, query and report) as all data are stored in the *back end*.
