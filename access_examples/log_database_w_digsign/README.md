# Access Examples

## HWLL_Log_Database

This is a simple Access database to track and manage Test Logs, it allow to insert, view and print logs including attachments. It offer as option the digital signature using GnuPG (gpg4win).

### Before Start

The database is split in front-end and back-end, at first run, open the Linked Table Manager and select the folder where the back-end has been located. The back-end can be located on a local or network share.

**Attachments**

It use a standard Microsoft Access feature to handle attachments, that are stored (and compressed) directly into the database file. As result the database file increas as increas the number of attachments. **Is not required to store the attachments into a dedicated folder.**

**Export Macro**

A set of export macro allow to export logs, attachments and relevant index quickly. Most of macro are *_byDate* and export only data that match the date range typed into *Select Date*.

**Multi-user**

Distribute the front-end file to all people that need to access the database remotely. To use the digital signature remotely, you need to setup gpg4win as server or have local copy of the certificates.

**Digital Signature**
Logs can be signed individually, this allow to identify if the content has been modified after the signature.

To use this feature follow these steps:
1) Install gpg4win
2) Insert the folder where gpg4win has been installed in the PATH
3) Using Kleopatra (the GUI for gpg4win) in the settings, remove PIN cache
4) Using Kleopatra import or build local certificate, each name listed in PEOPLE table of the Access database shall have an individual certificate with the same spelled as in that table. 
