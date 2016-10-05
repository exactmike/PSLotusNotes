# PSLotusNotes
Functions to retrieve Lotus Notes Objects via PowerShell using persistent connections
# Requirements
Notes Client or Domino Server installation for access to Lotus.NotesSession COM objects
# Usage
## Create one or more Notes Sessions and Database connections using New-NotesDatabaseConnection.  When doing so, specify a "friendly" name for each database.  
## Use Get-NotesUser (and other additional functions that may be added in the future) and specify the "friendly" name of the database in the -NotesDatabase parameter.  


