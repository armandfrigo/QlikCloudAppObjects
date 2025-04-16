# QlikCloudAppObjects
This PowerShell Script will extract Qlik Cloud application objects, including masterItems.

(Fast) Execution to fetch all app objects without details about Master Items:
.\QlikCloudFetchAppObjects.ps1 or .\QlikCloudFetchAppObjects.ps1 -masterItems $false

(Slow) Execution to fetch all app objects, including details about Master Items:
.\QlikCloudFetchAppObjects.ps1 -masterItems $true

How it works:
This script will:
1. fetch the list of spaces and create a mapping
2. fecth the list of applications
3. filter the list of applications based on space type (managed only) and space names (regex example)
4. for every app in scope, fetch the list of app objects. Note that master items don't have a visualization type and name
   qlik app object ls -a $app.resourceId --no-data --json
5. if param masterItems is $true, it will fetch details for the master items
   qlik app object properties -a $app.resourceId $object.qId --no-data --json
   Note: this is a mess since every extension vendor can use key names that are very close, like yaxiscolor and yAxiscolor. Powershell cannot handle case-sensitive Json attributes, hence I eventually used string functions. 
6. it will generate a CSV file and upload it to Azure Storage
   CSV header: 'Space Name,Space Type,App ID,App Name,Object ID,Is Master Object,Object Type,Object Title'

There is a retry mechanism, 3 times for each master item.

Testing:
You can use the local or server mode, update the paths accordingly.
There are placeholders for you to test for one app and/or one app object.

Pre-requisites:
PowerShell 5.1 or above
Azcopy, but you can adapt for Google CLI or AWS CLI, or copy locally too
