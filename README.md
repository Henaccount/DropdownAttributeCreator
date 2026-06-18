<pre>
Sample Code, AI generated, use at own risk!
  there are two tools now (Mode1 / Mode2):
    - Dropdown creation
    - Attribute filling

  This readme file explains both tool respectively:
  
ACC Attribute Tool - PowerShell 5.1 version
===========================================

Purpose
-------
Runs ACC/BIM 360 Docs custom attribute workflows using only Windows onboard technology:
PowerShell 5.1, built-in .NET, JSON text files, and REST API calls.

No Visual Studio, .NET SDK, Excel library, or third-party install is required.

Files
-----
AccAttributeTool.ps1                       Main PowerShell script
Create-AccDropdownAttributes.ps1           Backward-compatible wrapper
Run-AccAttributeTool.cmd                    Double-click launcher, auto mode
Run-DryRun.cmd                              Double-click dry-run launcher, auto mode
Run-PopulateFileAttributes.cmd             Double-click launcher for folder-derived file attribute updates
Run-PopulateFileAttributes-DryRun.cmd       Double-click dry-run launcher for folder-derived file attribute updates
sample-project.json                         Example for creating dropdown definitions
sample-populate-file-attributes.json        Example for populating file attributes from folder names

Authentication
--------------
Preferred: set these environment variables before running:

  set APS_CLIENT_ID=your_client_id
  set APS_CLIENT_SECRET=your_client_secret

Or run the tool and paste the credentials when prompted.

For quick testing, you can also provide a ready-made token:

  set APS_ACCESS_TOKEN=your_access_token

Default scopes are:

  data:read data:write

You can override scopes in the JSON file with a scopes array.

Mode 1: create dropdown definitions
-----------------------------------
Use sample-project.json as the template.

This mode creates ACC Docs dropdown custom attribute definitions.
Each attribute object creates one dropdown definition. The script sends type "array" and arrayValues.

Example:

{
  "mode": "create-dropdown-definitions",
  "projectId": "b.YOUR_PROJECT_ID_OR_GUID",
  "folderIdTemplate": "urn:adsk.wipprod:fs.folder:co.{0}",
  "folders": [
    {
      "folderId": "urn:adsk.wipprod:fs.folder:co.FOLDER_ID",
      "attributes": [
        { "name": "Status", "values": ["Draft", "For Review", "Approved"] }
      ]
    }
  ]
}

Mode 2: populate file attributes from folder names
-------------------------------------------------
Use sample-populate-file-attributes.json as the template.

This mode expects this folder structure below rootFolderId:

  customer1
      123.234.567
      123.234.568
  customer2
      123.234.569
      123.234.510

For all files in matching machine folders, the script writes:

  customerAttributeName = customer folder name
  machineAttributeName  = machine folder name

Machine folders are processed only when their name matches this default pattern:

  ^\d{3}\.\d{3}\.\d{3}$

Other folders below customer folders are skipped.

The two custom attributes must already exist as text/string custom attributes on the specified root folder, so that files below the root folder can receive these values.

Example:

{
  "mode": "populate-file-attributes-from-folders",
  "projectId": "b.YOUR_PROJECT_ID_OR_GUID",
  "rootFolderId": "urn:adsk.wipprod:fs.folder:co.ROOT_FOLDER_ID",
  "customerAttributeName": "Customer",
  "machineAttributeName": "Machine",
  "machineFolderNamePattern": "^\\d{3}\\.\\d{3}\\.\\d{3}$",
  "scopes": ["data:read", "data:write"]
}

Run
---
Double-click:

  Run-DryRun.cmd

Select the JSON file. Inspect the console output and CSV log.

When ready, double-click:

  Run-AccAttributeTool.cmd

For the new folder-derived update mode, double-click:

  Run-PopulateFileAttributes-DryRun.cmd

Then:

  Run-PopulateFileAttributes.cmd

Command-line examples
---------------------
Create dropdown definitions dry-run:

  powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\AccAttributeTool.ps1 -InputPath .\sample-project.json -DryRun

Populate file attributes dry-run:

  powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\AccAttributeTool.ps1 -InputPath .\sample-populate-file-attributes.json -Mode populate-file-attributes-from-folders -DryRun

Populate file attributes for real:

  powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\AccAttributeTool.ps1 -InputPath .\sample-populate-file-attributes.json -Mode populate-file-attributes-from-folders

Important notes
---------------
The APS app must be allowed/provisioned for the ACC/BIM 360 account/project and must have the required permissions.

The projectId can be supplied with or without the leading b. prefix. The script automatically uses the b. format for Data Management folder traversal and the non-b. format for BIM 360 Docs custom attribute endpoints.

The folder traversal uses the Data Management API folder contents endpoint. File updates use each item's latest tip version ID and the BIM 360 Docs custom-attributes:batch-update endpoint.

If token creation succeeds but API calls return 401 or 403, the cause is usually account/project access, app provisioning, folder permission, or auth-context support rather than the JSON parser.

  
</pre>
