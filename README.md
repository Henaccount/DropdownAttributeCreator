<pre>
Sample Code, AI generated, use at own risk!
  
ACC Dropdown Attribute Tool - PowerShell 5.1 version
===================================================

Purpose
-------
Creates ACC/BIM 360 Docs dropdown custom attribute definitions from a JSON text file.
No Visual Studio, .NET SDK, Excel library, or third-party install is required.

Files
-----
Create-AccDropdownAttributes.ps1  Main PowerShell script
Run-AccAttributeTool.cmd          Double-click launcher
Run-DryRun.cmd                    Double-click dry-run launcher
sample-project.json               Example input file

Authentication
--------------
Preferred: set these environment variables before running:

  set APS_CLIENT_ID=your_client_id
  set APS_CLIENT_SECRET=your_client_secret

Or run the tool and paste the credentials when prompted.

For quick testing, you can also provide a ready-made token:

  set APS_ACCESS_TOKEN=your_access_token

Input JSON
----------
Use sample-project.json as the template.

projectId:
  The ACC/BIM 360 Docs project ID expected by the API. Keep the leading b. if your project ID has one.

folderId:
  Either a full folder URN, such as:
    urn:adsk.wipprod:fs.folder:co.xxxxx
  or a bare folder ID/GUID if folderIdTemplate is provided.

folderIdTemplate:
  Optional. Used only when folderId does not start with urn:.
  Example:
    urn:adsk.wipprod:fs.folder:co.{0}

attributes:
  Each attribute object creates one dropdown definition.
  type is sent as "array" and values are sent as "arrayValues".

Run
---
Double-click:

  Run-DryRun.cmd

Select your JSON file. Inspect the console output and CSV log.

When ready, double-click:

  Run-AccAttributeTool.cmd

Command-line examples
---------------------
Dry-run:

  powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Create-AccDropdownAttributes.ps1 -InputPath .\sample-project.json -DryRun

Create attributes:

  powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Create-AccDropdownAttributes.ps1 -InputPath .\sample-project.json

Notes
-----
The APS app must be allowed/provisioned for the ACC/BIM 360 account/project and must have the required permissions. If token creation succeeds but API calls return 401/403, the problem is usually account/project access, app provisioning, or unsupported auth context rather than the JSON parser.

  
</pre>
