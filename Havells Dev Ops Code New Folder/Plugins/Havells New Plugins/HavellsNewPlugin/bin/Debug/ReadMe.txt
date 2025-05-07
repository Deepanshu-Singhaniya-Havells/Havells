Step 1: Open CMD from debug Folder.
setp 2: Run this command --:    Ilmerge  /keyfile:"HavellsNewPlugin.snk"  /copyattrs /allowDup /targetplatform:v4 /out:"..\Debug\MergFile\HavellsNewPlugin.dll"  "HavellsNewPlugin.dll" "Newtonsoft.Json.dll" "RestSharp.dll"
Step 3: now open MergFile folder and deploy this DLL.