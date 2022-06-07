# DrawingPropertyChangePropertyTypeBug_DocumentManager
SOLIDWORKS Document Manager : This repo showcases an issue with deleting a property and re-adding the same property with a different type. Works fine for parts but not drawings. 


# Steps to recompile and replicate issue
- Git clone the repo.
- Open sln in VS. 
- Restore nuget packages.
- Make sure swdocumentmgr.dll Embed Interop Types = false
- Create a file called DocManKey.txt at the root of the solution and put your document key manager there (must be 2021 or newer)
- Inspect the Vendor Date property in the part and the drawing files. Files are at the root. Observe how the type of the Vendor date property is Date. 
- Compile. 
- Run. 
- For both the part and drawing (from the debug folder), inspect the Vendor Date property in SW.
- Observe how the Vendor Property type for the part is text (which is true) and Date for the drawing (which is not the expected value).
