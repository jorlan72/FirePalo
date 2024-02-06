# FirePalo
FirePalo (Windows GUI) helps you convert rules and objects from Cisco FirePower to Palo Alto

(See the "Sceenshots from the application.docx")

Run "show access-control-config" from the FTD device and save output to a textfile.
Open the textfile in FirePalo.exe and it will create editable objects.
Finally, "commit" the changes and create a configuration in SET format that can be pasted into a Palo Alto device or Panorama.

This version will not convert device configuration like interfaces, routing or NAT.
Some manual work needed for User-ID, URL Categories and Application filters.

Download the PaloAppID.txt file and place it with the FirePalo.exe. It contains all the Palo Alto APP-ID's

FirePalo also lets you export sections of the configuration to edit in preferred editor and than import the result back (great for search and replace).
In addition you can easily lowercase or uppercase sections (or the entire configuration) and cut object names automatically to supported length.
Further, you can convert services to applications (as not all services from FTD are supported as a service).
Finally, you can add tags for objects, so that all rules using a certain object get the tag set.

Easily select if this is a standalone or Panorama configuration to be created (so that device group get included in the configuration).

FirePalo takes the output from the FTD and first turns it into a treeview. It then takes all the items in the treeview and creates objects you can edit, providing an unique ID for each object.
This binds everything to the correct rules and all edits will be in place when you finally turn the objects into a treeview again ("commit").
You can then look through the result as a treeview and make more changes if needed (and then doing a new commit).

When everything looks good, you can generate the final configuration in SET format and paste it into the Palo Alto device or Panorama CLI.

https://www.linkedin.com/pulse/migrating-firepower-palo-alto-j%25C3%25B8rgen-lanesskog-d4gyf
