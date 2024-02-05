# FirePalo
FirePalo (Windows GUI) helps you convert rules and objects from Cisco FirePower to Palo Alto

Run "show access-control-config" from the FTD device and save output to a textfile.
Open the textfile in FirePalo.exe and it will create editable objects.
Finally, "commit" the changes and create a configuration in SET format that can be pasted into a Palo Alto device or Panorama.

This version will not convert device configuration like interfaces, routing or NAT.
Some manual work needed for User-ID, URL Categories and Application filters.

Download the PaloAppID.txt file and place it with the FirePalo.exe. It contains all the Palo Alto APP-ID's
