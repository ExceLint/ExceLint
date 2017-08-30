These PRQ files are necessary for InstallShield to correctly locate redistributable libraries.  You should put them in:

C:\Program Files (x86)\InstallShield\2015LE\SetupPrerequisites

You will need to download the linked restributable libraries prior to building the installer.  Restart Visual Studio as Administrator, go to the InstallShield Visual Studio project, find the new entries in "Specify Application Data" -> "Redistributables", find your new libraries, right-click and select "Download Selected Item..."

Quit Visual Studio and restart as a regular user.  You should now be able to build with the missing prerequisites.
