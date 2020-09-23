General Description
This is a VB6  Add-In which allows you to indent code, find and repair/improve code for greater efficiency and readability and includes a user friendly replacement for VB's Find And Replace Tool.

WARNING
All care; No responsibility.  Some of the fixes that the Add-In performs are capable of damaging code. Please use on copies of your code or use the backup systems built into the program. 

INSTALLATION & DEPENDENCIES 

INSTALLATION
1. Load project into VB. 
2. Check that you have the right references (just try to run code) and select from Project|References menu item (if something is missing). See Dependencies for details and See Note A below
3. Compile the Dll. By default it should get written into the VB folder if you have used the standard install path('C:\Program Files\Microsoft Visual Studio\VB98' on my machine). (See Notes B & C). If you have a non-standard installation you can locate your VB folder ('VB98') and install it there or anywhere you find convenient (perhaps a sub-folder called 'Add-ins' below VB98. 
4. Close VB. 
5. Copy the help file 'Code Fixer 2.chm' to the same folder as the Dll. (See comment below) 
6. Restart VB 
7. If the Add-in does not appear automatically then open Add-in menu and click the menu item. 
8. If not in the menu then 
Open Add-In Manager. 
Find 'Code Fixer 2' and activate it. 
9.Take it for a test run on some code. (Download something good and something bad and run it on them just to see how it copes) 
--------------------------------------------------------------------------------
HELP FILE 
Because of its size the Help file is available for download at a separate site at PSC. 
Do a Quick Search for CODE FIXER HELP FILE and take the latest version. 
NOTE The help file may be slightly out-of-date as I only update it when I add new fixes or change the interface, not just for bug fixes (See History.txt for these) 
I have done this so that you can download updates without having to download the help file every time. 
--------------------------------------------------------------------------------
NOTES 
A. As Code Fixer is an Add-in running it in the IDE will not show anything. If you really want to see it without compiling; Run the code. Leave it running and open some other code in another instance of VB and see steps 7 & 8 above. If you run Code Fixer this way you can only have one other instance of VB open or it gets confused. Compiled you can run several instances of VB but you will be using a lot of memory. 
B. If you already have Code Fixer installed, you must first deactivate the old one. Open Add-in Manager and unload the current version (top CheckBox bottom right)) 
C. If you have installed my Extended Find Add-In you may as well deactivate it as Code Fixer 2 incorporates all its functionality (and contains a few minor bug fixes). 

DEPENDENCIES 
The Learning edition of VB6 doesn't support Add-ins.
VB6 SP5 or SP6; earlier SP versions have trouble with some ListView code used to set options
Reference to MicroSoft OFfice 8 (or better) All add-ins need access to this for button/menu/toolbar usage. 'As Is' the code expects version 8 just update to the version you have if 8 is not available. 
The Help file: Code Fixer is too powerful and complex to run with a 'suck it and see' approach. You should at least look through the help file before using it on major code.


DOCKING
The Find Tool component is designed to be permanently docked with the VB IDE( See 'Launch on Startup' checkbox on Settings screen). The tool docks with the IDE (best at top or bottom; toolbar and input boxes can be positioned at top or bottom of tool window) so the code is never hidden behind the Tool and presents any finds in a list. While this takes screen real estate, unlike VB's own find tool it doesn't hide the code. I find it easiest to use if docked at top of the IDE under the VB toolbars.
UserDocument tools remember where you placed them and what size you set them to. On initial run the tool will appear as a floating tool; just drag to the edge of the IDE area that you want it to use.
To re-size, drag the edge of the tool. To change docking position, drag the caption bar. On initial run the tool will appear as a floating tool; just drag to the edge of the IDE area that you want it to use.

BUG REPORTS
If you find a bug please let me know at the address below. Please include the version number, damaged code  (routine, module or whole code if not too large) with Code Fixer comments, a copy of the original working code and any additional comments. I do not have personal access to the net so may not be able to test your code fully but should be able to work out what is wrong. ;) I can usually turn around a bug in 2-3 days. Suggestions (other than bugs) may take longer depending on how useful/necessary I find it but are most welcome and if small may also only take a couple of days.

CONTACT
Roger Gilchrist
rojagilkrist@hotmail.com

