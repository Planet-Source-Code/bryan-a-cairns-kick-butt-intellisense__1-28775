Bryan Cairns
11/09/01

Before Attempting to Load this project, you need to do a few things.

1. Download and Register the CodeSense Ocx Control (Version 2.1.0.16)
   I know a lot of you hate downloading controls, but this one is definatly worth
   every second of your download.
   http://www.ticz.com/homes/users/nlewis/index.html?target=intro

2. Make sure you have MS DAO 3.51 Object Libray installed.
   If you have no idea if this is installed, don't worry, VB usally installs it
   as a default option.


Now for the good stuff...

This script editor is in the beginning stages, but is already very advanced.
If you are coding in one of the following languages, this this will save you a ton
of time.

VB Script
Java Script
Delphi Script
D+ (on Planet Source Code)
DM+ (on Planet Source Code)
Jel (on Planet Source Code)
WinScript (on Planet Source Code)

Current Features:
(Most of these come from the CodeSense Ocx you downloaded)
Intellisense 
Syntax Highlighting
Smart Indentation
Find / Replace
Code Bookmarking
Code Snipplets
Unlimited Undo / Redo
Line Numbering
And Much More


The Data Base
If you have no noticed, there is a small Access 97 database included with the project.
Please make sure this file is on the app.path.

This database will allow you to add / edit / remove items for your Intellisense, code snipplets,
and for your Subs and Functions

To see how the database is used, open it up, and look in the "CodeEdit" table...
You will see five fields:
Type - Used to determine what type of record this is
IconType - Used to determine the Icon in the Intellisense Drop Down List
Header - Used to determine the starting variable (map.something)
Data - This is the code added to the editor when this record is choosen
ExtraData - Not in use yet

So just for kicks, start the program and type in "Device" and then a period "."
you will see how the database is used to generate the dropdown list.

Bryan Cairns
cairnsb@html-helper.com


