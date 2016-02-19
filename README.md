# VBECustomColors
The VBA / VB editor (or VBE) is limited in the 16 colors it can use to render code text.  

The project aims to serve as a guide for adding your own custom colors to the palette.


<img src="https://raw.githubusercontent.com/dimitrimitropulos/VBECustomColors/master/ExampleColors.png">

There are two main things you need to know about getting your editor look like this:
1. The actual hex values for the color palette are stored in the editor dll or exe (VB6.exe, VBE7.dll, etc.).  To update the colors you have to open the editor dll or exe file in a hex editor (I use HxD) and manually make the changes there.
2. The registry holds what you can think of as a preset.  This is located somewhere around "C:\Program Files (x86)\Common Files\microsoft shared\VBA\" (in the folder "VBA7.1\VBE7.DLL" for Office 2016, for example).  You can either make changes to what colors are assigned to what syntax groups (like "Selection Text", "Execution Text", etc.) in the registry all at once, or you can do it manually in the editor ( Tools | Options | Editor Format)

# Quick Start
Want to get your Office 2016 editor looking like the picture above super fast?  Heres a quick and dirty guide:

0. Make sure anything like excel related to the editor is closed.

1. Locate your DLL: should be at `C:\Program Files (x86)\Common Files\microsoft shared\VBA\VBA7.1\VBE7.DLL`

   *  Make a copy and rename VBE7orig.DLL just in case you make a mistake.

2. Open the DLL in HxD

3. Search for the first instance of the following ( Search | Find, Data Type = "Hex-Values" Search Direction="All": 

   ff ff ff 00 c0 c0 c0 00 80 80 80 00 00 00 00 00 ff 00 00 00 80 00 00 00 ff ff 00 00 80 80 00 00 00 ff 00 00 00 80 00 00 00 ff ff 00 00 80 80 00 00 00 ff 00 00 00 80 00 ff 00 ff 00 80 00 80 00

   *  Replace that text with the following:

   00 00 00 00 1e 1e 1e 00 34 3a 40 00 3c 42 48 00 d4 d4 d4 00 ff ff ff 00 26 4f 78 00 56 9c d6 00 74 b0 df 00 79 4e 8b 00 9f 74 b1 00 e5 14 00 00 d6 9d 85 00 ce 91 78 00 60 8b 4e 00 b5 ce a8 00 

5. GO TO THE TOP (scroll up and click on the top value or something)

6. Search for the SECOND instance of ( hit F3 to find next):

   00 00 00 00 00 00 80 00 00 80 00 00 00 80 80 00 80 00 00 00 80 00 80 00 80 80 00 00 c0 c0 c0 00 80 80 80 00 00 00 ff 00 00 ff 00 00 00 ff ff 00 ff 00 00 00 ff 00 ff 00 ff ff 00 00 ff ff ff 00

    * Replace that text with the following:

   00 00 00 00 1e 1e 1e 00 34 3a 40 00 3c 42 48 00 d4 d4 d4 00 ff ff ff 00 26 4f 78 00 56 9c d6 00 74 b0 df 00 79 4e 8b 00 9f 74 b1 00 e5 14 00 00 d6 9d 85 00 ce 91 78 00 60 8b 4e 00 b5 ce a8 00

OPTIONAL:

5  Navigate in regedit to HKEY_CURRENT_USER\Software\Microsoft\VBA\6.0\Common\CodeBackColors

   * change value to : 2 7 1 13 15 2 2 2 11 9 0 0 0 0 0 0 

6  change CodeForeColors to : 13 5 12 1 6 15 8 5 1 1 0 0 0 0 0 0 

7  change FontFace to : Consolas


