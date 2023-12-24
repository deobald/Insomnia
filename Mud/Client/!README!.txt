Insomnia ReadMe [To be read with a simple text viewer with a fixed font and wordwrap on. Preferrably NotePad+.]

Keyboard Instructions
	> Movement: Arrow Keys
	> Sending commands over internet: <Enter> Key

Mouse Instructions
	> Click anywhere on the screen to stop typing and being playing again
	> Click on the "Exit" button to leave the game
	> Click on the lower text box to enter commands
	> Double-Click anywhere on the screen to exit the game (Temporary - Used as a failsafe)

Player Selection
	> Click on (or near) the character image to rotate through the characters
	> Click the "Enter" button to play using that character

TroubleShooting
	
PROBLEM: I get a "Component 'MSWINSCK.OCX' not correctly registered: File is missing or invalid." error message when I start up the game.
SOLUTION: Try any/all of the following fixes:
(1) Make sure you have the file "OC30.DLL" in your /Windows/System/ directory.
(2) Make sure you have the file "MSWINSCK.OCX" in your /Insomnia/ directory, or in the /System/ directory
(3) Make sure you have files "MSWINSCK" with all of these extensions: .DEP, .CNT, .OCA, .FTS, and (this one is optional) .HLP
(4) If none of these work, find your "MSWINSCK.OCX" file and use it's location in place of the following run sequence's <Winsock File>, and run:

	regsvr32.exe <Winsock File> (ie. C:\insomnia\mswinsck.ocx )
	[Regsvr32 should be in the \Windows\ or \Windows\System\ directory.]

FAQ
	Q: How large are Insomnia's maps?
	A: 50 * 50 tiles, 1.4 m * 1.4 m, or 56 in. by 56 in.

	Q: Can I work on the Insomnia project, or any other project that PseudoSoft is currently working on?
	A: Apply by visiting www.net.gull-lake.sk.ca/pp/OSoft/ or email std@leo.net.gull-lake.sk.ca

System Requirements

CPU: 486 or better (Pentium recommended)
RAM: 8 Megabytes (16 recommended)
Free Drive Space: 5 Megabytes or more
Video Card: 2 Megabytes of video RAM and 256 color, 640*480 resolution (64K color HIGHLY recommended)
Sound Card: Optional
Internet Connection: 14.4 kbps connection or better (33.6 recommended)

Disclaimer - PLEASE READ
	> PseudoSoft will not be held responsible for any unintentional damage done to computers with PseudoSoft software installed, or their data withheld.