# Redweed_Excel_Strategy_Game
A vintage strategy game converted play in Microsoft Excel

IMPORTANT! Open a new, blank workbook before double-clicking the game file.

Here's something I've been working on for a while that started with an idle thought at work, wondering if I could make the cells in an Excel spreadshhet
small enough to look like pixels on a screen. That thought developed into wondering if I could colour the pixels to mimic a print command to display words
then trying to mimick a basic game and finally if I could write the game code in VBA such that it be as similar to the original code of the game as possible.

The game I chose for this project is a strategy game called RedWeed published in 1984 by company M. C. Lothlorien for the Sinclair ZX Spectrum 48K
(a popular UK home computer from the 1980s). It is written in Sinclair BASIC and simulates the attack by Allied forces to destroy invading Martian war
machines.

To play, download the spreadsheet. Open a new, blank workbook then double-click on the game file to begin. 
If you double click the game file without opening a new, blank workbook first, the game will not run (if anyone knows why this is and has a solution,
please let me know)
You may have to resize the game window depending on the resolution of your screen. Instructions are included but effectively the game runs as follows:
Red weed grows
Attack redweed if possible.
Move allied forces
Attack Martians if possible
Martian movement
Martian attack
There is an additional phase at the start of the game where you can place mines in the hope of damaging the Martian tripods.

As you'll appreciate, Excel is not a game platform and consequently the game is quite fragile. Clicking another Excel spreadsheet while running the game
will cause it to crash. If the game is displaying text, it will display the text in the new active window then crash.

Some notes for those interested:
While the original game was keyboard only, this version is mouse only. I found that the game recognised mouse input but could not recognise keyboard input.
There is no sound. I developed this game in my lunch hours at work and did not want to subject my work colleagues to constant bleeps and bloops. Honestly,
I'm not even sure if sound could be incorporated as I never looked into it.
The speed of execution is much slower than the native machine it was originally written for. Please be patient. If the blue circle cursor is showing the 
computer is thinking. Also, I've found the final victory or defeat screen takes a couple of minutes to render. Give it time!
If you are skilled enough to destroy all the Martians invaders or if all your forces are destroyed or a Martian reaches the right of the screen, click X to 
see the final screen.
Click < or > to move through the instruction pages.
Click v or ^ to cycle through the game level choices. Click the number to select it. Note that the game is wildly unbalanced. It gets very hard very quickly
as you choose the later levels.
Note that you can only select a unit to move once per turn. If you select a unit then deselect it without moving, you cannot reselect it until the next turn.
I've updated the graphics, added simple animations and added more text options. This was because the original game was a bit plain and repetitive. This has
all been done in subfunctions so the main game code remains as similar to the original Sinclair BASIC code as possible.

For the Spectrum purists:
Yes, I know the algorithm to draw the Sun CIRCLE is incorrect. The original Spectrum algorithm is Z80 code which I can't read. I went with an algorithm that 
worked. If I study Z80 in future I'll update the CIRCLE draw algorithm.
I haven't imitated the Spectrum's infamous colour clash because life is too short. It would only affect the Martian laser flash on the title screen anyway.
Due to the PRINT command being so much slower than the original Spectrum the FLASH text command only really works on single characters. I've retained the 
FLASH for units in range when attacking Martians but the long line of FLASHing text in the game has been mostly redacted (though there is an attempt with 
the win/lose text message).
