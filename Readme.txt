*Notes for JimCamels RPGame*
(Put on wordwrap for easiest viewing)

Thank you for downloading my game source.
If you've played it, you've probably noticed that it still needs a lot of work

On my list of things to do, I have (in no particular order):
- Finish game server
- Finish Map Editor
- Make levels
- Find a permenant computer server
- Work on scripting engine
- Find game Admins
- Advertise the game

If you think you can help with any of these things, it would be much appreciated if you could contact me via email (jimcamel@jimcamel.8m.com) or ICQ (UIN: 25282667).

Music is Everlong by the Foo Fighters (www.foofighters.com)
Various code snippets used from code on Planet-Source-Code.com

*Notes about Map Format*
The map format I used is one my friend (Roger The Camel) and I devised for a game we were making a while ago. Each 16x16 pixel square part of the map has a 3 digit code. ie.

abc

a: Is the X Co-ordinate area from the backround area to take the graphics from. Ranges from 0-Z. Each increment is equal to 16 pixels
b: Is the Y Co-ordinate area from the backround area to take the graphics from. Ranges from 0-Z. Each increment is equal to 16 pixels
c: Is the special properties of the square of the map. Currently there is only 0, 1 and 2. 0 has no properties, 1 is a non-walkthroughable sector, 2 makes the character sit down.

*Known Bugs*
- The Music for some reason starts quite then turns itself up. I am not sure why this happens.
- The map editor is very buggy (it was only a 30 minute afterthough in this version). Sometimes the tiles placed on the map are not the ones selected in the editor. The non-walkthroughable areas lose their box whenever you try to draw a new one - or whenever you exit the editor (fixed for next release). And there is currently no way to make an area walkthroughable again if you make a mistake, other than editing the map using a text editor (will be fixed in the next release)
- In the menu, Restart and End both exit the program. This will be fixed when I can be bothered.
- The name of the player doesn't always center perfectly under the player.
- The server defaults to induced.iwarp.com. There is no actual server at the moment, so there isn't really much point in trying to connect anyway.

Thats about all I can think of for the moment. If you find another other bugs, please contact me.

Thank you.
JimCamel
jimcamel@jimcamel.8m.com
ICQ: 25282667