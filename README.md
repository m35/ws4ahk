# ws4ahk

I've been working on "EasyScript" this whole time. Seems like every time I look at the code, there are a dozen things I need to add/change/remove/improve. Even just this morning I made an improvement to the API.

The code has grown into a massive 1500 lines, although it is worth noting that a few hundred of those lines are for comments and error checking.

This script contains a ton of error checking and reporting. While it is really nice for development, not everyone may want to have that extra bulk in their released code. Plus, it seems like most are content with Sean's scripts which have no error checking. So I've designed the EasyScript code so the error checking is easily removed via a simple AHK script (not made yet).

There's still a bunch of things I want to clean up in the script. I would like to make the error handling more consistent between the different parts. Plus there is still more error checking to be added. But the script is fully functional.

So here's the first version: v0.01 beta (21Jul2007) I hope the API doesn't change anymore, but no promises.

I didn't think "EasyScript" was such a good name for it. The most accurate might be "Embedded Microsoft Windows Scripting Host for Autohotkey", but that's just silly. I'm leaning toward "Windows Scripting for AHK", abbreviated as WS4AHK or simply WS. A possible alternative is "COMscript", which emphasis its ability to access COM. Anyone have any thoughts/preferences?

There is also an example script. It's a conversion of ABCYourWay's DHTML edit program to use this new WS4AHK script.

