# ws4ahk

ws4ahk v0.20

- `WS_Exec()`/`WS_Eval()`: Better handling of errors.
Removed printf() style functionality, moved to codef() function.
Removed leftover Clipboard debug.
- `codef()`: New function to handle `printf()` style formatting of code.
Also fixes the bug if in hex mode.

Note that if you are using the printf style in `WS_Exec` or `WS_Eval` (I doubt anyone is), you will need to slightly modify your code to use the new `codef()` function. 