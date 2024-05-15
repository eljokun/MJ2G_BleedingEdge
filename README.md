# MJ2G_BleedingEdge
An upgraded version of MathJax-To-Go, STRICTLY FOR WINDOWS, still in development and highly untested, implementing WordHook: An experimental framework that uses the MJ2G code base to create what the Microsoft Word equation editor should have been.
I hope this helps any of you, especially academics.

# Conditions
There are no conditions really, but please consider crediting me in any works this helps you in, if you would like that, or buying me a coffee on ko-fi :) Enjoy!

# WordHook
WordHook uses the windows COM framework to interact with the word document you're interacting with. When a suitable equation is detected within the latex delimiters, $$, it will show a live conversion on a separate widget that is configurable. Either typing \done in your equation or manually pressing the done button then creates the SVG for that equation and converts the text for you, saving a LOT of time, and providing ease of access.
An optional Auto-Show feature, when enabled, makes it so that this widget will stay hidden and show up when it detects that you are typing an equation enclosed within delimiters, and go back after you're done with it.

Below is a demonstration:


https://github.com/eljokun/MJ2G_BleedingEdge/assets/93293178/b402e561-f4fa-4234-8295-bf0a606e8a2c


# New requirements
To use, MJ2G_BleedingEdge requires more dependencies than its MJ2G base, which are:
```
PySide6, re, time, tempfile, pyperclip, os, win32com, pythoncom, win32gui
```

# As is
I am developing this for my own use, as such, i will devote as much time as i want or can to this project, no guarantees are made whatsoever and it is provided as is. While i tried to make it work, it does for some people, but chances are it will not run on non-windows systems.
