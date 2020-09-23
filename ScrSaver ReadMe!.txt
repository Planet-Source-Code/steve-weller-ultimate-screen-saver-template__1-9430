Ultimate Screen Saver Template
by Steve Weller

Use this as a template for any screen saver you create.  This came about as an idea because I noticed that the other screen savers on Planet Source Code did not come with a configuration or password dialog box or a preview window.

Specs.:
1.  Tells Windows that a screen saver is active.
2.  Comes with configuration dialog box.
3.  Can be password-enabled (automatic under WinNT).
4.  Allows preview (in which you can put in smaller pictures, etc.).

Notes:
1.  In the Create EXE dialog click Options, the Make tab, and the Application Title box type SCRNSAVE:Screen Saver, replacing Screen Saver with the title of the screen saver.  Also, add the .scr extension so that you will not have to rename it later.
2.  If your Windows directory is not C:\Windows, then open the project file in NotePad, and change the Path32 value to your Windows directory (note that the Package and Deployment Wizard will automatically install to the Windows path with the correct macro).
3.  For the password to take effect, you must click Apply or OK.  This is not a bug, but this is by Windows design.