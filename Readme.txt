Yippee! Messenger
=================

This is a server/client IM messenger that currently:

1) allows users to chat to one another
2) uplaod an avatar picture
3) send files
4) allows he server to see what a user is up to on their workstation

It is intended for LAN networks, specifically in schools, where the server would be
running on the instructors PC.  The teacher could then see whatever the students are doing.
while allow the students to chat among themselves, and send files.

The server can also distribute a file to all the students but only one at a time.

Uses Winsock UDP and TCP.

See documentation (included) for message protocols


NOTES:
===================
Requires CRC.dll and zlib32.dll, just place in the same folder for client and server

note: 
===================
Rename CRC.dl_ and zlib32.dl_ to .DLL

To the best of my knowledge these files are not trojans or viruses, nor do they contain viruses.
However, by downloading this file you agree to the terms that I am not to be held responsible for any damage
that may occur to your computer because of this program.

These files are needed for viewing the workstation and uploading/downlaoding avatar pictures


Setting up:

Server
==================
- Run the server, click Options>Manage Users
- Under Add/Remove Users, enter a Student number (or ID number) that the client user will
  use to register onto the system and click Add.  The current default password is 12345.

Client
==================
- Run the client on a networked PC.
- Click File>Settings
- Enter the computer name or IP address of the machine running the server in "Server Name/IP Address"
- Click OK
- Close the client program and run it again.

This system uses port 9004 to communicate to you may need to open port 9004 on your firewall


To register a first time user, (whose student ID has already been entered into the system using the server see above):

Click File>Register
Choose "I do not have a login name yet" and click Next
Enter a new Login name and click Next.
Enter the student number you added at the server and the default password (12345) and click next.
If all went well, you can now edit your personal info, change passwords, and upload a picture.

Once registered, go back to the main window and click File>Sign in
Enter your login name that you used to register, and your password and click Log in.

Double click name to open a chat window to that user.

You can add friends to your friends list. Right click main window and click "Add friend by ID"
Enter the student ID of your friend, and the user will apear in your list. 

Icon will be grayed out if the user is offline.

You can also send offline messages.

Two accounts are already in database:

Login: User
Pass:  12345

Login: Dave
Pass:  12345

Features
================
- Filesharing
- Offline messages
- Workstation Monitoring

Todo:
=====
- Fix memory leaks especially in Monitoring
- Cancellation notification of filesharing
- database password protection
- failing gracefully
- move picture transfer to TCP method?

WIP Updates
===========
9/20/04
=======
- Fixed focus-stealing bug.  Kept calling frmMonitor.Show during MonitorUser
  Moved it out.

9/19/04
=======
- Workstation monitoring on the server crashes randomly. Could be a GPF?
- frmMonitor keeps stealing focus.  Server user can't type anything because of this.
 
 
9/18/04
=======
- I keep getting a "Run-time error 126" everytime the client tries to connect,
  but if I continue, the connection establishes normally.
  Oh well, time to brine out the old On Error Resume Next kludge.
- When filesharing large files, user is locked out while downloading.
  Can't use DoEvents because it will cause another DataArrival event while
  an ongoing DataArrival is running.
- Recompiled zlib.dll to work properly with the /D "ZLIB_DLL" option. RTFM!
- Got compression to work as a result. 2MB screenies are now compressed down to 67kB!
  While there is very little improvement in update speed, bandwidth-wise, it should
  be more efficient.

9/17/04
=======
- After nonstop work on Monitoring, finally got the memory transfer to run properly.
  No more pauses when updating.
- Changed FileSharing completely, now contained in one form, instead of a messy structure.
  As a bonus, I can send multiple files at the same time too...
- First XP compatibility check in a while... slight problem.  The titlebars' size is different
  from 98. Since controls' sizes and positions depend on the overall size of the form, 
  menubar and titlebar included, most stuff goes past the bottom.

  Need to find a way to get the correct size of title bars and modify all
  Form_Resize events that change the height and top of affected controls, most notably Chat.

9/16/04
=======
- Implemented Workstation Monitoring... finally! Slow though, as bytes are still
  copied piecemeal from the DIB array into the data array.
- Added Minimize to System Tray Icon to client.

9/15/04
=======
- Implemented FileSharing, but can't seem to send more than one file at the same time
- Added user chat Options.

9/12/04
=======
- Added CRC32 checking of user pics.  When pictures do not match, 
  server sends updated picture.  Uses CRC.DLL

Random update
=============
- Removed XP styles which was causing flickering with frames.
- Added "Add a friend" using name/student number

9/5/04
======
- Fixed picture transfer bug.  Data was being incorrectly parsed at the client.

What we need to do is establish a ping-pong method to retrieve data such as:

PIC:REQ,UIND,UNAME
PIC:0,5,NLEN,UNAME,DATA
PIC:GOT,0,UIND,UNAME
PIC:1,5,NLEN,UNAME,DATA
PIC:GOT,1,UIND,UNAME
PIC:2,5,NLEN,UNAME,DATA
PIC:GOT,2,UIND,UNAME
PIC:3,5,NLEN,UNAME,DATA
PIC:GOT,3,UIND,UNAME
PIC:4,5,NLEN,UNAME,DATA
PIC:GOT,4,UIND,UNAME
PIC:5,5,NLEN,UNAME,DATA
PIC:GOT,5,UIND,UNAME

Didn't work quite properly when compiled. Put a Sleep 10 before sending a GOT, and works fine...

- Fixed again!

9/4/04
======
- implemented add/remove friends via right-click
- implemented registering
- added offline messages
- added user timeout of 60 seconds, with ACK request on the 30th second
- fixed a logon bug
- added XP visual styles via a manifest using a resource (RES) file

9/3/04
======
- Implemented picture downloading from server to client
  buggy, works best with small files
- Implemented database for user logon
- Implemented username and password checking during logon
- Fixed bug where winsock doesn't listen for a connection
  Forgot to Bind local port before anything 
- Implemented Friends downloading & listing
- Implemented INI saving of client settings
