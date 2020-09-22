VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Begin VB.Form Form1 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Example by DarkKnight"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   3360
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2640
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Exit"
      Height          =   195
      Left            =   2280
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&LogOut"
      Height          =   195
      Left            =   1200
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Login"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000009&
      Caption         =   "Incoming"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3135
      Begin VB.TextBox Text4 
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Password"
      Height          =   675
      Left            =   1680
      TabIndex        =   0
      Top             =   0
      Width           =   1575
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   8
         Text            =   "Password"
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000009&
      Caption         =   "User Name"
      Height          =   675
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   1575
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Text            =   "DarkKnight"
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is pretty basic inet stuff.
'Inet1 opens a url in text4, then searches
'the source code of the webpage for certain
'strings that appear in the source code if you have
'a frozen account, bad password, or are logged in.
'Does that make since to you? :X
'DarkKnight is sexy
'http://www.fdo-files.com
'http://www.dark.has.it
'http://www.aimfilez.com
'Shouts to Midiman, DarkKnight, Xeon, Peter, Nemisis,
'And nobody else cause your all gay.

Private Sub Command1_Click()
On Error Resume Next
Text4 = Inet1.OpenURL("http://neopets.com/login.phtml?destination=/petcentral.phtml&username=" & Text1 & "&password=" & Text2)
If InStr(Text4, "This account has been <b>FROZEN</b>") Then 'Searches for the string "This account has been <b>FROZEN</b>" in text4 (on the site).
Form1.Caption = "This account is FROZEN" 'If the string is in text4 (on the site), then it means your account is frozen.
End If
If InStr(Text4, "Bad Password") Then 'Same as above but searches for "Bad Password" as string
Form1.Caption = "Bad Password" 'If the string is found then the password given is wrong
End If
If InStr(Text4, "user :") Then 'Searches for the "user :" string on the site
Form1.Caption = "Logged In!" 'If its there, then that means the password was correct and you are now logged in.
End If
If InStr(Text4, "Rated") Then
Form1.Caption = "Too many login attempts..."
MsgBox "You have attempted to login to this neopets account too many time. Please wait 24 hours before trying to login again."
'The above means that your IP can not try logging into that account for 24 hours. If your trying to make a password cracker
'for neopets (which I dont recommend) then you will need to use a proxy or a different method of getting the users password.
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
Inet1.OpenURL "http://neopets.com/logout.phtml"
Form1.Caption = "Logged Out!"
'This will log you out of the site by opening the logout page.
End Sub

Private Sub Command3_Click()
MsgBox "DarkKnight is Hot :D"
End 'exits everything!
End Sub

Private Sub text4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "This program was made by DarkKnight as a VB example. I'd like to give shout outs to-" & vbCrLf & "DarkKnight, Xeon, Midiman, Robbie, Nemisis, Peter, Fdo-Files, Anti-Aol, and Myself." & vbCrLf & vbCrLf & "DarkKnight is sexy"
'Ha...did I put my name on this a little too much? :D
End Sub

Private Sub Form_Load()
Form1.Show
Text4 = Inet1.OpenURL("http://neopets.com/loginpage.phtml") 'makes text4 get the data of the loginpage
Form1.Caption = "Checking Logon Status"
Call Pause(3)
If InStr(Text4, "You are already logged in!") Then 'checks to see if the string "You are already logged in!" (which would be on the neopets page)is in text4
Form1.Caption = "You Are Currenty Logged In" 'If the string is in text4 (on the neopets page), you are logged in.
Else
Form1.Caption = "Status: Not Logged In!" 'If the string isnt in text4 (on the site) then you are not logged in.
End If
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "This program was made by DarkKnight as a VB example. I'd like to give shout outs to-" & vbCrLf & "DarkKnight, Xeon, Midiman, Robbie, Nemisis, Peter, Fdo-Files, Anti-Aol, and Myself." & vbCrLf & vbCrLf & "DarkKnight is sexy"
'Ha...did I put my name on this a little too much? :D
End Sub
