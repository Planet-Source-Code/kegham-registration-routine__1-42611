VERSION 5.00
Begin VB.Form frmregister 
   BackColor       =   &H00C19222&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registration Routine"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   Icon            =   "FRMREG~1.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Text            =   "12F75E88J"
      Top             =   720
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   4080
      Top             =   1560
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Validate my key"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1080
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2760
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      X1              =   2400
      X2              =   2640
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Your key"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Your name"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmregister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Len(Text6.Text) < 4 Then
    MsgBox "The Name must be more than 4 characters.", vbInformation + vbOKOnly, "shit ! small name"
Me.Timer1.Enabled = False

Else

MousePointer = 11
begintimer

If Text2.Text = Text3.Text Then
Label2.Caption = "Validating code please wait ..."

End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "Vote for me please if you like this techniq", vbApplicationModal, "Thx a lot"
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Text6.Text = "" Then
MsgBox "                                  your name and code please", vbInformation, "                                                                                             "

Text6.TabStop = True
Me.Timer1.Enabled = False
Me.Show
Else
Label2.Caption = "Connecting website please wait ... "
MousePointer = 11
begintimer
End If
End If
End Sub

Private Sub Text6_Change()
On Error Resume Next

Dim Code1 As Single

For I = 1 To Len(Text6.Text) - 1
    Code1 = Format(Asc(Right(Text1.Text, Len(Text6.Text) - I)) * 2 + (31 / I) + (I + 3 / 7), "#.#")
    zip = zip & Code1
Next I
zip = Right(zip, 8)

For I = 1 To Len(zip) - 1
    Code1 = Format(Asc(Right(zip, Len(zip) - I)) * 2 + (1 / I) + (I + 1 / 7), "#00")
    final = final & Code1
Next I
final = Right(final, Len(final) - 4)
final = final & Asc(Text1)
Text3 = final


End Sub

Private Sub Timer1_Timer()
If Text2.Text <> Text3.Text Then
MsgBox "Code entered was not correct", vbInformation, "Registration terminated"
End
Else

If Text2.Text = Text3.Text Then
MsgBox "Registration code was successfull" & vbCrLf & "Restart the program to take effect", vbInformation, "Welcome " & Text6.Text

Timer1.Enabled = False
MousePointer = 0
End If

End If
End Sub

Sub begintimer()
Timer1.Enabled = True
End Sub
