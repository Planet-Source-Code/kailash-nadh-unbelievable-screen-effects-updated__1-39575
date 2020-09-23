VERSION 5.00
Begin VB.Form frmSet 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Screen Man Settings"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2700
   Icon            =   "frmset.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   2700
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmb 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Choose an Effect:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1305
   End
End
Attribute VB_Name = "frmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Tag = Me.cmb.ListIndex
frmMain.doeffect
'Set this form's tag to the effect number
End Sub

Private Sub Command3_Click()
about.Show vbModal, Me
End Sub

Private Sub Form_Load()
cmb.AddItem "Melt" '0
cmb.AddItem "Powder Blow" '1
cmb.AddItem "Powder" '2
cmb.AddItem "Evaporate" '3
cmb.AddItem "Water Color" '4
cmb.AddItem "Accumulate" '5
cmb.AddItem "Checks" '6
cmb.AddItem "Extreme Checks" '7
cmb.AddItem "Wind Blow" '8
cmb.AddItem "Pour Down" '9
On Error GoTo er2
cmb.ListIndex = GetSetting("MeltSCR", "Effect", "Effect")
Exit Sub
er2:
cmb.ListIndex = 0
End Sub
