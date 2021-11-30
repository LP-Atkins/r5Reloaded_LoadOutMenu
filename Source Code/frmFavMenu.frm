VERSION 5.00
Begin VB.Form frmFavMenu 
   Caption         =   "Reloaded Apex - Favourite Menu (by LongParsnip)"
   ClientHeight    =   885
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   431.493
   ScaleMode       =   0  'User
   ScaleWidth      =   11129.65
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnFav 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Empty"
      Height          =   375
      Index           =   5
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton btnFav 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Empty"
      Height          =   375
      Index           =   4
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   1935
   End
   Begin VB.CommandButton btnFav 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Empty"
      Height          =   375
      Index           =   3
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton btnFav 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Empty"
      Height          =   375
      Index           =   2
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1935
   End
   Begin VB.CommandButton btnFav 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Empty"
      Height          =   375
      Index           =   1
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton btnFav 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Empty"
      Height          =   375
      Index           =   0
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
   Begin VB.CommandButton btnBack 
      BackColor       =   &H00FFFFFF&
      Caption         =   "<- Back"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmFavMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Forms")

Private Sub Form_Load()
    Call FormOnTop(Me)
    
    'Set button text.
    If FavDesc0 <> vbNullString Then btnFav(0).Caption = FavDesc0
    If FavDesc1 <> vbNullString Then btnFav(1).Caption = FavDesc1
    If FavDesc2 <> vbNullString Then btnFav(2).Caption = FavDesc2
    If FavDesc3 <> vbNullString Then btnFav(3).Caption = FavDesc3
    If FavDesc4 <> vbNullString Then btnFav(4).Caption = FavDesc4
    If FavDesc5 <> vbNullString Then btnFav(5).Caption = FavDesc5
    
End Sub

Private Sub btnFav_Click(Index As Integer)

    Select Case Index
        Case 0: SendCommand (FavCmd0)
        Case 1: SendCommand (FavCmd1)
        Case 2: SendCommand (FavCmd2)
        Case 3: SendCommand (FavCmd3)
        Case 4: SendCommand (FavCmd4)
        Case 5: SendCommand (FavCmd5)
    End Select

End Sub

Private Sub btnBack_Click()
    frmSelect.Show
    frmSelect.Left = Me.Left
    frmSelect.Top = Me.Top
    Unload Me
End Sub
