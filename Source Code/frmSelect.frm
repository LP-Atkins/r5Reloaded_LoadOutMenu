VERSION 5.00
Begin VB.Form frmSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reloaded Apex - Loadout Menu (by LongParsnip)"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7155
   Icon            =   "frmSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrMove 
      Left            =   6720
      Top             =   480
   End
   Begin VB.CommandButton btnFavMenu 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Open Fav Menu"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton btnSave 
      BackColor       =   &H0080FFFF&
      Caption         =   "Save Fav"
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   600
      Width           =   975
   End
   Begin VB.ComboBox cbFav 
      Height          =   315
      ItemData        =   "frmSelect.frx":0442
      Left            =   4920
      List            =   "frmSelect.frx":0458
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   600
      Width           =   615
   End
   Begin VB.ComboBox cbAbility 
      Height          =   315
      Index           =   1
      ItemData        =   "frmSelect.frx":0474
      Left            =   4680
      List            =   "frmSelect.frx":04B4
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   240
      Width           =   1575
   End
   Begin VB.ComboBox cbWeapon 
      Height          =   315
      Index           =   1
      ItemData        =   "frmSelect.frx":05C5
      Left            =   1440
      List            =   "frmSelect.frx":060B
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton btnDeploy 
      BackColor       =   &H00FFFF00&
      Caption         =   "Deploy"
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
   Begin VB.ComboBox cbAbility 
      Height          =   315
      Index           =   0
      ItemData        =   "frmSelect.frx":06D8
      Left            =   3000
      List            =   "frmSelect.frx":0718
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
   Begin VB.ComboBox cbWeapon 
      Height          =   315
      Index           =   0
      ItemData        =   "frmSelect.frx":0829
      Left            =   0
      List            =   "frmSelect.frx":086F
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Ultimate"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Tactical"
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Secondary"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Primary"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Forms")
Option Explicit


Private Sub Form_Load()

    Call FormOnTop(Me)      'Makes the application render always on top.
    cbWeapon(0).Text = DEF_PRIMARY
    cbWeapon(1).Text = DEF_SECONDARY
    cbAbility(0).Text = DEF_TACTICAL
    cbAbility(1).Text = DEF_ULTIMATE
    cbFav.Text = " 1"       'Needs to have a leadicng space because VB combobox goes all weird if you don't have it (will be trimmed later).
    
    'Add custom weapons and abilities.
    Call addCustoms
    
    'Remember the settings when switching between forms.
    If PrimSel <> vbNullString Then cbWeapon(0).Text = PrimSel
    If SecSel <> vbNullString Then cbWeapon(1).Text = SecSel
    If TacSel <> vbNullString Then cbAbility(0).Text = TacSel
    If UltSel <> vbNullString Then cbAbility(1).Text = UltSel
    If FavSel <> vbNullString Then cbFav.Text = FavSel
    
    'Set window last position and start save timer.
    If xyPos Then
        Me.Left = xPos
        Me.Top = yPos
    End If
    tmrMove.Interval = 1000
    
End Sub

Private Sub btnDeploy_Click()
    Call SendCommand(AbilityCommand(cbAbility(1).Text, 1) & "; " & AbilityCommand(cbAbility(0).Text, 0) & "; " & WeaponCommand(cbWeapon(1).Text, 1) & "; " & WeaponCommand(cbWeapon(0).Text, 0))
End Sub

Private Sub btnFavMenu_Click()

    'Save values of comboboxes.
    PrimSel = cbWeapon(0).Text
    SecSel = cbWeapon(1).Text
    TacSel = cbAbility(0).Text
    UltSel = cbAbility(1).Text

    'Position fav form in the same place.
    frmFavMenu.Show
    frmFavMenu.Left = Me.Left
    frmFavMenu.Top = Me.Top
    Unload Me
    
End Sub

Private Sub btnSave_Click()
Dim Index As Long

    Index = CLng(Trim(cbFav.Text)) - 1
    
    'Write to the appropriate program variable.
    Select Case Index
        
        Case 0: FavDesc0 = Left$(cbWeapon(0).Text, 4) & " + " & Left$(cbWeapon(1).Text, 4)
                FavCmd0 = AbilityCommand(cbAbility(1).Text, 1) & "; " & AbilityCommand(cbAbility(0).Text, 0) & "; " & WeaponCommand(cbWeapon(1).Text, 1) & "; " & WeaponCommand(cbWeapon(0).Text, 0)
        
        Case 1: FavDesc1 = Left$(cbWeapon(0).Text, 4) & " + " & Left$(cbWeapon(1).Text, 4)
                FavCmd1 = AbilityCommand(cbAbility(1).Text, 1) & "; " & AbilityCommand(cbAbility(0).Text, 0) & "; " & WeaponCommand(cbWeapon(1).Text, 1) & "; " & WeaponCommand(cbWeapon(0).Text, 0)
                
        Case 2: FavDesc2 = Left$(cbWeapon(0).Text, 4) & " + " & Left$(cbWeapon(1).Text, 4)
                FavCmd2 = AbilityCommand(cbAbility(1).Text, 1) & "; " & AbilityCommand(cbAbility(0).Text, 0) & "; " & WeaponCommand(cbWeapon(1).Text, 1) & "; " & WeaponCommand(cbWeapon(0).Text, 0)
        
        Case 3: FavDesc3 = Left$(cbWeapon(0).Text, 4) & " + " & Left$(cbWeapon(1).Text, 4)
                FavCmd3 = AbilityCommand(cbAbility(1).Text, 1) & "; " & AbilityCommand(cbAbility(0).Text, 0) & "; " & WeaponCommand(cbWeapon(1).Text, 1) & "; " & WeaponCommand(cbWeapon(0).Text, 0)
        
        Case 4: FavDesc4 = Left$(cbWeapon(0).Text, 4) & " + " & Left$(cbWeapon(1).Text, 4)
                FavCmd4 = AbilityCommand(cbAbility(1).Text, 1) & "; " & AbilityCommand(cbAbility(0).Text, 0) & "; " & WeaponCommand(cbWeapon(1).Text, 1) & "; " & WeaponCommand(cbWeapon(0).Text, 0)
        
        Case 5: FavDesc5 = Left$(cbWeapon(0).Text, 4) & " + " & Left$(cbWeapon(1).Text, 4)
                FavCmd5 = AbilityCommand(cbAbility(1).Text, 1) & "; " & AbilityCommand(cbAbility(0).Text, 0) & "; " & WeaponCommand(cbWeapon(1).Text, 1) & "; " & WeaponCommand(cbWeapon(0).Text, 0)
    
    End Select
    
    
    'Write INI File.
    Call WriteINI

End Sub


'Adds the user custom weapons & abilities to the comboboxes.
Private Sub addCustoms()
Dim Key As Variant

    'Weapons
    For Each Key In colCustomWeapons
        cbWeapon(0).AddItem Key.Name
        cbWeapon(1).AddItem Key.Name
    Next Key
    
    'Abilities
    For Each Key In colCustomAbilities
        cbAbility(0).AddItem Key.Name
        cbAbility(1).AddItem Key.Name
    Next Key
    
End Sub


'Checks and saves the window position to the INI file.
Private Sub tmrMove_Timer()

    If Me.Left <> xPos Then
        Call PutINISetting("Settings", "xPos", Me.Left, INI_PATH)
        xPos = Me.Left
    End If
    
    If Me.Top <> yPos Then
        Call PutINISetting("Settings", "yPos", Me.Top, INI_PATH)
        yPos = Me.Top
    End If
        
End Sub
