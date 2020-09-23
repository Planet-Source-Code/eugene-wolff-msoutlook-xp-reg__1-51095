VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Override level1 security on the following extensions"
   ClientHeight    =   1275
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRemoveAccess 
      Caption         =   "Remove Access"
      Height          =   495
      Left            =   3450
      TabIndex        =   3
      Top             =   615
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cboFileExtensions 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   4785
      TabIndex        =   1
      Top             =   615
      Width           =   1215
   End
   Begin VB.TextBox txtReg 
      Height          =   300
      Left            =   105
      TabIndex        =   0
      Top             =   225
      Width           =   5910
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FT As Boolean   ' Check to see if this is the first entry

Private Sub cboFileExtensions_Click()
    If FT = True Then
        txtReg.Text = cboFileExtensions.Text
        FT = False
    Else
        txtReg.Text = txtReg.Text & ";" & cboFileExtensions.Text
    End If
    
End Sub

Private Sub cmdRemoveAccess_Click()
    ' Delete the list of extensions
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Office\10.0\Outlook\Security", "Level1Remove"
End Sub

Private Sub cmdUpdate_Click()
    ' Updates the list of extensions
    savestring HKEY_CURRENT_USER, "Software\Microsoft\Office\10.0\Outlook\Security", "Level1Remove", txtReg
End Sub

Private Sub Form_Load()
    ' Below is the list of extensions blocked by Outlook 2002
    With cboFileExtensions
        .AddItem ".ade"
        .AddItem ".adp"
        .AddItem ".asx"
        .AddItem ".bas"
        .AddItem ".bat"
        .AddItem ".chm"
        .AddItem ".cmd"
        .AddItem ".com"
        .AddItem ".cpl"
        .AddItem ".crt"
        .AddItem ".exe"
        .AddItem ".hlp"
        .AddItem ".hta"
        .AddItem ".inf"
        .AddItem ".ins"
        .AddItem ".isp"
        .AddItem ".js"
        .AddItem ".jse"
        .AddItem ".lnk"
        .AddItem ".mdb"
        .AddItem ".mde"
        .AddItem ".msc"
        .AddItem ".msi"
        .AddItem ".msp"
        .AddItem ".mst"
        .AddItem ".pcd"
        .AddItem ".pif"
        .AddItem ".prf"
        .AddItem ".reg"
        .AddItem ".scf"
        .AddItem ".scr"
        .AddItem ".sct"
        .AddItem ".shb"
        .AddItem ".shs"
        .AddItem ".url"
        .AddItem ".vb"
        .AddItem ".vbe"
        .AddItem ".vbs"
        .AddItem ".wsc"
        .AddItem ".wsf"
        .AddItem ".wsh"
    End With
    txtReg = getstring(HKEY_CURRENT_USER, "Software\Microsoft\Office\10.0\Outlook\Security", "level1remove")
    If txtReg.Text = "" Then
        FT = True
    Else
        FT = False
        cmdRemoveAccess.Visible = True
    End If
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub
