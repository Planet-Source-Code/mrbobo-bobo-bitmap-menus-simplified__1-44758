VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bobo Bitmap Menus"
   ClientHeight    =   3705
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   7995
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   300
      Width           =   7575
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5580
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0402
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0F36
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":14D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1A6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2B48
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":30E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":367C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3C16
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3D70
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":450E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4AA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5042
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":55DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5B76
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5CD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":626A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFilebase 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&New"
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Open"
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save"
         Index           =   2
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Save &As"
         Index           =   3
      End
      Begin VB.Menu mnuMRUSP1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMRU 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMRU 
         Caption         =   ""
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Setup..."
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileSP2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Properties..."
      End
      Begin VB.Menu mnuFileWordCount 
         Caption         =   "&Word Count..."
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuFileAssoc 
         Caption         =   "Program Settings"
         Begin VB.Menu mnuFileAssociate 
            Caption         =   "Associate with Plain text files"
            Index           =   0
         End
         Begin VB.Menu mnuFileAssociate 
            Caption         =   "Show Richtext file code"
            Index           =   1
         End
      End
      Begin VB.Menu mnuFileSP3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEditBase 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEdit 
         Caption         =   "Undo"
         Enabled         =   0   'False
         Index           =   0
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Cut"
         Enabled         =   0   'False
         Index           =   2
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Copy"
         Enabled         =   0   'False
         Index           =   3
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Paste"
         Enabled         =   0   'False
         Index           =   4
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Index           =   5
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Find..."
         Index           =   7
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Find Next"
         Index           =   8
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Replace..."
         Index           =   9
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Go To..."
         Index           =   10
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Highlight..."
         Index           =   12
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "UnHighlight"
         Enabled         =   0   'False
         Index           =   13
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Select All"
         Index           =   15
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Select Above"
         Index           =   16
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Select Below"
         Index           =   17
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   18
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Time/Date"
         Index           =   19
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "F&ormat"
      Begin VB.Menu mnuFormatWordWrap 
         Caption         =   "WordWrap"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFormatFont 
         Caption         =   "Font..."
      End
      Begin VB.Menu mnuFormatBackColor 
         Caption         =   "BackColor..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpTopics 
         Caption         =   "Help Topics"
      End
      Begin VB.Menu mnuHelpSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuRTFPopup 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuRTF 
         Caption         =   "Undo"
         Index           =   0
      End
      Begin VB.Menu mnuRTF 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuRTF 
         Caption         =   "Cut"
         Index           =   2
      End
      Begin VB.Menu mnuRTF 
         Caption         =   "Copy"
         Index           =   3
      End
      Begin VB.Menu mnuRTF 
         Caption         =   "Paste"
         Index           =   4
      End
      Begin VB.Menu mnuRTF 
         Caption         =   "Delete"
         Index           =   5
      End
      Begin VB.Menu mnuRTF 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuRTF 
         Caption         =   "Select All"
         Index           =   7
      End
      Begin VB.Menu mnuRTF 
         Caption         =   "Select Above"
         Index           =   8
      End
      Begin VB.Menu mnuRTF 
         Caption         =   "Select Below"
         Index           =   9
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********Copyright PSST Software 2003**********************
'Submitted to Planet Source Code - April 2003
'If you got it elsewhere - they stole it from PSC.

'Written by MrBobo - enjoy
'Please visit our website - www.psst.com.au


Option Explicit


Private Sub Form_Load()
    Set IL = ImageList1
    'Assign images to menu tags
    mnuFile(0).Tag = 1
    mnuFile(1).Tag = 2
    mnuFile(2).Tag = 3
    mnuFile(3).Tag = 4
    mnuFilePageSetup.Tag = 5
    mnuFilePrint.Tag = 6
    mnuFileProperties.Tag = 7
    mnuFileAssociate(0).Tag = 8
    mnuFileAssociate(1).Tag = 9
    mnuEdit(0).Tag = 10
    mnuEdit(2).Tag = 11
    mnuEdit(3).Tag = 12
    mnuEdit(4).Tag = 13
    mnuEdit(5).Tag = 14
    mnuEdit(7).Tag = 15
    mnuEdit(19).Tag = 16
    mnuHelpTopics.Tag = 17
    'convert VB menus to ownerdrawn menus
    ConvertOD Me
End Sub

Private Sub mnuFile_Click(Index As Integer)
    'Not really appropriate for these menus but
    'demonstrates the checked state of an icon
    mnuFile(Index).Checked = Not mnuFile(Index).Checked
End Sub

Private Sub mnuFileExit_Click()
    'just to prove the VB menus still work !!
    Unload Me
End Sub
Private Sub mnuFormatWordWrap_Click()
    ' checkmark only
    mnuFormatWordWrap.Checked = Not mnuFormatWordWrap.Checked

End Sub
'If you choose to add a menu dynamically at run time use the sub 'AddODMenu'
