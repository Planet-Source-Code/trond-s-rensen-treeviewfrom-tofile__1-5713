VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "TreeView Example"
   ClientHeight    =   7905
   ClientLeft      =   2460
   ClientTop       =   2340
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   8385
   Begin VB.Frame Frame1 
      Caption         =   "Selected Node Info"
      Height          =   3015
      Left            =   6240
      TabIndex        =   11
      Top             =   1320
      Width           =   2055
      Begin VB.Label lblInfo 
         Caption         =   "Click node for info.."
         Height          =   2415
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdAddGroup 
      Caption         =   "Add Group"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdAddNode 
      Caption         =   "Add Node"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdLBackup 
      Caption         =   "Load TreeView"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton cmdSaveText 
      Caption         =   "Save Text"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save TreeView"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6360
      Width           =   1335
   End
   Begin MSComctlLib.TreeView TVProjects 
      Height          =   4755
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   8387
      _Version        =   393217
      Indentation     =   353
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imlDriveFileList2"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList imlDriveFileList2 
      Left            =   6600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0354
            Key             =   "TextFile"
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtfEdit 
      Height          =   4785
      Left            =   3120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1440
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   8440
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      HideSelection   =   0   'False
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmMain.frx":0468
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "trond.sorensen@bi.no"
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   7560
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Load/save TreeView from/to tab indented texfile example by Trond Sørensen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      TabIndex        =   8
      Top             =   6840
      Width           =   7785
   End
   Begin VB.Label Label2 
      Caption         =   "What the textfile looks like. You can edit it here and save it. The TreeView will update itself from the new file."
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   720
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "What theTreeView looks like. You can add nodes and save it. The textfile will update itself from the new file."
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   2895
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddGroup_Click()
    Call AddNewGroup(TVProjects)
End Sub

Private Sub cmdAddNode_Click()
    Call AddNewNode(TVProjects)
End Sub

Private Sub cmdLBackup_Click()
    rtfEdit.LoadFile (App.Path) & "\" & "projects.txt"
    Call LoadTreeViewFromFile((App.Path) & "\" & "projects.txt", TVProjects)

End Sub

Private Sub cmdSave_Click()
    Call SaveTreeViewToFile((App.Path) & "\" & "projects.txt", TVProjects)
    rtfEdit.LoadFile (App.Path) & "\" & "projects.txt"
End Sub

Private Sub cmdSaveText_Click()
    rtfEdit.SaveFile (App.Path) & "\" & "projects.txt", rtfText
    Call LoadTreeViewFromFile((App.Path) & "\" & "projects.txt", TVProjects)
End Sub

Private Sub Form_Load()
    Call LoadTreeViewFromFile((App.Path) & "\" & "projects.txt", TVProjects)
    rtfEdit.LoadFile (App.Path) & "\" & "projects.txt"
    
End Sub

Private Sub TVProjects_NodeClick(ByVal Node As MSComctlLib.Node)

   Dim nodX As Node
   ' Set the variable to the SelectedItem.
   Set nodX = TVProjects.SelectedItem
   Dim strProps As String
   ' Retrieve properties of the node.
   strProps = "Text: " & nodX.Text & vbLf
   strProps = strProps & "Key: " & nodX.Key & vbLf
   strProps = strProps & "index: " & nodX.Index & vbLf
   strProps = strProps & "image: " & nodX.Image & vbLf
   On Error Resume Next ' Root node doesn't have a parent.
   strProps = strProps & "Parent: " & nodX.Parent.Text & vbLf
   strProps = strProps & "FirstSibling: " & nodX.FirstSibling.Text & vbLf
   strProps = strProps & "LastSibling: " & nodX.LastSibling.Text & vbLf
   strProps = strProps & "Next: " & nodX.Next.Text & vbLf
   strProps = strProps & "Root: " & nodX.Root & vbLf
   
   lblInfo.Caption = strProps
End Sub
