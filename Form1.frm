VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Megacomputing - Waren Sortiment"
   ClientHeight    =   7425
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9465
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   10800
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0CCA
            Key             =   "DISKS04"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1266
            Key             =   "TRAFFIC"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1F42
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2C1E
            Key             =   "TRASH02A"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":31BA
            Key             =   "TRASH02B"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3756
            Key             =   "MISC33"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3AF6
            Key             =   "MISC29"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3E96
            Key             =   "ARW05DN"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":42EA
            Key             =   "ARW05UP"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":473E
            Key             =   "FILES03B"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4AE2
            Key             =   "FILES04"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer3 
      Interval        =   400
      Left            =   10320
      Top             =   4440
   End
   Begin VB.Frame Frame2 
      Caption         =   "Info ... "
      Height          =   2655
      Left            =   4680
      TabIndex        =   14
      Top             =   360
      Width           =   4455
      Begin VB.Label lblcopy 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(c) 2003 by Robert Niedziela, Alle Rechte vorbehalten!"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   735
         Left            =   240
         TabIndex        =   16
         Top             =   1800
         Width           =   4215
      End
      Begin VB.Label lblHelp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hallo!"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   240
         TabIndex        =   15
         Tag             =   "Help"
         ToolTipText     =   "Hier erhalten Sie kurze Hilfe!"
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   10200
      Top             =   4920
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10920
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":507E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5D5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6A36
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7712
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":83EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "System "
      Height          =   4335
      Left            =   4680
      TabIndex        =   7
      Top             =   3120
      Width           =   4455
      Begin VB.TextBox txtMultiplikator 
         Height          =   285
         Left            =   1560
         TabIndex        =   21
         Text            =   "1,25"
         Top             =   2760
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox txtDiff 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1560
         TabIndex        =   19
         Text            =   "null"
         Top             =   1800
         Width           =   2775
      End
      Begin VB.Timer Timer2 
         Interval        =   400
         Left            =   240
         Top             =   3720
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Neue Artikel"
         Height          =   255
         Left            =   3000
         TabIndex        =   0
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox txtCategory 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Text            =   "null"
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtRecipe 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Text            =   "null"
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtAuthor 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Text            =   "null"
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox txtInstructions 
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Text            =   "null"
         Top             =   2160
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Speichern"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   3960
         Width           =   1215
      End
      Begin VB.TextBox txtEK 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Text            =   "null"
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Differenz: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   555
         TabIndex        =   20
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Gruppe:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Warenname: "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Serial: "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "VK: "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label label5 
         BackStyle       =   0  'Transparent
         Caption         =   "  EK:  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   1440
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   7050
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "17:33"
            Object.ToolTipText     =   "Uhrzeit"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   1
            TextSave        =   "08.01.2003"
            Object.ToolTipText     =   "Heutige Datum"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   5292
            MinWidth        =   5292
            Object.ToolTipText     =   "Total Einträge in dieser Kategorie"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   4145
            MinWidth        =   4145
            Object.ToolTipText     =   "Total Einträge Gesamt im Datenbank"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   4145
            MinWidth        =   4145
            Object.ToolTipText     =   "Total Kategorien"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trajan"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   7095
      Left            =   0
      TabIndex        =   17
      Top             =   360
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   12515
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      HotTracking     =   -1  'True
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   26
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "save"
            Description     =   "Save Recipe"
            Object.ToolTipText     =   "Artikel Speichern..."
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "exit"
            Object.ToolTipText     =   "Programm Beenden"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            Object.ToolTipText     =   "Drucken"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deleterecipe"
            Description     =   "Delete Recipe"
            Object.ToolTipText     =   "Löscht ausgewählte Artikel"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Delete Category"
            Object.ToolTipText     =   "Kategorie Löschen..."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "insertrecipe"
            Description     =   "Insert Recipe"
            Object.ToolTipText     =   "Neue Artikel Hinzufügen..."
            ImageIndex      =   6
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "insertcategory"
            Description     =   "Insert Category"
            Object.ToolTipText     =   "Neue Kategorie Hinzufügen.."
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "importrecipe"
            Object.ToolTipText     =   "Import Artikel"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "exportrecipe"
            Object.ToolTipText     =   "Export Artikel"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "backup"
            Description     =   "Backup Database"
            Object.ToolTipText     =   "Backup Database"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "restore"
            Description     =   "Restore Database"
            Object.ToolTipText     =   "Restore Database"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "searchepecurious"
            Description     =   "Search Epecurious"
            Object.ToolTipText     =   "Search Epecurious Recipes"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "searchfoodtv"
            Description     =   "Search FoodTv"
            Object.ToolTipText     =   "Search FoodTv Recipes"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "searchqvcrecipes"
            Description     =   "Search QVC Recipes"
            Object.ToolTipText     =   "Search QVC Recipes"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "searchhungrymonster"
            Object.ToolTipText     =   "Search Hungry Monster Recipes"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "searchcdkitchen"
            Object.ToolTipText     =   "Search CDKitchen Recipes"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "searchcopykat"
            Object.ToolTipText     =   "Search CopyKat Recipes"
            ImageIndex      =   21
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Datenbank"
      Begin VB.Menu mnuAddCat 
         Caption         =   "&Neue Warengruppe Hinzufügen"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "&Artikel Hinzufügen"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Artikel &Speichern"
      End
      Begin VB.Menu mnuDelCat 
         Caption         =   "Kategorie &Löschen"
      End
      Begin VB.Menu n1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Beenden"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************
'*** MC Store - The Store database     ***
'*** Coded by Robert Niedziela (c)2003 ***
'*** some code comes from PSC Thanks!  ***
'*** http://www.personal-webserver.de  ***
'*****************************************
'*****************************************





Private Sub Command1_Click()
mnuSave_Click
CloseRs
End Sub

Private Sub Command2_Click()
mnuAdd_Click
'Command2.Enabled = False
'txtMultiplikator.Visible = True

End Sub

Private Sub Form_Load()
'Compact Database
    Dim strSource As String
    Dim strTarget As String
    
    strSource = App.Path & "\dbstore.mdb"
    strTarget = App.Path & "\Compact.mdb"
    DBEngine.CompactDatabase strSource, strTarget

lblHelp.Caption = "Waren Sortiment a special version for the PSC Users, Visit my Homepage @ www.personal-webserver.de "


 'Delete Old Database
    Kill (strSource)
    frmMain.Caption = "Megacomputing - Waren Sortiment " & "Version " & App.Major & "." & App.Minor & "." & App.Revision
    'Copy the Compact.mdb DataBase back to Recipes.mdb
    FileCopy strTarget, strSource
    
    'Kill Old Compact
    Kill (strTarget)
    StartIt
    'Call function to load the DataBase catagories and Records
    'into the Treeview
    'ListColumns
    
   

' PB1.Max = RecipeBas.rstCategory.RecordCount
 '   PB1.Value = RecipeBas.rstCategory.AbsolutePosition
  


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
CloseRs
End
End Sub

Private Sub Form_Terminate()
CloseRs
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
CloseRs
End
End Sub

Private Sub lblHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
lblHelp.FontBold = True

lblHelp.ForeColor = &HFF&
End Sub

Private Sub mnuAdd_Click()
On Error Resume Next
    Dim i As Integer
   
    
    'User wants to add a Record
    'Makes sure user has choosen a Catagory to add a Record to
    If Mid(tv.SelectedItem.Key, 1, 3) <> "Cat" Then
        MsgBox "Um ein neues Artikel hinzufügen, müssen Sie vorerst eine Kategorie auswählen."
        Exit Sub
    End If
 Command1.Enabled = True
    Command2.Enabled = False
    txtMultiplikator.Visible = True
    txtInstructions.Enabled = False
    
    rstRecipes.AddNew
    mnuMain.Enabled = False
    mnuSearch.Enabled = False
    For i = 2 To Toolbar1.Buttons.Count Step 1
        Toolbar1.Buttons(i).Enabled = False
    Next i
    tv.Enabled = False
    Toolbar1.Buttons(1).Enabled = True
    
    txtCategory = tv.SelectedItem.Text

End Sub

Private Sub mnuDelCat_Click()

    'user wants to delete a Catagory
    DelCat
End Sub


Private Sub mnuExit_Click()
CloseRs
End
End Sub

Private Sub mnuAddCat_Click()
AddCat ("")
End Sub

Private Sub mnuNeu_Click()

End Sub

Private Sub mnuSave_Click()
On Error Resume Next
    If txtCategory = "" _
     Or txtRecipe = "" _
      Or txtAuthor = "" _
       Or txtEK = "" Then
            MsgBox "Bitte fühlen Sie die Textfelder erst aus!"
            Exit Sub
    End If
    
    tv.Nodes.Add tv.SelectedItem, tvwChild, intTtlRecipeCount & "_", txtRecipe, 4, 3
    intTtlRecipeCount = intTtlRecipeCount + 1
    rstRecipes.Update
    
    Me.mnuMain.Enabled = True
   ' Me.mnuSearch.Enabled = True
    Me.tv.Enabled = True
    Me.Toolbar1.Buttons(1).Enabled = False
    For i = 2 To Toolbar1.Buttons.Count Step 1
        Toolbar1.Buttons(i).Enabled = True
    Next i
    
    'let user know Record has been added
    MsgBox "Artikel wurde gespeichert!"
    Command1.Enabled = False
    Command2.Enabled = True
    txtInstructions.Visible = True
End Sub

Private Sub Timer1_Timer()
StatusBar1.Panels(1) = Time
    If StatusBar1.Panels(2) <> Date Then
        StatusBar1.Panels(2) = Date
    End If
    If StatusBar1.Panels(3) <> intCatRecipeCount & " Derzeit gesamt Kategorien" Then
        StatusBar1.Panels(3) = intCatRecipeCount & " Derzeit gesamt Kategorien"
    End If
    If StatusBar1.Panels(4) <> intTtlRecipeCount & " Total Einträge" Then
        StatusBar1.Panels(4) = intTtlRecipeCount & " Total Einträge"
    End If
    If StatusBar1.Panels(5) <> intTtlCategories & " Total Kategorien" Then
        StatusBar1.Panels(5) = intTtlCategories & " Total Kategorien"
    End If
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
Dim rechne1 As String
lblHelp.FontBold = False
If txtEK.DataChanged = True Then
  rechne1 = txtEK.Text * txtMultiplikator.Text
  txtInstructions.Text = rechne1
  End If
  rechnenDiff
  
End Sub




Private Sub Tv_Expand(ByVal Node As MSComctlLib.Node)
    Node.Selected = True
End Sub

Private Sub Tv_DragOver(Source As Control, X As Single, y As Single, State As Integer)
    On Error Resume Next

    Dim target As Node
    Dim highlight As Boolean

    ' See what node we're above.
    Set target = tv.HitTest(X, y)
    
    ' If it's the same as last time, do nothing.
    If target Is TargetNode Then Exit Sub
    Set TargetNode = target
    
    highlight = False
    If Not (TargetNode Is Nothing) Then
        ' See what kind of node were above.
        highlight = True
    End If
    
    If highlight Then
        Set tv.DropHighlight = TargetNode
    Else
        Set tv.DropHighlight = Nothing
    End If
End Sub
Private Sub Tv_DragDrop(Source As Control, X As Single, y As Single)
    On Error Resume Next
    Dim intRec As Integer

    If Mid(SourceNode.Key, 1, 3) = "Cat" Then Exit Sub
    If SourceNode.Key = "BOOK" Then Exit Sub

    ' If it's the same as last time, do nothing.
    If tv.SelectedItem Is tv.DropHighlight Then Exit Sub
    
    intRec = Replace(SourceNode.Key, "_", "") + 1

    If Not (tv.DropHighlight Is Nothing) Then
        ' It's a valid drop. Set source node's
        ' parent to be the target node.
        If Mid(tv.DropHighlight.Key, 1, 3) = "Cat" Then
            Set SourceNode.Parent = tv.DropHighlight
            rstRecipes.AbsolutePosition = intRec
            rstRecipes.Fields(0) = tv.DropHighlight.Text
            txtCategory = tv.DropHighlight.Text
        Else
            tv.Nodes.Remove (SourceNode.Index)
            tv.Nodes.Add tv.DropHighlight, tvwNext, _
            SourceNode.Key, SourceNode.Text, SourceNode.Image, _
            SourceNode.SelectedImage
            rstRecipes.AbsolutePosition = intRec
            rstRecipes.Fields(0) = tv.DropHighlight.Parent.Text
            txtCategory = tv.DropHighlight.Parent.Text
        End If
        rstRecipes.Update
        Set tv.DropHighlight = Nothing
    End If

    Set SourceNode = Nothing

End Sub


Private Sub Tv_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    ' Set the item being dragged.
    Set SourceNode = tv.HitTest(X, y)
End Sub

Private Sub Tv_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbLeftButton Then
        ' Start a new drag.
        If y = tv.Height Then
            
        End If
        
        ' Select this node. When no node is highlighted,
        ' this node will be displayed as selected. That
        ' shows where it will land if dropped.
        Set tv.SelectedItem = SourceNode

        ' Fire the Begin Drag
        tv.Drag vbBeginDrag
    End If
End Sub
Private Sub Tv_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim i As Integer
    Select Case Mid(Node.Key, 1, 3)
    Case "Cat"
        If Node.Checked = True Then
            For i = 1 To Node.Children
                tv.Nodes(Node.Index + i).Checked = True
            Next
        Else
            For i = 1 To Node.Children
                tv.Nodes(Node.Index + i).Checked = False
            Next
        End If
    Case "BOO"
        If Node.Checked = True Then
            For i = 1 To intTtlRecipeCount + intTtlCategories
                tv.Nodes(Node.Index + i).Checked = True
            Next
        Else
            For i = 1 To intTtlRecipeCount + intTtlCategories
                tv.Nodes(Node.Index + i).Checked = False
            Next
        End If
    End Select
        
End Sub


Private Sub Tv_NodeClick(ByVal Node As MSComctlLib.Node)
    'this is what does the stuff when a user click on
    'a Itemin the TreeView
        
        Select Case Mid(Node.Key, 1, 3)
            Case "BOO"
                'we do nothing here
            Case "Cat"
                'If user clicks Catagory then populate category count Variable
                intCatRecipeCount = Node.Children
            Case Else
                If rstRecipes.AbsolutePosition < 0 Then
                    rstRecipes.MoveFirst
                End If
                If Node.Key = "0_" Then
                    rstRecipes.MoveFirst
                Else
                    rstRecipes.AbsolutePosition = ((Replace(Node.Key, "_", "")) + 1)
                End If
                intCatRecipeCount = Node.Parent.Children
              
        End Select
End Sub

Private Sub txtAuthor_Change()
    If Len(txtAuthor) > 75 Then
        txtAuthor = Mid(txtAuthor, 1, 75)
        txtAuthor.SelStart = 75
        Beep
    End If
End Sub

Private Sub txtAuthor_LostFocus()
    txtAuthor = Proper(txtAuthor)
    If Not rstRecipes.EOF And Not rstRecipes.BOF Then
        rstRecipes.Update
    End If
End Sub

Private Sub txtCategory_Change()
    If Len(txtCategory) > 75 Then
        txtCategory = Mid(txtCategory, 1, 75)
        txtCategory.SelStart = 75
        Beep
    End If
End Sub

Private Sub txtEK_Change()
If Len(txtEK) > 75 Then
        txtEK = Mid(txtEK, 1, 75)
        txtEK.SelStart = 75
        Beep
    End If
    
End Sub

Private Sub txtEK_LostFocus()
If Not rstRecipes.EOF And Not rstRecipes.BOF Then
        rstRecipes.Update
    End If
End Sub

Private Sub txtInstructions_LostFocus()
    If Not rstRecipes.EOF And Not rstRecipes.BOF Then
        rstRecipes.Update
    End If
End Sub

Private Sub txtRecipe_Change()
    If Len(txtRecipe) > 75 Then
        txtRecipe = Mid(txtRecipe, 1, 75)
        txtRecipe.SelStart = 75
        Beep
    End If
End Sub

Private Sub txtRecipe_LostFocus()
    txtRecipe = Proper(txtRecipe)
    If Not rstRecipes.EOF And Not rstRecipes.BOF Then
        rstRecipes.Update
    End If
End Sub


Private Sub rechnenDiff()
 Dim rechne As String
 rechne = txtInstructions.Text - txtEK.Text
             txtDiff.Text = rechne
End Sub
