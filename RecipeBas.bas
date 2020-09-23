Attribute VB_Name = "RecipeBas"
'*****************************************
'*** MC Store - The Store database     ***
'*** Coded by Robert Niedziela (c)2003 ***
'*** some code comes from PSC Thanks!  ***
'*** http://www.personal-webserver.de  ***
'*****************************************
'*****************************************

Option Explicit
' Set up global Virables
'Connections
Public cnnRecipe As ADODB.Connection
Public cnnExport As ADODB.Connection

'Execute Commands
Public cmdChange As ADODB.Command

'RecordSets
Public rstExport As ADODB.Recordset
Public rstRecipes As ADODB.Recordset
Public rstCategory As ADODB.Recordset

' Set Category Count Variable
Public intCatCount As Integer
Public intCatRecipeCount As Integer
Public intTtlRecipeCount As Integer
Public intTtlCategories As Integer
Public strCnn As String
Public strXCnn As String

' Set Drag and Drop variables
Public SourceNode As Object
Public TargetNode As Object





'*******************************************
' Add Klasse!
'*******************************************

Public Function StartIt()
    Dim tvNode As Node
    Dim strCategory As String
    Dim i As Integer
    Dim booExists As Boolean
    Dim strSQL As String
    
    strCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & App.Path & _
             "\dbstore.mdb"";Mode=ReadWrite;Persist Security Info=False;" & _
             "Jet OLEDB:Compact Without Replica Repair=False"
    
    Set cnnRecipe = New ADODB.Connection
    Set rstRecipes = New ADODB.Recordset
    
    If cnnRecipe.State = adStateClosed Then
        cnnRecipe.CursorLocation = adUseClient
        cnnRecipe.Open strCnn
    End If
    
    rstRecipes.Open "Select * From Recipes Order By Category, WARENNAME", cnnRecipe, adOpenStatic, adLockBatchOptimistic, adCmdText
    

    'Reset Field Owners to Speed Up Refresh
   frmMain.txtCategory.DataField = ""
    frmMain.txtRecipe.DataField = ""
    frmMain.txtAuthor.DataField = ""
    frmMain.txtInstructions.DataField = ""
    frmMain.txtEK.DataField = ""
    
    
    'Clear TreeView
    frmMain.tv.Nodes.Clear

    'Put in the Main Node
    Set tvNode = frmMain.tv.Nodes.Add(, , "BOOK", "Index", 1)
    
    'Make sure Main Node is Visible
    intCatCount = 1
    With rstRecipes
        If .RecordCount = 0 Then Exit Function
        'Open The RecordSet for each Child Node (Catagory)
        'start at the first Record
        .MoveFirst
        strCategory = "Cat" & intCatCount
        Set tvNode = frmMain.tv.Nodes.Add("BOOK", tvwChild, _
                    strCategory, .Fields(0), 2, 3)
        While Not .EOF
            If frmMain.tv.Nodes(strCategory).Text <> .Fields(0) Then
                'Put the Child Nodes into the TreeView (Catagories)
                intCatCount = intCatCount + 1
                strCategory = "Cat" & intCatCount
                Set tvNode = frmMain.tv.Nodes.Add("BOOK", tvwChild, _
                            strCategory, .Fields(0), 2, 3)
                'Set strCategory Variable
            End If
            'Put in the children Node
            'Place "_" after absolute position so that it in not interprited as a true Index Number
            Set tvNode = frmMain.tv.Nodes.Add(strCategory, tvwChild, (.AbsolutePosition - 1) & "_", _
                         .Fields(1) & "", 4, 3)
            intTtlRecipeCount = intTtlRecipeCount + 1
            'Make sure the Child Nodes are Visible
            .MoveNext
        Wend
    End With
    intTtlCategories = intCatCount
     
    frmMain.tv.Nodes.Item("BOOK").Expanded = True
      
    ' Open connection.
    strSQL = "SELECT Category.Category From Category WHERE (((Category.Category) " & _
              "Not In (Select Distinct Category from Recipes where Category.Category = Recipes.Category)))"
          
             
    'Set cnnRecipe = New ADODB.Connection
    Set rstCategory = New ADODB.Recordset
    'cnnRecipe.Open strCnn
    rstCategory.CursorType = adOpenStatic
    rstCategory.Open strSQL, cnnRecipe, , , adCmdText
    ' Add Categories that do not exist in Recipes table but do exist in Category table
    If rstCategory.RecordCount > 0 Then
    
        With rstCategory
            .MoveFirst
            While Not rstCategory.EOF
             
                intCatCount = intCatCount + 1
               
                Set tvNode = frmMain.tv.Nodes.Add("BOOK", tvwChild, _
                            "Cat" & intCatCount, .Fields(0), 2, 3)
                .MoveNext
            Wend
         
        End With
    End If
    
    'Set Fields Back To Original DataSource
    Set frmMain.txtCategory.DataSource = rstRecipes
    Set frmMain.txtRecipe.DataSource = rstRecipes
    Set frmMain.txtAuthor.DataSource = rstRecipes
    Set frmMain.txtInstructions.DataSource = rstRecipes
    Set frmMain.txtEK.DataSource = rstRecipes
    

    'Set Fields Back To Original DB Owner
    frmMain.txtCategory.DataField = "Category"
    frmMain.txtRecipe.DataField = "Warenname"
    frmMain.txtAuthor.DataField = "Warenserial"
    frmMain.txtInstructions.DataField = "VK"
    frmMain.txtEK.DataField = "EK"
    
    
    Set rstRecipes.ActiveConnection = Nothing
    Set rstCategory.ActiveConnection = Nothing
    cnnRecipe.Close
    
    'Select Top Node
    frmMain.tv.Nodes(1).Selected = True


End Function


Public Sub AddCat(strCategory)
    'Function to add a Catagory to the DataBase
    On Error Resume Next
          
    Dim strMessage As String, strTitle As String, strSQLChange As String
    
    ' Show the Input Box for the Name of Catagory to add to DataBase
    If Trim(strCategory) <= "" Then
Again:
        strMessage = "Geben Sie jetzt die neue Kategorie ein!"
        strTitle = "Neue Kategorie"
        strCategory = LCase(InputBox(strMessage, strTitle))
        If Len(strCategory) = 0 Then Exit Sub
        If Len(strCategory) > 75 Then
            MsgBox "Sie können nicht mehr als 75 zeichen eingeben! " & _
                   Chr(13) & Chr(13) & "Fehler", vbOKOnly + vbCritical _
                   , "Input Error"
            GoTo Again
        End If
    End If
    If Trim(strCategory) = "" Then Exit Sub
    
    ' Open connection.
    'Set cnnRecipe = New ADODB.Connection
    cnnRecipe.Open strCnn

    strSQLChange = "INSERT INTO Category ( Category )" & _
                   "SELECT '" & strCategory & "' AS Expr1;"

    ' Create command object.
    Set cmdChange = New ADODB.Command
    Set cmdChange.ActiveConnection = cnnRecipe
    cmdChange.CommandText = strSQLChange
    cmdChange.Execute
    
    If Err.Number <> 0 Then
        ' Notify User of Error
        MsgBox "Kann ihre eingabe: [" & strCategory & "]nicht ins Datenbank speichern! beachten Sie das sie keine gleichen Namen verwenden sollten." & _
               " Kategorie existiert bereits.", vbCritical + vbOKOnly, "Kategorie Error"
    Else
        intCatCount = intCatCount + 1
        frmMain.tv.Nodes.Add "BOOK", tvwChild, "Cat" & (intCatCount), _
                             strCategory, 2, 3
        frmMain.tv.Nodes("Cat" & intCatCount).Selected = True
        intTtlCategories = intTtlCategories + 1
    End If
    MsgBox "Neue eintrag wurde Erfolgreich gespeichert!", vbInformation, "Eintrag"
    
    cnnRecipe.Close
    Set cmdChange = Nothing
End Sub

Public Function DelCat()
    'Function to Delete a Catagory
    Dim strSQLChange As String
    Dim intIndex As Integer
    Dim Msg, Style, Title, Response
   
    'Make user has choosen a Catagory to Delete
    If Mid(frmMain.tv.SelectedItem.Key, 1, 3) <> "Cat" Then
       MsgBox "Please First Choose a Catagory to Delete"
       Exit Function
    End If

    'Warn User He/She is about to Delete the Catagory
    Msg = "This will Delete Catagory " & Chr$(34) & frmMain.tv.SelectedItem & Chr$(34) _
          & vbCrLf & "All of It's Recipes" & vbCrLf & "Do you want to continue ?"
    Style = vbYesNo + vbCritical + vbDefaultButton2
    Title = "Warnning"
    Response = MsgBox(Msg, Style, Title)
    
    If Response = vbNo Then Exit Function

    'Delete The Catagory (Table)
    ' Open connection.
    'Set cnnRecipe = New ADODB.Connection
    cnnRecipe.Open strCnn

    ' Process Delete from Recipes Table
    strSQLChange = "DELETE From Recipes Where Category = '" & frmMain.tv.SelectedItem & "'"
    Set cmdChange = New ADODB.Command
    Set cmdChange.ActiveConnection = cnnRecipe
    cmdChange.CommandText = strSQLChange
    cmdChange.Execute
    
    ' Process Delete from Category Table
    strSQLChange = "DELETE From Category Where Category = '" & frmMain.tv.SelectedItem & "'"
    cmdChange.CommandText = strSQLChange
    cmdChange.Execute
    
    cnnRecipe.Close
    Set cmdChange = Nothing
   
    'Delete Category Node from TreeView
    frmMain.tv.Nodes.Remove (frmMain.tv.SelectedItem.Index)
    intTtlCategories = intTtlCategories - 1

    'Reset Recordset Cursor to Beginning of File
    frmMain.tv.Nodes.Item("BOOK").Selected = True
    
End Function
Public Function CloseRs()
    'Function to Close RecordSets and DataBase
    On Error Resume Next
    
    ' In Case user has made an update movenext to commit update to recordset
    ' For some reason the Update command does not do this when the form is
    ' unloaded so I used Move 0 to get around this issue.
    If Not rstRecipes.EOF And Not rstRecipes.BOF Then
        rstRecipes.Move 0
    End If
    cnnRecipe.Open strCnn                           'reopen connection
    Set rstRecipes.ActiveConnection = cnnRecipe     'reconnect recordset
    rstRecipes.UpdateBatch                          'Update Recipes
    
    Set rstRecipes.ActiveConnection = Nothing
    cnnRecipe.Close

End Function
Public Function Proper(X)
    Dim temp$, C$, OldC$, i As Integer
    If IsNull(X) Then
        Exit Function
    Else
        temp$ = CStr(LCase(X))
        ' Initialize OldC$ to a single space because first
        ' letter needs to be capitalized but has no preceding letter.
        OldC$ = " "
        For i = 1 To Len(temp$)
            C$ = Mid$(temp$, i, 1)
            If C$ >= "a" And C$ <= "z" And (OldC$ < "a" Or OldC$ > "z") Then
                Mid$(temp$, i, 1) = UCase$(C$)
            End If
            OldC$ = C$
        Next i
        Proper = temp$
    End If
End Function
Public Function Normalize()
    'On Error Resume Next

    'Change Mouse Pointer to HourGlass
    frmMain.MousePointer = 11
    
    'Set Virables
    Dim rs As Recordset
    Dim DB As Database
    Dim rsRecipes As Recordset
    Dim Table As TableDef
    Dim i As Integer
    

    'Open DataBase
    Set DB = OpenDatabase(App.Path & "\dbstore.mdb")

    With DB
        Set rsRecipes = DB.OpenRecordset("Recipes")
        'Move through the DataBase and get the Catagories (Tables)
        For Each Table In .TableDefs
            'We don't Want the Next Selven Ms Access Stuff in the Tables
            If Table.Name = "MSysACEs" Then GoTo TheNextTable
            If Table.Name = "MSysModules" Then GoTo TheNextTable
            If Table.Name = "MSysModules2" Then GoTo TheNextTable
            If Table.Name = "MSysObjects" Then GoTo TheNextTable
            If Table.Name = "MSysQueries" Then GoTo TheNextTable
            If Table.Name = "MSysRelationships" Then GoTo TheNextTable
            If Table.Name = "MSysAccessObjects" Then GoTo TheNextTable
            If Table.Name = "Recipes" Then GoTo TheNextTable
            If Table.Name = "Category" Then GoTo TheNextTable
            If Table.Name = "zzzNothing" Then GoTo TheNextTable
            'Open The RecordSet for each Child Node (Catagory)
            Set rs = DB.OpenRecordset(Table.Name)
            rs.MoveFirst
            'Move through the Records
            For i = 1 To rs.RecordCount
                rsRecipes.AddNew
                rsRecipes.Fields(0) = LCase(Table.Name)
                rsRecipes.Fields(1) = Proper(rs.Fields(0).Value)
                rsRecipes.Fields(2) = Proper(rs.Fields(1).Value)
                rsRecipes.Fields(3) = rs.Fields(2).Value & vbCrLf & _
                                      vbCrLf & vbCrLf & rs.Fields(3).Value
                'Move to the Next record
                rsRecipes.Update
                rs.MoveNext
            Next i
TheNextTable:
        Next Table
    End With
    
    'Call function to close the DataBase
    rs.Close
    rsRecipes.Close
    'Close DataBase
    DB.Close
    'Set Virables to Nothing
    Set rs = Nothing
    Set rsRecipes = Nothing
    Set DB = Nothing
    'Change Mouse Pointer to Standard
    frmMain.MousePointer = 0
End Function



