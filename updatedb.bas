Attribute VB_Name = "updatedb"



Public Function runCleanUP()
'On Error Resume Next
    Dim dang
    
     
     Call deleteTables
    'update all tables relations for sharepoint and make local
     Data1 = updatedb("Calls", "RID", True)
     Data2 = updatedb("ContactPerson", "RID", True)
     Data3 = updatedb("CountryList", "RID", False)
     Data4 = updatedb("Criteria", "RID", False)
     data5 = updatedb("DefineAccountManager", "RID", False)
     data6 = updatedb("DefineDepartment", "RID", False)
     data7 = updatedb("DefineIndustry", "RID", False)
     data8 = updatedb("DefineMailing", "RID", False)
     data8 = updatedb("DefineProject", "RID", False)
     data9 = updatedb("DefineSalesPipeStatus", "RID", False)
     data10 = updatedb("DefineState", "RID", False)
     Data11 = updatedb("DefineStatus", "RID", False)
     Data12 = updatedb("Mailing", "RID", True)
     data13 = updatedb("Relation", "RID", False)
     data14 = updatedb("SalesPipe", "RID", True)
     data15 = updatedb("Selections", "RID", False)
     Call makeimagelocal

     'delete salespipe without reference (legacy problems)
     Data16 = deleteBrokenSalesPipe()
          
     're-create relations
     Data11 = createRelation("Calls", "RID")
     Data12 = createRelation("ContactPerson", "RID")
     Data112 = createRelation("Mailing", "RID")
     data114 = createRelation("SalesPipe", "RID")
     data115 = fixPrimaryKeys()
     'repair duplicate value error on index

     
     'update fields in calls
     next1233 = updateCalls
     newformssetup = updateForm
     'CurrentDb.Properties.Delete ("StartupForm")
  
           'update fields to long and richt text
     Call toLongText("Relation", "SearchName")
     Call toLongText("Relation", "RName")
     Call toLongText("Relation", "PhoneNo")
     Call toLongText("Relation", "FAXNo")
     Call toLongText("Relation", "Email")
     Call toLongText("Relation", "website")
     Call toLongText("Relation", "Activities")
     Call toLongText("Relation", "Address")
     Call toLongText("Relation", "ZipCode")
     Call toLongText("Relation", "City")
     Call toLongText("Relation", "StateProvincie")
     Call toLongText("Relation", "Country")

    
     Call toLongText("Calls", "AccountManager")
     Call toLongText("Calls", "Content")
     Call toLongText("Calls", "Attachment")
     Call toLongText("Calls", "Description")
     
     Call CheckEntries("Calls")
     Call CheckEntries("Relation")
     
     

     
     'Always the last commamnd!!
        RunCommand acCmdShareOnSharePoint
     
End Function
Public Function makeimagelocal()

DoCmd.OpenForm "Relation", acFormEdit
Forms!Relation.Controls("Image765").PictureType = 0
DoCmd.Save
DoCmd.Close


End Function

Public Function deleteTables()

Dim db As DAO.Database
Dim tdf As DAO.TableDef
Set db = CurrentDb
For Each tdf In db.TableDefs
    ' ignore system and temporary tables
    If Not (tdf.Name Like "MSys*" Or tdf.Name Like "~*") Then
        
        If Not ( _
        tdf.Name = "Relation" _
        Or tdf.Name = "Calls" _
        Or tdf.Name = "ContactPerson" _
        Or tdf.Name = "CountryList" _
        Or tdf.Name = "Criteria" _
        Or tdf.Name = "DefineAccountManager" _
        Or tdf.Name = "DefineDepartment" _
        Or tdf.Name = "DefineIndustry" _
        Or tdf.Name = "DefineMailing" _
        Or tdf.Name = "DefineProject" _
        Or tdf.Name = "DefineSalesPipeStatus" _
        Or tdf.Name = "DefineState" _
        Or tdf.Name = "DefineStatus" _
        Or tdf.Name = "Mailing" _
        Or tdf.Name = "SalesPipe" _
        Or tdf.Name = "Selections" _
        ) Then
        Debug.Print tdf.Name
        DoCmd.DeleteObject acTable, tdf.Name
        End If
        
        
    End If
Next
Set tdf = Nothing
Set db = Nothing

End Function

Public Function fixPrimaryKeys()
On Error Resume Next
   CurrentDb.Execute "CREATE UNIQUE INDEX SomeIndex1 ON DefineAccountManager(ID) WITH PRIMARY"
   CurrentDb.Execute "CREATE UNIQUE INDEX SomeIndex2 ON DefineDepartment(ID) WITH PRIMARY"
   CurrentDb.Execute "CREATE UNIQUE INDEX SomeIndex3 ON DefineIndustry(ID) WITH PRIMARY"
   CurrentDb.Execute "CREATE UNIQUE INDEX SomeIndex4 ON DefineMailing(ID) WITH PRIMARY"
   CurrentDb.Execute "CREATE UNIQUE INDEX SomeIndex5 ON DefineProject(ID) WITH PRIMARY"
   CurrentDb.Execute "CREATE UNIQUE INDEX SomeIndex6 ON DefineSalesPipeStatus(ID) WITH PRIMARY"
   CurrentDb.Execute "CREATE UNIQUE INDEX SomeIndex7 ON ContactPerson(ID) WITH PRIMARY"

End Function

Public Function test()

     Call toLongText("Calls", "AccountManager")
     Call toLongText("Calls", "Attachment")
     Call toLongText("Calls", "Description")
     Call toLongText("Calls", "Content")


End Function
Sub toLongText(ByVal tableName As String, ByVal fieldName As String)
Dim sqlString As String
Dim db As DAO.Database
'Set db = CurrentDb
Set db = CurrentDb()
sqlString = "ALTER TABLE " & tableName & " ALTER COLUMN " & fieldName & " MEMO"

DoCmd.RunSQL sqlString


db.Close

Set db = Nothing

DoCmd.OpenTable tableName, acViewDesign
DoCmd.Save acTable, tableName
DoCmd.Close

Call toRichText(tableName, fieldName)

End Sub
Function toRichText(ByVal tableName As String, ByVal fieldName As String)

'MsgBox fieldName

Dim db As DAO.Database
Dim tdf As DAO.TableDef
Dim fld As DAO.Field
Dim tableN As String
Dim fieldN As String
Dim newProp As String
Dim prp As DAO.Property


tableN = tableName
fieldN = fieldName

Set db = CurrentDb
Set tdf = db.TableDefs(tableN)
Set fld = tdf.Fields(fieldN)
db.Properties.Refresh

With fld.Properties("TextFormat")
    If .Value = acTextFormatPlain Then
        .Value = acTextFormatHTMLRichText
    End If
 End With

db.Close

Set db = Nothing

Exit Function

End Function

Public Function deleteBrokenSalesPipe()

    CurrentProject.Connection.Execute "DELETE RID From SalesPipe WHERE RID NOT IN (SELECT ID FROM Relation)"

End Function

Public Function test12()
    CheckIllChar ("Relation")
End Function

Public Function CheckIllChar(ByVal checkerstr As String)
'Check for space area which is not a space
 
Dim abData() As Byte
Dim ChrPos As String
Dim DatString As String
Dim Chrsel As String
Dim HexNr As String

Dim Counter As Integer
Dim MyString As String

DatString = checkerstr


'MsgBox DatString
    

Start = 1

Do
  pos = InStr(Start, DatString, " ", vbTextCompare)
  If pos > 0 Then
    Start = pos + 1  'alternatively: start = pos + Len(srch)
    Chrsel = MID(DatString, pos, 1)
    'MsgBox Chrsel
    
    abData = StrConv(Chrsel, vbFromUnicode)
    
    'MsgBox DatString
    
                Dim i As Integer
                For i = 0 To UBound(abData)
                    HexNr = Hex(abData(i))
                    If HexNr = B Then
                        MsgBox DatString
                    End If
                Next
  End If
Loop While pos > 0




End Function
Public Function randomtest()
    'Debug.Print ChrW(&H6C)
    Debug.Print Chr(76)
End Function
Public Function testchr()
    Dim chartype As Characters

    'chartype = ChrW(&HB)) 'ChrW(&HB)
        MsgBox ChrW(&HB78)
    
    Dim abData() As Byte
    abData = StrConv("", vbFromUnicode)

    Dim i As Integer
    For i = 0 To UBound(abData)
    Debug.Print Hex(abData(i)) & " ";
    Next
    
    Dim MyCalculate As String
    

End Function

Public Function runReplace()

 CheckEntries ("Copy Of error")
 
End Function


Public Function CheckEntries(ByVal WhichTable As String)
    
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset(WhichTable)
    Dim f As DAO.Field
    Dim Editstring As String
    Dim EditDate As String
    Dim ClmName As String
    
    Do While Not rs.EOF
                
         For Each f In rs.Fields
        If f.Type <> vbLong Then
        
            ClmName = f.Name
            'fix Date
                
                If ClmName = "CallDate" Or ClmName = "CallBackDate" Then
                      With rs
                      If Not IsNull(.Fields(ClmName)) Then
                      EditDate = .Fields(ClmName)
                       If Len(EditDate) < 10 Then
                        dt = CDate(EditDate)
                        myDate = dt + TimeValue("9:00")

                         .Edit
                        .Fields(ClmName) = myDate
                        .Update
                       End If
                      End If
                      
                    End With
                      
                Else
                            
                If ClmName <> "ID" Then
                If ClmName <> "RID" Then
      
                    With rs
         
                      If Not IsNull(.Fields(ClmName)) Then
                        Editstring = .Fields(ClmName)
                        EditstringChk = .Fields(ClmName)
                        'MsgBox Editstring
                        'CheckIllChar (Editstring)
                        Editstring = findInvisChar(Editstring)
                        Editstring = StripAccentb(Editstring)
                        Editstring = Replace(Editstring, ChrW(65533), "")
                        Editstring = Replace(Editstring, ChrW(37156), "")
                        Editstring = Replace(Editstring, ChrW(9633), "")
                        Editstring = Replace(Editstring, "", "")
                        Editstring = Replace(Editstring, "", "")
                        Editstring = Trim(Editstring)
                                                
                        If EditstringChk <> Editstring Then
                        .Edit
                        .Fields(ClmName) = Editstring
                        .Update
                        End If
                        

                
                      End If
    
                    End With
                    
                   
                    
                    End If
            End If
            End If
            End If
            Next
    
         rs.MoveNext
        Loop
               
End Function
        
        
Function findInvisChar(sInput As String) As String
Dim sSpecialChars As String
Dim i As Long
Dim sReplaced As String
Dim ln As Long


sSpecialChars = "" & Chr(1) & Chr(2) & Chr(3) & Chr(4) & Chr(5) & Chr(6) & Chr(7) & Chr(8) & Chr(9) & Chr(10) & Chr(11) & Chr(12) & Chr(13) & Chr(14) & Chr(15) & Chr(16) & Chr(17) & Chr(18) & Chr(19) & Chr(20) & Chr(21) & Chr(22) & Chr(23) & Chr(24) & Chr(25) & Chr(26) & Chr(27) & Chr(28) & Chr(29) & Chr(30) & Chr(31) & Chr(32) & ChrW(&HA0) & ChrW(&H3) & ChrW(&HB) 'This is your list of characters to be removed
'For loop will repeat equal to the length of the sSpecialChars string
'loop will check each character within sInput to see if it matches any character within the sSpecialChars string
For i = 1 To Len(sSpecialChars)
    ln = Len(sInput) 'sets the integer variable 'ln' equal to the total length of the input for every iteration of the loop
    sInput = Replace$(sInput, MID$(sSpecialChars, i, 1), "")
    If ln <> Len(sInput) Then sReplaced = sReplaced & MID$(sSpecialChars, i, 1)
    If ln <> Len(sInput) Then sReplaced = sReplaced & IIf(MID$(sSpecialChars, i, 1) = Chr(10), "<Line Feed>", MID$(sSpecialChars, i, 1)) & IIf(MID$(sSpecialChars, i, 1) = Chr(1), "<Start of Heading>", MID$(sSpecialChars, i, 1)) & IIf(MID$(sSpecialChars, i, 1) = Chr(9), "<Character Tabulation, Horizontal Tabulation>", MID$(sSpecialChars, i, 1)) & IIf(MID$(sSpecialChars, i, 1) = Chr(13), "<Carriage Return>", MID$(sSpecialChars, i, 1)) & IIf(MID$(sSpecialChars, i, 1) = Chr(28), "<File Separator>", MID$(sSpecialChars, i, 1)) & IIf(MID$(sSpecialChars, i, 1) = Chr(29), "<Group separator>", MID$(sSpecialChars, i, 1)) & IIf(MID$(sSpecialChars, i, 1) = Chr(30), "<Record Separator>", MID$(sSpecialChars, i, 1)) & IIf(MID$(sSpecialChars, i, 1) = Chr(31), "<Unit Separator>", MID$(sSpecialChars, i, 1)) & IIf(MID$(sSpecialChars, i, 1) = ChrW(&HA0), "<Non-Breaking Space>", MID$(sSpecialChars, i, 1)) 'Currently will remove all control character but only tell the user about Bell and Line Feed
Next

'MsgBox sReplaced & " These were identified and removed"
findInvisChar = sInput


End Function 'end of function


Function StripAccentb(text) As Variant

Dim A As String * 1
Dim B As String * 1
Dim i As Integer
Dim S As String
'Const AccChars = "ŠŽšžŸÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖÙÚÛÜÝàáâãäåçèéêëìíîïðñòóôõöùúûüýÿ"
'Const RegChars = "SZszYAAAAAACEEEEIIIIDNOOOOOUUUUYaaaaaaceeeeiiiidnooooouuuuyy"
Const AccChars = "" ' using less characters is faster
Const RegChars = ".."
S = text
For i = 1 To Len(AccChars)
A = MID(AccChars, i, 1)
B = MID(RegChars, i, 1)
S = Replace(S, A, B)
'Debug.Print (S)
Next


StripAccentb = S

Exit Function
End Function
Public Function updateForm()
    
   'Address
    DoCmd.OpenForm "Relation", acFormEdit
    Forms!Relation![Staat / Provincie].RowSource = "SELECT DefineState.State FROM DefineState;"
    Forms!Relation![Land].RowSource = "SELECT CountryList.Country, CountryList.Country FROM CountryList;"
  
   
    'Relation status
    Forms!Relation![Project].RowSource = "SELECT DefineIndustry.Industry FROM DefineIndustry;"
    Forms!Relation![cboIndustry].RowSource = "SELECT DefineProject.Project FROM DefineProject;"
    Forms!Relation![Account manager].RowSource = "SELECT DefineAccountManager.AccountManager FROM DefineAccountManager;"
    Forms!Relation![Status].RowSource = "SELECT DefineStatus.Status FROM DefineStatus;"
    DoCmd.Close acForm, "Relation", acSaveYes
    
    
    'contactperson
    DoCmd.OpenForm "Relation_ContactPersonTab", acFormEdit
    Forms!Relation_ContactPersonTab![CPDepartment].RowSource = "SELECT DefineDepartment.Department FROM DefineDepartment;"
    DoCmd.Close acForm, "Relation_ContactPersonTab", acSaveYes
    
    
    'Calls
    DoCmd.OpenForm "Relation_Calls", acFormEdit
    Forms!Relation_Calls![AccountManager].RowSource = "SELECT DefineAccountManager.AccountManager FROM DefineAccountManager;"
    Forms!Relation_Calls![Title].ControlSource = "Description"
    DoCmd.Close acForm, "Relation_Calls", acSaveYes
    
    'Mailing
    DoCmd.OpenForm "Relation_Mailing", acFormEdit
    Forms!Relation_Mailing![Mailing_mailing_ID].RowSource = "SELECT DefineMailing.Mailing FROM DefineMailing;"
    DoCmd.Close acForm, "Relation_Mailing", acSaveYes
    
    'Salespipe
    DoCmd.OpenForm "Relation_SalesPipe", acFormEdit
    Forms!Relation_SalesPipe!Status.RowSource = "SELECT DefineSalesPipeStatus.Status FROM DefineSalesPipeStatus;"
    DoCmd.Close acForm, "Relation_SalesPipe", acSaveYes
    


End Function
Public Function createRelation(ByVal tableName As String, ByVal fieldName As String)
    On Error GoTo errhandler


    Dim db As DAO.Database
    Dim newRelation As DAO.Relation
    Dim relatingField As DAO.Field
    Dim relationUniqueName As String
    Dim Attributes As String
        
    Attributes = dbRelationUpdateCascade + dbRelationDeleteCascade + dbRelationLeft
    

    primaryTableName = "Relation"
    primaryFieldName = "ID"
    foreignTableName = tableName
    foreignFieldName = fieldName
    
    relationUniqueName = primaryTableName + "_" + primaryFieldName + _
                         "__" + foreignTableName + "_" + foreignFieldName
    
    Set db = CurrentDb()
    
    'Arguments for CreateRelation(): any unique name,
    'primary table, related table, attributes.
    Set newRelation = db.createRelation(relationUniqueName, _
                            primaryTableName, foreignTableName, Attributes)
    'The field from the primary table.
    Set relatingField = newRelation.CreateField(primaryFieldName)
    'Matching field from the related table.
    relatingField.ForeignName = foreignFieldName
    'Add the field to the relation's Fields collection.
    newRelation.Fields.Append relatingField
    'Add the relation to the database.
    db.Relations.Append newRelation
    
    Set db = Nothing
    
    createRelation = True

        
Exit Function

errhandler:
    Debug.Print Err.Description + " (" + relationUniqueName + ")"
    createRelation = False



End Function
Public Function updateCalls()
    On Error GoTo errhandler
    bFieldExists = FieldExists("Calls", "Title") ' Custom field_exists in table function

    If bFieldExists Then
    CurrentDb.TableDefs("Calls").Fields("[Title]").Name = "Description"
    End If

errhandler:
    Debug.Print Err.Description + "Calls Description already exists"


End Function

Public Function updatedb(ByVal tableName As String, ByVal fieldName As String, ByVal addColumn As Boolean)
    
    On Error Resume Next
    
    Dim strField As String
    Dim curDatabase As Object
    Dim tblTest As Object
    Dim fldNew As Object
    Dim prp As DAO.Property
        
    Set curDatabase = CurrentDb
    MakeTableLocal (tableName)
    
    If addColumn = True Then
        Set tblTest = curDatabase.TableDefs(tableName)
        'Set rst = db.OpenRecordset(tableName)
                
        strField = fieldName
        strTable = tableName

        bFieldExists = FieldExists(strTable, strField) ' Custom field_exists in table function

        If bFieldExists Then
        'Set prpNew = curDatabase.TableDefs(strTable).Fields(strField).CreateProperty("DisplayControl", dbText, acComboBox)
        'Set prpNew = curDatabase.TableDefs(strTable).Fields(strField).CreateProperty("RowSourceType", dbText, "Table/Query")
        Set prpNew = curDatabase.TableDefs(strTable).Fields(strField).CreateProperty("RowSource", dbText, "SELECT [Relation].[ID] FROM [Relation] ORDER BY [ID];")
        curDatabase.TableDefs(strTable).Fields(strField).Properties.Append prpNew
        curDatabase.TableDefs(strTable).Fields(strField).Properties("DisplayControl") = acComboBox
        'curDatabase.TableDefs(strTable).Fields(strField).Properties("RowSourceType") = "Table/Query"
        'curDatabase.TableDefs(strTable).Fields(strField).Properties("RowSource") = "SELECT [Relation].[ID] FROM [Relation] ORDER BY [ID];"
        End If

    'rst.Close ' Recordset must release the table data before we can alter the table!

        If bFieldExists = False Then
            Set fldNew = tblTest.CreateField(strField, dbInteger)
            tblTest.Fields.Append fldNew
            Set prp = fldNew.CreateProperty("DisplayControl", dbLong, AcControlType.acComboBox)
            fldNew.Properties.Append prp
            Set prp = fldNew.CreateProperty("RowSourceType", dbText, "Table/Query")
            fldNew.Properties.Append prp
            Set prp = fldNew.CreateProperty("RowSource", dbText, "SELECT [Relation].[ID] FROM [Relation] ORDER BY [ID];")
            fldNew.Properties.Append prp
            'CurrentDb.Execute "UPDATE tableName SET strField = tableName.RID"
            
        End If
    End If
Set db = Nothing
End Function

Public Function FieldExists(ByVal tableName As String, ByVal fieldName As String) As Boolean
    Dim nLen As Long

    On Error GoTo Failed
    With DBEngine(0)(0).TableDefs(tableName)
    .Fields.Refresh
    nLen = Len(.Fields(fieldName).Name)

    If nLen > 0 Then FieldExists = True

    End With
    Exit Function
Failed:
    If Err.Number = 3265 Then Err.Clear ' Error 3265 : Item not found in this collection.
    FieldExists = False
End Function


Sub MakeTableLocal(tableName As String, Optional deleteOriginal As Boolean = True)
    
    If Not (tableName Like "MSys*" Or tableName Like "~*") Then
          
    Dim DbPath As Variant, TblName As Variant

    'get path of linked table
    DbPath = DLookup("Database", "MSysObjects", "Name='" & tableName & "' And Type=6")
    'Get the real name of the linked table (in case it has been given an alias in the link)
    TblName = DLookup("ForeignName", "MSysObjects", "Name='" & tableName & "' And Type=6")
    If IsNull(DbPath) Then
        'Either a local table, or the wrong table name has been supplied, exit the sub
        Exit Sub
    End If

    'delete linked table
    If deleteOriginal Then
        DoCmd.DeleteObject acTable, tableName
    Else
        'If we're not deleting the existing table we'll have to rename the imported table to avoid
        'overwriting it etc
        tableName = tableName & " - local"
    End If
    
    'import the table as a local, unlinked table
    DoCmd.TransferDatabase acImport, "Microsoft Access", DbPath, acTable, TblName, tableName
    End If

End Sub



Public Function openscript()
   
    
    DoCmd.RunCommand acCmdHideMessageBar
   
    Dim strBackEndPath As String
    Dim strFinalPath As String
    Dim lenPath As Integer
    Dim i As Integer
    Dim j As Integer

    strBackEndPath = CurrentDb.TableDefs("Relation").Connect
    
    
    ' Now remove the datebase  & password prefix
    j = InStrRev(strBackEndPath, "DATABASE=") + 9
    strFinalPath = MID(strBackEndPath, j)
    'MsgBox strFinalPath
    i = InStrRev(strFinalPath, "LIST") - 2
    strFinalPath = MID(strFinalPath, 1, i)
    'MsgBox strFinalPath


    verified = TestURL(strFinalPath)

    If verified = True Then DoCmd.OpenForm "Relation"
    If verified = False Then MsgBox "Connection with Database not possible, Please login First"
    If verified = False Then CreateObject("Shell.Application").Open "http://office.com"
   
End Function


Public Function TestURL(ByVal url As String) As Boolean
    On Error GoTo errhandler
    Dim Request As Object
    Dim ff As Integer
    Dim rc As Variant

    Set Request = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    With Request
      .Open "GET", url, False
      .Send
      rc = .StatusText
    End With
    Set Request = Nothing
    'MsgBox rc
    If rc = "OK" Then TestURL = True


    Exit Function
errhandler:      MsgBox Err.Description
   

End Function
