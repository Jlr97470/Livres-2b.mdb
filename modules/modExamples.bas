Attribute VB_Name = "modExamples"
Option Compare Database   'Use database order for string comparisons
Option Explicit           'Force me to declare all variables

Public Function CreateErrTable()
  ' This function creates a table containing a list of all the valid
  '  Access application error codes -- using ADO

  ' Declare variables used in this function
  Dim cnMyConnect As ADODB.Connection, catThisDB As New ADOX.Catalog
  Dim tblErrTable As New ADOX.Table, idxPKey As New ADOX.Index
  Dim rcdErrRecSet As New ADODB.Recordset, lngErrCode As Long, intMsgRtn As Integer
  Dim varReturnVal As Variant, varErrString As Variant

  ' Create Errors table with Error Code and Error String fields.
  ' Initialize the MyConnection connection variable to the current connection
  Set cnMyConnect = CurrentProject.Connection

  ' Trap error if table doesn't exist
  On Error Resume Next         ' Skip to next statement if an error occurs
  rcdErrRecSet.Open "SELECT * FROM ErrTable;", cnMyConnect, _
    adOpenKeyset, adLockOptimistic, adCmdText

  Select Case Err   ' See if an error was raised

  Case 0        ' No error - table must alread exist
    On Error GoTo 0   ' Turn off error trapping
    intMsgRtn = MsgBox("ErrTable already exists.  Do you want to delete and rebuild all rows?", _
        vbQuestion + vbYesNo, "Microsoft Press")
    If intMsgRtn = vbYes Then    ' Reply was YES - delete rows and rebuild
      cnMyConnect.Execute "DELETE ErrTable.* FROM ErrTable;"  ' Run quick SQL to delete rows
    Else                 ' Reply was NO - done
      rcdErrRecSet.Close   ' Close the table
      Exit Function        ' And exit
    End If

  Case -2147217865    ' Couldn't find table, so build it
    On Error GoTo 0   ' Turn off error trapping
    ' Open the catalog for this database
    catThisDB.ActiveConnection = cnMyConnect
    ' Create a new table to contain the error rows
    With tblErrTable
        .Name = "ErrTable"
        ' Create a field in ErrTable to contain the error code
        ' and append it to the Columns collection of the table
        .Columns.Append "ErrorCode", adInteger
        ' Create a field in ErrTable to contain the error description
        ' and append it to the Columns collection of the table
        .Columns.Append "ErrorString", adLongVarWChar
    End With
    ' Append the new table to the Tables collection in
    ' the current catalog
    catThisDB.Tables.Append tblErrTable
    ' Set the description for the first field
    tblErrTable("ErrorCode").Properties("Description") = "Error code number"
    ' Set the description for the second field
    tblErrTable("ErrorString").Properties("Description") = "Error message"
    ' Now, let's create a Primary Key
    idxPKey.Name = "PrimaryKey"
    idxPKey.PrimaryKey = True
    idxPKey.Unique = True
    idxPKey.Columns.Append "ErrorCode"
    ' And append it to the indexes collection of the table
    tblErrTable.Indexes.Append idxPKey

    ' Set recordset to Errors Table recordset.
    rcdErrRecSet.Open "[ErrTable]", cnMyConnect, adOpenKeyset, adLockOptimistic

  Case Else         ' Can't identify the error - write message and bail
    MsgBox "Unknown error in CreateErrTable " & Err & ", " & Error$(Err), vbExclamation
    Exit Function

  End Select

  'Initialize the progress meter on the status bar
  varReturnVal = SysCmd(acSysCmdInitMeter, "Building Error Table", 32767)
  ' Turn on hourglass to show this may take a while
  DoCmd.Hourglass True

  ' Start a transaction to make it go fast
  cnMyConnect.BeginTrans

  ' Loop through Microsoft Access error codes, skipping error codes
  ' that generate "User-defined error" or "Reserved Error" message.
  For lngErrCode = 1 To 32767
    varErrString = AccessError(lngErrCode)
    If IsNothing(varErrString) Or varErrString = "Application-defined or object-defined error" Then
        ' If AccessError returned nothing, then try Error
        varErrString = Error(lngErrCode)
    End If
    If Not IsNothing(varErrString) Then
        If varErrString <> "Application-defined or object-defined error" Then
            ' Add each error code and string to Errors table.
            rcdErrRecSet.AddNew
            rcdErrRecSet("ErrorCode") = lngErrCode
            ' Some error messages are longer than 255 - truncate
            rcdErrRecSet("ErrorString") = varErrString
            rcdErrRecSet.Update
        End If
    End If
    ' Update the status meter
    varReturnVal = SysCmd(acSysCmdUpdateMeter, lngErrCode)
  ' Process next error code
  Next lngErrCode

  ' Commit all added rows
  cnMyConnect.CommitTrans

  ' Close recordset.
  rcdErrRecSet.Close
  ' Turn off the hourglass - we're done
  DoCmd.Hourglass False
  ' And reset the status bar
  varReturnVal = SysCmd(acSysCmdClearStatus)
  ' Select the new table in the database window to refresh the list
  DoCmd.SelectObject acTable, "ErrTable", True
  ' pop up a confirmation dialog
  MsgBox "Errors table created."
End Function

Public Function ShowTables()
' An example of looping through all the Tables
'  and their indexes -- using ADO
On Error Resume Next  ' Skip over errors in system tables

Dim cn As ADODB.Connection, cat As New ADOX.Catalog
Dim tbl As ADOX.Table, idx As ADOX.Index, col As ADOX.Column
  ' Get a copy of the current connection
  Set cn = CurrentProject.Connection
  ' Point our catalog object at the current connection
  cat.ActiveConnection = cn
  ' Loop through the tables
  For Each tbl In cat.Tables
    If tbl.Type = "Table" Then
        Debug.Print tbl.Name
        ' Loop through the Indexes
        For Each idx In tbl.Indexes
            Debug.Print " ", idx.Name
            For Each col In idx.Columns
                Debug.Print " ", " ", col.Name
            Next col
        Next idx
    End If
  Next tbl

End Function

Public Sub ListFields()
' An example of looping through all the Tables
'  and their fields -- using ADO
On Error Resume Next  ' Skip over errors in system tables

Dim cn As ADODB.Connection, cat As New ADOX.Catalog
Dim tbl As ADOX.Table, col As ADOX.Column, varVal As Variant
Dim prp As ADOX.Property, varType As Variant
On Error Resume Next

  ' Get a copy of the current connection
  Set cn = CurrentProject.Connection
  ' Point the catalog at the current connection
  cat.ActiveConnection = cn
  ' Loop through all the tables
  For Each tbl In cat.Tables
    If tbl.Type = "Table" Then
        Debug.Print tbl.Name
        ' Loop through all the columns and their properties
        For Each col In tbl.Columns
            Debug.Print " ", col.Name
            ' Use a Variant in case the property value is invalid
            varVal = " "
            varVal = col.Attributes
            Debug.Print " ", " ", "Attributes", varVal
            varVal = " "
            varVal = col.DefinedSize
            Debug.Print " ", " ", "DefinedSize", varVal
            varVal = " "
            varVal = col.NumericScale
            Debug.Print " ", " ", "NumericScale", varVal
            varVal = " "
            varVal = col.Precision
            Debug.Print " ", " ", "Precision", varVal
            varVal = " "
            varVal = col.RelatedColumn
            Debug.Print " ", " ", "RelatedColumn", varVal
            varVal = " "
            varVal = col.SortOrder
            Debug.Print " ", " ", "SortOrder", varVal
            varVal = " "
            varVal = col.Type
            Debug.Print " ", " ", "Type", varVal
            For Each prp In col.Properties
                varVal = " "
                varType = " "
                varVal = prp.Value
                varType = prp.Type
                Debug.Print " ", " ", prp.Name, varVal, varType
            Next prp
        Next col
    End If
  Next tbl

End Sub

Public Sub LoadStores()
' Code to load stores from the attached "US Stores.txt" file
Dim cn As ADODB.Connection, rcdI As New ADODB.Recordset, rcdO As New ADODB.Recordset
Dim intLine As Integer, intAddit As Integer, intComma As Integer
    Set cn = CurrentProject.Connection
    rcdI.Open "[USStores]", cn, adOpenForwardOnly
    rcdO.Open "tblUSStoresXmpl", cn, adOpenKeyset, adLockOptimistic
    Do Until rcdI.EOF
        ' If on a blank line,
        If IsNothing(rcdI!Field1) Then
            ' We're between rows - set line counter to 0
            intLine = 0
            If intAddit Then
                rcdO.Update
                intAddit = False
            End If
        Else
            ' Keep a counter: 1 = name, 2 = address, 3 = City/State/Zip, 4 = phone, 5 = email
            intLine = intLine + 1
            Select Case intLine
                Case 1
                    rcdO.AddNew
                    intAddit = True
                    rcdO!StoreName = rcdI!Field1
                Case 2
                    rcdO!Address = rcdI!Field1
                Case 3
                    intComma = InStr(rcdI!Field1, ",")
                    If intComma = 0 Then intComma = 2
                    rcdO!City = Left(rcdI!Field1, intComma - 1)
                    rcdO!StateOrProvince = Mid(rcdI!Field1, intComma + 2, 2)
                    rcdO!PostalCode = Mid(rcdI!Field1, intComma + 5)
                Case 4
                    rcdO!PhoneNumber = "(" & Left(rcdI!Field1, 3) & ") " & Mid(rcdI!Field1, 5, 8)
                Case Is > 4
                    If InStr(rcdI!Field1, "@") <> 0 Then
                        rcdO!Email = rcdI!Field1
                    End If
                    If Left(rcdI!Field1, 5) = "http:" Then
                        rcdO!WebAddress = rcdI!Field1
                    End If
                    If Left(rcdI!Field1, 5) = "Fax: " Then
                        rcdO!FaxNumber = "(" & Mid(rcdI!Field1, 6, 3) & ") " & Mid(rcdI!Field1, 11, 8)
                    End If
            End Select
        End If
        rcdI.MoveNext
    Loop
    If intAddit Then rcdO.Update
    rcdI.Close
    rcdO.Close

End Sub

Sub FixAllowDesign()
Dim objFrm As AccessObject, frm As Form
    For Each objFrm In CurrentProject.AllForms
        DoCmd.OpenForm FormName:=objFrm.Name, View:=acDesign
        Set frm = Forms(objFrm.Name)
        If frm.AllowDesignChanges = True Then
            frm.AllowDesignChanges = False
            DoCmd.RunCommand acCmdSave
        End If
        DoCmd.Close acForm, frm.Name
    Next objFrm
End Sub
