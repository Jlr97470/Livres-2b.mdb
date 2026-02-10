Attribute VB_Name = "zmodDumpTbl"
Option Compare Database
Option Explicit

Public Sub DumpADOXTables()
' This routine examines all tables in the ADOX catalog
' It writes rows to a temp table for each field in the table
' and flags which fields are in the Primary Key
' The resulting table is useful as a crossref to see where fields
'  are defined and used.

' For the equivalent DAO module,
'  see "DumpTablesAndQueries" in zmodDumpTbl
'  in the Entertain sample database

Dim cn As ADODB.Connection, tdf As ADOX.Table, fld As ADOX.Column
Dim idx As ADOX.Index, rcd As New ADODB.Recordset, cat As New ADOX.Catalog
Dim strPKeyNames(1 To 10) As String, intPKeyCount As Integer
Dim intI As Integer, varRtn As Variant, intCount As Integer, intRowNum As Integer
    ' Blow past any errors
    On Error Resume Next
    ' Set a variable pointing to the current connection
    Set cn = CurrentProject.Connection
    ' Use the execute method to clean out all rows in the demo table
    cn.Execute "Delete * From zTableFields"
    ' Open a recordset to append new rows only
    rcd.Open "zTableFields", cn, adOpenKeyset, adLockOptimistic

    ' Open the catalog
    cat.ActiveConnection = cn

    ' Be nice, show a temperature bar as we process tables
    ' Call SysCmd to set up the count
    varRtn = SysCmd(acSysCmdInitMeter, "Processing tables...", cat.Tables.Count)

    ' Go through all the tables
    For Each tdf In cat.Tables
        ' Skip the System and Hidden tables and the demo table
        If (tdf.Type = "Table" Or tdf.Type = "LINK") And _
            (tdf.Name <> "zTableFields" And Left(tdf.Name, 1) <> "~") Then
            ' Set counter for Primary Key fields to zero
            intPKeyCount = 0
            ' Set artificial sequence number to zero
            intRowNum = 0
            ' Loop through all the indexes looking for the Primary Key
            '  (It should be named PrimaryKey, but it might not be!)
            For Each idx In tdf.Indexes
                If idx.PrimaryKey Then
                    ' Found the Primary Key - save all the field names for later
                    For Each fld In idx.Columns
                        ' We're set up to save only the first 10 fields in the PKey
                        If intPKeyCount = 10 Then Exit For
                        ' Increment the counter
                        intPKeyCount = intPKeyCount + 1
                        ' Save the field name at that counter position
                        strPKeyNames(intPKeyCount) = fld.Name
                    Next fld
                ' This Exit For stops our scan for PKey once we found it
                Exit For
                End If
            Next idx

            ' Now, examine all the fields and dump them to the demo table
            For Each fld In tdf.Columns
                ' Set up to add a new row
                rcd.AddNew
                ' Save table name
                rcd!TableName = tdf.Name
                ' Generate a sequence number
                intRowNum = intRowNum + 1
                rcd!Sequence = intRowNum
                ' Save the field name
                rcd!FieldName = fld.Name
                ' Grab the description
                rcd!Description = fld.Properties("Description")
                ' Scan the list of saved PrimaryKey field names
                '  to see if this field is in the PKey
                For intI = 1 To intPKeyCount
                    If fld.Name = strPKeyNames(intI) Then
                        ' Found IT!, so mark the PKey column
                        rcd!PKey = True
                        Exit For
                    End If
                Next intI
                ' Now, figure out the field type as a text string
                Select Case fld.Type
                    Case adBoolean
                        rcd!Type = "Yes/No"
                        rcd!Length = 1
                    Case adUnsignedTinyInt
                        rcd!Type = "Byte"
                        rcd!Length = 1
                    Case adSmallInt
                        rcd!Type = "Integer"
                        rcd!Length = 2
                    Case adInteger
                        If fld.Properties("Autoincrement") Then
                            rcd!Type = "Auto Number"
                            Else
                            rcd!Type = "Long Integer"
                        End If
                        rcd!Length = 4
                    Case adCurrency
                        rcd!Type = "Currency"
                        rcd!Length = 8
                    Case adDecimal
                        rcd!Type = "Decimal"
                        rcd!Length = 8
                    Case adSingle
                        rcd!Type = "Single"
                        rcd!Length = 4
                    Case adDouble
                        rcd!Type = "Double"
                        rcd!Length = 8
                    Case adDate
                        rcd!Type = "Date/Time"
                        rcd!Length = 8
                    Case adWChar
                        rcd!Type = "Text"
                        rcd!Length = fld.DefinedSize
                    Case adLongVarBinary
                        rcd!Type = "OLE Object"
                    Case adLongVarWChar
                        If fld.Properties("Jet OLEDB:Hyperlink") Then
                            rcd!Type = "Hyperlink"
                        Else
                            rcd!Type = "Memo"
                        End If
                    Case adGUID
                        rcd!Type = "Replication ID"
                        rcd!Length = 16
                    Case Else
                        rcd!Type = "Unknown"
                End Select
                ' Done -- save the row
                rcd.Update
            ' This gets the next field
            Next fld
        ' This End IF matches test for system, hidden tables
        End If

    ' Call SysCmd to update the count
    intCount = intCount + 1
    varRtn = SysCmd(acSysCmdUpdateMeter, intCount)
    ' This gets the next tabledef
    Next tdf
    ' Clear the status bar
    varRtn = SysCmd(acSysCmdClearStatus)

End Sub
