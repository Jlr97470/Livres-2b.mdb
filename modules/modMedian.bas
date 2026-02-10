Attribute VB_Name = "modMedian"
Option Compare Database   'Use database order for string comparisons
Option Explicit


Function DMedian(strFieldName As String, strDomainName As String, Optional varWhere As Variant) As Variant
' Function to calculate the Median numeric value of a column in a domain
' Input: Name of the field, Name of the domain, Predicate (optional) - all strings
' Output: Median numeric value or Null if the field is a string or error encountered
Dim cn As ADODB.Connection, rcd As New ADODB.Recordset
Dim strSQL As String, lngRecords As Long, intVarType As Integer
Dim varValue As Variant, lngRows As Long, strWhere As String
'On Error GoTo DMedianBail
    ' Set default return of Null
    DMedian = Null
    ' Check for possible "missing" third argument
    If Not IsMissing(varWhere) Then
        If varType(varWhere) = vbString Then strWhere = varWhere
    End If
    ' Point to the current connection
    Set cn = CurrentProject.Connection
    ' Open the domain sorted by the field
    strSQL = "SELECT [" & strFieldName & "] FROM [" & strDomainName & "]"
    If Len(strWhere) <> 0 Then
        strSQL = strSQL & " WHERE " & strWhere
    End If
    strSQL = strSQL & " Order By [" & strFieldName & "];"
    rcd.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    ' If no rows, then return Null
    If rcd.EOF Then
        rcd.Close
        Exit Function
    End If
    ' Get the number of rows
    rcd.MoveLast
    lngRecords = rcd.RecordCount
    rcd.MoveFirst
    intVarType = varType(rcd(strFieldName))
    ' If the data type of the field isn't a number, then return Null
    If (intVarType < vbInteger Or intVarType > vbDate) And (intVarType <> vbDecimal) Then
        rcd.Close
        Exit Function
    End If
    ' Calculate the "middle" of the recordset
    lngRows = lngRecords \ 2
    ' If an even number, then do an average
    If lngRecords Mod 2 = 0 Then
        rcd.Move lngRows - 1
        varValue = rcd(strFieldName)
        rcd.MoveNext
        varValue = (varValue + rcd(strFieldName)) / 2
        Else
        rcd.Move lngRows
        varValue = rcd(strFieldName)
    End If
    DMedian = varValue
    ' Be sure to return the proper data type
    Select Case intVarType
        Case vbInteger
            DMedian = CInt(varValue)
        Case vbLong
            DMedian = CLng(varValue)
        Case vbSingle
            DMedian = CSng(varValue)
        Case vbDouble
            DMedian = CDbl(varValue)
        Case vbCurrency
            DMedian = CCur(varValue)
        Case vbDate
            DMedian = CDate(varValue)
        Case vbDecimal
            DMedian = CDec(varValue)
    End Select
    rcd.Close

DMedianBail:
    Exit Function

End Function
