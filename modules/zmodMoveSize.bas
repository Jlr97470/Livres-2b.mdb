Attribute VB_Name = "zmodMoveSize"
Option Compare Database
Option Explicit

Function CopyCommandBar(strOrigCBName As String, strNewCBName As String) As Boolean
    ' This procedure copies command bar named in strOrigCBName variable to new
    ' command bar named in strNewCBName variable.
    Dim cbrOriginal As CommandBar, intOrigCount As Integer
    Dim cbrCopy As CommandBar, cbrCtl As CommandBarControl

    On Error GoTo CBCopy_Err

    Set cbrOriginal = CommandBars(strOrigCBName)

intOrigCount = cbrOriginal.Controls.Count
    Set cbrCopy = CommandBars.Add(strNewCBName)
    ' Make sure new command bar is visible.
    cbrCopy.Visible = True

    For Each cbrCtl In cbrOriginal.Controls
        cbrCtl.Copy cbrCopy
    Next cbrCtl
Exit Function

CBCopy_Err:
    MsgBox Err.Description
    Exit Function
End Function

Sub ResizeForms()
Dim intI As Integer, strFormName As String

    For intI = 0 To CurrentProject.AllForms.Count - 1
        strFormName = CurrentProject.AllForms(intI).Name
        DoCmd.OpenForm FormName:=strFormName, View:=acDesign
        DoCmd.MoveSize 180, 180, 10800, 6480
        Forms(strFormName).Filter = ""
        DoCmd.RunCommand acCmdSave
        DoCmd.Close acForm, strFormName
    Next intI

End Sub
Sub ResizeQd()
Dim intI As Integer, strQueryName As String

    For intI = 0 To CurrentData.AllQueries.Count - 1
        strQueryName = CurrentData.AllQueries(intI).Name
        If Left(strQueryName, 1) <> "~" Then
            DoCmd.OpenQuery QueryName:=strQueryName, View:=acDesign
            DoCmd.MoveSize 360, 360, 10800, 5760
            DoCmd.RunCommand acCmdSave
            DoCmd.Close acQuery, strQueryName
        End If
    Next intI

End Sub

Sub ResizeReports()
Dim intI As Integer, strReportName As String

    For intI = 0 To CurrentProject.AllReports.Count - 1
        strReportName = CurrentProject.AllReports(intI).Name
        DoCmd.OpenReport ReportName:=strReportName, View:=acDesign
        DoCmd.MoveSize 180, 90, 10800, 6480
        If Left(strReportName, 3) = "rpt" Then
            Reports(strReportName).Toolbar = "Book Print Tools"
        End If
        Reports(strReportName).Filter = ""
        DoCmd.RunCommand acCmdSave
        DoCmd.Close acReport, strReportName
    Next intI

End Sub
