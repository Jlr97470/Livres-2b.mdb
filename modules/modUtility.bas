Attribute VB_Name = "modUtility"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                     UTILITY FUNCTIONS                        '
'                                                              '
'       This module contains useful functions that you         '
'       can use in expressions on your forms and reports.      '
'                                                              '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Compare Database   'Use database order for string comparisons
Option Explicit

Public Function IsLoaded(ByVal strFormName As String) As Integer
' From Northwind sample database:
' Returns True if the specified form is open in Form view or Datasheet view.

' These variables are used to test the return values of the SysCmd function
'  and the CurrentView property of the requested form.
    Const conObjStateClosed = 0
    Const conDesignView = 0

    'The first If statement uses the SysCmd function to check the current
    'state of the requested form. It can be in one of four states: not open
    'or nonexistent, open, new, or changed but not saved.
    If SysCmd(acSysCmdGetObjectState, acForm, strFormName) <> conObjStateClosed Then

        'The second If statement checks for the current view of the requested
        'form, assuming the previous If statement found it to be open. If
        'the form is currently open in Form View, the function will return
        'true. If the form is in Design View, the function will return false.
        If Forms(strFormName).CurrentView <> conDesignView Then
            IsLoaded = True
        End If

    End If

End Function


Public Function IsNothing(varToTest As Variant) As Integer
'  Tests for a "logical" nothing based on data type
'  Empty and Null = Nothing
'  Number = 0 is Nothing
'  Zero length string is Nothing
'  Date/Time is never Nothing

    IsNothing = True

    Select Case varType(varToTest)
        Case vbEmpty
            Exit Function
        Case vbNull
            Exit Function
        Case vbBoolean
            If varToTest Then IsNothing = False
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency
            If varToTest <> 0 Then IsNothing = False
        Case vbDate
            IsNothing = False
        Case vbString
            If (Len(varToTest) <> 0 And varToTest <> " ") Then IsNothing = False
    End Select

End Function

Public Function Larger(lngA() As Long) As Integer
'  Input:  An array of Long values of any length
'
'  Output: Integer index of the largest value
'

    Dim lngBig As Long, intX As Integer, intI As Integer

    intI = LBound(lngA)
    lngBig = lngA(intI)
    Larger = intI

    For intX = intI + 1 To UBound(lngA)
        If lngA(intX) > lngBig Then
            lngBig = lngA(intX)
            Larger = intX
        End If
    Next intX

End Function

Public Function Smaller(lngA() As Long) As Integer
'  Input:  An array of Long values of any length
'
'  Output: Integer index of the smallest value
'

    Dim lngSmall As Long, intX As Integer, intI As Integer

    intI = LBound(lngA)
    lngSmall = lngA(intI)
    Smaller = intI

    For intX = intI + 1 To UBound(lngA)
        If lngA(intX) < lngSmall Then
            lngSmall = lngA(intX)
            Smaller = intX
        End If
    Next intX

End Function

Public Function FormAssist()
' Generic function to respond to OnAction on
'   custom menu or toolbar help command
' Function checks for an active form, then
'   looks for a "FormHelp" handler subroutine
'   on that form.
Dim frm As Form

    ' Set an error trap
    On Error Resume Next

    ' Try to locate a form that has the focus
    Set frm = Screen.ActiveForm
    If Err <> 0 Then
        ' Error means no active form,
        '  so open standard Office Assistant
        Application.Assistant.Help
        Exit Function
    End If

    ' No error, so try to call the FormHelp
    '   method of the active form
    frm.FormHelp
    If Err <> 0 Then
        ' Error means no FormHelp method for
        '  the current form,
        '  so open standard Office Assistant
        Application.Assistant.Help
    End If
End Function
