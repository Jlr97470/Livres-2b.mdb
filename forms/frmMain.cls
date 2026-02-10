Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Use database order for string comparisons
Option Explicit

Sub Close_Click()
    ' We're done -- the Form_Close routine will clean up
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdAuthors_Click()
    ' Authors...
    DoCmd.OpenForm "frmAuthorList"
End Sub

Private Sub cmdBooks_Click()
    ' Books...
    DoCmd.OpenForm "frmBookList"
End Sub

Private Sub cmdCustomers_Click()
    ' Customers...
    DoCmd.OpenForm "frmCustomers"
    ' If someone signed on, ...
    If glngCurrentCustID <> 0 Then
        ' .. position form at that customer
        Forms!frmCustomers.Recordset.FindFirst "[CustomerID] = " & glngCurrentCustID
    End If

End Sub

Private Sub cmdOrders_Click()
    ' Orders...
    DoCmd.OpenForm "frmOrderOptions"
End Sub

Private Sub cmdStores_Click()
    ' Stores...
    DoCmd.OpenForm "frmStoreList"
End Sub

Private Sub Form_Close()
    ' Reset the global values
    glngCurrentCustID = 0
    gintPrintPreview = acPreview
    gintIsAdmin = False
    ' Select an object in the DB window to reveal it again
    DoCmd.SelectObject acTable, "tblAuthors", True
End Sub

Public Sub FormHelp()
' Custom assistant help method for this form
Dim balCurrent As Balloon

    ' Create a new assistant baloon
    Set balCurrent = Assistant.NewBalloon

    With balCurrent
        ' Set balloon type to bullets
        .BalloonType = msoBalloonTypeBullets
        ' Set mode to auto-close as soon as the user clicks anywhere else
        .Mode = msoModeAutoDown
        ' Add an OK button just for looks
        .Button = msoButtonSetOK
        ' Define the balloon title
        .Heading = "Sample Balloon Help"
        ' This is a tip, so use the right icon
        .Icon = msoIconTip
        ' Set up title text inside the balloon
        .Text = "This is a sample baloon help tied to this form.  Many forms in this application have" & _
            " sample help available.  Click on the Office Assistant button on the toolbar to check it out!"
        ' Activate the balloon
        .Show
        ' Code will stop here until user dismisses the baloon
    End With
End Sub

Private Sub Form_Load()
    ' Halt screen flashing
    Application.Echo False
    DoEvents
    ' Select an object in the database window
    DoCmd.SelectObject acForm, "frmMain", True
    ' And hide the window
    DoCmd.RunCommand acCmdWindowHide
    ' Show 'em what we've done
    Application.Echo True

End Sub

Private Sub Form_Open(Cancel As Integer)
    If glngCurrentCustID = 0 Then
        ' The following is a good example of a MsgBox that
        '  takes care to format the instructions on separate lines
        If vbOK = MsgBox("No customer is currently 'signed on' to this application." & _
            vbCrLf & vbCrLf & "Are you sure you want to continue?" & _
            vbCrLf & vbCrLf & "Click OK to continue without signing on -- you won't be able to create any new orders." & _
            vbCrLf & vbCrLf & "Click Cancel to return to the database window.  " & _
            vbCrLf & "Start the application and sign on by opening 'frmSignon'.", _
            vbQuestion + vbOKCancel, "Microsoft Press") Then
            gintPrintPreview = acPreview
            Exit Sub
        End If
        Cancel = True
        DoCmd.SelectObject acTable, "tblAuthors", True
    End If
End Sub

Private Sub Reports_Click()
    ' Reports...
    DoCmd.OpenForm "frmReports", acNormal
End Sub
