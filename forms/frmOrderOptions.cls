Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Use database order for string comparisons
Option Explicit

Private Sub Close_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmd1_Click()
    ' Orders for currently signed-on customer
    gstrWhereOrder = "[CustomerID] = " & glngCurrentCustID
    If 0 = DCount("*", "tblOrders", gstrWhereOrder) Then
        If vbYes = MsgBox("No orders exist for the currently signed on customer." & _
            "  Do you want to create a new order?", _
            vbQuestion + vbYesNo, "Microsoft Press") Then
            cmd3_Click
            Exit Sub
            Else
            Exit Sub
        End If
    End If

    DoCmd.Close acForm, Me.Name
    DoCmd.OpenForm FormName:="frmOrders", WhereCondition:=gstrWhereOrder

End Sub

Private Sub cmd2_Click()
    ' Unprinted orders for currently signed on customer...
    gstrWhereOrder = "[CustomerID] = " & glngCurrentCustID & " AND [Printed] = False"
    If 0 = DCount("*", "tblOrders", gstrWhereOrder) Then
        If vbYes = MsgBox("No open orders exist for the currently signed on customer." & _
            "  Do you want to create a new order?", _
            vbQuestion + vbYesNo, "Microsoft Press") Then
            cmd3_Click
            Exit Sub
            Else
            Exit Sub
        End If
    End If

    DoCmd.Close acForm, Me.Name
    DoCmd.OpenForm FormName:="frmOrders", WhereCondition:=gstrWhereOrder

End Sub

Private Sub cmd3_Click()
    ' New order for currently signed on customer...
    gstrWhereOrder = "[CustomerID] = " & glngCurrentCustID

    DoCmd.Close acForm, Me.Name
    DoCmd.OpenForm FormName:="frmOrders", WhereCondition:=gstrWhereOrder, _
        OpenArgs:="New"

End Sub

Private Sub cmd4_Click()
    ' All orders for customer selected above...
    gstrWhereOrder = "[CustomerID] = " & Me!CustomerID
    If 0 = DCount("*", "tblOrders", gstrWhereOrder) Then
        MsgBox "No orders exist for the selected customer.", _
            vbExclamation, "Microsoft Press"
        Exit Sub
    End If

    DoCmd.Close acForm, Me.Name
    DoCmd.OpenForm FormName:="frmOrders", WhereCondition:=gstrWhereOrder

End Sub

Private Sub cmd5_Click()
    ' Unprinted orders for customer selected above...
    gstrWhereOrder = "[CustomerID] = " & Me!CustomerID & " AND [Printed] = False"
    If 0 = DCount("*", "tblOrders", gstrWhereOrder) Then
        If vbYes = MsgBox("No open orders exist for the selected customer." & _
            "  Do you want to see all orders for this customer?", _
            vbQuestion + vbYesNo, "Microsoft Press") Then
            cmd4_Click
            Exit Sub
            Else
            Exit Sub
        End If
    End If

    DoCmd.Close acForm, Me.Name
    DoCmd.OpenForm FormName:="frmOrders", WhereCondition:=gstrWhereOrder

End Sub

Private Sub cmd6_Click()
    ' Review all orders...
    gstrWhereOrder = ""
    DoCmd.Close acForm, Me.Name
    DoCmd.OpenForm "frmOrders"
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
        .Heading = "Sample Balloon Help for Order Options"
        ' This is a tip, so use the right icon
        .Icon = msoIconTip
        ' Set up labels to display helpful hints
        .Labels(1).Text = "If you are signed on, you can edit your orders or create new ones using the options on the left side of the form."
        .Labels(2).Text = "Use the options on the right side of the form to select a specific customer and review orders."
        ' Set up title text inside the balloon
        .Text = "Instructions for Order Options:"
        ' Activate the balloon
        .Show
        ' Code will stop here until user dismisses the baloon
    End With
End Sub
Private Sub Form_Load()
    ' Set form options based on whether a customer is signed on ...
    If IsNothing(glngCurrentCustID) Then
        Me!txtCurrCust = "No customer signed on."
        Me!cmd1.Enabled = False
        Me!cmd2.Enabled = False
        Me!cmd3.Enabled = False
        If IsLoaded("frmOrders") Then
            Me!CustomerID = Forms!frmOrders!CustomerID
            Else
            Me!CustomerID = DLookup("[CustomerID]", "tblCustomers")
        End If
        Else
        Me!txtCurrCust = DLookup("[LastUserName]", "tblLocalValues")
        Me!CustomerID = glngCurrentCustID
    End If
End Sub
