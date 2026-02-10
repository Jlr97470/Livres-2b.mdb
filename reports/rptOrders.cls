Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub GroupHeader0_Format(Cancel As Integer, FormatCount As Integer)
    Me.Page = 1
End Sub

Private Sub PageHeader_Format(Cancel As Integer, FormatCount As Integer)
    If IsNothing(Me!StoreName) Then
        Me!StoreName.Visible = False
        Me!StAddress.Visible = False
        Me!StCSZ.Visible = False
        Me!txtMSPName.Visible = True
        Me!txtMSPAddress.Visible = True
        Me!txtMSPCSZ.Visible = True
    Else
        Me!StoreName.Visible = True
        Me!StAddress.Visible = True
        Me!StCSZ.Visible = True
        Me!txtMSPName.Visible = False
        Me!txtMSPAddress.Visible = False
        Me!txtMSPCSZ.Visible = False
    End If
    If IsNothing(Me!ShipCountry) Then
        Me!StCountry.Visible = False
        Me!txtMSPCountry.Visible = False
        Me!Country.Visible = False
        Me!ShipCountry.Visible = False
    Else
        Me!StCountry.Visible = Me!StoreName.Visible
        Me!txtMSPCountry.Visible = Me!txtMSPName.Visible
        Me!Country.Visible = True
        Me!ShipCountry.Visible = True
    End If
    If Me!PayBy < 3 Then
        Me!CCNumber.Visible = False
        Else
        Me!CCNumber.Visible = True
    End If

End Sub
