Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Use database order for string comparisons
'This form starts the application

Private Sub Form_Timer()
    ' Form stays open for 5 seconds, then closes...
    DoCmd.Close acForm, Me.Name
    ' .. and opens the signon form
    DoCmd.OpenForm "frmSignon"
End Sub
