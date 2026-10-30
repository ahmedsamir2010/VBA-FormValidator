VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormValidatorTestEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Optional regression-test fixture, not a production dependency.
Private WithEvents mValidator As FormValidator
Public RaiseOnCustom As Boolean
Public ReenterOnCustom As Boolean
Public RequireValue As Boolean
Public SawBusy As Boolean

Public Property Set Validator(ByVal Value As FormValidator)
    Set mValidator = Value
End Property

Private Sub mValidator_CustomValidation(ByVal control As MSForms.Control, _
    ByVal Value As String, ByVal RuleName As String, ByRef Handled As Boolean, _
    ByRef IsValid As Boolean, ByRef ErrorMessage As String)

    SawBusy = mValidator.IsBusy
    If RaiseOnCustom Then Err.Raise vbObjectError + 2201, "Test callback", "Expected callback error"
    If ReenterOnCustom Then
        Dim unused As Boolean
        unused = mValidator.ValidateAll
    End If
    Select Case RuleName
        Case "ContainsExcel"
            Handled = True
            IsValid = (InStr(1, Value, "Excel", vbTextCompare) > 0)
            ErrorMessage = "Include Excel"
        Case "ConditionalRequired"
            Handled = True
            IsValid = (Not RequireValue Or Len(Value) > 0)
            ErrorMessage = "Conditionally required"
    End Select
End Sub
