VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' FormValidator 2.0.0 -- standalone, Windows Office / MSForms.
' Import this class into a VBA project containing a UserForm.
' No extra checked reference is needed for Dictionary or RegExp.
'
' Existing AddField positional arguments are unchanged. New options are appended.
' Deliberate changes: FullMatch defaults to True; MultiLine defaults to False.
' IgnoreCase still defaults to True for compatibility with the original class.
' Regex objects are created and syntax-checked when a pattern rule is registered.
' Required-only fields do not need a RegExp engine.
'
' Rules are evaluated in registration order; only the first error per field is
' stored. Errors describe the last completed validation, not unvalidated edits.
' Each selected bound value of a multi-select ListBox is checked separately.
' Numeric parsing follows the Office/Windows locale; values are never rewritten.
'
' For custom rules, declare the validator WithEvents and handle CustomValidation.
' The handler must set Handled=True for its RuleName; an unhandled rule raises a
' configuration error instead of silently accepting the value.
'
' During ResetForm, native Change events may fire. In those handlers use:
'   If v Is Nothing Then Exit Sub
'   If v.IsBusy Then Exit Sub
'   v.ValidateField Me.txtName
'
' Keep one validator per form. Reusing a control name from a different form is
' rejected explicitly. Do not rename registered controls; remove/re-add them.

Public Event CustomValidation(ByVal control As MSForms.Control, _
    ByVal Value As String, ByVal RuleName As String, ByRef Handled As Boolean, _
    ByRef IsValid As Boolean, ByRef ErrorMessage As String)

Private Const ERR_VALIDATOR As Long = vbObjectError + 2100
Private Const DEFAULT_ERROR As String = "This field is invalid"

Private mFields As Object
Private mErrors As Object
Private mOrder As Collection
Private mErrorBorderColor As Long
Private mStrictRegistration As Boolean
Private mSkipDisabled As Boolean
Private mSkipHidden As Boolean
Private mBusy As Boolean

Private Sub Class_Initialize()
    Set mFields = NewDictionary()
    Set mErrors = NewDictionary()
    Set mOrder = New Collection
    mErrorBorderColor = RGB(211, 47, 47)
End Sub

Private Sub Class_Terminate()
    ' Best-effort UI cleanup only: the form may already be unloading.
    On Error Resume Next
    ClearErrorsCore
    Set mFields = Nothing
    Set mErrors = Nothing
    Set mOrder = Nothing
    On Error GoTo 0
End Sub

Public Property Get Version() As String
    Version = "2.0.0"
End Property

Public Property Get IsBusy() As Boolean
    IsBusy = mBusy
End Property

Public Property Get ErrorBorderColor() As Long
    ErrorBorderColor = mErrorBorderColor
End Property

Public Property Let ErrorBorderColor(ByVal Value As Long)
    AssertIdle
    mErrorBorderColor = Value
    Dim key As Variant
    Dim field As Object
    For Each key In mOrder
        If mErrors.Exists(key) Then
            Set field = mFields(key)
            ApplyBorderState field, False
        End If
    Next key
End Property

' False preserves the old behavior for ValidateField on an unregistered control.
Public Property Get StrictRegistration() As Boolean
    StrictRegistration = mStrictRegistration
End Property

Public Property Let StrictRegistration(ByVal Value As Boolean)
    AssertIdle
    mStrictRegistration = Value
End Property

' Both skip options default to False. They inspect the control itself, not its
' ancestor containers or which MultiPage page is currently selected.
Public Property Get SkipDisabled() As Boolean
    SkipDisabled = mSkipDisabled
End Property

Public Property Let SkipDisabled(ByVal Value As Boolean)
    AssertIdle
    mSkipDisabled = Value
End Property

Public Property Get SkipHidden() As Boolean
    SkipHidden = mSkipHidden
End Property

Public Property Let SkipHidden(ByVal Value As Boolean)
    AssertIdle
    mSkipHidden = Value
End Property

Public Sub AddField(ByVal control As MSForms.Control, ByVal Pattern As String, _
    Optional ByVal ErrorMsg As String = "This field is invalid", _
    Optional ByVal Required As Boolean = True, _
    Optional ByVal TrimValue As Boolean = True, _
    Optional ByVal ColorBorderOnError As Boolean = True, _
    Optional ByVal FullMatch As Boolean = True, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = False, _
    Optional ByVal MustBeTrue As Boolean = False, _
    Optional ByVal FieldLabel As String = "", _
    Optional ByVal RequiredErrorMsg As String = "")

    AssertIdle
    RequireSupportedControl control
    If MustBeTrue Then
        If Not IsBooleanControl(control) Then
            Fail "MustBeTrue requires a CheckBox, OptionButton or ToggleButton."
        End If
    End If

    Dim previous As Object
    Dim field As Object
    Dim rules As Collection
    Dim rule As Object
    Dim borderStyle As Long
    Dim borderColor As Long
    Dim specialEffect As Long

    Set previous = FindField(control, False)
    Set field = NewDictionary()
    Set rules = New Collection
    ErrorMsg = NonEmptyMessage(ErrorMsg)

    ' Prepare first. A bad replacement leaves the old rule/error/style intact.
    If Len(Pattern) > 0 Then
        Set rule = PatternRule(control, Pattern, ErrorMsg, FullMatch, IgnoreCase, MultiLine)
        rules.Add rule
    End If

    field.Add "Control", control
    field.Add "Required", Required
    field.Add "TrimValue", TrimValue
    field.Add "MustBeTrue", MustBeTrue
    field.Add "ErrorMsg", ErrorMsg
    If Len(RequiredErrorMsg) = 0 Then RequiredErrorMsg = ErrorMsg
    field.Add "RequiredErrorMsg", RequiredErrorMsg
    If Len(FieldLabel) = 0 Then FieldLabel = control.Name
    field.Add "Label", FieldLabel
    field.Add "Rules", rules
    field.Add "ColorBorderOnError", ColorBorderOnError
    field.Add "Highlighted", False

    If previous Is Nothing Then
        field.Add "HasBorder", TryReadBorder(control, borderStyle, borderColor)
        field.Add "BorderStyle", borderStyle
        field.Add "BorderColor", borderColor
        field.Add "HasSpecialEffect", TryReadSpecialEffect(control, specialEffect)
        field.Add "SpecialEffect", specialEffect
        mFields.Add control.Name, field
        mOrder.Add control.Name, control.Name
    Else
        field.Add "HasBorder", previous("HasBorder")
        field.Add "BorderStyle", previous("BorderStyle")
        field.Add "BorderColor", previous("BorderColor")
        field.Add "HasSpecialEffect", previous("HasSpecialEffect")
        field.Add "SpecialEffect", previous("SpecialEffect")
        ApplyBorderState previous, True
        Set mFields(control.Name) = field
    End If
    If mErrors.Exists(control.Name) Then mErrors.Remove control.Name
End Sub

' Add an additional regex rule without replacing the field's existing rules.
Public Sub AddPatternRule(ByVal control As MSForms.Control, ByVal Pattern As String, _
    Optional ByVal ErrorMsg As String = "This field is invalid", _
    Optional ByVal FullMatch As Boolean = True, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = False)

    AssertIdle
    Dim field As Object
    Dim rule As Object
    Set field = FindField(control, True)
    If Len(Pattern) = 0 Then Fail "AddPatternRule requires a non-empty pattern."
    Set rule = PatternRule(control, Pattern, ErrorMsg, FullMatch, IgnoreCase, MultiLine)
    AppendRule field, rule
End Sub

' MaxLength=-1 means no upper limit. Len counts VBA string characters.
Public Sub AddLengthRule(ByVal control As MSForms.Control, ByVal MinLength As Long, _
    Optional ByVal MaxLength As Long = -1, _
    Optional ByVal ErrorMsg As String = "The text length is invalid")

    AssertIdle
    Dim field As Object
    Dim rule As Object
    Set field = FindField(control, True)
    If MinLength < 0 Or MaxLength < -1 Then Fail "Invalid length limits."
    If MaxLength <> -1 Then
        If MaxLength < MinLength Then Fail "MaxLength must be at least MinLength."
    End If
    Set rule = NewRule("Length", ErrorMsg)
    rule.Add "Minimum", MinLength
    rule.Add "Maximum", MaxLength
    AppendRule field, rule
End Sub

' Numeric comparisons use CDbl and the current Office/Windows locale.
Public Sub AddNumberRangeRule(ByVal control As MSForms.Control, _
    ByVal Minimum As Double, ByVal Maximum As Double, _
    Optional ByVal ErrorMsg As String = "Enter a number within the allowed range")

    AssertIdle
    Dim field As Object
    Dim rule As Object
    Set field = FindField(control, True)
    If Maximum < Minimum Then Fail "Maximum must be at least Minimum."
    Set rule = NewRule("NumberRange", ErrorMsg)
    rule.Add "Minimum", Minimum
    rule.Add "Maximum", Maximum
    AppendRule field, rule
End Sub

' Useful for password confirmation. Case-sensitive unless explicitly changed.
' Both fields must be scalar controls; the peer need not be registered.
Public Sub AddMatchRule(ByVal control As MSForms.Control, _
    ByVal OtherControl As MSForms.Control, _
    Optional ByVal ErrorMsg As String = "The values do not match", _
    Optional ByVal IgnoreCase As Boolean = False)

    AssertIdle
    Dim field As Object
    Dim rule As Object
    Set field = FindField(control, True)
    RequireSupportedControl OtherControl
    If IsMultiSelectList(control) Or IsMultiSelectList(OtherControl) Then
        Fail "AddMatchRule requires controls with a single value."
    End If
    Set rule = NewRule("Match", ErrorMsg)
    rule.Add "OtherControl", OtherControl
    rule.Add "IgnoreCase", IgnoreCase
    AppendRule field, rule
End Sub

' RunOnEmpty=True supports conditional requirements on optional fields.
Public Sub AddCustomRule(ByVal control As MSForms.Control, ByVal RuleName As String, _
    Optional ByVal ErrorMsg As String = "This field is invalid", _
    Optional ByVal RunOnEmpty As Boolean = False)

    AssertIdle
    Dim field As Object
    Dim rule As Object
    Set field = FindField(control, True)
    If Len(Trim$(RuleName)) = 0 Then Fail "A custom rule needs a RuleName."
    Set rule = NewRule("Custom", ErrorMsg)
    rule.Add "Name", RuleName
    rule("RunOnEmpty") = RunOnEmpty
    AppendRule field, rule
End Sub

Public Function HasField(ByVal control As MSForms.Control) As Boolean
    Dim field As Object
    Set field = FindField(control, False)
    HasField = Not (field Is Nothing)
End Function

Public Function FieldCount() As Long
    FieldCount = mFields.Count
End Function

Public Sub RemoveField(ByVal control As MSForms.Control)
    AssertIdle
    Dim field As Object
    Set field = FindField(control, False)
    If field Is Nothing Then Exit Sub
    ApplyBorderState field, True
    mFields.Remove control.Name
    mOrder.Remove control.Name
    If mErrors.Exists(control.Name) Then mErrors.Remove control.Name
End Sub

Public Sub ClearFields()
    AssertIdle
    ClearErrorsCore
    mFields.RemoveAll
    Set mOrder = New Collection
End Sub

Public Sub ClearErrors()
    AssertIdle
    ClearErrorsCore
End Sub

Public Sub ClearFieldError(ByVal control As MSForms.Control)
    AssertIdle
    Dim field As Object
    Set field = FindField(control, False)
    If field Is Nothing Then Exit Sub
    ApplyBorderState field, True
    If mErrors.Exists(control.Name) Then mErrors.Remove control.Name
End Sub

Public Function ValidateField(ByVal control As MSForms.Control) As Boolean
    AssertIdle
    Dim field As Object
    Dim message As String
    Dim valid As Boolean
    Dim failureNumber As Long
    Dim failureSource As String
    Dim failureMessage As String
    Set field = FindField(control, mStrictRegistration)
    If field Is Nothing Then
        ValidateField = True
        Exit Function
    End If

    On Error GoTo Failed
    mBusy = True
    valid = CheckField(field, message)
    If valid Then
        If mErrors.Exists(control.Name) Then mErrors.Remove control.Name
    Else
        mErrors(control.Name) = message
    End If
    ApplyBorderState field, valid
    mBusy = False
    ValidateField = valid
    Exit Function
Failed:
    failureNumber = Err.Number
    failureSource = Err.Source
    failureMessage = Err.Description
    mBusy = False
    Err.Raise failureNumber, failureSource, failureMessage
End Function

' Commit errors and borders only when every rule has finished successfully.
' A configuration/callback exception leaves the previous error state intact.
Public Function ValidateAll() As Boolean
    AssertIdle
    Dim pendingErrors As Object
    Dim field As Object
    Dim key As Variant
    Dim message As String
    Dim failureNumber As Long
    Dim failureSource As String
    Dim failureMessage As String
    Set pendingErrors = NewDictionary()

    On Error GoTo Failed
    mBusy = True
    For Each key In mOrder
        Set field = mFields(key)
        message = vbNullString
        If Not CheckField(field, message) Then pendingErrors.Add key, message
    Next key
    Set mErrors = pendingErrors
    For Each key In mOrder
        Set field = mFields(key)
        ApplyBorderState field, Not mErrors.Exists(key)
    Next key
    mBusy = False
    ValidateAll = (mErrors.Count = 0)
    Exit Function
Failed:
    failureNumber = Err.Number
    failureSource = Err.Source
    failureMessage = Err.Description
    mBusy = False
    Err.Raise failureNumber, failureSource, failureMessage
End Function

Public Function HasError(ByVal control As MSForms.Control) As Boolean
    Dim field As Object
    Set field = FindField(control, False)
    If field Is Nothing Then Exit Function
    HasError = mErrors.Exists(control.Name)
End Function

Public Function GetErrorMessage(ByVal control As MSForms.Control) As String
    If HasError(control) Then GetErrorMessage = mErrors(control.Name)
End Function

Public Function ErrorCount() As Long
    ErrorCount = mErrors.Count
End Function

' Fields are listed in AddField registration order, including after live checks.
Public Function ErrorSummary(Optional ByVal separator As String = vbNewLine, _
    Optional ByVal IncludeFieldNames As Boolean = False) As String

    Dim key As Variant
    Dim field As Object
    Dim result As String
    Dim part As String
    For Each key In mOrder
        If mErrors.Exists(key) Then
            part = mErrors(key)
            If IncludeFieldNames Then
                Set field = mFields(key)
                part = field("Label") & ": " & part
            End If
            If Len(result) > 0 Then result = result & separator
            result = result & part
        End If
    Next key
    ErrorSummary = result
End Function

Public Property Get FirstInvalidControl() As MSForms.Control
    Dim key As Variant
    Dim field As Object
    For Each key In mOrder
        If mErrors.Exists(key) Then
            Set field = mFields(key)
            Set FirstInvalidControl = field("Control")
            Exit Property
        End If
    Next key
End Property

' Call after showing the form. Tries invalid controls in registration order.
' Does not automatically switch MultiPage pages or reveal hidden containers.
Public Function FocusFirstError() As Boolean
    AssertIdle
    Dim key As Variant
    Dim field As Object
    Dim control As MSForms.Control
    For Each key In mOrder
        If mErrors.Exists(key) Then
            Set field = mFields(key)
            Set control = field("Control")
            If TryFocus(control) Then
                FocusFirstError = True
                Exit Function
            End If
        End If
    Next key
End Function

' Text controls become empty; selections are cleared; Boolean controls become
' False; SpinButton/ScrollBar return to Min. Data-bound controls also update their
' bound source. Failures are reported after attempting all registered controls;
' this is not a transaction and already-reset values are not rolled back.
Public Sub ResetForm()
    AssertIdle
    Dim key As Variant
    Dim field As Object
    Dim control As MSForms.Control
    Dim detail As String
    Dim failures As String
    Dim failureNumber As Long
    Dim failureSource As String
    Dim failureMessage As String
    On Error GoTo Failed
    mBusy = True
    For Each key In mOrder
        Set field = mFields(key)
        Set control = field("Control")
        detail = TryResetControl(control)
        If Len(detail) > 0 Then
            If Len(failures) > 0 Then failures = failures & vbNewLine
            failures = failures & control.Name & ": " & detail
        End If
    Next key
    ClearErrorsCore
    mBusy = False
    On Error GoTo 0
    If Len(failures) > 0 Then Fail "Some controls could not be reset:" & vbNewLine & failures
    Exit Sub
Failed:
    failureNumber = Err.Number
    failureSource = Err.Source
    failureMessage = Err.Description
    mBusy = False
    Err.Raise failureNumber, failureSource, failureMessage
End Sub

' --------------------------- Internal validation ---------------------------

Private Function CheckField(ByVal field As Object, ByRef message As String) As Boolean
    Dim control As MSForms.Control
    Dim values As Collection
    Dim rules As Collection
    Dim rule As Object
    Dim raw As Variant
    Dim text As String
    Set control = field("Control")
    message = vbNullString

    If mSkipDisabled Then
        If Not control.Enabled Then
            CheckField = True
            Exit Function
        End If
    End If
    If mSkipHidden Then
        If Not control.Visible Then
            CheckField = True
            Exit Function
        End If
    End If
    If field("MustBeTrue") Then
        If Not BooleanIsTrue(control) Then
            message = field("ErrorMsg")
            Exit Function
        End If
    End If

    Set values = ReadControlValues(control)
    Set rules = field("Rules")
    For Each raw In values
        text = CStr(raw)
        If field("TrimValue") Then text = TrimEdges(text)
        If Len(text) = 0 Then
            If field("Required") Then
                message = field("RequiredErrorMsg")
                Exit Function
            End If
        End If
        For Each rule In rules
            If Len(text) > 0 Or rule("RunOnEmpty") Then
                If Not CheckRule(control, field, rule, text, message) Then Exit Function
            End If
        Next rule
    Next raw
    CheckField = True
End Function

Private Function CheckRule(ByVal control As MSForms.Control, ByVal field As Object, _
    ByVal rule As Object, ByVal text As String, ByRef message As String) As Boolean

    Dim valid As Boolean
    Dim regex As Object
    Dim matches As Object
    Dim other As MSForms.Control
    Dim otherValues As Collection
    Dim peerText As String
    Dim compareMode As VbCompareMethod
    Dim handled As Boolean
    Dim customMessage As String

    Select Case CStr(rule("Kind"))
        Case "Pattern"
            Set regex = rule("Regex")
            If rule("FullMatch") Then
                Set matches = regex.Execute(text)
                If matches.Count > 0 Then
                    valid = (matches(0).FirstIndex = 0 And matches(0).Length = Len(text))
                End If
            Else
                valid = regex.Test(text)
            End If
        Case "Length"
            valid = (Len(text) >= rule("Minimum"))
            If rule("Maximum") <> -1 Then valid = valid And (Len(text) <= rule("Maximum"))
        Case "NumberRange"
            valid = NumberInRange(text, rule("Minimum"), rule("Maximum"))
        Case "Match"
            Set other = rule("OtherControl")
            Set otherValues = ReadControlValues(other)
            If otherValues.Count <> 1 Then Fail "A matching rule requires one peer value."
            peerText = CStr(otherValues(1))
            If field("TrimValue") Then peerText = TrimEdges(peerText)
            compareMode = vbBinaryCompare
            If rule("IgnoreCase") Then compareMode = vbTextCompare
            valid = (StrComp(text, peerText, compareMode) = 0)
        Case "Custom"
            customMessage = rule("ErrorMsg")
            RaiseEvent CustomValidation(control, text, CStr(rule("Name")), handled, valid, customMessage)
            If Not handled Then
                Fail "No handler accepted custom rule '" & rule("Name") & "' for '" & control.Name & "'."
            End If
            If Not valid Then
                If Len(customMessage) = 0 Then customMessage = rule("ErrorMsg")
                message = customMessage
                Exit Function
            End If
        Case Else
            Fail "Unknown internal validation rule."
    End Select
    If Not valid Then message = rule("ErrorMsg")
    CheckRule = valid
End Function

Private Function PatternRule(ByVal control As MSForms.Control, ByVal Pattern As String, _
    ByVal ErrorMsg As String, ByVal FullMatch As Boolean, ByVal IgnoreCase As Boolean, _
    ByVal MultiLine As Boolean) As Object

    Dim rule As Object
    Dim regex As Object
    Dim detail As String
    Dim unused As Boolean
    On Error GoTo EngineUnavailable
    Set regex = CreateObject("VBScript.RegExp")
    On Error GoTo InvalidPattern
    With regex
        .Global = False
        .IgnoreCase = IgnoreCase
        .MultiLine = MultiLine
        .Pattern = Pattern
        unused = .Test(vbNullString)
        If FullMatch Then
            ' Strict end assertion also rejects a final newline. Non-capturing
            ' wrapping preserves backreferences and handles a|ab correctly.
            .Pattern = "^(?:" & Pattern & ")(?![\s\S])"
        Else
            .Pattern = Pattern
        End If
        unused = .Test(vbNullString)
    End With
    On Error GoTo 0
    Set rule = NewRule("Pattern", ErrorMsg)
    rule.Add "Regex", regex
    rule.Add "FullMatch", FullMatch
    rule.Add "Pattern", Pattern
    Set PatternRule = rule
    Exit Function
EngineUnavailable:
    detail = Err.Description
    On Error GoTo 0
    Fail "RegExp is unavailable while registering '" & control.Name & "'. " & detail
    Exit Function
InvalidPattern:
    detail = Err.Description
    On Error GoTo 0
    Fail "Invalid regex for '" & control.Name & "': " & Pattern & ". " & detail
End Function

Private Function NewRule(ByVal Kind As String, ByVal ErrorMsg As String) As Object
    Dim rule As Object
    Set rule = NewDictionary()
    rule.Add "Kind", Kind
    rule.Add "ErrorMsg", NonEmptyMessage(ErrorMsg)
    rule.Add "RunOnEmpty", False
    Set NewRule = rule
End Function

Private Sub AppendRule(ByVal field As Object, ByVal rule As Object)
    Dim rules As Collection
    Dim control As MSForms.Control
    Set rules = field("Rules")
    Set control = field("Control")
    rules.Add rule
    ApplyBorderState field, True
    If mErrors.Exists(control.Name) Then mErrors.Remove control.Name
End Sub

Private Function ReadControlValues(ByVal control As MSForms.Control) As Collection
    Dim values As Collection
    Dim row As Long
    Dim column As Long
    Dim detail As String
    Set values = New Collection
    On Error GoTo Unreadable
    If IsMultiSelectList(control) Then
        column = control.BoundColumn
        For row = 0 To control.ListCount - 1
            If control.Selected(row) Then
                If column = 0 Then
                    values.Add CStr(row)
                Else
                    values.Add ValueText(control.List(row, column - 1))
                End If
            End If
        Next row
        If values.Count = 0 Then values.Add vbNullString
    Else
        values.Add ValueText(control.Value)
    End If
    Set ReadControlValues = values
    Exit Function
Unreadable:
    detail = Err.Description
    On Error GoTo 0
    Fail "Cannot read '" & control.Name & "': " & detail
End Function

Private Function ValueText(ByVal Value As Variant) As String
    If IsError(Value) Then Fail "An error value cannot be validated as text."
    If IsObject(Value) Or IsArray(Value) Then Fail "A scalar value is required."
    If IsNull(Value) Or IsEmpty(Value) Then Exit Function
    ValueText = CStr(Value)
End Function

Private Function NumberInRange(ByVal text As String, ByVal Minimum As Double, _
    ByVal Maximum As Double) As Boolean
    Dim number As Double
    On Error GoTo NotNumeric
    If Not IsNumeric(text) Then Exit Function
    number = CDbl(text)
    NumberInRange = (number >= Minimum And number <= Maximum)
NotNumeric:
End Function

Private Function BooleanIsTrue(ByVal control As MSForms.Control) As Boolean
    Dim Value As Variant
    Value = control.Value
    If IsNull(Value) Then Exit Function
    BooleanIsTrue = CBool(Value)
End Function

Private Function TrimEdges(ByVal text As String) As String
    Dim first As Long
    Dim last As Long
    first = 1
    last = Len(text)
    Do While first <= last
        If Not IsEdgeSpace(Mid$(text, first, 1)) Then Exit Do
        first = first + 1
    Loop
    Do While last >= first
        If Not IsEdgeSpace(Mid$(text, last, 1)) Then Exit Do
        last = last - 1
    Loop
    If last >= first Then TrimEdges = Mid$(text, first, last - first + 1)
End Function

Private Function IsEdgeSpace(ByVal character As String) As Boolean
    Select Case AscW(character)
        Case 9 To 13, 32, 160
            IsEdgeSpace = True
    End Select
End Function

Private Function IsMultiSelectList(ByVal control As MSForms.Control) As Boolean
    If TypeName(control) = "ListBox" Then
        IsMultiSelectList = (control.MultiSelect <> fmMultiSelectSingle)
    End If
End Function

Private Function IsBooleanControl(ByVal control As MSForms.Control) As Boolean
    Select Case TypeName(control)
        Case "CheckBox", "OptionButton", "ToggleButton"
            IsBooleanControl = True
    End Select
End Function

Private Sub RequireSupportedControl(ByVal control As MSForms.Control)
    If control Is Nothing Then Fail "The control cannot be Nothing."
    Select Case TypeName(control)
        Case "TextBox", "ComboBox", "ListBox", "CheckBox", "OptionButton", _
             "ToggleButton", "SpinButton", "ScrollBar"
            ' Supported scalar and multi-selection value sources.
        Case Else
            Fail "Unsupported control '" & control.Name & "' (" & TypeName(control) & ")."
    End Select
End Sub

Private Function FindField(ByVal control As MSForms.Control, _
    ByVal RequireRegistered As Boolean) As Object
    Dim field As Object
    Dim registered As MSForms.Control
    If control Is Nothing Then Fail "The control cannot be Nothing."
    If Not mFields.Exists(control.Name) Then
        If RequireRegistered Then Fail "Control '" & control.Name & "' is not registered."
        Exit Function
    End If
    Set field = mFields(control.Name)
    Set registered = field("Control")
    If Not (registered Is control) Then
        Fail "Control name '" & control.Name & "' belongs to another control. Use one validator per form."
    End If
    Set FindField = field
End Function

' ---------------------------- Optional UI state ----------------------------

Private Function TryReadBorder(ByVal control As MSForms.Control, _
    ByRef style As Long, ByRef color As Long) As Boolean
    On Error GoTo Unsupported
    style = control.BorderStyle
    color = control.BorderColor
    TryReadBorder = True
Unsupported:
End Function

Private Function TryReadSpecialEffect(ByVal control As MSForms.Control, _
    ByRef effect As Long) As Boolean
    On Error GoTo Unsupported
    effect = control.SpecialEffect
    TryReadSpecialEffect = True
Unsupported:
End Function

Private Sub ApplyBorderState(ByVal field As Object, ByVal IsValid As Boolean)
    Dim control As MSForms.Control
    If Not field("HasBorder") Then Exit Sub
    Set control = field("Control")
    On Error GoTo Unavailable
    If IsValid Then
        If field("Highlighted") Then
            control.BorderStyle = field("BorderStyle")
            control.BorderColor = field("BorderColor")
            If field("HasSpecialEffect") Then control.SpecialEffect = field("SpecialEffect")
            field("Highlighted") = False
        End If
    ElseIf field("ColorBorderOnError") Then
        field("Highlighted") = True
        control.BorderStyle = fmBorderStyleSingle
        control.BorderColor = mErrorBorderColor
    End If
Unavailable:
    ' Border support is cosmetic; validation errors are never suppressed here.
End Sub

Private Sub ClearErrorsCore()
    Dim key As Variant
    Dim field As Object
    For Each key In mOrder
        Set field = mFields(key)
        ApplyBorderState field, True
    Next key
    mErrors.RemoveAll
End Sub

Private Function TryFocus(ByVal control As MSForms.Control) As Boolean
    On Error GoTo CannotFocus
    If Not control.Visible Then Exit Function
    If Not control.Enabled Then Exit Function
    control.SetFocus
    TryFocus = True
CannotFocus:
End Function

Private Function TryResetControl(ByVal control As MSForms.Control) As String
    Dim row As Long
    On Error GoTo CannotReset
    Select Case TypeName(control)
        Case "TextBox"
            control.Value = vbNullString
        Case "ComboBox"
            control.ListIndex = -1
            If control.Style = fmStyleDropDownCombo Then control.Value = vbNullString
        Case "ListBox"
            If IsMultiSelectList(control) Then
                For row = 0 To control.ListCount - 1
                    control.Selected(row) = False
                Next row
            Else
                control.ListIndex = -1
            End If
        Case "CheckBox", "OptionButton", "ToggleButton"
            control.Value = False
        Case "SpinButton", "ScrollBar"
            control.Value = control.Min
        Case Else
            Fail "Unsupported reset operation."
    End Select
    Exit Function
CannotReset:
    TryResetControl = Err.Description
    If Len(TryResetControl) = 0 Then TryResetControl = "Reset failed."
End Function

Private Function NewDictionary() As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbBinaryCompare
    Set NewDictionary = result
End Function

Private Function NonEmptyMessage(ByVal message As String) As String
    If Len(message) = 0 Then message = DEFAULT_ERROR
    NonEmptyMessage = message
End Function

Private Sub AssertIdle()
    If mBusy Then
        Fail "Validation/reset is already running. Check IsBusy in control event handlers."
    End If
End Sub

Private Sub Fail(ByVal message As String)
    Err.Raise ERR_VALIDATOR, "FormValidator", message
End Sub
