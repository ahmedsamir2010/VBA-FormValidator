Attribute VB_Name = "FormValidatorTests"
Option Explicit

' Import FormValidator.cls and FormValidatorTestEvents.cls, then this module.
' Pass an empty UserForm from the Immediate Window, for example:
'   RunFormValidatorTests UserForm1
' Tests create/remove only controls whose names begin with FVTest.
' The form is not shown. No workbook cells or VBProject settings are changed.

Private mAssertions As Long

Public Sub RunFormValidatorTests(ByVal host As Object)
    Dim v As FormValidator
    Dim sink As FormValidatorTestEvents
    Dim txt As MSForms.TextBox
    Dim peer As MSForms.TextBox
    Dim check As MSForms.CheckBox
    Dim list As MSForms.ListBox
    Dim combo As MSForms.ComboBox
    Dim spin As MSForms.SpinButton
    Dim scroll As MSForms.ScrollBar
    Dim label As MSForms.Label
    Dim first As MSForms.Control
    Dim created As Collection
    Dim name As Variant
    Dim oldColor As Long
    Dim oldStyle As Long
    Dim oldEffect As Long
    Dim number As Long
    Dim detail As String
    Dim ignored As Boolean
    Dim failureNumber As Long
    Dim failureMessage As String
    On Error GoTo Failed
    mAssertions = 0
    Set created = New Collection
    Set v = New FormValidator

    Set txt = host.Controls.Add("Forms.TextBox.1", "FVTestText")
    created.Add txt.Name
    Set peer = host.Controls.Add("Forms.TextBox.1", "FVTestPeer")
    created.Add peer.Name
    Set check = host.Controls.Add("Forms.CheckBox.1", "FVTestCheck")
    created.Add check.Name
    Set list = host.Controls.Add("Forms.ListBox.1", "FVTestList")
    created.Add list.Name
    Set combo = host.Controls.Add("Forms.ComboBox.1", "FVTestCombo")
    created.Add combo.Name
    Set spin = host.Controls.Add("Forms.SpinButton.1", "FVTestSpin")
    created.Add spin.Name
    Set scroll = host.Controls.Add("Forms.ScrollBar.1", "FVTestScroll")
    created.Add scroll.Name
    Set label = host.Controls.Add("Forms.Label.1", "FVTestLabel")
    created.Add label.Name

    AssertTrue v.ValidateAll, "An empty validator passes"
    AssertTrue v.Version = "2.0.0", "Version"

    ' Full string matching, alternatives, anchors, newlines and case settings.
    v.AddField txt, "\d{3}", "Three digits", TrimValue:=False
    txt.Value = "xx123yy"
    AssertTrue Not v.ValidateField(txt), "Partial match rejected by default"
    txt.Value = "123"
    AssertTrue v.ValidateField(txt), "Full match accepted"
    v.AddField txt, "\d{3}", FullMatch:=False
    txt.Value = "xx123yy"
    AssertTrue v.ValidateField(txt), "Explicit partial matching"
    v.AddField txt, "^\d{3}$", TrimValue:=False, MultiLine:=True
    txt.Value = "123" & vbLf & "bad"
    AssertTrue Not v.ValidateField(txt), "A good line cannot hide a bad line"
    txt.Value = "123" & vbLf
    AssertTrue Not v.ValidateField(txt), "Strict full match rejects final newline"
    v.AddField txt, "a|ab"
    txt.Value = "ab"
    AssertTrue v.ValidateField(txt), "Full match handles alternative backtracking"
    v.AddField txt, "([a-z])\1", IgnoreCase:=False
    txt.Value = "aa"
    AssertTrue v.ValidateField(txt), "Wrapping preserves capture numbers"
    v.AddField txt, "[A-Z]{3}", IgnoreCase:=False
    txt.Value = "abc"
    AssertTrue Not v.ValidateField(txt), "Case-sensitive rule"
    v.AddField txt, "[A-Z]{3}", IgnoreCase:=True
    AssertTrue v.ValidateField(txt), "Explicit case-insensitive rule"

    ' Required/optional values, whitespace, error counts and cosmetic opt-out.
    v.AddField txt, "", "Required", RequiredErrorMsg:="Enter a value"
    txt.Value = vbTab & ChrW(160) & vbCrLf & " "
    AssertTrue Not v.ValidateField(txt), "Whitespace-only required value rejected"
    AssertTrue v.GetErrorMessage(txt) = "Enter a value", "Required-specific message"
    AssertTrue v.ErrorCount = 1, "ErrorCount counts failed fields"
    v.AddField txt, "\d+", Required:=False
    AssertTrue v.ValidateField(txt), "Empty optional field skips ordinary rules"
    txt.Value = "abc"
    AssertTrue Not v.ValidateField(txt), "Non-empty optional field still checked"
    v.ClearErrors
    AssertTrue txt.Value = "abc", "ClearErrors keeps data"
    AssertTrue v.ErrorCount = 0, "ClearErrors clears state"
    v.ClearFields
    oldColor = RGB(10, 40, 90)
    oldStyle = fmBorderStyleNone
    txt.BorderColor = oldColor
    txt.BorderStyle = oldStyle
    txt.SpecialEffect = fmSpecialEffectSunken
    oldStyle = txt.BorderStyle
    oldEffect = txt.SpecialEffect
    v.AddField txt, "\d+", "Digits", ColorBorderOnError:=False
    AssertTrue Not v.ValidateField(txt), "Opted-out field still validated"
    AssertTrue txt.BorderColor = oldColor And txt.BorderStyle = oldStyle, "Coloring opt-out"
    v.ClearFields

    ' Rule changes and deletion preserve the border from the first registration.
    v.AddField txt, "\d+", "Digits"
    AssertTrue Not v.ValidateField(txt), "Prepare highlighted field"
    AssertTrue txt.BorderColor = v.ErrorBorderColor, "Invalid border colored"
    v.ErrorBorderColor = RGB(99, 22, 11)
    AssertTrue txt.BorderColor = RGB(99, 22, 11), "Color setter updates highlighted fields"
    On Error Resume Next
    v.AddField txt, "[", "Bad replacement"
    number = Err.Number
    detail = Err.Description
    Err.Clear
    On Error GoTo Failed
    AssertTrue number <> 0 And InStr(1, detail, txt.Name) > 0, "Bad regex names its control"
    AssertTrue v.HasError(txt), "Bad replacement preserves previous errors"
    AssertTrue txt.BorderColor = v.ErrorBorderColor, "Bad replacement preserves border"
    v.AddField txt, "[a-z]+", "Letters"
    AssertTrue Not v.HasError(txt), "Successful replacement clears stale error"
    AssertTrue txt.BorderColor = oldColor And txt.BorderStyle = oldStyle, "Replacement restores original"
    AssertTrue txt.SpecialEffect = oldEffect, "Replacement restores original SpecialEffect"
    AssertTrue v.ValidateField(txt), "Replacement rule used"
    txt.Value = "123"
    AssertTrue Not v.ValidateField(txt), "Replacement can fail"
    v.RemoveField txt
    AssertTrue txt.BorderColor = oldColor And txt.BorderStyle = oldStyle, "Removal restores original"
    AssertTrue Not v.HasField(txt), "Removal unregisters field"
    v.AddField txt, "[a-z]+"
    AssertTrue Not v.ValidateField(txt), "Prepare ClearFields"
    v.ClearFields
    AssertTrue txt.BorderColor = oldColor And txt.BorderStyle = oldStyle, "ClearFields restores original"
    AssertTrue v.FieldCount = 0 And v.ErrorCount = 0, "ClearFields clears state"

    ' Stable error order, explicit clearing and strict registration.
    txt.Value = ""
    peer.Value = ""
    v.AddField txt, "", "First", FieldLabel:="Name"
    v.AddField peer, "", "Second", FieldLabel:="Email"
    ignored = v.ValidateField(peer)
    ignored = v.ValidateField(txt)
    AssertTrue v.ErrorSummary("|") = "First|Second", "Summary follows registration order"
    AssertTrue v.ErrorSummary("|", True) = "Name: First|Email: Second", "Labels in summary"
    Set first = v.FirstInvalidControl
    AssertTrue first Is txt, "First invalid control"
    v.ClearFieldError txt
    AssertTrue Not v.HasError(txt) And v.HasError(peer), "Clear just one field's error"
    v.RemoveField txt
    AssertTrue v.ValidateField(txt), "Legacy unregistered pass"
    v.StrictRegistration = True
    On Error Resume Next
    ignored = v.ValidateField(txt)
    number = Err.Number
    Err.Clear
    On Error GoTo Failed
    AssertTrue number <> 0, "Strict mode rejects unregistered control"
    v.StrictRegistration = False
    On Error Resume Next
    v.AddField label, ""
    number = Err.Number
    Err.Clear
    On Error GoTo Failed
    AssertTrue number <> 0, "Unsupported controls rejected during registration"
    v.ClearFields

    ' Multiple rules and comparisons.
    v.AddField txt, "[a-z]+", "Letters"
    v.AddLengthRule txt, 3, 5, "Length"
    v.AddPatternRule txt, ".*x.*", "Include x"
    txt.Value = "ab"
    AssertTrue Not v.ValidateField(txt), "Length rule applied"
    AssertTrue v.GetErrorMessage(txt) = "Length", "First failing rule's own message"
    txt.Value = "abcd"
    AssertTrue Not v.ValidateField(txt), "Additional regex applied"
    AssertTrue v.GetErrorMessage(txt) = "Include x", "Extra regex message"
    txt.Value = "abxd"
    AssertTrue v.ValidateField(txt), "All rules pass"
    v.AddField txt, ""
    v.AddNumberRangeRule txt, 1, 10, "Range"
    txt.Value = "11"
    AssertTrue Not v.ValidateField(txt), "Number upper bound"
    txt.Value = "not a number"
    AssertTrue Not v.ValidateField(txt), "Invalid number fails without runtime error"
    txt.Value = "1E999"
    AssertTrue Not v.ValidateField(txt), "Numeric overflow fails without runtime error"
    txt.Value = "5"
    AssertTrue v.ValidateField(txt), "Number inside range"
    v.AddField txt, "", TrimValue:=False
    v.AddMatchRule txt, peer
    txt.Value = "Secret "
    peer.Value = "Secret"
    AssertTrue Not v.ValidateField(txt), "Match respects untrimmed password"
    txt.Value = "secret"
    AssertTrue Not v.ValidateField(txt), "Match is case-sensitive by default"
    txt.Value = "Secret"
    AssertTrue v.ValidateField(txt), "Matching values pass"
    v.ClearFields

    ' Boolean requirements and multi-selection bound values.
    check.Value = False
    v.AddField check, "", "Accept", MustBeTrue:=True
    AssertTrue Not v.ValidateField(check), "Unchecked mandatory agreement fails"
    check.Value = True
    AssertTrue v.ValidateField(check), "Checked agreement passes"
    v.AddField check, ""
    check.Value = False
    AssertTrue v.ValidateField(check), "Required alone permits explicit False"
    list.MultiSelect = fmMultiSelectMulti
    list.ColumnCount = 2
    list.BoundColumn = 2
    list.AddItem "First label"
    list.List(0, 1) = "100"
    list.AddItem "Second label"
    list.List(1, 1) = "bad"
    v.AddField list, "\d+", "Choose valid numbers"
    AssertTrue Not v.ValidateField(list), "No multi-selection fails required"
    list.Selected(0) = True
    AssertTrue v.ValidateField(list), "Selected bound column validated"
    list.Selected(1) = True
    AssertTrue Not v.ValidateField(list), "Every selected value must pass"
    list.Selected(1) = False
    list.BoundColumn = 0
    AssertTrue v.ValidateField(list), "BoundColumn zero validates selected row index"

    ' Typed reset: no blanket Value="" and no discarded reset failure.
    combo.Style = fmStyleDropDownList
    combo.AddItem "Cairo"
    combo.ListIndex = 0
    spin.Min = 5
    spin.Max = 15
    spin.Value = 10
    scroll.Min = 3
    scroll.Max = 8
    scroll.Value = 6
    v.AddField combo, ""
    v.AddField spin, ""
    v.AddField scroll, ""
    v.AddField txt, ""
    txt.Value = "Keep until reset"
    v.ResetForm
    AssertTrue txt.Value = "", "Reset clears text"
    AssertTrue check.Value = False, "Reset clears Boolean"
    AssertTrue Not list.Selected(0) And Not list.Selected(1), "Reset clears all selections"
    AssertTrue combo.ListIndex = -1, "Reset clears list-only ComboBox"
    AssertTrue spin.Value = spin.Min And scroll.Value = scroll.Min, "Reset uses numeric minima"
    AssertTrue v.ErrorCount = 0 And Not v.IsBusy, "Reset clears errors and busy flag"
    AssertTrue v.FieldCount = 6, "Reset keeps registration"
    v.ClearFields

    ' Skip options are explicit and apply to the control's own flags.
    v.AddField txt, ""
    txt.Value = ""
    txt.Enabled = False
    AssertTrue Not v.ValidateField(txt), "Disabled controls checked by default"
    v.SkipDisabled = True
    AssertTrue v.ValidateField(txt), "Optional disabled skip"
    txt.Enabled = True
    v.SkipDisabled = False
    txt.Visible = False
    AssertTrue Not v.ValidateField(txt), "Hidden controls checked by default"
    v.SkipHidden = True
    AssertTrue v.ValidateField(txt), "Optional hidden skip"
    txt.Visible = True
    v.SkipHidden = False
    v.ClearFields

    ' Custom callbacks, empty values, exception atomicity and recursion guard.
    Set sink = New FormValidatorTestEvents
    Set sink.Validator = v
    v.AddField txt, "", Required:=False
    v.AddCustomRule txt, "ConditionalRequired", RunOnEmpty:=True
    sink.RequireValue = True
    AssertTrue Not v.ValidateField(txt), "Custom conditional requirement checks blank"
    AssertTrue sink.SawBusy, "Busy flag visible inside event"
    sink.RequireValue = False
    AssertTrue v.ValidateField(txt), "Conditional optional value accepted"
    v.AddField txt, ""
    v.AddCustomRule txt, "ContainsExcel"
    txt.Value = "Word"
    AssertTrue Not v.ValidateField(txt), "Custom rule failure"
    AssertTrue v.GetErrorMessage(txt) = "Include Excel", "Callback message returned"
    txt.Value = "Excel training"
    AssertTrue v.ValidateField(txt), "Custom rule success"
    v.AddField peer, "", "Peer is required"
    peer.Value = ""
    AssertTrue Not v.ValidateAll, "Prepare an existing error snapshot"
    peer.Value = "Now valid"
    sink.RaiseOnCustom = True
    On Error Resume Next
    ignored = v.ValidateAll
    number = Err.Number
    Err.Clear
    On Error GoTo Failed
    AssertTrue number <> 0, "Callback exceptions propagate"
    AssertTrue v.HasError(peer) And v.ErrorCount = 1, "Exception preserves old error snapshot"
    AssertTrue Not v.IsBusy, "Exception releases busy flag"
    sink.RaiseOnCustom = False
    sink.ReenterOnCustom = True
    On Error Resume Next
    ignored = v.ValidateAll
    number = Err.Number
    Err.Clear
    On Error GoTo Failed
    AssertTrue number <> 0 And Not v.IsBusy, "Recursive validation is rejected and cleaned up"
    sink.ReenterOnCustom = False
    Set sink.Validator = Nothing
    v.AddField txt, ""
    v.AddCustomRule txt, "Unhandled"
    On Error Resume Next
    ignored = v.ValidateField(txt)
    number = Err.Number
    Err.Clear
    On Error GoTo Failed
    AssertTrue number <> 0, "Unhandled custom rule cannot silently pass"

    v.ClearFields
    Set sink = Nothing
    Set v = Nothing
    For Each name In created
        host.Controls.Remove CStr(name)
    Next name
    Debug.Print "FormValidator: " & mAssertions & " assertions passed."
    MsgBox "FormValidator: " & mAssertions & " assertions passed.", vbInformation
    Exit Sub
Failed:
    failureNumber = Err.Number
    failureMessage = Err.Description
    On Error Resume Next
    If Not sink Is Nothing Then Set sink.Validator = Nothing
    If Not v Is Nothing Then v.ClearFields
    If Not created Is Nothing Then
        For Each name In created
            host.Controls.Remove CStr(name)
        Next name
    End If
    Set sink = Nothing
    Set v = Nothing
    On Error GoTo 0
    Err.Raise failureNumber, "FormValidatorTests", failureMessage
End Sub

Private Sub AssertTrue(ByVal condition As Boolean, ByVal description As String)
    If Not condition Then Err.Raise vbObjectError + 2200, "FormValidatorTests", description
    mAssertions = mAssertions + 1
End Sub
