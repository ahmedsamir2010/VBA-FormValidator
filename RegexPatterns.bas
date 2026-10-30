Attribute VB_Name = "RegexPatterns"
Option Explicit

'====================================================================================================
' RegexPatterns
' One place to keep every regex pattern (and its default user-facing message) used across the
' project, so you never retype the same pattern string in multiple forms.
'
' USAGE:
'   validator.AddField txtEmail, RX_EMAIL, MSG_EMAIL
'   validator.AddField txtPhone, RX_PHONE_EG, MSG_PHONE_EG
'
' Add new patterns here as the project needs them -- everything stays in this one module.
' Patterns are code (English), messages are the Arabic text shown to the end user.
'====================================================================================================

Public Const RX_EMAIL           As String = "([\w\-\.]+)@([\w\-\.]+)\.[a-zA-Z]{2,5}"
Public Const RX_PHONE_EG        As String = "^01\d{9}$"                  ' Egyptian mobile, e.g. 01012345678
Public Const RX_NATIONAL_ID_EG  As String = "^\d{14}$"                   ' Egyptian national ID, 14 digits
Public Const RX_NUMERIC         As String = "^\d+$"
Public Const RX_DECIMAL         As String = "^\d+(\.\d+)?$"
Public Const RX_CURRENCY        As String = "^$|^\d+(\.\d{1,2})?$"
Public Const RX_YEAR            As String = "^(\d{4}|\d{2})$"
Public Const RX_DATE_DMY        As String = "^\d{1,2}[\s\-\.\/]\d{1,2}[\s\-\.\/]\d{2,4}$"
Public Const RX_ENGLISH_ONLY    As String = "^[a-zA-Z\s]+$"
Public Const RX_ARABIC_ONLY     As String = "^[\u0600-\u06FF\s]+$"
Public Const RX_ALPHANUMERIC    As String = "^[a-zA-Z0-9]+$"
Public Const RX_URL             As String = "^(https?:\/\/)?([\w\-]+\.)+[\w\-]{2,}(\/\S*)?$"

Public Const MSG_EMAIL          As String = "من فضلك أدخل بريد إلكتروني صحيح"
Public Const MSG_PHONE_EG       As String = "من فضلك أدخل رقم موبايل مصري صحيح (11 رقم يبدأ بـ 01)"
Public Const MSG_NATIONAL_ID_EG As String = "من فضلك أدخل رقم قومي صحيح (14 رقم)"
Public Const MSG_NUMERIC        As String = "من فضلك أدخل أرقام فقط"
Public Const MSG_DECIMAL        As String = "من فضلك أدخل قيمة رقمية صحيحة"
Public Const MSG_CURRENCY       As String = "من فضلك أدخل مبلغ صحيح"
Public Const MSG_YEAR           As String = "من فضلك أدخل سنة صحيحة"
Public Const MSG_DATE_DMY       As String = "من فضلك أدخل تاريخ صحيح (يوم/شهر/سنة)"
Public Const MSG_ENGLISH_ONLY   As String = "من فضلك أدخل حروف إنجليزية فقط"
Public Const MSG_ARABIC_ONLY    As String = "من فضلك أدخل حروف عربية فقط"
Public Const MSG_ALPHANUMERIC   As String = "من فضلك أدخل حروف وأرقام فقط"
Public Const MSG_URL            As String = "من فضلك أدخل رابط صحيح"
Public Const MSG_REQUIRED       As String = "هذا الحقل مطلوب"
