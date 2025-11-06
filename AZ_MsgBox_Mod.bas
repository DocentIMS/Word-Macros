Attribute VB_Name = "AZ_MsgBox_Mod"
Option Explicit
Public Enum NewMsgBoxStyle
    None = 0 'DEFAULT
    Success = 2 ^ 0
    Exclamation = 2 ^ 1
    Critical = 2 ^ 2
    Information = 2 ^ 3
    Question = 2 ^ 4
    ZoomIcon = 2 ^ 5
End Enum
