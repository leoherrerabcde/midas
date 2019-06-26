Attribute VB_Name = "modGlobalConstants"
'---------------------------------------------------------------------------------------
' Module    : modGlobalConstants
' Author    : lherrera
' Date      : 19/01/2011
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit


Global Const CT_EXTENSION_DEFAULT = "*.*"
Global Const CT_NULL_INDEX = -1

Public Enum KeyActionConstant
    CloseForm = 1
End Enum

Public Enum BitModeConstant
    bmcModeBitEight = 1
    bmcModeBitSixten
    bmcModeBitTwentyFour
    bmcModeBitThirtyTwo
End Enum

Public Enum NumericBaseConstant
    nbcBinary = 1
    nbcOctal
    nbcDecimal
    nbcHexa
    nbcAscii
End Enum
