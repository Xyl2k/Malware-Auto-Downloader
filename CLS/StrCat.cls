VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStrCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private ms_BigStr As String
Private ml_Pos As Long
Private ml_MaxLength As Long

Public Property Get StrVal() As String
Attribute StrVal.VB_Description = "Return Current String Content"
Attribute StrVal.VB_UserMemId = 0
       StrVal = Left$(ms_BigStr, ml_Pos)
End Property




Public Property Let MaxLength(ByVal lNewValue As Long)

    ms_BigStr = Space$(lNewValue)
    ml_Pos = 0

End Property

Public Function AddStr(NewStr As String)

Mid$(ms_BigStr, ml_Pos + 1) = NewStr
ml_Pos = ml_Pos + Len(NewStr)


End Function

