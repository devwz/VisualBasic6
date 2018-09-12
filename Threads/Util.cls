VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Util"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Event Done()

Public strLog As String
Dim i As Long

Public Function Log(ByVal v As Long) As Boolean

    Sleep 1000
    
    Debug.Print i
    
    i = i + 1
    
    If i > v Then
        i = 0
 
        Log = True
        RaiseEvent Done
        
    End If

End Function
