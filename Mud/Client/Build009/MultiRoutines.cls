VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MultiRoutines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub FuncTimeOut(TOInterval As Single, Index As Integer)

Dim TOStart As Single
TOStart = Timer
Do: DoEvents: Loop Until Timer - TOStart >= TOInterval

End Sub
