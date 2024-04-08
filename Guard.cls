VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Guard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Private mlohWndWord As Long         ' Word window handler

Private Sub Class_Initialize()
    
    Load frmWatchAppl

End Sub

Public Sub AssignAppl(pstApplClass, pstApplCaption As String)

    Dim lloPidWord As Long          ' Word process id
     
    ' Find the application process and window
    
    mlohWndWord = FindWindow(pstApplClass, pstApplCaption)
    lloPidWord = 0
    
    Call GetWindowThreadProcessId(mlohWndWord, lloPidWord)
    
    frmWatchAppl.loPidWord = lloPidWord

End Sub

Private Sub Class_Terminate()
    
    Unload frmWatchAppl

End Sub

Public Sub StartProcessDoc()
    
    frmWatchAppl.StartWatch

End Sub

Public Sub EndProcessDoc()
    
   frmWatchAppl.EndWatch
    
End Sub

