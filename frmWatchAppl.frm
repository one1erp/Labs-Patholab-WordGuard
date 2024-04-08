VERSION 5.00
Begin VB.Form frmWatchAppl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "משגיחון"
   ClientHeight    =   885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   960
   ControlBox      =   0   'False
   Icon            =   "frmWatchAppl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   885
   ScaleWidth      =   960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer tmrWatchAppl 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   240
      Top             =   240
   End
End
Attribute VB_Name = "frmWatchAppl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public loPidWord As Long          ' Word process id

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const SYNCHRONIZE = &H100000
Private Const PROCESS_TERMINATE As Long = &H1

Private mdtStartDocTime As Date     ' Start time of document process
Private mboProcessDoc As Boolean    ' Flag - document is being currently processing

Private Sub Form_Load()
    
    mboProcessDoc = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CloseWordAppl
End Sub

Private Sub tmrWatchAppl_Timer()
    
    If mboProcessDoc And (Abs(DateDiff("n", Now, mdtStartDocTime)) > 3) Then
            Unload Me
    End If

End Sub

Private Function CloseWordAppl()
        
    Dim llohPidWord As Long       ' Word process handle
    
    On Error Resume Next        ' if word already closed
    
    llohPidWord = OpenProcess(SYNCHRONIZE Or PROCESS_TERMINATE, ByVal 0&, loPidWord)


    Call TerminateProcess(llohPidWord, 0&)
    Call CloseHandle(llohPidWord)

End Function

Public Sub StartWatch()
    
    mdtStartDocTime = Now()
    mboProcessDoc = True
    tmrWatchAppl.Enabled = True

End Sub

Public Sub EndWatch()
    
    mboProcessDoc = False
    tmrWatchAppl.Enabled = False

End Sub

