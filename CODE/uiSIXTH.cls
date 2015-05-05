VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "uiSIXTH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'=======================================================================================
'SIXTH: a VB6 Forth-like compiler & runtime; Copyright (C) Kroc Camen, 2015
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'=======================================================================================
'CLASS :: uiSIXTH

'We have to use a class to be able to sink the events from the SIXTH virtual machine, _
 therefore this class exists to act as an intermediary between the VM and console. _
 You won't need this class if you're embedding SIXTH in your own application as you _
 can just talk to the VM directly

'/// CLASS /////////////////////////////////////////////////////////////////////////////

Private WithEvents VM As vmSIXTH
Attribute VM.VB_VarHelpID = -1

'/// EVENTS ////////////////////////////////////////////////////////////////////////////

'CLASS Initialize
'=======================================================================================
Private Sub Class_Initialize()
    'Connect to the console (or display one, if running from IDE)
    Call Con.Initialize
    
    'Disclaimer
    Let Con.Title = "SIXTH"
    
    Set VM = New vmSIXTH
    
    Let Con.ForeColor = conWhiteHi
    Call Con.WriteLine("  SIXTH: A VB6 Forth-like compiler & runtime; Version " & App.Major & "." & App.Minor & "," & App.Revision)
    Call Con.WriteLine("  Copyright (C) Kroc Camen, 2015")
    Call Con.WriteLine
    
GetInput:
    Call Con.WriteLine("> ", False)
    Let Con.ForeColor = conWhite
    
    Dim Msg As String
    Let Msg = Con.ReadLine
    Call Con.WriteLine
    
    Dim Err As SIXTH_ERROR
    Let Err = VM.Parse(Msg)
    
    Call Con.WriteLine
    Let Con.ForeColor = conWhite
    
    Select Case Err
        Case SIXTH_ERROR.ERR_NONE
            Let Con.ForeColor = conWhiteHi
            Call Con.WriteLine("  OK")
        
        Case SIXTH_ERROR.ERR_STACKOVERFLOW
            Let Con.ForeColor = conRedHi
            Call Con.WriteLine("  STACK OVERFLOW!")
            
        Case SIXTH_ERROR.ERR_STACKUNDERFLOW
            Let Con.ForeColor = conRedHi
            Call Con.WriteLine("  STACK UNDERFLOW!")
        
        Case SIXTH_ERROR.ERR_OVERFLOW
            Let Con.ForeColor = conRedHi
            Call Con.WriteLine("  MATH OVERFLOW!")
            
        Case SIXTH_ERROR.ERR_DIVIDEBYZERO
            Let Con.ForeColor = conRedHi
            Call Con.WriteLine("  DIVIDE BY ZERO!")
            
    End Select
    
    Let Con.ExitCode = Err
    Set VM = Nothing
End Sub

'CLASS Terminate
'=======================================================================================
Private Sub Class_Terminate()
    'Shutdown the virtual machine
    Set VM = Nothing
End Sub

'EVENT <VM> Log : A message from the SIXTH virtual machine
'=======================================================================================
Private Sub VM_Log( _
 Message As String _
)
    Call Con.WriteLine(Message)
End Sub
