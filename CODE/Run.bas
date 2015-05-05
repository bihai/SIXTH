Attribute VB_Name = "Run"
Option Explicit
'=======================================================================================
'SIXTH: a VB6 Forth-like compiler & runtime; Copyright (C) Kroc Camen, 2015
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'=======================================================================================
'MODULE :: Run

Private UI As uiSIXTH

'MAIN
'=======================================================================================
Sub Main()
    'Create our interface to the virtual machine that will handle the output
    Set UI = New uiSIXTH
    
    
    'Shutdown the VM & UI
    Set UI = Nothing
End Sub
