VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vmSIXTH"
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
'CLASS :: vmSIXTH

'This is a Forth-like system implemented in VB6. You should read about Forth here: _
 <en.wikipedia.org/wiki/Forth_(programming_language)>

'Traditional Forth systems parse the code every time the program is run. Therefore _
 even comments exist as Words (functions) within the program. Since we're using VB6 _
 and we're not knocking against the real hardware we need to make things as fast as _
 possible. Therefore SIXTH converts source code text into numerical tokens, ignoring _
 comments entirely, and executes the token stream
 
'Secondly a Forth program typically re-defines itself each time it is run, that is, _
 the Words (functions) are created and assembled as the code runs. In SIXTH, we're _
 going for a compiled style of execution, rather than interpretive. When Words are _
 defined, the definition is kept and executed rather than being wiped and redefined _
 each run. For speed, we don't want to re-assemble our program every time it is run

'/// API ///////////////////////////////////////////////////////////////////////////////
'/// Tap into Window's functions

'In VB6 True is -1 and False is 0, but in the Win32 API it's 1 for True
Private Enum BOOL
    API_TRUE = 1
    API_FALSE = 0
End Enum

'Windows Version:
'---------------------------------------------------------------------------------------
'Some Unicode APIs vary between XP & Vista+

'Structure for obtaining the Windows version _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms724834%28v=vs.85%29.aspx>
Private Type OSVERSIONINFO
    SizeOfMe As Long
    MajorVersion As Long
    MinorVersion As Long
    BuildNumber As Long
    PlatformID As Long
    ServicePack As String * 128
End Type

'Get the Windows version _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms724451%28v=vs.85%29.aspx>
Private Declare Function api_GetVersionEx Lib "kernel32" Alias "GetVersionExA" ( _
    ByRef VersionInfo As OSVERSIONINFO _
) As BOOL

'---------------------------------------------------------------------------------------

'Get the Locale Identifier (LCID) of this app _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd318127(v=vs.85).aspx>
'This is used for Windows XP support as Vista+ use Locale Name strings
Private Declare Function api_GetThreadLocale Lib "kernel32" Alias "GetThreadLocale" ( _
) As Long

'Unicode & Locale-aware case conversion (Windows XP) _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd318700(v=vs.85).aspx>
Private Declare Function api_LCMapString Lib "kernel32" Alias "LCMapStringW" ( _
    ByVal LocaleID As Long, _
    ByVal MapFlags As LCMAP, _
    ByVal SourceStringPointer As Long, _
    ByVal SourceStringLength As Long, _
    ByVal DestinationStringPointer As Long, _
    ByVal DestinationStringLength As Long _
) As Long

'Unicode & Locale-aware case conversion (Windows Vista+) _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd318702(v=vs.85).aspx>
Private Declare Function api_LCMapStringEx Lib "kernel32" Alias "LCMapStringEx" ( _
    ByVal LocaleNamePointer As Long, _
    ByVal MapFlags As LCMAP, _
    ByVal SourceStringPointer As Long, _
    ByVal SourceStringLength As Long, _
    ByVal DestinationStringPointer As Long, _
    ByVal DestinationStringLength As Long, _
    ByVal VersionInfoPointer As Long, _
    ByVal ReservedPointer As Long, _
    ByVal SortHandle As Long _
) As Long

Private Enum LCMAP
    LCMAP_LINGUISTIC_CASING = &H1000000 'Manage accents &c. in lower/upper casing
    LCMAP_LOWERCASE = &H100&            'Convert to lower case
    LCMAP_UPPERCASE = &H200&            'Convert to UPPER CASE
    LCMAP_SORTKEY = &H400&              'Generate a sort key
    LCMAP_BYTEREV = &H800&              'Reverse the bytes
    LCMAP_HIRAGANA = &H100000           'Change Japanese Katakana to Hiragana
    LCMAP_KATAKANA = &H200000           'Change Japanese Hiragana to Katakana
    LCMAP_HALFWIDTH = &H400000          'Use standard (half-width) characters
    LCMAP_FULLWIDTH = &H800000          'Use wide-characters (i.e. Latin in Chinese)
    'Change traditional Chinese to simplified Chinese and vice-versa
    LCMAP_SIMPLIFIED_CHINESE = &H2000000
    LCMAP_TRADITIONAL_CHINESE = &H4000000
End Enum

'Comparing Strings:
'--------------------------------------------------------------------------------------

'Binary string comparison (case-sensitive) _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd317762(v=vs.85).aspx>
Private Declare Function api_CompareStringOrdinal Lib "kernel32" Alias "CompareStringOrdinalW" ( _
    ByVal String1Pointer As Long, _
    ByVal String1Length As Long, _
    ByVal String2Pointer As Long, _
    ByVal String2Length As Long, _
    ByVal IgnoreCase As BOOL _
) As CSTR_

'Unicode & Locale-aware string comparison (Windows XP) _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd317759(v=vs.85).aspx>
Private Declare Function api_CompareString Lib "kernel32" Alias "CompareStringW" ( _
    ByVal LocaleID As Long, _
    ByVal CompareFlags As CompareString_Flags, _
    ByVal String1Pointer As Long, _
    ByVal String1Length As Long, _
    ByVal String2Pointer As Long, _
    ByVal String2Length As Long _
) As CSTR_

'Unicode & Locale-aware string comparison (Windows Vista+) _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd317761(v=vs.85).aspx>
Private Declare Function api_CompareStringEx Lib "kernel32" Alias "CompareStringEx" ( _
    ByVal LocaleNamePointer As Long, _
    ByVal CompareFlags As CompareString_Flags, _
    ByVal String1Pointer As Long, _
    ByVal String1Length As Long, _
    ByVal String2Pointer As Long, _
    ByVal String2Length As Long, _
    ByVal VersionInfoPointer As Long, _
    ByVal ReservedPointer As Long, _
    ByVal Param As Long _
) As CSTR_

Private Enum CompareString_Flags
    'Ignore case, as linguistically appropriate
    LINGUISTIC_IGNORECASE = &H10&
    'Ignore nonspacing characters, as linguistically appropriate
    LINGUISTIC_IGNOREDIACRITIC = &H20&
    'Legacy ignore case -- less accurate than LINGUISTIC_IGNORECASE
    NORM_IGNORECASE = &H1&
    'Treat equivilent Japanese Hiragana and Katakana as the same
    NORM_IGNOREKANATYPE = &H10000
    'Legacy version of `LINGUISTIC_IGNOREDIACRITIC`
    NORM_IGNORENONSPACE = &H2&
    'Ignore symbols and punctuation
    NORM_IGNORESYMBOLS = &H4&
    'Ignore the difference between half-width and full-width characters
    NORM_IGNOREWIDTH = &H20000
    'Use linguistic rules for casing, instead of file system rules (default)
    NORM_LINGUISTIC_CASING = &H8000000
    'Treat punctuation the same as symbols
    SORT_STRINGSORT = &H1000
End Enum

'Comparison API return values; _
 You can subtract 2 from these to get C-style values
Private Enum CSTR_
    CSTR_LESS_THAN = 1
    CSTR_EQUAL = 2
    CSTR_GREATER_THAN = 3
End Enum

'Unicode character properties: _
 ---------------------------------------------------------------------------------------
'Returns an array of Unicode character properties for a string _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd318119(v=vs.85).aspx>
Private Declare Function api_GetStringTypeW Lib "kernel32" Alias "GetStringTypeW" ( _
    ByVal InfoType As CT, _
    ByVal SourceStringPointer As Long, _
    ByVal SourceStringLength As Long, _
    ByVal CharTypePointer As Long _
) As BOOL

Private Enum CT
    CT_CTYPE1 = &H1                     'Character type info (unicode properties)
    CT_CTYPE2 = &H2                     'Bi-directional layout information
    CT_CTYPE3 = &H3                     'Text-processing information
End Enum

Private Enum C1
    C1_UPPER = 2 ^ 0                    'Uppercase
    C1_LOWER = 2 ^ 1                    'Lowercase
    C1_DIGIT = 2 ^ 2                    'Decimal digit
    C1_SPACE = 2 ^ 3                    'Space characters
    C1_PUNCT = 2 ^ 4                    'Punctuation
    C1_CNTRL = 2 ^ 5                    'Control characters
    C1_BLANK = 2 ^ 6                    'Blank characters
    C1_XDIGIT = 2 ^ 7                   'Hexadecimal digits
    C1_ALPHA = 2 ^ 8                    'Any linguistic character
    C1_DEFINED = 2 ^ 9                  'Defined, but not one of the other C1_* types
    
    'Shorthand for "alpha-numeric"
    C1_ALPHANUM = C1_ALPHA Or C1_DIGIT
    'All kinds of blank characters you would want to strip off the ends
    C1_WHITESPACE = C1_SPACE Or C1_BLANK Or C1_CNTRL
    'Visible ("Printable") characters, this includes spaces, tabs &c.
    C1_VISIBLE = C1_SPACE Or C1_PUNCT Or C1_BLANK Or C1_ALPHANUM
End Enum

'Cast String to Array: _
 ---------------------------------------------------------------------------------------
'By manipulating some pointers we can cast a VB string directly into an array without _
 having to make a copy of the string, nor waste memory on a separate array. _
 This is obviously very fast, especially for long strings
'<vbforums.com/showthread.php?729385>

'This is VB6's internal structure used for VB arrays _
 <msdn.microsoft.com/en-us/library/ms221482(v=vs.85).aspx>
Private Type SAFEARRAY1D
    cDims      As Integer               'The count of dimensions
    fFeatures  As Integer               'Flags used by the SafeArray
    cbElements As Long                  'The size of an array element
    cLocks     As Long                  'Lock count
    pvData     As Long                  'Pointer to the data
    cElements  As Long                  'The number of elements in the dimension
    lLbound    As Long                  'The lower bound of the dimension
End Type

'- An array that is allocated on the stack
Private Const FADF_AUTO      As Integer = &H1
'- An array that may not be resized or reallocated
Private Const FADF_FIXEDSIZE As Integer = &H10

'Undocumented VB6 API to get at the array pointer
Private Declare Function msvbvm60_VarPtr Lib "msvbvm60.dll" Alias "VarPtr" ( _
    ByRef ArrayVar() As Any _
) As Long

'Undocumented VB6 API to write 4-bytes (Long) of memory
Private Declare Sub msvbvm60_PutMem4 Lib "msvbvm60.dll" Alias "PutMem4" ( _
    ByVal Addr As Long, _
    ByVal NewVal As Long _
)

Private CastStrSA1D As SAFEARRAY1D      'A dummy VB6 array header
Private CastStrPtr  As Long             'Pointer to our dummy array
Private CastStrArr() As Integer         'Empty array that will be hacked

'/// CLASS /////////////////////////////////////////////////////////////////////////////

'Public error numbers:
Public Enum SIXTH_ERROR
    ERR_NONE
    
    ERR_STACKUNDERFLOW                  'Tried to pop when the stack is empty
    ERR_STACKOVERFLOW                   'Stack is full!
    ERR_OVERFLOW                        'Mathematical overflow
    ERR_DIVIDEBYZERO                    'Divide by Zero!
End Enum

'Some APIs are different for XP/Vista. We check the Windows version and, for speed _
 reasons, this variable is set to 0 for Vista and above or 1 otherwise
Private WindowsVersionCheck As Long

'The Stack:
'---------------------------------------------------------------------------------------

'SIXTH, like Forth, is a stack-machine. This means that it performs calculations _
 using a Last-In-First-Out ("LIFO") list where values are "pushed" onto the top of _
 the stack (pushing the previous values down the list) and "popped" (or "pulled") _
 off the top, lifting the value below to the top

'For simplicity, the stack always consists of 64-bit Double values beacuse these _
 can hold 8, 16 & 32-bit signed and unsigned numbers and also 32-bit floating-points

'The stack is 1-based so that we can easily determine an empty stack (0) _
 rather than managing a 0th element that may or may not be unpopulated

Private Stack(1 To 256) As Double       'Stack storage
Private SP As Long                      'Stack Pointer

'Word Definitions:
'---------------------------------------------------------------------------------------
'A 'Word' is a function, named so because any text that is not white-space is taken as _
 a Word. E.g. "cat sat on mat" would be a program that ran the "cat", "sat", "on" and _
 "mat" functions in succession

'User defined Words are created using ":" to begin the definition, followed by a Word _
 to define the name of the function and then the code to execute, ending with ";"

'Numbers are also Words! They are special Words that automatically push the number's _
 value on the stack, e.g. `2 2 + .` adds 2 and 2

Private Type SIXTH_Word
    Text As String                      'What the Word looks like
    Kind As SIXTH_KIND                  'What kind of Word is it?
    Param As Double                     'The associated value (e.g. number / ID)
End Type

Private Enum SIXTH_KIND
    KIND_UNKNOWN                        'Unset
    
    KIND_LITERAL                        'A number, Param = value
    KIND_CODE                           'Program code, Param = ID
    
    '...                                'Strings, Variables, Constants &c...
End Enum

'This is where the user-defined Words are stored, each program instruction is either _
 a built-in Word, or an index to this list
Private Words() As SIXTH_Word
Private WordCount As Long

'Built-in Words:
'---------------------------------------------------------------------------------------

'These are the built-in Words (functions) of SIXTH. Each program instruction is simply _
 a Word ID to execute. The built in words are negative and the user's custom Words _
 begin at 1 upwards and are stored in the `Words()` array defined above
Public Enum SIXTH_WORDS
    WORD_END = 0                        'End the program
    
    WORD_ADD = -1                       ' +
    WORD_DIV = -2                       ' /
    WORD_DROP = -3                      'Drop top of stack
    WORD_DUP = -4                       'Duplicate top of stack
    WORD_ECHO = -5                      'Display top of stack
    WORD_MUL = -6                       ' *
    WORD_POW = -7                       ' ^
    WORD_ROT = -8                       'Rotate top three of stack
    WORD_SUB = -9                       ' -
    WORD_SWAP = -10                     'Swap top two of stack
End Enum

Private Const STR_ADD As String = "+"
Private Const STR_DIV As String = "/"
Private Const STR_DROP As String = "DROP"
Private Const STR_DUP As String = "DUP"
Private Const STR_ECHO As String = "."
Private Const STR_END As String = "END"
Private Const STR_MUL As String = "*"
Private Const STR_POW As String = "^"
Private Const STR_ROT As String = "ROT"
Private Const STR_SUB As String = "-"
Private Const STR_SWAP As String = "SWAP"

Private Program() As SIXTH_WORDS
Private ProgramLength As Long           '1-based length of the program
Private PC As Long                      'Program counter, 0-based

'Parsing / Assembling:
'---------------------------------------------------------------------------------------

'A look-up table of lower-case to upper-case Unicode characters
Private CharToUpperCase(0 To 65535) As Integer
'A look-up table of Unicode characters that are considered white-space
Private CharIsWhiteSpace(0 To 65535) As Long

'/// EVENTS ////////////////////////////////////////////////////////////////////////////

Event Log( _
    ByRef Message As String _
)

'CLASS Initialize
'=======================================================================================
Private Sub Class_Initialize()
    'Set up the dummy array that will be used to cast a String to an Array
    With CastStrSA1D
        Let .cDims = 1
        Let .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        Let .cbElements = 2&
        Let .cLocks = 1&
        Let .lLbound = 0&
    End With
    
    'Windows version check:
    '-----------------------------------------------------------------------------------
    'Get the Windows version as Windows Vista and above use alternative APIs for _
     locale-dependent methods (such as `Equals`)
    Dim VersionInfo As OSVERSIONINFO
    Let VersionInfo.SizeOfMe = Len(VersionInfo)
    'NOTE: If the app is in compatibility mode, this will return the compatible _
     Windows version, not the actual version. Windows 8.1 and above also cause _
     this API to always return as Windows 8.0 unless the app opts-in to correct _
     version info via manifest, but this isn't an issue to bluString as we are _
     only concerned with testing for Windows Vista+
    If api_GetVersionEx(VersionInfo) = API_TRUE Then
        If VersionInfo.MajorVersion < 6 Then Let WindowsVersionCheck = 1
    End If
    
    'Upper-case look-up table:
    '-----------------------------------------------------------------------------------
    'Begin with a straight forward list of all 2-byte UCS-2 code points
    Dim i As Long
    For i = 0 To &HFFFF&
        Dim Unicode(0 To &HFFFF&) As Integer
        If i >= 32768 Then Let Unicode(i) = i - 65536 Else Let Unicode(i) = i
    Next i
    
    'Use the Windows API to convert these to their linguistic upper-case equivilents
    'Check for Windows Vista and above:
    If WindowsVersionCheck = 0 Then
        'NOTE: The use of null on the first parameter causes Windows to use the _
         user's current locale, so we don't need to manually fetch the locale name
        Call api_LCMapStringEx( _
             LocaleNamePointer:=0&, _
                      MapFlags:=LCMAP_UPPERCASE Or LCMAP_LINGUISTIC_CASING, _
           SourceStringPointer:=VarPtr(Unicode(0)), _
            SourceStringLength:=UBound(Unicode) + 1, _
      DestinationStringPointer:=VarPtr(CharToUpperCase(0)), _
       DestinationStringLength:=UBound(CharToUpperCase) + 1, _
            VersionInfoPointer:=0&, ReservedPointer:=0&, SortHandle:=0& _
        )
    Else
        'Use the Windows XP compatible API
        Call api_LCMapString( _
                    LocaleID:=api_GetThreadLocale(), _
                    MapFlags:=LCMAP_UPPERCASE Or LCMAP_LINGUISTIC_CASING, _
         SourceStringPointer:=VarPtr(Unicode(0)), _
          SourceStringLength:=UBound(Unicode) + 1, _
    DestinationStringPointer:=VarPtr(CharToUpperCase(0)), _
     DestinationStringLength:=UBound(CharToUpperCase) + 1 _
        )
    End If
    
    'Unicode properties look-up tables:
    '-----------------------------------------------------------------------------------
    'Use the Windows API to convert the list of Unicode points into their properties. _
     This API doesn't allow overwriting the original array
    Dim Properties(0 To &HFFFF&) As Integer
    Call api_GetStringTypeW( _
                   InfoType:=CT_CTYPE1, _
        SourceStringPointer:=VarPtr(Unicode(0)), _
         SourceStringLength:=UBound(Unicode) + 1, _
            CharTypePointer:=VarPtr(Properties(0)) _
    )
    'We won't need the array of Unicode points any more
    Erase Unicode
    
    'Walk the array of Unicode properties and populate the look-up tables for _
     specfic properties...
    For i = 0 To 65535
        Let CharIsWhiteSpace(i) = _
            ((Properties(i) And C1_WHITESPACE) <> 0) _
         Or ((Properties(i) And C1_DEFINED) = 0)
    Next i
    Erase Properties
    
    Call Me.Clear
End Sub

'CLASS Terminate
'=======================================================================================
Private Sub Class_Terminate()
    'Free up the program
    Call Clear
End Sub

'/// PUBLIC INTERFACE //////////////////////////////////////////////////////////////////

'Clear : Clear the program
'=======================================================================================
Public Sub Clear()
    'Reset the machine's run-state
    Call Reset
    'Clear the user-Words
    Erase Words: Let WordCount = 0
    'Clear the program code
    Erase Program
    ReDim Program(0 To 0) As SIXTH_WORDS
End Sub

'LoadProgram
'=======================================================================================
'FilePath       | Path to the source code file to read and assemble
'---------------¦-----------------------------------------------------------------------
'Returns        | An error number
'=======================================================================================
Public Function LoadProgram( _
    ByRef FilePath As String _
) As SIXTH_ERROR
    'Clear any existing program and reset the VM
    Call Clear
       
    'Read the file
    '-----------------------------------------------------------------------------------
    On Error GoTo ErrFile
    Dim FileNumber As Integer: Let FileNumber = FreeFile
    Open FilePath For Input Lock Write As #FileNumber
    
    Dim Length As Long
    Let Length = VBA.LOF(FileNumber)
    
    Dim Source As String
    Let Source = StrConv(InputB(Length, FileNumber), vbUnicode)
    
ErrFile:
    Close #FileNumber
    On Error GoTo 0
    
    'If the file is empty, nothing can be assembled!
    If Source = vbNullString Then Exit Function
    
    Call Assemble(Source)
End Function

'Parse : Takes a VB String and assembles it as SIXTH code
'=======================================================================================
'Source         | Any text that will be treated as SIXTH code
'---------------¦-----------------------------------------------------------------------
'Returns        | An error number (e.g. if the text failed to parse/assemble)
'=======================================================================================
Public Function Parse( _
    ByRef Source As String _
) As SIXTH_ERROR
    'If the text is empty, nothing can be parsed!
    If Source = vbNullString Then Exit Function
    
    Let Parse = Assemble(Source)
    
    'Now execute the instructions...
    If Parse = ERR_NONE Then Let Parse = Me.Run()
End Function

'Reset : Does not clear the current program, but resets the VM ready to run
'=======================================================================================
Public Sub Reset()
    'Rewind the program
    Let PC = 0
    'Reset the stack data
    Erase Stack: Let SP = 0
End Sub

'Run
'=======================================================================================
'Returns        | An error number
'=======================================================================================
Public Function Run( _
) As SIXTH_ERROR
    'Is there a program present?
    If ProgramLength = 0 Then Exit Function
    
    'These are temporary holders for values popped off the stack for calculations
    Dim A As Double
    Dim B As Double
    Dim C As Double
    
    'Rewind the program
    Call Reset
    
    'Execute:
    '-----------------------------------------------------------------------------------
    Do
        Dim Word As SIXTH_WORDS
        Let Word = Program(PC)
        
        'NOTE: We'll be doing as much as possible in-line (e.g. stack operations), _
         purely for speed reasons. In tight loops, function calls are slow
            
        If Word = WORD_ADD Then
            '...........................................................................
            'Pop the first value
            If SP = 0 Then GoTo ErrStackUnderflow
            Let A = Stack(SP): Let SP = SP - 1
            'Pop the second value
            If SP = 0 Then GoTo ErrStackUnderflow
            Let B = Stack(SP)
            'Do the addition and push to the stack
            On Error GoTo ErrOverflow
            Let Stack(SP) = A + B
            On Error GoTo 0
            
        ElseIf Word = WORD_DIV Then
            '...........................................................................
            'Pop the first value
            If SP = 0 Then GoTo ErrStackUnderflow
            Let A = Stack(SP): Let SP = SP - 1
            'Pop the second value
            If SP = 0 Then GoTo ErrStackUnderflow
            Let B = Stack(SP)
            'Are we going to divide by zero?
            If B = 0 Then GoTo ErrDivideByZero
            'Do the division and push to the stack
            Let Stack(SP) = A / B
        
        ElseIf Word = WORD_DROP Then
            '...........................................................................
            If SP = 0 Then GoTo ErrStackUnderflow
            Let SP = SP - 1
            
        ElseIf Word = WORD_DUP Then
            'Duplicate the top of the stack
            '...........................................................................
            'Is there anything to duplicate!?
            If SP = 0 Then GoTo ErrStackUnderflow
            'Push the stack down and duplicate the value
            Let SP = SP + 1: If SP = 257 Then GoTo ErrStackOverflow
            Let Stack(SP) = Stack(SP - 1)
            
        ElseIf Word = WORD_ECHO Then
            '...........................................................................
            'Log the top of the stack
            If SP = 0 Then GoTo ErrStackUnderflow
            RaiseEvent Log(CStr(Stack(SP)))
            
        ElseIf Word = WORD_END Then
            '...........................................................................
            Exit Do
        
        ElseIf Word = WORD_MUL Then
            '...........................................................................
            'Pop the first value
            If SP = 0 Then GoTo ErrStackUnderflow
            Let A = Stack(SP): Let SP = SP - 1
            'Pop the second value
            If SP = 0 Then GoTo ErrStackUnderflow
            Let B = Stack(SP)
            'Do the multiplication and push to the stack
            On Error GoTo ErrOverflow
            Let Stack(SP) = A * B
            On Error GoTo 0
        
        ElseIf Word = WORD_POW Then
            '...........................................................................
            'Pop the first value
            If SP = 0 Then GoTo ErrStackUnderflow
            Let A = Stack(SP): Let SP = SP - 1
            'Pop the second value
            If SP = 0 Then GoTo ErrStackUnderflow
            Let B = Stack(SP)
            'Do the power and push to the stack
            On Error GoTo ErrOverflow
            Let Stack(SP) = A ^ B
            On Error GoTo 0
        
        ElseIf Word = WORD_ROT Then
            '...........................................................................
            'Is the stack at least three deep in the first place?
            If SP < 3 Then GoTo ErrStackUnderflow
            'Take aside the 3rd value
            Let A = Stack(SP - 3)
            'Move the 2nd down to the 3rd
            Let Stack(SP - 3) = Stack(SP - 2)
            'Move the 1st down to the 2nd
            Let Stack(SP - 1) = Stack(SP)
            'Put the 3rd on top
            Let Stack(SP) = A
            
        ElseIf Word = WORD_SUB Then
            '...........................................................................
            'Pop the first value
            If SP = 0 Then GoTo ErrStackUnderflow
            Let A = Stack(SP): Let SP = SP - 1
            'Pop the second value
            If SP = 0 Then GoTo ErrStackUnderflow
            Let B = Stack(SP)
            'Do the subtraction and push to the stack
            On Error GoTo ErrOverflow
            Let Stack(SP) = A - B
            On Error GoTo 0
            
        ElseIf Word = WORD_SWAP Then
            '...........................................................................
            'Is the stack at least two deep in the first place?
            If SP < 2 Then GoTo ErrStackUnderflow
            'Put aside the top value
            Let A = Stack(SP)
            'Swap in the 2nd value
            Let Stack(SP) = Stack(SP - 1)
            'Put the old top value below
            Let Stack(SP - 1) = A
            
        Else
            '...........................................................................
            'A custom word
            
            Dim Kind As SIXTH_KIND
            Let Kind = Words(Word).Kind
            
            If Kind = KIND_LITERAL Then
                Let SP = SP + 1: If SP = 257 Then GoTo ErrStackOverflow
                Let Stack(SP) = Words(Word).Param
            End If
            
        End If
        
        'Next instruction in the program...
        Let PC = PC + 1
    Loop
    Exit Function
    
    '-----------------------------------------------------------------------------------
ErrStackUnderflow:
    Let Run = ERR_STACKUNDERFLOW:   Exit Function
ErrStackOverflow:
    Let Run = ERR_STACKOVERFLOW:    Exit Function
ErrOverflow:
    Let Run = ERR_OVERFLOW:         Exit Function
ErrDivideByZero:
    Let Run = ERR_DIVIDEBYZERO:     Exit Function
End Function

'/// PRIVATE ///////////////////////////////////////////////////////////////////////////

'Assemble : Assemble source code into a program
'=======================================================================================
'Source         | Any text that will be treated as SIXTH code
'---------------¦-----------------------------------------------------------------------
'Returns        | An error number (e.g. if the text failed to parse/assemble)
'=======================================================================================
Private Function Assemble( _
    ByRef Source As String _
) As SIXTH_ERROR
    'If the text is empty, nothing can be parsed!
    If Source = vbNullString Then Exit Function
    
    '-----------------------------------------------------------------------------------
    'Initialise:
    '-----------------------------------------------------------------------------------
    
    'Map the VB String to an array so we can compare character codes quickly
    Let CastStrPtr = msvbvm60_VarPtr(CastStrArr())
    Call msvbvm60_PutMem4(CastStrPtr, VarPtr(CastStrSA1D))
    
    'Take the string, and attach it to our dummy array header
    With CastStrSA1D
        Let .pvData = StrPtr(Source)
        Let .cElements = Len(Source)
    End With
    
    '-----------------------------------------------------------------------------------
    'Parse Words:
    '-----------------------------------------------------------------------------------
    'We will walk through the source code letter-by-letter and break out the words
    
    Dim Line As Long: Let Line = 1
    Dim Col As Long: Let Col = 1
    
    'In Forth, all whitespace is ignored and considered a word-break, including _
     within strings and comments. In SIXTH we want to include whitespace when walking _
     strings and comments because a.) we support Unicode in text and b.) we want to _
     discard comments wholesale and not tokenise them as indivdual words
    Dim Whitespace As Long
    Dim CommentMode As Long
    
    Dim Word As String
    Dim WordUp As String
    
    Dim i As Long
    Do
        Dim Char As Integer
        Let Char = CastStrArr(i)
        
        'If outside of strings and comments, we will take whitespace as word-break
        If Whitespace = 0 Then
            '...........................................................................
            'Is the current letter whitespace?
            If CharIsWhiteSpace(Char) Then
                'If a Word is defined, break here
                If LenB(Word) > 0 Then GoSub AssembleWord
            Else
                'Add the letter to the current Word
                Let Word = Word & ChrW$(Char)
                Let WordUp = WordUp & ChrW$(CharToUpperCase(Char))
            End If
            
        'Currently parsing a string or comment...
        Else
            '...........................................................................
        
        End If
        
        'Move on to the next character
        Let i = i + 1
    Loop While i < Len(Source)
    
    If LenB(Word) > 0 Then GoSub AssembleWord
    
    'Clean up our array hacking (for casting a string to an array) _
     otherwise VB will crash when it frees up the memory
    Call msvbvm60_PutMem4(CastStrPtr, 0&)
    Erase CastStrArr
    
    Exit Function
    
    '-----------------------------------------------------------------------------------
    'Assemble:
    '-----------------------------------------------------------------------------------
AssembleWord:
    Dim Instruction As SIXTH_WORDS
    
    'Begin with the built-in instructions
    If WordUp = STR_ADD Then
        Let Instruction = WORD_ADD
    ElseIf WordUp = STR_DIV Then Let Instruction = WORD_DIV
    ElseIf WordUp = STR_DROP Then Let Instruction = WORD_DROP
    ElseIf WordUp = STR_DUP Then Let Instruction = WORD_DUP
    ElseIf WordUp = STR_ECHO Then Let Instruction = WORD_ECHO
    ElseIf WordUp = STR_END Then Let Instruction = WORD_END
    ElseIf WordUp = STR_MUL Then Let Instruction = WORD_MUL
    ElseIf WordUp = STR_POW Then Let Instruction = WORD_POW
    ElseIf WordUp = STR_ROT Then Let Instruction = WORD_ROT
    ElseIf WordUp = STR_SUB Then Let Instruction = WORD_SUB
    ElseIf WordUp = STR_SWAP Then Let Instruction = WORD_SWAP
    Else
        'Check to see if the Word is already defined
        Dim w As Long
        For w = 1 To WordCount
            If Equals(Word, Words(w).Text) Then
                Let Instruction = w
                GoTo AddInstruction
            End If
        Next w
        
        'Define a custom word
        Let WordCount = WordCount + 1
        ReDim Preserve Words(1 To WordCount)
        
        If IsNumeric(Word) Then
            Let Words(WordCount).Kind = KIND_LITERAL
            Let Words(WordCount).Text = Word
            Let Words(WordCount).Param = CDbl(Word)
            Let Instruction = WordCount
        Else
            Stop
        End If
    End If
    
    'So that we don't have to check for the end of the program all the time, we always _
     make the program 1 instruction longer and cap it with END. Thus the program will _
     naturally come to a halt without having to check for an out-of-bounds array check _
     every execute cycle
AddInstruction:
    
    'Assign the instruction
    Let Program(ProgramLength) = Instruction
    'Increase the length of the program
    Let ProgramLength = ProgramLength + 1
    ReDim Preserve Program(0 To ProgramLength) As SIXTH_WORDS
    'Cap the program with the END instruction
    Let Program(ProgramLength) = WORD_END
    
    'Prepare for the next Word
    Let Word = vbNullString
    Let WordUp = vbNullString
    Return
End Function

'Equals : Case-insensitive string comparison
'=======================================================================================
'Word1          | First string
'Word2          | String to compare with the first
'---------------¦-----------------------------------------------------------------------
'Returns        | Boolean (as Long, for speed reasons)
'=======================================================================================
Private Function Equals( _
    ByRef Word1 As String, _
    ByRef Word2 As String _
) As Long
    'Check for Windows Vista and above:
    If WindowsVersionCheck = 0 Then
        'NOTE: The use of null on the first parameter causes Windows to use the _
         user's current locale, so we don't need to manually fetch the locale name
        Let Equals = (api_CompareStringEx( _
             LocaleNamePointer:=0&, _
                  CompareFlags:=LINGUISTIC_IGNORECASE, _
                String1Pointer:=StrPtr(Word1), String1Length:=-1, _
                String2Pointer:=StrPtr(Word2), String2Length:=-1, _
            VersionInfoPointer:=0&, _
               ReservedPointer:=0&, _
                         Param:=0& _
        ) = CSTR_EQUAL)
    Else
        'Fall back to the XP-compatible API
        Let Equals = (api_CompareString( _
                  LocaleID:=api_GetThreadLocale(), _
              CompareFlags:=NORM_IGNORECASE, _
            String1Pointer:=StrPtr(Word1), String1Length:=-1, _
            String2Pointer:=StrPtr(Word2), String2Length:=-1 _
        ) = CSTR_EQUAL)
    End If
End Function
