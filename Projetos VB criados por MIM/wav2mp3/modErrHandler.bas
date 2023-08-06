Attribute VB_Name = "modErrHandler"
Option Explicit
'This API function installs your custom exception handler.


Declare Function SetUnhandledExceptionFilter Lib "kernel32" (ByVal lpTopLevelExceptionFilter As Long) As Long
    'This API function is used to raise exceptions.


Declare Sub RaiseException Lib "kernel32" (ByVal dwExceptionCode As Long, ByVal dwExceptionFlags As Long, ByVal nNumberOfArguments As Long, lpArguments As Long)
    'Possible return values for the Unhandled Exception Filter.
Public Const EXCEPTION_CONTINUE_EXECUTION = -1
Public Const EXCEPTION_CONTINUE_SEARCH = 0
Public Const EXCEPTION_EXECUTE_HANDLER = 1
'Maximum number of parameters an Exception_Record can have
Public Const EXCEPTION_MAXIMUM_PARAMETERS = 15
'Structure that contains processor-specific register data


Type CONTEXT
    FltF0 As Double
    FltF1 As Double
    FltF2 As Double
    FltF3 As Double
    FltF4 As Double
    FltF5 As Double
    FltF6 As Double
    FltF7 As Double
    FltF8 As Double
    FltF9 As Double
    FltF10 As Double
    FltF11 As Double
    FltF12 As Double
    FltF13 As Double
    FltF14 As Double
    FltF15 As Double
    FltF16 As Double
    FltF17 As Double
    FltF18 As Double
    FltF19 As Double
    FltF20 As Double
    FltF21 As Double
    FltF22 As Double
    FltF23 As Double
    FltF24 As Double
    FltF25 As Double
    FltF26 As Double
    FltF27 As Double
    FltF28 As Double
    FltF29 As Double
    FltF30 As Double
    FltF31 As Double
    IntV0 As Double
    IntT0 As Double
    IntT1 As Double
    IntT2 As Double
    IntT3 As Double
    IntT4 As Double
    IntT5 As Double
    IntT6 As Double
    IntT7 As Double
    IntS0 As Double
    IntS1 As Double
    IntS2 As Double
    IntS3 As Double
    IntS4 As Double
    IntS5 As Double
    IntFp As Double
    IntA0 As Double
    IntA1 As Double
    IntA2 As Double
    IntA3 As Double
    IntA4 As Double
    IntA5 As Double
    IntT8 As Double
    IntT9 As Double
    IntT10 As Double
    IntT11 As Double
    IntRa As Double
    IntT12 As Double
    IntAt As Double
    IntGp As Double
    IntSp As Double
    IntZero As Double
    Fpcr As Double
    SoftFpcr As Double
    Fir As Double
    Psr As Long
    ContextFlags As Long
    Fill(4) As Long
End Type
    'Structure that describes an exception.


Type EXCEPTION_RECORD
    ExceptionCode As Long
    ExceptionFlags As Long
    pExceptionRecord As Long ' Pointer to an EXCEPTION_RECORD structure
    ExceptionAddress As Long
    NumberParameters As Long
    ExceptionInformation(EXCEPTION_MAXIMUM_PARAMETERS) As Long
End Type
'Structure that contains exception information that can be used b
'     y a debugger.


Type EXCEPTION_DEBUG_INFO
    pExceptionRecord As EXCEPTION_RECORD
    dwFirstChance As Long
End Type
'The EXCEPTION_POINTERS structure contains an exception record wi
'     th a
'machine-independent description of an exception and a context re
'     cord
'with a machine-dependent description of the processor context at
'     the
'time of the exception.


Type EXCEPTION_POINTERS
    pExceptionRecord As EXCEPTION_RECORD
    ContextRecord As CONTEXT
End Type
'Standard Exception Codes
Public Const EXCEPTION_ACCESS_VIOLATION = &HC0000005
Public Const EXCEPTION_DATATYPE_MISALIGNMENT = &H80000002
Public Const EXCEPTION_BREAKPOINT = &H80000003
Public Const EXCEPTION_SINGLE_STEP = &H80000004
Public Const EXCEPTION_ARRAY_BOUNDS_EXCEEDED = &HC000008C
Public Const EXCEPTION_FLT_DENORMAL_OPERAND = &HC000008D
Public Const EXCEPTION_FLT_DIVIDE_BY_ZERO = &HC000008E
Public Const EXCEPTION_FLT_INEXACT_RESULT = &HC000008F
Public Const EXCEPTION_FLT_INVALID_OPERATION = &HC0000090
Public Const EXCEPTION_FLT_OVERFLOW = &HC0000091
Public Const EXCEPTION_FLT_STACK_CHECK = &HC0000092
Public Const EXCEPTION_FLT_UNDERFLOW = &HC0000093
Public Const EXCEPTION_INT_DIVIDE_BY_ZERO = &HC0000094
Public Const EXCEPTION_INT_OVERFLOW = &HC0000095
Public Const EXCEPTION_PRIVILEGED_INSTRUCTION = &HC0000096
Public Const EXCEPTION_IN_PAGE_ERROR = &HC0000006
Public Const EXCEPTION_ILLEGAL_INSTRUCTION = &HC000001D
Public Const EXCEPTION_NONCONTINUABLE_EXCEPTION = &HC0000025
Public Const EXCEPTION_STACK_OVERFLOW = &HC00000FD
Public Const EXCEPTION_INVALID_DISPOSITION = &HC0000026
Public Const EXCEPTION_GUARD_PAGE_VIOLATION = &H80000001
Public Const EXCEPTION_INVALID_HANDLE = &HC0000008
Public Const EXCEPTION_CONTROL_C_EXIT = &HC000013A
'This is a friendly declaration of the CopyMemory function. It is
'     used to copy
'data into an EXTENSION_RECORD structure from a pointer to anothe
'     r structure.


Declare Sub CopyExceptionRecord Lib "kernel32" Alias "RtlMoveMemory" (pDest As EXCEPTION_RECORD, ByVal LPEXCEPTION_RECORD As Long, ByVal lngBytes As Long)
'******************************
' GetExceptionText
'******************************
' This function receives an exception code value and returns the
' text description of the exception.
'


Public Function GetExceptionText(ByVal ExceptionCode As Long) As String


    Dim strExceptionString As String
    


    Select Case ExceptionCode
        Case EXCEPTION_ACCESS_VIOLATION
        strExceptionString = "Access Violation"
        Case EXCEPTION_DATATYPE_MISALIGNMENT
        strExceptionString = "Data Type Misalignment"
        Case EXCEPTION_BREAKPOINT
        strExceptionString = "Breakpoint"
        Case EXCEPTION_SINGLE_STEP
        strExceptionString = "Single Step"
        Case EXCEPTION_ARRAY_BOUNDS_EXCEEDED
        strExceptionString = "Array Bounds Exceeded"
        Case EXCEPTION_FLT_DENORMAL_OPERAND
        strExceptionString = "Float Denormal Operand"
        Case EXCEPTION_FLT_DIVIDE_BY_ZERO
        strExceptionString = "Divide By Zero"
        Case EXCEPTION_FLT_INEXACT_RESULT
        strExceptionString = "Floating Point Inexact Result"
        Case EXCEPTION_FLT_INVALID_OPERATION
        strExceptionString = "Invalid Operation"
        Case EXCEPTION_FLT_OVERFLOW
        strExceptionString = "Float Overflow"
        Case EXCEPTION_FLT_STACK_CHECK
        strExceptionString = "Float Stack Check"
        Case EXCEPTION_FLT_UNDERFLOW
        strExceptionString = "Float Underflow"
        Case EXCEPTION_INT_DIVIDE_BY_ZERO
        strExceptionString = "Integer Divide By Zero"
        Case EXCEPTION_INT_OVERFLOW
        strExceptionString = "Integer Overflow"
        Case EXCEPTION_PRIVILEGED_INSTRUCTION
        strExceptionString = "Privileged Instruction"
        Case EXCEPTION_IN_PAGE_ERROR
        strExceptionString = "In Page Error"
        Case EXCEPTION_ILLEGAL_INSTRUCTION
        strExceptionString = "Illegal Instruction"
        Case EXCEPTION_NONCONTINUABLE_EXCEPTION
        strExceptionString = "Non Continuable Exception"
        Case EXCEPTION_STACK_OVERFLOW
        strExceptionString = "Stack Overflow"
        Case EXCEPTION_INVALID_DISPOSITION
        strExceptionString = "Invalid Disposition"
        Case EXCEPTION_GUARD_PAGE_VIOLATION
        strExceptionString = "Guard Page Violation"
        Case EXCEPTION_INVALID_HANDLE
        strExceptionString = "Invalid Handle"
        Case EXCEPTION_CONTROL_C_EXIT
        strExceptionString = "Control-C Exit"
        Case Else
        strExceptionString = "Unknown (&H" & Right("00000000" & Hex(ExceptionCode), 8) & ")"
    End Select


GetExceptionText = strExceptionString
End Function


'******************************
' MyExceptionFilter
'******************************
' This function will be called when an unhandled exception occurs
'     .
' It raises an error so that it can be trapped with an ON ERROR s
'     tatement
' in the procedure that caused the exception.


Public Function MyExceptionFilter(ByRef ExceptionPtrs As EXCEPTION_POINTERS) As Long


    Dim Rec As EXCEPTION_RECORD
    Dim strException As String
    
    'Get the current exception record.
    Rec = ExceptionPtrs.pExceptionRecord
    
    'If Rec.pExceptionRecord is not zero, then it is a nested excepti
    '     on and
    'Rec.pExceptionRecord points to another EXCEPTION_RECORD structur
    '     e. Follow
    'the pointers back to the original exception.


    Do Until Rec.pExceptionRecord = 0
        'A friendly declaration of CopyMemory.
        CopyExceptionRecord Rec, Rec.pExceptionRecord, Len(Rec)
    Loop


    
    'Translate the exception code into a user-friendly string.
    strException = GetExceptionText(Rec.ExceptionCode)
    
    'Raise an error to return control to the calling procedure.
    Err.Raise 10000, "MyExceptionFilter", strException
End Function

