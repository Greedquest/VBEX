VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VarDelegate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'
' Var Delegate
' ========
'
' Black Magic...
'    Lambda.cls writes source code and has that code compiled and executed all
' at run-time.  Delegate.cls modifies it's own compiled code at run-time.  In
' summary, we will replace the address of a member function (ProtoProc) with 
' the address of a given global function. 
'    Replacing the member function address is manageable for Sub routines with 
' no parameters.  In order to handle functions or any procedure that takes 
' parameters we must begin injecting compiled ASM code actually modify procedure
' I don't know ASM so all ASM code was provided from
'
' http://www.codeproject.com/Articles/19239/Compilers-Demystified-Function-Pointers-in-Visual
'
'    and the code there isn't complete.  It doesn't handle procedures with
' different signatures.  As we want one delegate to rule them all someone will
' need to learn some ASM and pick up the torch.
'
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (destination As Any, source As Any, ByVal Length As Long)
' http://www.ex-designz.net/apidetail.asp?api_id=235
Private Declare Function GlobalFree Lib "kernel32.dll" _
        (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32.dll" _
        (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Const GMEM_FIXED = &H0

' Black magic spell
' It's compiled ASM binary code represented as a hexadecimal string
' We will inject into a function and it will do shit I don't understand
'
Private Const hexStub4ProcedureCall = "585950B822114000FFE0"


'
' Private Members
' ---------------
Private Const OFFSET_TO_VFT = &H1C
Private Const VFT_ELEM_SIZE As Integer = 4

Private pProcAddress As Long ' address to copy to pRunVFT
Private pStubAddr As Long
Private Type tVFTRecord
    keyAddr As Long
    valAddr As Long
End Type
Private delegateProcRec As tVFTRecord


'
' Proto PROC
' ----------
'
' This procedure will be transfigured with black magic
' it needs to be at the top ALWAYS!!!!!
' casue black magic...
Public Sub ProtoProc(ByVal x As Variant)
    Exceptions.Unimplemented
End Sub
'
' Constructors
' ------------
'
Private Sub Class_Initialize()
    
    ' get VFT table
    Dim vft As Long ' location of vftable
    CopyMemory vft, ByVal ObjPtr(Me), VFT_ELEM_SIZE
    
    ' calculate `ProtoProc` location on VFT table
    delegateProcRec.keyAddr = vft + OFFSET_TO_VFT
    
    ' store Address of `ProtoProc` from VFT table
    CopyMemory delegateProcRec.valAddr, ByVal delegateProcRec.keyAddr, VFT_ELEM_SIZE
    
End Sub
Private Sub Class_Terminate()

    RestoreProtoProc
    
End Sub
Public Function Create(ByVal procAddr As Long) As VarDelegate

    Dim result As New VarDelegate
    result.ProcedureAddress = procAddr
    Set Create = result
    
End Function
Friend Property Let ProcedureAddress(ByVal procAddr As Long)
    
    pProcAddress = procAddr
    
End Property
Public Sub Run(ParamArray args() As Variant)
    
    ' > One more important thing to remember when modifying vftables,
    ' > is that they are shared by all instances of that class.
    '
    ' ergo we must modify the global definition before we run and
    ' return it to normal afterwards.
    MoveDelegateToProtoProc
    ProtoProc CVar(args)
    RestoreProtoProc
    ' This is about as un-threadsafe as one can be...
    
End Sub
'
' Private Methods
' ---------------
'
Private Sub WriteToProtoProc(ByVal procAddr As Long)

    CopyMemory ByVal delegateProcRec.keyAddr, procAddr, VFT_ELEM_SIZE
      
End Sub
Private Sub MoveDelegateToProtoProc()

    DoBlackMagic pProcAddress
    
End Sub
Private Sub RestoreProtoProc()

    DoWhiteMagic
    
End Sub
Private Sub DoBlackMagic(ByVal fnAddress As Long)

    Dim nBytes As Long 'number of code bytes to allocate
    nBytes = Len(hexStub4ProcedureCall) \ 2
    
    Dim bytes() As Byte 'array of code bytes converted from hex
    bytes = ReadByteMagic(nBytes, hexStub4ProcedureCall)
    
    ' replace the forwarding address in the native code
    ' this is hacker 101 shit
    CopyMemory bytes(VFT_ELEM_SIZE + 1), fnAddress, VFT_ELEM_SIZE
    
    'allocate memory to store the code bytes
    pStubAddr = GlobalAlloc(GMEM_FIXED, nBytes)
    
    CopyMemory ByVal pStubAddr, bytes(LBound(bytes)), nBytes 'copy given code bytes
    WriteToProtoProc pStubAddr  'replace member address with stub address
    
End Sub
Private Sub DoWhiteMagic()
        
    WriteToProtoProc delegateProcRec.valAddr
    GlobalFree pStubAddr
    pStubAddr = 0
    
End Sub
''
' ReadByteMagic: Convert compiled ASM code back to binary code
'
Private Function ReadByteMagic(ByVal byteCount As Long, ByVal hexCode As String) As Byte()
    
    Dim results() As Byte
    ReDim results(1 To byteCount)
    
    Dim i As Long
    For i = 1 To byteCount
        'convert each pair of hex chars to a byte code
        results(i) = Val("&H" & Mid$(hexCode, i * 2 - 1, 2))
    Next i
    
    ReadByteMagic = results

End Function


