Attribute VB_Name = "TestVarDelegate"
Option Explicit
Option Private Module

'@TestModule
Private Assert As New Rubberduck.AssertClass
Private globalVar As String

Private Const SEP  As String = " "
Sub DelegateSub(ByVal str As String, ByVal n As Integer)

    Dim result As String

    Dim i As Integer
    For i = 1 To n
        result = result & str & SEP
    Next
    
    globalVar = result
    
End Sub
Sub DelegateSubAlt(ByVal str As String, ByVal n As Integer)
    
    Dim result As String
    
    Dim i As Integer
    For i = 1 To n
        result = result & "Alt" & str & SEP
    Next
    
    globalVar = result
    
End Sub
Sub VarWrapper(ByVal args As Variant)

    DelegateSub args(0), args(1)
    
End Sub
Sub VarWrapperAlt(ByVal args As Variant)

    DelegateSubAlt args(0), args(1)
    
End Sub
'@TestMethod
Sub TestArrayArgs()
    
    globalVar = ""
    Assert.AreEqual "", globalVar, "Check initial conditions"
    
    Dim f As VarDelegate
    Set f = VarDelegate.Make(AddressOf VarWrapper)
    f.Run "Spam", 3

    Assert.AreEqual "Spam Spam Spam ", globalVar

End Sub
'@TestMethod
Sub TestArgsWithMultipleInstances()
    
    globalVar = ""
    Assert.AreEqual "", globalVar, "Check initial conditions"
    
    Dim f1 As VarDelegate
    Set f1 = VarDelegate.Make(AddressOf VarWrapper)
    f1.Run "Spam", 3
    
    Assert.AreEqual "Spam Spam Spam ", globalVar
    
    Dim f2 As VarDelegate
    Set f2 = VarDelegate.Make(AddressOf VarWrapperAlt)
    f2.Run "ernative", 3
    
    Assert.AreEqual "Alternative Alternative Alternative ", globalVar

End Sub
