Attribute VB_Name = "TestStr"
'@TestModule
Option Explicit
Option Private Module
Private Assert As New Rubberduck.AssertClass
'
'
' Constructors
' ------------
'
'@TestMethod
Public Sub TestStrJoin()

    Dim s As str
    Set s = str.Join(List.Create("Hello", "World"), ", ")
    
    Assert.AreEqual "Hello, World", s.Show

End Sub
'@TestMethod
Public Sub TestStrMake()

    Dim s As str
    Set s = str.Make("Hello, World")
    
    Assert.AreEqual "Hello, World", s.Show

End Sub
'@TestMethod
Public Sub TestStrRepeat()

    Dim s As str
    Set s = str.Repeat("Spam", 3)
    
    Assert.AreEqual "SpamSpamSpam", s.Show

End Sub
'@TestMethod
Public Sub TestStrFormat()

    Dim s As str
    Set s = str.Format("{0}, {2}, {1}", "a", 2, 4.5)
    
    Assert.AreEqual "a, 4.5, 2", s.Show

End Sub
'@TestMethod
Public Sub TestStrEscape()

    Dim s As str
    Set s = str.Escape("&Phil's parrot said ""I'm not dead""")
    
    Assert.AreEqual "`&Phil`'s` parrot` said` `""I`'m` not` dead`""", s.Show

End Sub
'@TestMethod
Public Sub StrIterable()

    BatteryIterable.Battery str.Make("Hello, World")

End Sub

