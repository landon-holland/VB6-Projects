Attribute VB_Name = "modGlobal"
Option Explicit

'Declare
Type student
    gpa As Single
    id As Long
    last As String
End Type

Global arrstudents(1 To 100) As student
