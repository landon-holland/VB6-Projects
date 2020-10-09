Attribute VB_Name = "modEmployees"
Option Explicit

Type employee
    firstname As String
    lastname As String
    age As Integer
    id As Integer
    paytype As String
    wage As Single
    phonenumber As String
End Type

Global arremployees(1 To 6) As employee

Global entries As Integer
Global currententry As Integer
Global entrytochange As Integer
Global entrytodelete As Integer
