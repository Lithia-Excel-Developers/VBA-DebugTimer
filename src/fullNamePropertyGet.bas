Attribute VB_Name = "fullNamePropertyGet"
Option Explicit

Private FirstName As String
Private LastName As String


Property Get FullName() As String
    FullName = FirstName & " " & LastName
End Property


Sub FullNameTest()
   
    FirstName = "Kevin"
    LastName = "Conner"
    
    Debug.Print FullName

    
End Sub
