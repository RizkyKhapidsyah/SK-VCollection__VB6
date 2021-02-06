Attribute VB_Name = "VShared"

Public Function DictionaryTreeTest()
    Dim AddressBook As New VDictionaryTree
    Dim Item As VDictionaryTree

    Dim Strings(10) As String
    
    Strings(1) = "First Name"
    Strings(2) = "Last Name"
    Strings(3) = "Age"
    Strings(4) = "Favorite Fruit"
    Strings(5) = "Height"
    Strings(6) = "Salary"
    
    AddressBook.Add "Larry"
    AddressBook.Add "Harry"
    AddressBook.Add "Jim"
    AddressBook.Add "Soonyee"
    
    For Each Item In AddressBook
        Debug.Print Item.Value
    Next

    AddressBook!Coworkers!John!Phone = "201-345-3456"
    Debug.Print AddressBook!Coworkers!John!Phone
    
    Debug.Print AddressBook("Coworkers")("John")("Phone")
    
    AddressBook!Coworkers!Larry!JobTitle = "Boss"
    Set AddressBook!Coworkers!John!Supervisor = AddressBook!Coworkers!Larry
    
    Debug.Print AddressBook!Coworkers!John!Supervisor!JobTitle
    
    AddressBook!People!Friends.Remove 1

    AddressBook!People!Friends!Jim = "Hello"
    Debug.Print AddressBook!People!Friends!Jim
    
    For Each Item In AddressBook!Coworkers
        Debug.Print Item.Key
    Next
End Function

Public Function LetOrSet(InputValue As Variant, OutputValue As Variant)
    If IsObject(InputValue) = True Then
        Set OutputValue = InputValue
    Else
        Let OutputValue = InputValue
    End If
End Function

Public Sub Main()
DictionaryTreeTest
MsgBox "test completed succesfully"
End
End Sub
