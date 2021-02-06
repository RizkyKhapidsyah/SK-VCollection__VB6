VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VDictionaryTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'AUTHOR:  Stephen Goguen
'EMAIL: gsteve@hotmail.com

'OVERVIEW:
'When I decided to make this, I wanted to make a feature rich collection
'class, but I wanted to keep it friendly for the novice.  The interface
'should only expose a handful of methods/properties, but each method and
'property should be feature rich.

'Let's look at an example...

'   Dim Organizer As New VDictionaryTree
'
'   'Basic useage
'   Organizer = "My personal organizer"
'   Organizer!Friends = "All of my friends contact info"
'   Debug.Print Organizer   'Guess what this prints...
'   Debug.Print Organizer("Friends")    'Same as Organizer!Friends

'   'Can your collection class do this???
'   Organizer!Friends!Chuck!Phone = "201-536-3451"
'   Organizer!Friends!Jon!Phone = "201-456-5361"
'   Set Organizer!Friends!Chuck!Brother = Jon

Option Explicit

Private mKey As String
Private mValue As Variant
Dim mCollection As Collection

'PROPERTY:  Get Key
'DESCRIPTION:  Public Key used to retreive key in Key/Value pair.
Public Property Get Key() As String
    Key = mKey
End Property

'PROPERTY:  Let Key
'DESCRIPTION:  Used by relative classes and functions to set the key
'in the Key/Value pair.
Friend Property Let Key(Key As String)
    Let mKey = Key
End Property

'PROPERTY:  Get Value
'DESCRIPTION:  Returns the a Value or Sub dictionary given a Key
Public Property Get Value(Optional Index As Variant) As Variant
Attribute Value.VB_UserMemId = 0
    
    Dim AddAgent As VAgent
    Dim AddField As VDictionaryTree
    
    On Error Resume Next
    
    'Are we getting a value???
    If IsMissing(Index) = True Then
        
        'Are we returning an object?
        If IsObject(mValue) = True Then
            Set Value = mValue
        Else
            Let Value = mValue
        End If
        
        'Never return Empty...  It breaks other code...
        'This should be irrelivant soon...
        If IsEmpty(Value) = True Then
            Let Value = Null
        End If
    
    Else
        
        'mValue could be an array, a recordset, or any other object that
        'allows Object(Index).  Test this first...
        Value = mValue(Index)
        
        If Err.Number <> 0 Then
            Err.Clear
            
            'Allocate our collection if not allocated
            CheckCollectionAndAllocate
            
            'Attempt to retrieve index from collection...
            Set Value = mCollection(Index)
            
            If Err.Number <> 0 Then
                
                'This field does not exist, return an AddAgent which
                'allows the user to query multiple sub-fields or add
                'multiple sub-fields intelligently.
                Set AddAgent = New VAgent
                AddAgent.SetParent Me
                AddAgent.Key = Index
                Set AddField = AddAgent
                Set Value = AddField
            End If
        End If
    End If
End Property

'FUNCTION:  Add
'DESCRIPTION:  Adds an item to the Dictionary tree, given
Public Function Add(Value As Variant, Optional Key As Variant, Optional Before As Variant, Optional After As Variant) As Variant
    Dim AddValue As VDictionaryTree
    CheckCollectionAndAllocate
    Set AddValue = New VDictionaryTree
    If IsMissing(Key) = False Then
        AddValue.Key = Key
    End If
    AddValue.Value = Value
    
    mCollection.Add AddValue, Key, Before, After
    Set Add = AddValue
End Function

Public Function Remove(Index As Variant)
    On Error Resume Next
    mCollection.Remove Index
End Function

'PROPERTY Let Value
'DESCRIPTION:  Sets the Value in a Key/Pair
Public Property Let Value(Optional Index As Variant, Val As Variant)
    
    Dim Temp As Variant
    Dim AddValue As VDictionaryTree
    
    On Error Resume Next
    If IsMissing(Index) = True Then
        Let mValue = Val
    Else
        CheckCollectionAndAllocate
        If IsObject(Val) = True Then
            Set Temp = Val
        Else
            Let Temp = Val
        End If
        Set AddValue = New VDictionaryTree
        AddValue.Key = Index
        AddValue.Value = Temp
        mCollection.Add AddValue, Index
        If Err.Number <> 0 Then
            mCollection.Remove Index
            mCollection.Add AddValue, Index
        End If
    End If
End Property

Public Property Set Value(Optional Index As Variant, Val As Variant)
    Dim AddValue As VDictionaryTree
    Dim Temp As Variant
    On Error Resume Next
    
    If IsMissing(Index) = True Then
        Set mValue = Val
    Else
        CheckCollectionAndAllocate
        If IsObject(Val) = True Then
            Set Temp = Val
        Else
            Let Temp = Val
        End If
        Set AddValue = New VDictionaryTree
        AddValue.Key = Index
        Set AddValue.Value = Temp
        mCollection.Add AddValue, Index
        If Err.Number <> 0 Then
            mCollection.Remove Index
            mCollection.Add AddValue, Index
        End If
    End If
End Property

Private Function CheckCollectionAndAllocate()
    
    Dim Count As Long
    
    On Error Resume Next
    Count = mCollection.Count
    
    If Err.Number <> 0 Then
        Set mCollection = New Collection
    End If
    
End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    CheckCollectionAndAllocate
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCollection.[_NewEnum]
End Property


Public Property Get Count() As Long
    On Error Resume Next
    Count = mCollection.Count
End Property

