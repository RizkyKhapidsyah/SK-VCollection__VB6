VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VAgent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements VDictionaryTree

Dim mParent As VDictionaryTree
Dim mKey As String

Public Function SetParent(Parent As VDictionaryTree)
    Set mParent = Parent
End Function

Public Function Add(Value As Variant, Optional Key As Variant, Optional Before As Variant, Optional After As Variant) As Variant
    Dim subTree As VDictionaryTree
    Set subTree = mParent.Add(Null, mKey, Before, After)
    Set Add = subTree.Add(Value, Key, Before, After)
    
End Function

Public Property Get Key() As String
    Key = mKey
End Property

Friend Property Let Key(Key As String)
    Let mKey = Key
End Property

Public Property Get Value(Optional Index As Variant) As Variant
Attribute Value.VB_UserMemId = 0
    Dim Agent As VAgent
    Dim AddValue As VDictionaryTree
    
    If IsMissing(Index) = True Then
        Let Value = Null
    Else
        Set Agent = New VAgent
        Agent.Key = Index
        Agent.SetParent Me
        Set AddValue = Agent
        Set Value = Agent
    End If
End Property

Public Property Let Value(Optional Index As Variant, Val As Variant)
    Me.Add Val, Index
End Property

Public Property Set Value(Optional Index As Variant, Val As Variant)
    Me.Add Val, Index
End Property

Public Function Remove(Index As Variant)
    
End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Dim mCol As New Collection
    Set NewEnum = mCol.[_NewEnum]
End Property

Public Property Get Count() As Long
    Let Count = 0
End Property

Public Property Get Item(Index As Variant) As VDictionaryTree

End Property

Public Property Set Item(Index As Variant, subTree As VDictionaryTree)

End Property

Private Function VDictionaryTree_Add(Value As Variant, Optional Key As Variant, Optional Before As Variant, Optional After As Variant) As Variant
    Set VDictionaryTree_Add = Me.Add(Value, Key, Before, After)
End Function

Private Property Get VDictionaryTree_Count() As Long
    Let VDictionaryTree_Count = Me.Count
End Property

Private Property Set VDictionaryTree_Item(Index As Variant, RHS As VDictionaryTree)

End Property

Private Property Get VDictionaryTree_Item(Index As Variant) As VDictionaryTree

End Property

Private Property Get VDictionaryTree_Key() As String
    Let VDictionaryTree_Key = Me.Key
End Property

Friend Property Let VDictionaryTree_Key(Key As String)
    Let Me.Key = Key
End Property

Private Property Get VDictionaryTree_NewEnum() As stdole.IUnknown
    Set VDictionaryTree_NewEnum = Me.NewEnum
End Property

Private Function VDictionaryTree_Remove(Index As Variant) As Variant
    VDictionaryTree_Remove = Me.Remove(Index)
End Function

Private Property Set VDictionaryTree_Value(Optional Index As Variant, RHS As Variant)
    Set Me.Value(Index) = RHS
End Property

Private Property Let VDictionaryTree_Value(Optional Index As Variant, RHS As Variant)
    Let Me.Value(Index) = RHS
End Property

Private Property Get VDictionaryTree_Value(Optional Index As Variant) As Variant
    LetOrSet Me.Value(Index), VDictionaryTree_Value
End Property
