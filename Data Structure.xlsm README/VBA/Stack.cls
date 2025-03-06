VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private StackCollection As Collection

Public Property Get IsEmpty() As Boolean
    
    '@Description: Returns True if the stack is empty, False otherwise
    '@Return: Boolean
    '@Remarks: Checks if collection is initialized and has any items
    
    IsEmpty = (StackCollection.Count = 0)

End Property

Public Property Get Count() As Long
    
    '@Description: Returns the number of items in the stack
    '@Return: Long
    '@Remarks: Returns 0 if collection is not initialized
    
    Count = StackCollection.Count

End Property

Public Sub Push(Item As Variant)
    
    '@Description: Adds an item to the top of the stack
    '@Parameter: Item - Can be any valid data type (string, number, array, object)
    '@Remarks: Handles both object and non-object data types
    
    StackCollection.Add Item
    
End Sub

Public Function Pop() As Variant
    
    '@Description: Removes and returns the top item from the stack
    '@Return: Variant - The top item in the stack
    '@Error: Raises error 91 if stack is empty

    If Me.IsEmpty Then
        Err.Raise 91, "Stack.Pop", "Cannot pop from empty stack."
    End If
    
    If IsObject(StackCollection.Item(StackCollection.Count)) Then
        Set Pop = StackCollection.Item(StackCollection.Count)
    Else
        Pop = StackCollection.Item(StackCollection.Count)
    End If

    StackCollection.Remove StackCollection.Count

End Function

Public Function Peek() As Variant
    
    '@Description: Returns the top item without removing it
    '@Return: Variant - The top item in the stack
    '@Error: Raises error 91 if stack is empty
    
    If Me.IsEmpty Then
        Err.Raise 91, "Stack.Peek", "Cannot peek empty stack."
    End If
    
    If IsObject(StackCollection.Item(StackCollection.Count)) Then
        Set Peek = StackCollection.Item(StackCollection.Count)
    Else
        Peek = StackCollection.Item(StackCollection.Count)
    End If

End Function

Public Sub Clear()
    '@Description: Removes all items from the stack
    Set StackCollection = New Collection
End Sub

Public Function ToArray() As Variant()
    
    '@Description: Converts stack items to a 2D array
    '@Return: Variant() - 2D array with items in first column
    '@Remarks: Returns empty array if stack is empty
    
    If Me.IsEmpty Then
        ToArray = Array()
        Exit Function
    End If

    Dim Result() As Variant
    Dim I As Long
    
    ReDim Result(1 To StackCollection.Count, 1 To 1)
    
    For I = 1 To StackCollection.Count
        If IsObject(StackCollection(I)) Then
            Set Result(I, 1) = StackCollection(I)
        Else
            Result(I, 1) = StackCollection(I)
        End If
    Next I
    
    ToArray = Result

End Function

Public Function ToVector() As Variant()

    '@Description: Converts stack items to a 1D array (vector)
    '@Return: Variant() - 1D array containing stack items
    '@Remarks: Returns empty array if stack is empty
    
    If Me.IsEmpty Then
        ToVector = Array()
        Exit Function
    End If
    
    Dim Result() As Variant
    Dim I As Long
    ReDim Result(1 To StackCollection.Count)
    
    For I = 1 To StackCollection.Count
        If IsObject(StackCollection(I)) Then
            Set Result(I) = StackCollection(I)
        Else
            Result(I) = StackCollection(I)
        End If
    Next I
    
    ToVector = Result

End Function

Private Sub Class_Initialize()
    Set StackCollection = New Collection
End Sub

Private Sub Class_Terminate()
    Set StackCollection = Nothing
End Sub
