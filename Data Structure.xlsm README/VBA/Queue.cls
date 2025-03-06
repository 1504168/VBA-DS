VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Queue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private QueueCollection As Collection

Public Property Get IsEmpty() As Boolean
    
    '@Description: Returns True if the queue is empty, False otherwise
    '@Return: Boolean
    '@Remarks: Checks if collection is initialized and has any items
    
    IsEmpty = (QueueCollection.Count = 0)

End Property

Public Property Get Count() As Long
    
    '@Description: Returns the number of items in the queue
    '@Return: Long
    '@Remarks: Returns 0 if collection is not initialized
    
    Count = QueueCollection.Count

End Property

Public Sub Push(Item As Variant)
    
    '@Description: Adds an item to the end of the queue
    '@Parameter: Item - Can be any valid data type (string, number, array, object)
    '@Remarks: Handles both object and non-object data types
    
    QueueCollection.Add Item
    
End Sub

Public Function Pop() As Variant
    
    '@Description: Removes and returns the first item from the queue
    '@Return: Variant - The first item in the queue
    '@Error: Raises error 91 if queue is empty

    If Me.IsEmpty Then
        Err.Raise 91, "Queue.Pop", "Cannot pop from empty queue."
    End If
    
    If IsObject(QueueCollection.Item(1)) Then
        Set Pop = QueueCollection.Item(1)
    Else
        Pop = QueueCollection.Item(1)
    End If

    QueueCollection.Remove 1

End Function

Public Function Peek() As Variant
    
    '@Description: Returns the first item without removing it
    '@Return: Variant - The first item in the queue
    '@Error: Raises error 91 if queue is empty
    
    If Me.IsEmpty Then
        Err.Raise 91, "Queue.Peek", "Cannot peek empty queue."
    End If
    
    If IsObject(QueueCollection.Item(1)) Then
        Set Peek = QueueCollection.Item(1)
    Else
        Peek = QueueCollection.Item(1)
    End If

End Function

'@Description: Removes all items from the queue
Public Sub Clear()
    Set QueueCollection = New Collection
End Sub

Public Function ToArray() As Variant()
    
    '@Description: Converts queue items to a 2D array
    '@Return: Variant() - 2D array with items in first column
    '@Remarks: Returns empty array if queue is empty
    
    If Me.IsEmpty Then
        ToArray = Array()
        Exit Function
    End If

    Dim Result() As Variant
    Dim I As Long
    
    ReDim Result(1 To QueueCollection.Count, 1 To 1)
    
    For I = 1 To QueueCollection.Count
        If IsObject(QueueCollection(I)) Then
            Set Result(I, 1) = QueueCollection(I)
        Else
            Result(I, 1) = QueueCollection(I)
        End If
    Next I
    
    ToArray = Result

End Function

Public Function ToVector() As Variant()

    '@Description: Converts queue items to a 1D array (vector)
    '@Return: Variant() - 1D array containing queue items
    '@Remarks: Returns empty array if queue is empty
    
    If Me.IsEmpty Then
        ToVector = Array()
        Exit Function
    End If
    
    Dim Result() As Variant
    Dim I As Long
    ReDim Result(1 To QueueCollection.Count)
    
    For I = 1 To QueueCollection.Count
        If IsObject(QueueCollection(I)) Then
            Set Result(I) = QueueCollection(I)
        Else
            Result(I) = QueueCollection(I)
        End If
    Next I
    
    ToVector = Result

End Function

Private Sub Class_Initialize()
    Set QueueCollection = New Collection
End Sub

Private Sub Class_Terminate()
    Set QueueCollection = Nothing
End Sub
