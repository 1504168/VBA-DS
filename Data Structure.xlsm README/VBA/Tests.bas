Attribute VB_Name = "Tests"
Option Explicit

Private Sub TestQueue()
    
    Dim myQueue As New Queue
    myQueue.Push "First"
    myQueue.Push 42

    Debug.Print myQueue.IsEmpty                  ' Prints: False
    Debug.Print myQueue.Count                    ' Prints: 2
    Debug.Print myQueue.Pop                      ' Prints: First and removes it

    Dim Arr As Variant
    Arr = myQueue.ToArray                        ' Returns 1x1 2D array
    Arr = myQueue.ToVector                       ' Returns 1D array
    
End Sub

Public Sub TestStack()
    
    Dim Stack As New Stack
    
    ' Test IsEmpty and Count
    Debug.Assert Stack.IsEmpty = True
    Debug.Assert Stack.Count = 0
    
    ' Test Push and Count
    Stack.Push "First"
    Stack.Push 42
    Stack.Push Range("A1")
    
    Debug.Assert Stack.IsEmpty = False
    Debug.Assert Stack.Count = 3
    
    ' Test Peek
    Debug.Assert TypeName(Stack.Peek) = "Range" ' Last item should be Range
    Debug.Assert Stack.Count = 3 ' Peek shouldn't change count
    
    ' Test Pop
    Dim RangeItem As Range
    Set RangeItem = Stack.Pop
    Debug.Assert TypeName(RangeItem) = "Range"
    Debug.Assert Stack.Count = 2
    
    Debug.Assert Stack.Pop = 42
    Debug.Assert Stack.Count = 1
    
    Debug.Assert Stack.Pop = "First"
    Debug.Assert Stack.Count = 0
    Debug.Assert Stack.IsEmpty = True
    
    ' Test error handling
    On Error Resume Next
    Stack.Pop ' Try to pop from empty stack
    Debug.Assert Err.Number = 91
    On Error GoTo 0
    
    ' Test Clear
    Stack.Push "A"
    Stack.Push "B"
    Stack.Clear
    Debug.Assert Stack.IsEmpty = True
    
    ' Test ToArray
    Stack.Push "First"
    Stack.Push "Second"
    Stack.Push "Third"
    
    Dim Arr As Variant
    Arr = Stack.ToArray
    Debug.Assert UBound(Arr, 1) = 3 ' 3 rows
    Debug.Assert UBound(Arr, 2) = 1 ' 1 column
    Debug.Assert Arr(1, 1) = "First"
    Debug.Assert Arr(3, 1) = "Third"
    
    ' Test ToVector
    Dim Vec As Variant
    Vec = Stack.ToVector
    Debug.Assert UBound(Vec) = 3
    Debug.Assert Vec(1) = "First"
    Debug.Assert Vec(3) = "Third"
    
    Debug.Print "All Stack tests passed!"
    
End Sub

Public Sub TestEmptyStackOperations()
    Dim Stack As New Stack
    
    ' Test ToArray with empty stack
    Dim Arr As Variant
    Arr = Stack.ToArray
    Debug.Assert IsArray(Arr)
    Debug.Assert UBound(Arr) = -1 ' Empty array
    
    ' Test ToVector with empty stack
    Dim Vec As Variant
    Vec = Stack.ToVector
    Debug.Assert IsArray(Vec)
    Debug.Assert UBound(Vec) = -1 ' Empty array
    
    ' Test Peek on empty stack
    On Error Resume Next
    Stack.Peek
    Debug.Assert Err.Number = 91
    On Error GoTo 0
    
    Debug.Print "All empty Stack tests passed!"
    
End Sub
