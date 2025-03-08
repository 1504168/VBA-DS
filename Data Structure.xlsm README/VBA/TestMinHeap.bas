Attribute VB_Name = "TestMinHeap"
Option Explicit

Public Sub TestingMaxHeap()
    
    Dim Heap As IHeap
    Set Heap = New MinHeap
    
    With Heap
        .BufferSize = 12
        .AreValueAndPrioritySame = True
        .Add 45
        .Add 31
        .Add 20
        .Add 14
        .Add 7
        .Add 12
        .Add 18
        .Add 11
        .Add 7
        .Add 32
    End With
    
    Debug.Assert Heap.Top = 45
    Debug.Assert Heap.Count = 10
    
    Debug.Assert Heap.Pop = 45
    Debug.Assert Heap.Count = 9
    
    Dim V As Variant
    V = Application.WorksheetFunction.Transpose(Heap.Values)
    
End Sub


Public Sub TestingPop()
    
    Dim Heap As IHeap
    Set Heap = New MinHeap
    
    With Heap
        .Add "A", 45
        .Add "B", 31
        .Add "C", 14
        .Add "D", 20
        .Add "E", 7
        .Add "F", 18
        .Add "G", 11
        .Add "H", 7
        .Add "I", 32
        .Add "J", 7
        .Add "K", 7
    End With
    
    Do While Not Heap.IsEmpty
        Debug.Print Heap.Pop
    Loop
    
End Sub

Public Sub TestFromVector()
    
    Dim Heap As IHeap
    Set Heap = New MinHeap
    
'    Dim CurrentEl As Variant
'    For Each CurrentEl In Sheet2.Range("C5:L5").Value
'        Debug.Assert CurrentEl <> 79
'        Heap.Add CurrentEl, CurrentEl
'    Next CurrentEl

    Dim Res As Variant
    Res = Heap.Sort(ToVector(Sheet2.Range("C5:L5").Value), , True)
    ActiveCell.Resize(10, 1).Value = Res
    
End Sub

Public Function ToVector(Items As Variant) As Variant
    
    Dim TotalItems As Long
    TotalItems = UBound(Items, 1) * UBound(Items, 2)
    Dim Count As Long
    
    Dim Result As Variant
    ReDim Result(1 To TotalItems)
    
    Dim CurrentEl As Variant
    For Each CurrentEl In Items
        Count = Count + 1
        Result(Count) = CurrentEl
    Next CurrentEl
    
    ToVector = Result
    
End Function

