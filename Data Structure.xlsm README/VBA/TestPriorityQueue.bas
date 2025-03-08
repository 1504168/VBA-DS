Attribute VB_Name = "TestPriorityQueue"
Option Explicit

Public Sub TestKthLargestElement()
    
    Dim Arr As Variant
    Arr = ToVector(Sheet2.Range("C5:L5").Value)
    
    Dim PQ As PriorityQueue
    Set PQ = New PriorityQueue
    Dim Counter As Long
    For Counter = 1 To 10
        Debug.Print "Max: " & PQ.KthLargestElement(Arr, Counter), "Min: " & PQ.KthSmallestElement(Arr, Counter)
    Next Counter
    
    Dim FirstN As Variant
    FirstN = PQ.KSmallestElements(Arr, 3)
    
End Sub
