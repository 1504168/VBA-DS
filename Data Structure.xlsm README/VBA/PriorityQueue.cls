VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PriorityQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This class is just a wrapper class. It is delegating all the calls to
' Heap (Default is MaxHeap) class. You can specify the Heap class using Heap property.

Private Type TPriorityQueue
    Heap As IHeap
End Type

Private this As TPriorityQueue

Public Property Get Heap() As IHeap
    Set Heap = this.Heap
End Property

Public Property Set Heap(ByVal RHS As IHeap)
    Set this.Heap = RHS
End Property

Public Property Let BufferSize(ByVal RHS As Long)
    this.Heap.BufferSize = RHS
End Property

Public Property Let AreValueAndPrioritySame(ByVal RHS As Boolean)
    this.Heap.AreValueAndPrioritySame = RHS
End Property

Public Property Get IsEmpty() As Boolean
    IsEmpty = this.Heap.IsEmpty
End Property

Public Property Get Count() As Long
    Count = this.Heap.Count
End Property

Public Sub Add(ByVal Value As Variant, Optional ByVal Priority As Variant)
    this.Heap.Add Value, Priority
End Sub

Public Function Pop() As Variant
    Pop = this.Heap.Pop
End Function

Public Function Top() As Variant
    Top = this.Heap.Top
End Function

Public Sub FromVector(ByVal ItemVector As Variant, Optional ByVal PriorityVector As Variant)
    this.Heap.FromVector ItemVector, PriorityVector
End Sub

Public Function Sort(ByVal ItemsToSortVector As Variant _
                     , Optional ByVal PriorityVector As Variant _
                      , Optional ByVal IsAscending As Boolean = True) As Variant
    Sort = this.Heap.Sort(ItemsToSortVector, PriorityVector, IsAscending)
End Function


Public Function KLargestElements(ByVal ItemVector As Variant _
                                  , ByVal K As Long _
                                   , Optional ByVal PriorityVector As Variant) As Variant
    
    ' This function will return you the First K largest elements in a vector.
    ' It doesn't return the distinct one. Make sure to remove duplicate if you have duplicates, prior to run this.
    ' Optionally you can set each item priority as well.
    ' Time Complexity: O(n*logK)
    ' Space Complexity: O(K+1) ~= O(K)
    
    If K <= 0 Then
        Err.Raise 13, "K should be greater than zero."
    End If
    
    If Not IsArray(ItemVector) And K > 1 Then
        Err.Raise 13, "There isn't enough element in the vector."
    ElseIf IsArray(ItemVector) And K > UBound(ItemVector) - LBound(ItemVector) + 1 Then
        Err.Raise 13, "There isn't enough element in the vector."
    End If
    
    If K = 1 And Not IsArray(ItemVector) Then
        
        If IsObject(ItemVector) Then
            Set KLargestElements = Array(ItemVector)
        Else
            KLargestElements = Array(ItemVector)
        End If
        Exit Function
        
    End If
    
    Set this.Heap = New MinHeap
    this.Heap.BufferSize = K + 1
    
    Dim ValueAndPrioritySame As Boolean
    ValueAndPrioritySame = IsMissing(PriorityVector)
    
    this.Heap.AreValueAndPrioritySame = ValueAndPrioritySame
    
    Dim Index As Long
    For Index = LBound(ItemVector) To UBound(ItemVector)
        
        AddToHeap ItemVector, PriorityVector, ValueAndPrioritySame, Index
        
        If this.Heap.Count > K Then
            this.Heap.Pop
        End If
        
    Next Index
    
    Dim Result As Variant
    ReDim Result(1 To K)
    
    For Index = 1 To K
        
        If IsObject(this.Heap.Top) Then
            Set Result(K - Index + 1) = this.Heap.Pop
        Else
            Result(K - Index + 1) = this.Heap.Pop
        End If
        
    Next Index
    
    KLargestElements = Result
    
    Set this.Heap = Nothing

End Function

Public Function KthLargestElement(ByVal ItemVector As Variant _
                                  , ByVal K As Long _
                                   , Optional ByVal PriorityVector As Variant) As Variant
    
    ' This function will return you the Kth largest element in a vector.
    ' It doesn't return the distinct one. Make sure to remove duplicate if you have duplicates, prior to run this.
    ' Optionally you can set each item priority as well.
    ' Time Complexity: O(n*logK)
    ' Space Complexity: O(K+1) ~= O(K)
    ' Explanation Ref: https://www.geeksforgeeks.org/kth-largest-element-in-an-array/
    
    If K <= 0 Then
        Err.Raise 13, "K should be greater than zero."
    End If
    
    If Not IsArray(ItemVector) And K > 1 Then
        Err.Raise 13, "There isn't enough element in the vector."
    ElseIf IsArray(ItemVector) And K > UBound(ItemVector) - LBound(ItemVector) + 1 Then
        Err.Raise 13, "There isn't enough element in the vector."
    End If
    
    If K = 1 And Not IsArray(ItemVector) Then
        If IsObject(ItemVector) Then
            Set KthLargestElement = ItemVector
        Else
            KthLargestElement = ItemVector
        End If
        Exit Function
    End If
    
    Set this.Heap = New MinHeap
    this.Heap.BufferSize = K + 1
    
    Dim ValueAndPrioritySame As Boolean
    ValueAndPrioritySame = IsMissing(PriorityVector)
    
    this.Heap.AreValueAndPrioritySame = ValueAndPrioritySame
    
    Dim Index As Long
    For Index = LBound(ItemVector) To UBound(ItemVector)
        
        AddToHeap ItemVector, PriorityVector, ValueAndPrioritySame, Index
        
        If this.Heap.Count > K Then
            this.Heap.Pop
        End If
        
    Next Index
    
    If IsObject(this.Heap.Top) Then
       Set KthLargestElement = this.Heap.Top
    Else
        KthLargestElement = this.Heap.Top
    End If
    
    Set this.Heap = Nothing

End Function

Public Function KSmallestElements(ByVal ItemVector As Variant _
                                  , ByVal K As Long _
                                   , Optional ByVal PriorityVector As Variant) As Variant
    
    ' This function will return you the First K smallest elements in a vector.
    ' It doesn't return the distinct one. Make sure to remove duplicate if you have duplicates, prior to run this.
    ' Optionally you can set each item priority as well.
    ' Time Complexity: O(n*logK)
    ' Space Complexity: O(K+1) ~= O(K)
    
    If K <= 0 Then
        Err.Raise 13, "K should be greater than zero."
    End If
    
    If Not IsArray(ItemVector) And K > 1 Then
        Err.Raise 13, "There isn't enough element in the vector."
    ElseIf IsArray(ItemVector) And K > UBound(ItemVector) - LBound(ItemVector) + 1 Then
        Err.Raise 13, "There isn't enough element in the vector."
    End If
    
    If K = 1 And Not IsArray(ItemVector) Then
        
        If IsObject(ItemVector) Then
            Set KSmallestElements = Array(ItemVector)
        Else
            KSmallestElements = Array(ItemVector)
        End If
        Exit Function
        
    End If
    
    Set this.Heap = New MaxHeap
    this.Heap.BufferSize = K + 1
    
    Dim ValueAndPrioritySame As Boolean
    ValueAndPrioritySame = IsMissing(PriorityVector)
    
    this.Heap.AreValueAndPrioritySame = ValueAndPrioritySame
    
    Dim Index As Long
    For Index = LBound(ItemVector) To UBound(ItemVector)
        
        AddToHeap ItemVector, PriorityVector, ValueAndPrioritySame, Index
        
        If this.Heap.Count > K Then
            this.Heap.Pop
        End If
        
    Next Index
    
    Dim Result As Variant
    ReDim Result(1 To K)
    
    For Index = 1 To K
        
        If IsObject(this.Heap.Top) Then
            Set Result(K - Index + 1) = this.Heap.Pop
        Else
            Result(K - Index + 1) = this.Heap.Pop
        End If
        
    Next Index
    
    KSmallestElements = Result
    
    Set this.Heap = Nothing

End Function

Public Function KthSmallestElement(ByVal ItemVector As Variant _
                                  , ByVal K As Long _
                                   , Optional ByVal PriorityVector As Variant) As Variant
    
    ' This function will return you the Kth smallest element in a vector.
    ' It doesn't return the distinct one. Make sure to remove duplicate if you have duplicates, prior to run this.
    ' Optionally you can set each item priority as well.
    ' Time Complexity: O(n*logK)
    ' Space Complexity: O(K+1) ~= O(K)
    ' Explanation Ref: https://www.geeksforgeeks.org/kth-smallest-largest-element-in-unsorted-array/
    
    If K <= 0 Then
        Err.Raise 13, "K should be greater than zero."
    End If
    
    If Not IsArray(ItemVector) And K > 1 Then
        Err.Raise 13, "There isn't enough element in the vector."
    ElseIf IsArray(ItemVector) And K > UBound(ItemVector) - LBound(ItemVector) + 1 Then
        Err.Raise 13, "There isn't enough element in the vector."
    End If
    
    If K = 1 And Not IsArray(ItemVector) Then
        
        If IsObject(ItemVector) Then
            Set KthSmallestElement = ItemVector
        Else
            KthSmallestElement = ItemVector
        End If
        Exit Function
        
    End If
    
    Set this.Heap = New MaxHeap
    this.Heap.BufferSize = K + 1
    
    Dim ValueAndPrioritySame As Boolean
    ValueAndPrioritySame = IsMissing(PriorityVector)
    
    this.Heap.AreValueAndPrioritySame = ValueAndPrioritySame
    
    Dim Index As Long
    For Index = LBound(ItemVector) To UBound(ItemVector)
        
        AddToHeap ItemVector, PriorityVector, ValueAndPrioritySame, Index
        
        If this.Heap.Count > K Then
            this.Heap.Pop
        End If
        
    Next Index
    
    If IsObject(this.Heap.Top) Then
        Set KthSmallestElement = this.Heap.Top
    Else
        KthSmallestElement = this.Heap.Top
    End If
    
    Set this.Heap = Nothing

End Function


Private Sub Class_Initialize()
    Set this.Heap = New MaxHeap
End Sub

Private Sub Class_Terminate()
    Set this.Heap = Nothing
End Sub


'============================================
' Helper function and sub
'============================================

Private Sub AddToHeap(ByVal ItemVector As Variant _
                      , PriorityVector As Variant _
                       , ValueAndPrioritySame As Boolean _
                        , ByVal Index As Long)
    
    If ValueAndPrioritySame Then
        this.Heap.Add ItemVector(Index)
    Else
        this.Heap.Add ItemVector(Index), PriorityVector(Index)
    End If
        
End Sub
