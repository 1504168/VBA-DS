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

Private This As TPriorityQueue

Public Property Get Heap() As IHeap
    Set Heap = This.Heap
End Property

Public Property Set Heap(ByVal RHS As IHeap)
    Set This.Heap = RHS
End Property

Public Property Let BufferSize(ByVal RHS As Long)
    This.Heap.BufferSize = RHS
End Property

Public Property Let AreValueAndPrioritySame(ByVal RHS As Boolean)
    This.Heap.AreValueAndPrioritySame = RHS
End Property

Public Property Get IsEmpty() As Boolean
    IsEmpty = This.Heap.IsEmpty
End Property

Public Property Get Count() As Long
    Count = This.Heap.Count
End Property

Public Sub Add(ByVal Value As Variant, Optional ByVal Priority As Variant)
    This.Heap.Add Value, Priority
End Sub

Public Function Pop() As Variant
    Pop = This.Heap.Pop
End Function

Public Function Top() As Variant
    Top = This.Heap.Top
End Function

Public Sub FromVector(ByVal ItemVector As Variant, Optional ByVal PriorityVector As Variant)
    This.Heap.FromVector ItemVector, PriorityVector
End Sub

Public Function Sort(ByVal ItemsToSortVector As Variant _
                     , Optional ByVal PriorityVector As Variant _
                      , Optional ByVal IsAscending As Boolean = True) As Variant
    Sort = This.Heap.Sort(ItemsToSortVector, PriorityVector, IsAscending)
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
    
    Set This.Heap = New MinHeap
    This.Heap.BufferSize = K + 1
    
    Dim ValueAndPrioritySame As Boolean
    ValueAndPrioritySame = IsMissing(PriorityVector)
    
    This.Heap.AreValueAndPrioritySame = ValueAndPrioritySame
    
    Dim Index As Long
    For Index = LBound(ItemVector) To UBound(ItemVector)
        
        AddToHeap ItemVector, PriorityVector, ValueAndPrioritySame, Index
        
        If This.Heap.Count > K Then
            This.Heap.Pop
        End If
        
    Next Index
    
    Dim Result As Variant
    ReDim Result(1 To K)
    
    For Index = 1 To K
        
        If IsObject(This.Heap.Top) Then
            Set Result(K - Index + 1) = This.Heap.Pop
        Else
            Result(K - Index + 1) = This.Heap.Pop
        End If
        
    Next Index
    
    KLargestElements = Result
    
    Set This.Heap = Nothing

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
    
    Set This.Heap = New MinHeap
    This.Heap.BufferSize = K + 1
    
    Dim ValueAndPrioritySame As Boolean
    ValueAndPrioritySame = IsMissing(PriorityVector)
    
    This.Heap.AreValueAndPrioritySame = ValueAndPrioritySame
    
    Dim Index As Long
    For Index = LBound(ItemVector) To UBound(ItemVector)
        
        AddToHeap ItemVector, PriorityVector, ValueAndPrioritySame, Index
        
        If This.Heap.Count > K Then
            This.Heap.Pop
        End If
        
    Next Index
    
    If IsObject(This.Heap.Top) Then
       Set KthLargestElement = This.Heap.Top
    Else
        KthLargestElement = This.Heap.Top
    End If
    
    Set This.Heap = Nothing

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
    
    Set This.Heap = New MaxHeap
    This.Heap.BufferSize = K + 1
    
    Dim ValueAndPrioritySame As Boolean
    ValueAndPrioritySame = IsMissing(PriorityVector)
    
    This.Heap.AreValueAndPrioritySame = ValueAndPrioritySame
    
    Dim Index As Long
    For Index = LBound(ItemVector) To UBound(ItemVector)
        
        AddToHeap ItemVector, PriorityVector, ValueAndPrioritySame, Index
        
        If This.Heap.Count > K Then
            This.Heap.Pop
        End If
        
    Next Index
    
    Dim Result As Variant
    ReDim Result(1 To K)
    
    For Index = 1 To K
        
        If IsObject(This.Heap.Top) Then
            Set Result(K - Index + 1) = This.Heap.Pop
        Else
            Result(K - Index + 1) = This.Heap.Pop
        End If
        
    Next Index
    
    KSmallestElements = Result
    
    Set This.Heap = Nothing

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
    
    Set This.Heap = New MaxHeap
    This.Heap.BufferSize = K + 1
    
    Dim ValueAndPrioritySame As Boolean
    ValueAndPrioritySame = IsMissing(PriorityVector)
    
    This.Heap.AreValueAndPrioritySame = ValueAndPrioritySame
    
    Dim Index As Long
    For Index = LBound(ItemVector) To UBound(ItemVector)
        
        AddToHeap ItemVector, PriorityVector, ValueAndPrioritySame, Index
        
        If This.Heap.Count > K Then
            This.Heap.Pop
        End If
        
    Next Index
    
    If IsObject(This.Heap.Top) Then
        Set KthSmallestElement = This.Heap.Top
    Else
        KthSmallestElement = This.Heap.Top
    End If
    
    Set This.Heap = Nothing

End Function


Private Sub Class_Initialize()
    Set This.Heap = New MaxHeap
End Sub

Private Sub Class_Terminate()
    Set This.Heap = Nothing
End Sub


'============================================
' Helper function and sub
'============================================

Private Sub AddToHeap(ByVal ItemVector As Variant _
                      , PriorityVector As Variant _
                       , ValueAndPrioritySame As Boolean _
                        , ByVal Index As Long)
    
    If ValueAndPrioritySame Then
        This.Heap.Add ItemVector(Index)
    Else
        This.Heap.Add ItemVector(Index), PriorityVector(Index)
    End If
        
End Sub
