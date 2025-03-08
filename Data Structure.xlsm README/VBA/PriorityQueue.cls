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

Private Sub Class_Initialize()
    Set this.Heap = New MaxHeap
End Sub

Private Sub Class_Terminate()
    Set this.Heap = Nothing
End Sub
