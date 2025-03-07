VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IHeap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Let BufferSize(ByVal RHS As Long)
    
End Property

Public Property Let AreValueAndPrioritySame(ByVal RHS As Boolean)

End Property

Public Property Get IsEmpty() As Boolean

End Property

Public Property Get Count() As Long

End Property

Public Property Get Values() As Variant

End Property

Public Property Get Priorities() As Variant

End Property

Public Sub Add(ByVal Value As Variant, Optional ByVal Priority As Variant)

End Sub

Public Function Pop() As Variant

End Function

Public Function Top() As Variant

End Function

Public Sub FromVector(ByVal ItemVector As Variant, Optional ByVal PriorityVector As Variant)

End Sub

Public Function Sort(ByVal ItemsToSortVector As Variant _
                     , Optional ByVal PriorityVector As Variant _
                      , Optional ByVal IsAscending As Boolean = True) As Variant

End Function


