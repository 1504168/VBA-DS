VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MaxHeap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' MaxHeap is a Complete Binary Tree (CBT).
' Any node priority is largest among all its descendant nodes.
' In 1 based index if i is the node index then:
' Parent node will be in (i/2)~=lowest integer index.
' Left child will be in 2*i index.
' Right child will be in 2*i+1 index.
' First item will always have the highest priority. Pop will be only done on the first index.


Implements IHeap

Private Const BUFFER_SIZE As Long = 32

Private Type TMaxHeap
    ValueRowIndex As Long
    InsertionRowIndex As Long
    PriorityRowIndex As Long
    ItemCount As Long
    Container() As Variant
    AreValueAndPrioritySame As Boolean
    BufferSize As Long
End Type
Private This As TMaxHeap

Public Property Get BufferSize() As Long
    BufferSize = This.BufferSize
End Property

Public Property Let BufferSize(ByVal RHS As Long)
    
    If This.ItemCount > 0 Then
        Err.Raise 13, "Heap is not empty."
    End If
    
    This.BufferSize = RHS
    
    Erase This.Container
    ReDim This.Container(This.InsertionRowIndex To This.PriorityRowIndex, 1 To This.BufferSize)
    
End Property

Public Property Get AreValueAndPrioritySame() As Boolean
    AreValueAndPrioritySame = This.AreValueAndPrioritySame
End Property

Public Property Let AreValueAndPrioritySame(ByVal RHS As Boolean)
    
    If This.ItemCount > 0 Then
        Err.Raise 13, "Heap is not empty."
    End If
    
    This.AreValueAndPrioritySame = RHS
    
    If This.AreValueAndPrioritySame Then
        This.PriorityRowIndex = 2
    Else
        This.PriorityRowIndex = 3
    End If
    
    Erase This.Container
    ReDim This.Container(This.InsertionRowIndex To This.PriorityRowIndex, 1 To This.BufferSize)
    
End Property

Private Sub Class_Initialize()
    
    ' Default behaviour is that value and priority are different.
    This.ItemCount = 0
    This.InsertionRowIndex = 1
    This.ValueRowIndex = 2
    This.PriorityRowIndex = 3
    This.AreValueAndPrioritySame = False
    This.BufferSize = BUFFER_SIZE
    
    ReDim This.Container(This.InsertionRowIndex To This.PriorityRowIndex, 1 To This.BufferSize)
    
End Sub

Private Sub Class_Terminate()
    
    ' Reset to default.
    This.ItemCount = 0
    This.InsertionRowIndex = -1
    This.ValueRowIndex = -1
    This.PriorityRowIndex = -1
    This.AreValueAndPrioritySame = False
    This.BufferSize = BUFFER_SIZE
    Erase This.Container
    
End Sub

Private Property Let IHeap_AreValueAndPrioritySame(ByVal RHS As Boolean)
    Me.AreValueAndPrioritySame = RHS
End Property

Private Property Let IHeap_BufferSize(ByVal RHS As Long)
    Me.BufferSize = RHS
End Property

Private Sub IHeap_FromVector(ByVal ItemVector As Variant, Optional ByVal PriorityVector As Variant)
    
    ' This will create a MaxHeap structure in O(n) Time complexity. This process is called heapify.
    ' Ref: https://www.geeksforgeeks.org/building-heap-from-array.
    
    ' This may not keep the same order as you would get if you add one item from left -> right.
    ' for example if the initial ItemVector = [82,54,18,42,17,29,100,90,20,79]
    ' then this will return :   [100, 90, 82, 54, 79, 29, 18, 42, 20, 17]
    ' Where as if you add then: [100, 90, 82, 54, 79, 18, 29, 42, 20, 17]
    ' so index 6 and 7 has different order.
    
    This.ItemCount = UBound(ItemVector) - LBound(ItemVector) + 1
    ReDim This.Container(This.InsertionRowIndex To This.PriorityRowIndex, 1 To This.ItemCount)
    
    If Not This.AreValueAndPrioritySame And IsMissing(PriorityVector) And Not IsObject(ItemVector(LBound(ItemVector))) Then
        PriorityVector = ItemVector
    End If
    
    Dim Counter As Long
    For Counter = This.ItemCount To 1 Step -1
        
        This.Container(This.InsertionRowIndex, Counter) = Counter
        This.Container(This.ValueRowIndex, Counter) = ItemVector(LBound(ItemVector) + Counter - 1)
        If Not This.AreValueAndPrioritySame Then
            This.Container(This.PriorityRowIndex, Counter) = PriorityVector(LBound(PriorityVector) + Counter - 1)
        End If
        ShiftDown Counter
        
    Next Counter
    
End Sub

Private Property Get IHeap_IsEmpty() As Boolean
    IHeap_IsEmpty = (This.ItemCount = 0)
End Property

Private Property Get IHeap_Count() As Long
    IHeap_Count = This.ItemCount
End Property

Private Function IHeap_Sort(ByVal ItemsToSortVector As Variant _
                            , Optional ByVal PriorityVector As Variant _
                             , Optional ByVal IsAscending As Boolean = True) As Variant
    
    ' First add all the item to the heap using FromVector (Heapify)
    ' Time complexity of that is O(n)
    ' And then for each node remove from the heap. And one each removal it takes O(logn) for rearrangement.
    ' So total time complexity if O(n)+ O(n logn)
    
    
    IHeap_FromVector ItemsToSortVector, PriorityVector
    
    Dim TotalItems As Long
    TotalItems = This.ItemCount
    
    Dim Result As Variant
    ReDim Result(1 To TotalItems, 1 To 1)
    
    Dim Counter As Long
    For Counter = 1 To TotalItems
        
        Dim CurrentItem As Variant
        If IsObject(IHeap_Top) Then
            Set CurrentItem = IHeap_Pop
        Else
            CurrentItem = IHeap_Pop
        End If
        
        If IsAscending Then
            If IsObject(CurrentItem) Then
                Set Result(TotalItems - Counter + 1, 1) = CurrentItem
            Else
                Result(TotalItems - Counter + 1, 1) = CurrentItem
            End If
        Else
            If IsObject(CurrentItem) Then
                Set Result(Counter, 1) = CurrentItem
            Else
                Result(Counter, 1) = CurrentItem
            End If
        End If
        
    Next Counter
    
    IHeap_Sort = Result
    
End Function

Private Property Get IHeap_Values() As Variant
    
    ' This will return the Current order of the heap values.
    
    If This.ItemCount = 0 Then Exit Property
    
    Dim Result As Variant
    ReDim Result(1 To This.ItemCount, 1 To 1)
    
    Dim Counter As Long
    For Counter = 1 To This.ItemCount
        If IsObject(This.Container(This.ValueRowIndex, Counter)) Then
            Set Result(Counter, 1) = This.Container(This.ValueRowIndex, Counter)
        Else
            Result(Counter, 1) = This.Container(This.ValueRowIndex, Counter)
        End If
    Next Counter
    
    IHeap_Values = Result
    
End Property

Private Property Get IHeap_Priorities() As Variant
    
    ' This will return the Current order of the heap priority.
    
    If This.ItemCount = 0 Then Exit Property
    
    Dim Result As Variant
    ReDim Result(1 To This.ItemCount, 1 To 1)
    
    Dim Counter As Long
    For Counter = 1 To This.ItemCount
        If IsObject(This.Container(This.PriorityRowIndex, Counter)) Then
            Set Result(Counter, 1) = This.Container(This.PriorityRowIndex, Counter)
        Else
            Result(Counter, 1) = This.Container(This.PriorityRowIndex, Counter)
        End If
    Next Counter
    
    IHeap_Priorities = Result
    
End Property

Private Sub IHeap_Add(ByVal Value As Variant, Optional ByVal Priority As Variant)
    
    ' It will add item on the end of the array. And then ShiftUp for MaxHeap constraints.
    ' Time complexity is : O(logn)
    ' It will keep checking each parent node and if parent node has low priority then move down to that children
    ' node position.
    
    This.ItemCount = This.ItemCount + 1
    
    ExtendContainerSizeIfNeeded
    
    This.Container(This.InsertionRowIndex, This.ItemCount) = This.ItemCount
    
    If IsObject(Value) Then
        Set This.Container(This.ValueRowIndex, This.ItemCount) = Value
    Else
        This.Container(This.ValueRowIndex, This.ItemCount) = Value
    End If
    
    If Not This.AreValueAndPrioritySame Then
        If IsMissing(Priority) And Not IsObject(Value) Then
            This.Container(This.PriorityRowIndex, This.ItemCount) = Value
        Else
            This.Container(This.PriorityRowIndex, This.ItemCount) = Priority
        End If
    End If
    
    ' No need to check for Shiftup
    If This.ItemCount = 1 Then Exit Sub
    
    ShiftUp This.ItemCount, Int(This.ItemCount / 2)
    
End Sub

Private Sub ShiftUp(ByVal ChildIndex As Long, ByVal ParentIndex As Long)
    
    ' Keep shifting up while we don't reach root or we don't need to shift up any further.
    ' Time complexity: O(logn)
    
    Do While ParentIndex > 0
        
        Dim ParentPriority As Variant
        ParentPriority = This.Container(This.PriorityRowIndex, ParentIndex)
        
        Dim ChildPriority As Variant
        ChildPriority = This.Container(This.PriorityRowIndex, ChildIndex)
        
        If IsSecondOneHasHighPriority(ParentIndex, ChildIndex) Then
            Swap ParentIndex, ChildIndex
            ChildIndex = ParentIndex
            ParentIndex = Int(ChildIndex / 2)
        Else
            Exit Do
        End If
        
    Loop
    
End Sub

Private Function IsSecondOneHasHighPriority(ByVal FirstOneIndex, ByVal SecondOneIndex As Long) As Boolean
    
    Dim FirstOnePriority As Variant
    FirstOnePriority = This.Container(This.PriorityRowIndex, FirstOneIndex)
    
    Dim SecondOnePriority As Variant
    SecondOnePriority = This.Container(This.PriorityRowIndex, SecondOneIndex)
    
    Dim IsHighPriority As Boolean
    IsHighPriority = False
    Dim IsEqualPriority As Boolean
    IsEqualPriority = True
    
    If IsArray(FirstOnePriority) Then
            
        Dim Index As Long
        For Index = LBound(FirstOnePriority) To UBound(FirstOnePriority)
            If FirstOnePriority(Index) < SecondOnePriority(Index) Then
                IsHighPriority = True
                IsEqualPriority = False
                Exit For
            ElseIf FirstOnePriority(Index) > SecondOnePriority(Index) Then
                IsHighPriority = False
                IsEqualPriority = False
                Exit For
            End If
        Next Index
            
    Else
        IsHighPriority = (FirstOnePriority < SecondOnePriority)
        IsEqualPriority = (FirstOnePriority = SecondOnePriority)
    End If
    
    If IsEqualPriority Then
        IsHighPriority = (This.Container(This.InsertionRowIndex, FirstOneIndex) > This.Container(This.InsertionRowIndex, SecondOneIndex))
    End If
    
    IsSecondOneHasHighPriority = IsHighPriority
    
End Function

Private Sub Swap(ByVal ParentIndex As Long, ByVal ChildIndex As Long)
    
    Dim RowIndex As Long
    For RowIndex = LBound(This.Container, 1) To UBound(This.Container, 1)
        
        Dim Temp As Variant
        If IsObject(This.Container(RowIndex, ParentIndex)) Then
        
            Set Temp = This.Container(RowIndex, ParentIndex)
            Set This.Container(RowIndex, ParentIndex) = This.Container(RowIndex, ChildIndex)
            Set This.Container(RowIndex, ChildIndex) = Temp
    
        Else
    
            Temp = This.Container(RowIndex, ParentIndex)
            This.Container(RowIndex, ParentIndex) = This.Container(RowIndex, ChildIndex)
            This.Container(RowIndex, ChildIndex) = Temp
    
        End If
        
    Next RowIndex
    
End Sub

'@Recursive method.
Private Sub ShiftDown(ByVal ParentIndex As Long)
    
    ' Start with parent and check if left or right has high priority. Swap parent with that item
    ' index and then consider that child as parent index again. Keep doing this for all children node.
    ' Worst time complexity is: O(logn)
    
    If ParentIndex * 2 > This.ItemCount Then Exit Sub
    
    Dim HighPriorityIndex As Long
    HighPriorityIndex = ParentIndex
    
    Dim LeftChildIndex As Long
    LeftChildIndex = 2 * ParentIndex
    
    Dim RightChildIndex As Long
    RightChildIndex = LeftChildIndex + 1
    
    ' Initially HighPriorityIndex is at the parent. Check if Left child has high priority or not
    If LeftChildIndex <= This.ItemCount Then
        
        If IsSecondOneHasHighPriority(HighPriorityIndex, LeftChildIndex) Then
            HighPriorityIndex = LeftChildIndex
        End If
        
    End If
    
    ' If left node has high priority then HighPriorityIndex would have been updated with LeftChildIndex.
    ' Now compare LeftChildIndex with Right one or with parent with right one.
    If RightChildIndex <= This.ItemCount Then
        If IsSecondOneHasHighPriority(HighPriorityIndex, RightChildIndex) Then
            HighPriorityIndex = RightChildIndex
        End If
    End If
    
    If HighPriorityIndex <> ParentIndex Then
        ' Swap parent index with HighPriorityIndex and call recursively for the HighPriorityIndex one
        Swap ParentIndex, HighPriorityIndex
        ShiftDown HighPriorityIndex
    End If
    
End Sub

Private Sub ExtendContainerSizeIfNeeded()
    
    ' Make room for extra items.
    
    If This.ItemCount > UBound(This.Container, 2) - LBound(This.Container, 2) + 1 Then
        ReDim Preserve This.Container(This.InsertionRowIndex To This.PriorityRowIndex, 1 To This.ItemCount - 1 + This.BufferSize)
    End If
    
End Sub

Private Function IHeap_Pop() As Variant
    
    ' This will remove the largest priority item and then rearrange again.
    ' Time complexity: O(logn)
    
    If This.ItemCount <= 0 Then
        Err.Raise 91, "Heap is empty."
    End If
    
    If IsObject(This.Container(This.ValueRowIndex, 1)) Then
        Set IHeap_Pop = This.Container(This.ValueRowIndex, 1)
    Else
        IHeap_Pop = This.Container(This.ValueRowIndex, 1)
    End If
    
    ' Swap with last item.
    Swap 1, This.ItemCount
    
    ' Reduce item count. We have the top element at the backup section.
    This.ItemCount = This.ItemCount - 1
    
    ' Start shifting down from the root node.
    ShiftDown 1
    
    
End Function

Private Function IHeap_Top() As Variant
    
    ' This will just show you the top element. It won't remove the item from the top.
    If This.ItemCount <= 0 Then
        Err.Raise 91, "Heap is empty."
    End If
    
    If IsObject(This.Container(This.ValueRowIndex, 1)) Then
        Set IHeap_Top = This.Container(This.ValueRowIndex, 1)
    Else
        IHeap_Top = This.Container(This.ValueRowIndex, 1)
    End If
    
End Function


