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
Private this As TMaxHeap

Public Property Get BufferSize() As Long
    BufferSize = this.BufferSize
End Property

Public Property Let BufferSize(ByVal RHS As Long)
    
    If this.ItemCount > 0 Then
        Err.Raise 13, "Heap is not empty."
    End If
    
    this.BufferSize = RHS
    
    Erase this.Container
    ReDim this.Container(this.InsertionRowIndex To this.PriorityRowIndex, 1 To this.BufferSize)
    
End Property

Public Property Get AreValueAndPrioritySame() As Boolean
    AreValueAndPrioritySame = this.AreValueAndPrioritySame
End Property

Public Property Let AreValueAndPrioritySame(ByVal RHS As Boolean)
    
    If this.ItemCount > 0 Then
        Err.Raise 13, "Heap is not empty."
    End If
    
    this.AreValueAndPrioritySame = RHS
    
    If this.AreValueAndPrioritySame Then
        this.PriorityRowIndex = 2
    Else
        this.PriorityRowIndex = 3
    End If
    
    Erase this.Container
    ReDim this.Container(this.InsertionRowIndex To this.PriorityRowIndex, 1 To this.BufferSize)
    
End Property

Private Sub Class_Initialize()
    
    ' Default behaviour is that value and priority are different.
    this.ItemCount = 0
    this.InsertionRowIndex = 1
    this.ValueRowIndex = 2
    this.PriorityRowIndex = 3
    this.AreValueAndPrioritySame = False
    this.BufferSize = BUFFER_SIZE
    
    ReDim this.Container(this.InsertionRowIndex To this.PriorityRowIndex, 1 To this.BufferSize)
    
End Sub

Private Sub Class_Terminate()
    
    ' Reset to default.
    this.ItemCount = 0
    this.InsertionRowIndex = -1
    this.ValueRowIndex = -1
    this.PriorityRowIndex = -1
    this.AreValueAndPrioritySame = False
    this.BufferSize = BUFFER_SIZE
    Erase this.Container
    
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
    
    this.ItemCount = UBound(ItemVector) - LBound(ItemVector) + 1
    ReDim this.Container(this.InsertionRowIndex To this.PriorityRowIndex, 1 To this.ItemCount)
    
    If Not this.AreValueAndPrioritySame And IsMissing(PriorityVector) And Not IsObject(ItemVector(LBound(ItemVector))) Then
        PriorityVector = ItemVector
    End If
    
    Dim Counter As Long
    For Counter = this.ItemCount To 1 Step -1
        
        this.Container(this.InsertionRowIndex, Counter) = Counter
        this.Container(this.ValueRowIndex, Counter) = ItemVector(LBound(ItemVector) + Counter - 1)
        If Not this.AreValueAndPrioritySame Then
            this.Container(this.PriorityRowIndex, Counter) = PriorityVector(LBound(PriorityVector) + Counter - 1)
        End If
        ShiftDown Counter
        
    Next Counter
    
End Sub

Private Property Get IHeap_IsEmpty() As Boolean
    IHeap_IsEmpty = (this.ItemCount = 0)
End Property

Private Property Get IHeap_Count() As Long
    IHeap_Count = this.ItemCount
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
    TotalItems = this.ItemCount
    
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
    
    If this.ItemCount = 0 Then Exit Property
    
    Dim Result As Variant
    ReDim Result(1 To this.ItemCount, 1 To 1)
    
    Dim Counter As Long
    For Counter = 1 To this.ItemCount
        If IsObject(this.Container(this.ValueRowIndex, Counter)) Then
            Set Result(Counter, 1) = this.Container(this.ValueRowIndex, Counter)
        Else
            Result(Counter, 1) = this.Container(this.ValueRowIndex, Counter)
        End If
    Next Counter
    
    IHeap_Values = Result
    
End Property

Private Property Get IHeap_Priorities() As Variant
    
    ' This will return the Current order of the heap priority.
    
    If this.ItemCount = 0 Then Exit Property
    
    Dim Result As Variant
    ReDim Result(1 To this.ItemCount, 1 To 1)
    
    Dim Counter As Long
    For Counter = 1 To this.ItemCount
        If IsObject(this.Container(this.PriorityRowIndex, Counter)) Then
            Set Result(Counter, 1) = this.Container(this.PriorityRowIndex, Counter)
        Else
            Result(Counter, 1) = this.Container(this.PriorityRowIndex, Counter)
        End If
    Next Counter
    
    IHeap_Priorities = Result
    
End Property

Private Sub IHeap_Add(ByVal Value As Variant, Optional ByVal Priority As Variant)
    
    ' It will add item on the end of the array. And then ShiftUp for MaxHeap constraints.
    ' Time complexity is : O(logn)
    ' It will keep checking each parent node and if parent node has low priority then move down to that children
    ' node position.
    
    this.ItemCount = this.ItemCount + 1
    
    ExtendContainerSizeIfNeeded
    
    this.Container(this.InsertionRowIndex, this.ItemCount) = this.ItemCount
    
    If IsObject(Value) Then
        Set this.Container(this.ValueRowIndex, this.ItemCount) = Value
    Else
        this.Container(this.ValueRowIndex, this.ItemCount) = Value
    End If
    
    If Not this.AreValueAndPrioritySame Then
        If IsMissing(Priority) And Not IsObject(Value) Then
            this.Container(this.PriorityRowIndex, this.ItemCount) = Value
        Else
            this.Container(this.PriorityRowIndex, this.ItemCount) = Priority
        End If
    End If
    
    ' No need to check for Shiftup
    If this.ItemCount = 1 Then Exit Sub
    
    ShiftUp this.ItemCount, Int(this.ItemCount / 2)
    
End Sub

Private Sub ShiftUp(ByVal ChildIndex As Long, ByVal ParentIndex As Long)
    
    ' Keep shifting up while we don't reach root or we don't need to shift up any further.
    ' Time complexity: O(logn)
    
    Do While ParentIndex > 0
        
        Dim ParentPriority As Variant
        ParentPriority = this.Container(this.PriorityRowIndex, ParentIndex)
        
        Dim ChildPriority As Variant
        ChildPriority = this.Container(this.PriorityRowIndex, ChildIndex)
        
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
    FirstOnePriority = this.Container(this.PriorityRowIndex, FirstOneIndex)
    
    Dim SecondOnePriority As Variant
    SecondOnePriority = this.Container(this.PriorityRowIndex, SecondOneIndex)
    
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
        IsHighPriority = (this.Container(this.InsertionRowIndex, FirstOneIndex) > this.Container(this.InsertionRowIndex, SecondOneIndex))
    End If
    
    IsSecondOneHasHighPriority = IsHighPriority
    
End Function

Private Sub Swap(ByVal ParentIndex As Long, ByVal ChildIndex As Long)
    
    Dim RowIndex As Long
    For RowIndex = LBound(this.Container, 1) To UBound(this.Container, 1)
        
        Dim Temp As Variant
        If IsObject(this.Container(RowIndex, ParentIndex)) Then
        
            Set Temp = this.Container(RowIndex, ParentIndex)
            Set this.Container(RowIndex, ParentIndex) = this.Container(RowIndex, ChildIndex)
            Set this.Container(RowIndex, ChildIndex) = Temp
    
        Else
    
            Temp = this.Container(RowIndex, ParentIndex)
            this.Container(RowIndex, ParentIndex) = this.Container(RowIndex, ChildIndex)
            this.Container(RowIndex, ChildIndex) = Temp
    
        End If
        
    Next RowIndex
    
End Sub

'@Recursive method.
Private Sub ShiftDown(ByVal ParentIndex As Long)
    
    ' Start with parent and check if left or right has high priority. Swap parent with that item
    ' index and then consider that child as parent index again. Keep doing this for all children node.
    ' Worst time complexity is: O(logn)
    
    If ParentIndex * 2 > this.ItemCount Then Exit Sub
    
    Dim HighPriorityIndex As Long
    HighPriorityIndex = ParentIndex
    
    Dim LeftChildIndex As Long
    LeftChildIndex = 2 * ParentIndex
    
    Dim RightChildIndex As Long
    RightChildIndex = LeftChildIndex + 1
    
    ' Initially HighPriorityIndex is at the parent. Check if Left child has high priority or not
    If LeftChildIndex <= this.ItemCount Then
        
        If IsSecondOneHasHighPriority(HighPriorityIndex, LeftChildIndex) Then
            HighPriorityIndex = LeftChildIndex
        End If
        
    End If
    
    ' If left node has high priority then HighPriorityIndex would have been updated with LeftChildIndex.
    ' Now compare LeftChildIndex with Right one or with parent with right one.
    If RightChildIndex <= this.ItemCount Then
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
    
    If this.ItemCount > UBound(this.Container, 2) - LBound(this.Container, 2) + 1 Then
        ReDim Preserve this.Container(this.InsertionRowIndex To this.PriorityRowIndex, 1 To this.ItemCount - 1 + this.BufferSize)
    End If
    
End Sub

Private Function IHeap_Pop() As Variant
    
    ' This will remove the largest priority item and then rearrange again.
    ' Time complexity: O(logn)
    
    If this.ItemCount <= 0 Then
        Err.Raise 91, "Heap is empty."
    End If
    
    If IsObject(this.Container(this.ValueRowIndex, 1)) Then
        Set IHeap_Pop = this.Container(this.ValueRowIndex, 1)
    Else
        IHeap_Pop = this.Container(this.ValueRowIndex, 1)
    End If
    
    ' Swap with last item.
    Swap 1, this.ItemCount
    
    ' Reduce item count. We have the top element at the backup section.
    this.ItemCount = this.ItemCount - 1
    
    ' Start shifting down from the root node.
    ShiftDown 1
    
    
End Function

Private Function IHeap_Top() As Variant
    
    ' This will just show you the top element. It won't remove the item from the top.
    If this.ItemCount <= 0 Then
        Err.Raise 91, "Heap is empty."
    End If
    
    If IsObject(this.Container(this.ValueRowIndex, 1)) Then
        Set IHeap_Top = this.Container(this.ValueRowIndex, 1)
    Else
        IHeap_Top = this.Container(this.ValueRowIndex, 1)
    End If
    
End Function


