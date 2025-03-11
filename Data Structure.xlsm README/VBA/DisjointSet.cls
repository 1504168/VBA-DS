VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DisjointSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

'@ModuleDescription: DisjointSet data structure(ds) can be used in multiple problem.
'                    This is the most common data structure in graph.
'                    If we need to find if two node is in the same component or not then
'                    We can use this ds. It's also useful to find the group.
'                    Minimum Spanning Tree(MST) also use this ds. Kruskal algorithm use this ds
'                    for finding MST.
'                    We have mainly two methods. FindUltimateParent and Union.
'                    Time complexity for both methods is O(4*alpha) = O(1)
'                    See details explanation and Implementation: https://youtu.be/aBxjDBC4M1U?si=1hvP6Z5qjSqlNhAR


'@TODO: Delete Size if not needed.

Private Type TDisjointSet
    Rank() As Long
    Size() As Long
    Parent() As Long
    NodeCount As Long
    StartNodeIdx As Long
    NumberOfSet As Long
    IgnoredUnionOptCount As Long
End Type

Private This As TDisjointSet

Public Property Get IgnoredUnionOptCount() As Long
    
    ' We can use this to see how many connection or edge we can remove by still keeping the component connected.
    ' See real Application here: https://youtu.be/FYrl7iz9_ZU?si=2HVp3HolTElU-5wm
    IgnoredUnionOptCount = This.IgnoredUnionOptCount
    
End Property

Public Property Get NumberOfSet() As Long
    
    ' The real example of this is how many component we have in a group.
    ' Each set is one component.
    Dim Count As Long
    Dim Idx As Long
    For Idx = LBound(This.Rank) To UBound(This.Rank)
        If This.Rank(Idx) > 0 And This.Parent(Idx) = Idx Then
            Count = Count + 1
        End If
    Next Idx
    
    NumberOfSet = Count
    
End Property

Public Property Get StartNodeIdx() As Long
    StartNodeIdx = This.StartNodeIdx
End Property

Public Property Let StartNodeIdx(ByVal RHS As Long)
    
    This.StartNodeIdx = RHS
    If This.NodeCount > 0 Then InitializeContainers
    
End Property

Public Property Get Rank() As Variant
    Rank = This.Rank
End Property

Public Property Get Size() As Variant
    Size = This.Size
End Property

Public Property Get Parent() As Variant
    Parent = This.Parent
End Property

Public Property Get NodeCount() As Long
    NodeCount = This.NodeCount
End Property

Public Property Let NodeCount(ByVal RHS As Long)
    
    This.NodeCount = RHS
    If This.StartNodeIdx <> -1 Then InitializeContainers
    
End Property

Public Property Get Self() As DisjointSet
    Set Self = Me
End Property

Public Function Create(ByVal NodeCount As Long _
                       , ByVal StartNodeIdx As Long) As DisjointSet
                       
    ' Constructor. Initialize with Number of Node and Start node number.
    ' This will initialize all the necessary node.
    
    With New DisjointSet
        .NodeCount = NodeCount
        .StartNodeIdx = StartNodeIdx
        Set Create = .Self
    End With

End Function

Private Sub InitializeContainers()
    
    Dim EndIndex As Long
    EndIndex = This.NodeCount + This.StartNodeIdx - 1
    
    ReDim This.Parent(This.StartNodeIdx To EndIndex)
    ReDim This.Rank(This.StartNodeIdx To EndIndex)
    ReDim This.Size(This.StartNodeIdx To EndIndex)
    
    Dim Idx As Long
    For Idx = This.StartNodeIdx To EndIndex
        This.Parent(Idx) = Idx
        This.Rank(Idx) = 0
        This.Size(Idx) = 1
    Next Idx
    
End Sub


Public Function FindUltimateParent(ByVal NodeIdx As Long) As Long
    
    '@Recursive
    
    ' Find the top most parent. Below one is the base case.
    If This.Parent(NodeIdx) = NodeIdx Then
        FindUltimateParent = NodeIdx
        Exit Function
    End If
    
    Dim UltimateParentIdx As Long
    ' Call recursively.
    UltimateParentIdx = FindUltimateParent(This.Parent(NodeIdx))
    ' Path compression.
    This.Parent(NodeIdx) = UltimateParentIdx
    
    FindUltimateParent = UltimateParentIdx
    
End Function

Public Sub Union(ByVal NodeIdx1 As Long, ByVal NodeIdx2 As Long)
    
    Dim UltNode1 As Long
    UltNode1 = FindUltimateParent(NodeIdx1)
    
    Dim UltNode2 As Long
    UltNode2 = FindUltimateParent(NodeIdx2)
    
    If UltNode1 = UltNode2 Then
        ' It's already connected. So ignore it.
        This.IgnoredUnionOptCount = This.IgnoredUnionOptCount + 1
        Exit Sub
    End If
    
    If This.Rank(UltNode1) < This.Rank(UltNode2) Then
        ' Connect smaller rank to higher rank
        This.Parent(UltNode1) = UltNode2
        This.Size(UltNode2) = This.Size(UltNode2) + This.Size(UltNode1)
        
    ElseIf This.Rank(UltNode1) > This.Rank(UltNode2) Then
        This.Parent(UltNode2) = UltNode1
        This.Size(UltNode1) = This.Size(UltNode1) + This.Size(UltNode2)
    Else
        ' Both of them have same rank. Connect any one of them.
        ' Connect UltNode2 with UltNode1
        This.Parent(UltNode2) = UltNode1
        ' Increase the UltNode1 by 1
        This.Rank(UltNode1) = This.Rank(UltNode1) + 1
        This.Size(UltNode1) = This.Size(UltNode1) + This.Size(UltNode2)
        
    End If
    
End Sub

Public Function IsInSameSet(ByVal NodeIdx1 As Long, ByVal NodeIdx2 As Long) As Boolean
    IsInSameSet = (FindUltimateParent(NodeIdx1) = FindUltimateParent(NodeIdx2))
End Function

Public Function GetSetsDetails(Optional ByVal SetNodesDelimiter As String = vbNullString) As Variant
    
    ' This will return all the different set info.
    ' First row will have the header. It is a 4 columns 2D array.
    ' SetNodesDelimiter is provided then it will concat the node index with that delimiter.
    ' If not provided then it will keep them in a Collection
    
    Dim Map As Collection
    Set Map = New Collection
    
    Dim Idx As Long
    On Error Resume Next
    For Idx = LBound(This.Rank) To UBound(This.Rank)
        
        ' Do path compression if not done yet.
        Dim Ultp As Long
        Ultp = FindUltimateParent(Idx)
        
        ' if the key already exists it will skip that due to On Error Resume Next
        Map.Add New Collection, CStr(Ultp)
        
        Dim Temp As Collection
        Set Temp = Map.Item(CStr(Ultp))
        
        If Temp.Count = 0 Then
            Temp.Add Ultp
            Temp.Add This.Rank(Ultp)
            Temp.Add This.Size(Ultp)
            Temp.Add Idx
        Else
            Temp.Add Idx
        End If
        
    Next Idx
    
    Dim Result As Variant
    ReDim Result(1 To Map.Count + 1, 1 To 4)
    Result(1, 1) = "Parent Node Idx"
    Result(1, 2) = "Parent Node Rank"
    Result(1, 3) = "Total Item In Set"
    Result(1, 4) = "Set Items Index"
    
    For Idx = 1 To Map.Count
        
        Set Temp = Map.Item(Idx)
        Result(Idx + 1, 1) = Temp.Item(1)
        Result(Idx + 1, 2) = Temp.Item(2)
        Result(Idx + 1, 3) = Temp.Item(3)
        
        RemoveFirstNItemFromCollection Temp, 3
        
        If SetNodesDelimiter = vbNullString Then
            Set Result(Idx + 1, 4) = Temp
        Else
            Result(Idx + 1, 4) = ConcatenateCollection(Temp, SetNodesDelimiter)
        End If
        
    Next Idx
    
    GetSetsDetails = Result
    
End Function

Private Sub Class_Initialize()
    
    This.StartNodeIdx = -1
    This.NodeCount = 0
    This.IgnoredUnionOptCount = 0
    
End Sub


'========================
' Helper Function and Sub
'========================

Private Function ConcatenateCollection(ByVal GivenCollection As Collection _
                                      , Optional ByVal Delimiter As String = ",") As String

    '@Description("This will concatenate all the item of a Collection if not an object Collection")
    '@Dependency("No Dependency")
    '@ExampleCall : ConcatenateCollection(ValidNameColl)
    '@Date : 21 January 2023 04:07:21 PM
    '@PossibleError:

    Dim Result As String
    Dim CurrentItem As Variant
    For Each CurrentItem In GivenCollection
        Result = Result & CStr(CurrentItem) & Delimiter
    Next CurrentItem

    If Result = vbNullString Then
        ConcatenateCollection = vbNullString
    Else
        ConcatenateCollection = Left$(Result, Len(Result) - Len(Delimiter))
    End If

End Function


Private Sub RemoveFirstNItemFromCollection(ByRef FromColl As Collection, ByVal N As Long)
    
    If N <= 0 Then Exit Sub
    If N > FromColl.Count Then N = FromColl.Count
    
    Dim Counter As Long
    For Counter = N To 1 Step -1
        FromColl.Remove Counter
    Next Counter
    
End Sub

