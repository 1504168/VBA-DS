Attribute VB_Name = "TestDisjointSet"
Option Explicit

Private Sub TestUnionByRank()
    
    Dim ds As DisjointSet
    Set ds = DisjointSet.Create(7, 1)
    
    With ds
        .Union 1, 2
        .Union 2, 3
        .Union 4, 5
        .Union 6, 7
        .Union 5, 6
'        Debug.Print .NumberOfSet
        Debug.Print "Is in same set: " & .IsInSameSet(3, 7)
        Dim V As Variant
        V = .GetSetsDetails(" | ")
        .Union 3, 7
        Debug.Print "Is in same set: " & .IsInSameSet(3, 7)
        Debug.Print .NumberOfSet
        V = .GetSetsDetails(" | ")
    End With
    
End Sub

