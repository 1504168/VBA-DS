Attribute VB_Name = "modUtility"
Option Explicit

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
