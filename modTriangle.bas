Attribute VB_Name = "modTriangle"
Option Explicit

Public Type Point
    X As Single
    Y As Single
End Type

Private Function Sign(P1 As Point, P2 As Point, P3 As Point)
    Sign = (P1.X - P3.X) * (P2.Y - P3.Y) - (P2.X - P3.X) * (P1.Y - P3.Y)
End Function

Public Function IsPointInTriangle(P As Point, V1 As Point, V2 As Point, V3 As Point) As Boolean
    Dim b1 As Boolean, b2 As Boolean, b3 As Boolean
    b1 = (Sign(P, V1, V2) < 0)
    b2 = (Sign(P, V2, V3) < 0)
    b3 = (Sign(P, V3, V1) < 0)
    IsPointInTriangle = ((b1 = b2) And (b2 = b3))
End Function

