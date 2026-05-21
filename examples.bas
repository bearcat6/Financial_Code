Sub ExampleSplineCurve()

  Dim x(1 To 5) As Double
  Dim y(1 To 5) As Double
  Dim cSpline As clsSplineCurve

  x(1) = 0#
  x(2) = 1#
  x(3) = 2#
  x(4) = 3#
  x(5) = 5#

  y(1) = 0.001
  y(2) = 0.003
  y(3) = 0.006
  y(4) = 0.01
  y(5) = 0.015

  Set cSpline = New clsSplineCurve
  cSpline.AllowFlatExtrapolation = False

  If Not cSpline.Build(5, x, y) Then
    MsgBox cSpline.LastError
    Exit Sub
  End If

  Debug.Print cSpline.Value(2.5)

End Sub
