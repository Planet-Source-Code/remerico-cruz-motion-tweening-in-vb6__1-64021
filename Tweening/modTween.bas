Attribute VB_Name = "modTween"
' This is where all the tweening magic happens.... :)
' Sorry if the code's a bit messy

' Portions of the code was ported from the C# motion class project
' http://www.codeproject.com/csharp/tweencs.asp

Dim counter As Long
Dim timeStart As Long
Dim timeDest As Long
Dim animType As String
        
Dim Arr_StartPos(1) As Long
Dim Arr_DestPos(1) As Long

Dim t As Double
Dim d As Double
Dim b As Double
Dim c As Double

Dim objHolder As Object
Dim objTimer As Timer


Sub StartTween(xTimer As Timer, xControl As Object, xDestXPos As Long, xDestYPos As Long, xAnimType As String, xTimeInterval As Long)

  counter = 0
  timeStart = counter
  timeDest = xTimeInterval
  animType = xAnimType

  xTimer.Interval = 1

  Set objHolder = xControl
  Set objTimer = xTimer

  Arr_StartPos(0) = objHolder.Left
  Arr_StartPos(1) = objHolder.Top
  Arr_DestPos(0) = xDestXPos
  Arr_DestPos(1) = xDestYPos
  
  objTimer.Enabled = False
  objTimer.Enabled = True

End Sub


Public Sub Timer_Tick()

  If objHolder.Left = Arr_DestPos(0) And objHolder.Top = Arr_DestPos(1) Then
    objTimer.Enabled = False
  
  Else
    objHolder.Left = tween(0)
    objHolder.Top = tween(1)
    counter = counter + 1

  End If

End Sub


Private Function tween(prop As Long) As Long

  t = counter - timeStart
  b = Arr_StartPos(prop)
  c = Arr_DestPos(prop) - Arr_StartPos(prop)
  d = timeDest - timeStart

  tween = getFormula(animType, t, b, d, c)

End Function


Private Function getFormula(xAnimType As String, xT As Double, xB As Double, xD As Double, xC As Double) As Long
  
  Dim xTra As Double

' adjust formula to selected algoritm from combobox

  Select Case xAnimType
    Case "linear"
      ' simple linear tweening - no easing
      getFormula = (xC * xT / xD + xB)

    Case "easeinquad"
      ' quadratic (t^2) easing in - accelerating from zero velocity
      xT = xT / xD
      getFormula = (xC * xT * xT + xB)
                        
    Case "easeoutquad"
      ' quadratic (t^2) easing out - decelerating to zero velocity
      xT = xT / xD
      getFormula = (-xC * xT * (xT - 2) + xB)
               
    Case "easeincubic"
      'cubic easing in - accelerating from zero velocity
      xT = xT / xD
      getFormula = (xC * (xT) * xT * xT + xB)
      
    Case "easeinquart"
      ' quartic easing in - accelerating from zero velocity
      xT = xT / xD
      getFormula = (xC * (xT) * xT * xT * xT + xB)

    Case Else
      getFormula = 0
      
  End Select
  
End Function
