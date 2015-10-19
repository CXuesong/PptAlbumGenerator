Option Explicit

Public Sub PAG_PostProcess(p As Presentation)
Dim eachSlide As Slide
Dim eachShape As Shape
For Each eachSlide In p.Slides
    For Each eachShape In eachSlide.Shapes
        Dim animation As Effect
        Set animation = eachSlide.TimeLine.MainSequence.FindFirstAnimationFor(eachShape)
        If Not animation Is Nothing Then
            If animation.EffectType = msoAnimEffectCustom Then
                Dim duration As Single
                Dim path As String
                duration = animation.Timing.duration
                path = animation.Behaviors(2).MotionEffect.path
                animation.EffectType = msoAnimEffectPathRight
                animation.Behaviors(1).MotionEffect.path = path
                animation.Behaviors(1).Timing.duration = duration
            End If
        End If
    Next
Next
End Sub

Public Function PAG_MakePathAnimation(shp As Shape) As Effect
    Dim parentSlide As Slide
    Set parentSlide = shp.Parent
    Set PAG_MakePathAnimation = parentSlide.TimeLine.MainSequence.AddEffect(shp, _
                          effectId:= MsoAnimEffect.msoAnimEffectPathRight, _
                          trigger:= MsoAnimTriggerType.msoAnimTriggerWithPrevious)
End Function
