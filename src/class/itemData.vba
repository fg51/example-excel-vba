Option Explicit


Public name_ As ItemNames
Public leadTime_ As String

Public Sub initialize(ByVal aName As ItemNames, ByVal aLeadTime As String)
    Me.name_ = aName
    Me.leadTime_ = aLeadTime
End Sub

Property Get name() As ItemNames
    name = Me.name_
End Property

Property Get leadtime() As String
    leadtime = Me.leadTime_
End Property
