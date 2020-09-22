VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CoOrdinate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
Public Key As String
Private mvarX As Long
Private mvarY As Long
Public Property Let Y(ByVal vData As Long)
    mvarY = vData
End Property
Public Property Get Y() As Long
    Y = mvarY
End Property
Public Property Let X(ByVal vData As Long)
    mvarX = vData
End Property
Public Property Get X() As Long
    X = mvarX
End Property
