VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AStarNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarY As Long
Private mvarX As Long
Private mvargval As Single
Private mvarhval As Single
Public Property Let hval(ByVal vData As Single)
    mvarhval = vData
End Property
Public Property Get hval() As Single
    hval = mvarhval
End Property
Public Property Let gval(ByVal vData As Single)
    mvargval = vData
End Property
Public Property Get gval() As Single
    gval = mvargval
End Property
Public Property Let X(ByVal vData As Long)
    mvarX = vData
End Property
Public Property Get X() As Long
    X = mvarX
End Property
Public Property Let Y(ByVal vData As Long)
    mvarY = vData
End Property
Public Property Get Y() As Long
    Y = mvarY
End Property
