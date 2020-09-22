VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFeatures"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public Enum SizeEditConst
    sed_Automatic = 0
    sed_Manual = 1
End Enum

Public Enum SizeNodeConst
    nod_NW = 0
    nod_N = 1
    nod_NE = 2
    nod_E = 3
    nod_SE = 4
    nod_S = 5
    nod_SW = 6
    nod_W = 7
End Enum

Public Enum BackStyleConst
    bsc_Opaque = 1
    bsc_Transparent = 0
End Enum

Public Enum BorderConst
    brd_None = 0
    brd_FixedSingle = 1
End Enum

Public Enum AppearanceConst
    apr_3D = 1
    apr_Flat = 0
End Enum

Public Enum VirtualBorderConst
    brd_Thin = 0
    brd_Large = 1
End Enum

