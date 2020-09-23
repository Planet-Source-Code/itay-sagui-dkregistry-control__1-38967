VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItem"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarhKey As Long 'local copy

'local variable(s) to hold property value(s)
Private mvarstrPath As String 'local copy
Private mvarstrValue As String 'local copy
Private mvarvData As Variant 'local copy

'local variable(s) to hold property value(s)
Private mvarstrType As String 'local copy

Public Property Let strType(ByVal strType As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strType = 5
    If InStr(1, DATA_TYPES, strType) = 0 And strType <> Space$(260) Then
        mvarstrType = "REG_SZ"
    Else
        mvarstrType = strType
    End If
End Property

Public Property Get strType() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.strType
    strType = mvarstrType
End Property

Public Property Let vData(ByVal vData As Variant)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.vData = 5
    mvarvData = vData
End Property

Public Property Get vData() As Variant
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.vData
    vData = mvarvData
End Property

Public Property Let strValue(ByVal New_strValue As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.strValue = 5
    mvarstrValue = New_strValue
End Property

Public Property Get strValue() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.strValue
    strValue = mvarstrValue
End Property

Public Property Let strPath(ByVal New_strPath As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.strPath = 5
    mvarstrPath = New_strPath
End Property

Public Property Get strPath() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.strPath
    strPath = mvarstrPath
End Property

Public Property Let hKey(ByVal New_hKey As Long)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.hKey = 5
    mvarhKey = New_hKey
End Property

Public Property Get hKey() As Long
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.hKey
    hKey = mvarhKey
End Property
