VERSION 5.00
Object = "*\AdkRegistryProj.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin dkRegistryProj.dkRegistry dkRegistry1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Col As New colItems
Dim Item As New clsItem

Private Const HKEY_CLASSES_ROOT     As Long = &H80000000
Private Const HKEY_CURRENT_CONFIG   As Long = &H80000005
Private Const HKEY_CURRENT_USER     As Long = &H80000001
Private Const HKEY_DYN_DATA         As Long = &H80000006
Private Const HKEY_LOCAL_MACHINE    As Long = &H80000002
Private Const HKEY_PERF_ROOT        As Long = HKEY_LOCAL_MACHINE
Private Const HKEY_PERFORMANCE_DATA As Long = &H80000004
Private Const HKEY_USERS            As Long = &H80000003

Private Sub Form_Load()
    With Item
        .hKey = HKEY_CURRENT_USER
        .strPath = "Software\WinGlass"
    End With
    dkRegistry1.SaveKey Item, App.Path & "\Itay.reg"
End Sub
