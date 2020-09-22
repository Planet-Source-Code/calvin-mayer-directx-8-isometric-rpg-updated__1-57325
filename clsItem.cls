VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public m_Name As String 'name of item
Public m_Value As Long 'the selling price
Public m_Icon As New clsTexture 'the item's picture
Public m_Quantity As Integer 'how many of the item
Public m_RestockTime As Integer 'how long till 1 more item appears

Public Sub LoadItemFromDB(itemName As String, filepath As String)
    m_Name = itemName
End Sub

