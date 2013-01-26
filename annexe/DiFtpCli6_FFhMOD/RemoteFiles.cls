VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RemoteFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"File"
Attribute VB_Ext_KEY = "Member0" ,"File"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'variable locale pour la collection
Private mCol As Collection

Friend Function Add(FileName As String, FileSize As Long, FileDate As Date, FileType As FileTypes, Description As String, Optional sKey As String) As RemoteFile
    'créer un nouvel objet
    Dim objNewMember As RemoteFile
    Set objNewMember = New RemoteFile


    'définir les propriétés passées dans la méthode
    objNewMember.FileName = FileName
    objNewMember.FileSize = FileSize
    objNewMember.FileDate = FileDate
    objNewMember.FileType = FileType
    objNewMember.Description = Description
    If Len(sKey) = 0 Then
      objNewMember.Key = "K" & FileName
        mCol.Add objNewMember
    Else
        objNewMember.Key = sKey
        mCol.Add objNewMember, sKey
    End If

    'renvoyé l'objet créé
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As RemoteFile
Attribute Item.VB_UserMemId = 0
    'utilisée pour faire référence à un élément de la collection
    'vntIndexKey contient l'index ou la clé, elle est donc
    'déclarée comme un Variant
    'Syntaxe: Set toto = x.Item(xyz) oo Set toto = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'utilisée lors de l'extraction du nombre d'éléments dans
    'la collection. Syntaxe: Debug.Print x.Count
    Count = mCol.Count
End Property


Friend Sub Remove(vntIndexKey As Variant)
    'utilisée pour la suppression d'un élément de la collection
    'vntIndexKey contient l'index ou la clé, elle est donc
    'déclarée comme un Variant
    'Syntaxe: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'cette propriété vous permet d'énumérer cette
    'collection avec la syntaxe For...Each
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'crée la collection lors de la création de cette classe
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'détruit la collection lorsque cette classe est terminée
    Set mCol = Nothing
End Sub

