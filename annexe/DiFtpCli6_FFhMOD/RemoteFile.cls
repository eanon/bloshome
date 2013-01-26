VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RemoteFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Enum FileTypes
  FtpDirectory = 0
  FtpFile = 1
  FtpLink = 2
  FtpUnkwnon = 4
End Enum
'variables locales de stockage des valeurs de propriétés
Private mvarFileName As String 'copie locale
Private mvarSize As Long 'copie locale
Private mvarDate As Date 'copie locale
Private mvarType As Integer 'copie locale
Private mvarKey As String
Private mvarDescription As String

Public Property Let FileDate(ByVal vData As Date)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.Date = 5
    mvarDate = vData
End Property
Public Property Get FileDate() As Date
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.Date
    FileDate = mvarDate
End Property

Public Property Let FileType(ByVal vData As FileTypes)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.Type = 5
    mvarType = vData
End Property
Public Property Get FileType() As FileTypes
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.Type
    FileType = mvarType
End Property

Public Property Let FileSize(ByVal vData As Long)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.Size = 5
    mvarSize = vData
End Property
Public Property Get FileSize() As Long
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.Size
    FileSize = mvarSize
End Property

Public Property Let FileName(ByVal vData As String)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.FileName = 5
    mvarFileName = vData
End Property
Public Property Get FileName() As String
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.FileName
    FileName = mvarFileName
End Property
Public Property Let Key(ByVal vData As String)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.FileName = 5
    mvarKey = vData
End Property
Public Property Get Key() As String
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.FileName
    Key = mvarKey
End Property

Public Property Get Description() As String
  Description = mvarDescription
End Property

Public Property Let Description(ByVal vNewValue As String)
  mvarDescription = vNewValue
End Property
