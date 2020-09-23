VERSION 5.00
Begin VB.Form FMain 
   Caption         =   " Example"
   ClientHeight    =   2340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   2160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cunregister 
      Caption         =   "Unregister"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CheckBox Cslient 
      Caption         =   "Slient"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Cregister 
      Caption         =   "Register"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton Cextract 
      Caption         =   "Extract"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'IMPORTANT: Make sure the VB Resource Editor is enabled if you wish to edit the resource file
'This is done by going
'Add-Ins > Add-In Manager > VB 6 Resource Editor
'Make sure the fist two check boxs are ticked and you will be able to acces it from your toolbar

'NOTE: The .OCX in the resource file is "Mswinsck.ocx" (Microsoft Winsock Control 6.0)

Dim SlientREG As Boolean

Private Sub Cextract_Click()

    ExtractResource "DLL", 101, SystemDirectory & "\Mswinsck.ocx"

End Sub

Private Sub Cregister_Click()

    Register SystemDirectory & "\Mswinsck.ocx", SlientREG

End Sub

Private Sub Cslient_Click()

    If Cslient.Value = 1 Then
        SlientREG = True
    Else
        SlientREG = False
    End If

End Sub

Private Sub Cunregister_Click()

    Unregister SystemDirectory & "\Mswinsck.ocx", SlientREG

End Sub
