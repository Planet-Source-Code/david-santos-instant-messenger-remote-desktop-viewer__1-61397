VERSION 5.00
Begin VB.Form frmViewInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User's Info"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   Icon            =   "frmViewInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Personal Information"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin VB.TextBox txtName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   6
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox txteMail 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   5
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtOther 
         Enabled         =   0   'False
         Height          =   1125
         Left            =   1080
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1080
         Width           =   2895
      End
      Begin VB.PictureBox picMyPic 
         Height          =   1815
         Left            =   4200
         ScaleHeight     =   1755
         ScaleWidth      =   1755
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   1
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Name"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&E-mail"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Other"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmViewInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

