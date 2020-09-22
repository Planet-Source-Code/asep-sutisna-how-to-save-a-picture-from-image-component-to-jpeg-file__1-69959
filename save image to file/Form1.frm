VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Height          =   2655
      Left            =   120
      ScaleHeight     =   2595
      ScaleWidth      =   4275
      TabIndex        =   3
      Top             =   120
      Width           =   4335
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   120
         ScaleHeight     =   2295
         ScaleWidth      =   4005
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   120
         Width           =   4005
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Resize"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get Picture"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   240
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   6000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Code by Asep Sutisna
'Email: aspstudio@yahoo.com
'How to save a picture from image component without FSO or API

Private Sub Command1_Click()
'save picture in picture box as a jpeg file
SaveJPG Picture1.Image, App.Path & "\look_at_me.jpg"
'confirm that file is saved
MsgBox "Image saved as '" & App.Path & "\look_at_me.jpg'"
Command1.Enabled = False
Command3.Enabled = True
End Sub

Private Sub Command2_Click()
'get picture from image component with the same height and width as image component
Picture1.PaintPicture Image1.Picture, 0, 0, Image1.Width, Image1.Height
Picture1.Width = Image1.Width
Picture1.Height = Image1.Height
'set picturebox component
Set Picture1.Picture = Picture1.Image
'enable save button
Command2.Enabled = False
Command1.Enabled = True
End Sub

Private Sub Command3_Click()
If Image1.Picture.Width > Image1.Picture.Height Then
        Image1.Width = Picture1.Height
        Image1.Height = Image1.Picture.Height / Image1.Picture.Width * Image1.Width
Else
        Image1.Height = Picture1.Height
        Image1.Width = Image1.Picture.Width / Image1.Picture.Height * Image1.Height
End If
Command3.Enabled = False
Command2.Enabled = True
End Sub

Private Sub Form_Load()
'make sure that picturebox component is autoredraw
Picture1.AutoRedraw = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "Code by Asep Sutisna" & vbCrLf & _
"Email: aspstudio@yahoo.com" & vbCrLf & vbCrLf & _
"The picture is property of www.swishriders.com" & vbCrLf & vbCrLf & _
"How to save a picture from image component without FSO or API"
End Sub

