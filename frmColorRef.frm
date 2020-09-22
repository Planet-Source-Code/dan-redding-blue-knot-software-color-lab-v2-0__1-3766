VERSION 5.00
Begin VB.Form frmColorRef 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Color Lab"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5160
   Icon            =   "frmColorRef.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   321
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   344
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picAnchor 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   1
      Left            =   4545
      ScaleHeight     =   525
      ScaleWidth      =   525
      TabIndex        =   82
      ToolTipText     =   "Click to make current color an anchor color for the blend bar"
      Top             =   4245
      Width           =   525
   End
   Begin VB.PictureBox picAnchor 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   0
      Left            =   4545
      ScaleHeight     =   525
      ScaleWidth      =   525
      TabIndex        =   81
      ToolTipText     =   "Click to make current color an anchor color for the blend bar"
      Top             =   45
      Width           =   525
   End
   Begin VB.VScrollBar vL 
      Height          =   375
      Left            =   3780
      Max             =   240
      SmallChange     =   2
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   4200
      Width           =   195
   End
   Begin VB.VScrollBar vS 
      Height          =   375
      Left            =   2640
      Max             =   240
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   4200
      Width           =   195
   End
   Begin VB.VScrollBar vH 
      Height          =   375
      Left            =   1500
      Max             =   239
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   4200
      Width           =   195
   End
   Begin VB.VScrollBar vB 
      Height          =   375
      Left            =   3780
      Max             =   255
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   3540
      Width           =   195
   End
   Begin VB.VScrollBar vG 
      Height          =   375
      Left            =   2640
      Max             =   255
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   3540
      Width           =   195
   End
   Begin VB.VScrollBar vR 
      Height          =   375
      Left            =   1500
      Max             =   255
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   3540
      Width           =   195
   End
   Begin VB.Timer tmr5by5 
      Left            =   2040
      Top             =   0
   End
   Begin VB.CommandButton cmd5by5 
      Height          =   375
      Left            =   900
      Picture         =   "frmColorRef.frx":548A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Take a 5x5 average sample from the screen"
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox pic5x5 
      Height          =   1485
      Left            =   2820
      ScaleHeight     =   1425
      ScaleWidth      =   1425
      TabIndex        =   40
      Top             =   420
      Visible         =   0   'False
      Width           =   1485
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   24
         Left            =   1080
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   65
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   23
         Left            =   840
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   64
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   22
         Left            =   600
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   63
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   21
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   62
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   20
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   61
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   19
         Left            =   1080
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   60
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   18
         Left            =   840
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   59
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   17
         Left            =   600
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   58
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   16
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   57
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   56
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   1080
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   55
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   840
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   54
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   600
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   53
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   52
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   51
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   1080
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   50
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   840
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   49
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   600
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   48
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   47
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   46
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   1080
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   45
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   120
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   840
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   44
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   120
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   600
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   43
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   120
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   42
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   120
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   41
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.TextBox txtH 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   8
      ToolTipText     =   "Hue (Tint) value for above color"
      Top             =   4200
      Width           =   555
   End
   Begin VB.TextBox txtL 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      ToolTipText     =   "Luminence (Brightness) value for above color"
      Top             =   4200
      Width           =   555
   End
   Begin VB.TextBox txtS 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2100
      TabIndex        =   9
      ToolTipText     =   "SAturation (Richness) value for above color"
      Top             =   4200
      Width           =   555
   End
   Begin VB.Timer tmrPick 
      Left            =   1560
      Top             =   0
   End
   Begin VB.CommandButton cmdPick 
      Height          =   375
      Left            =   900
      Picture         =   "frmColorRef.frx":55D4
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Pick the Background Color from Screen"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtVB 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2940
      MaxLength       =   6
      TabIndex        =   4
      ToolTipText     =   "VB Hex Code - 0 to FFFFFF"
      Top             =   2700
      Width           =   1275
   End
   Begin VB.TextBox txtHTML 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      MaxLength       =   6
      TabIndex        =   3
      ToolTipText     =   "HTML Hex code - 000000 to FFFFFF"
      Top             =   2160
      Width           =   1275
   End
   Begin VB.TextBox txtG 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2100
      TabIndex        =   6
      ToolTipText     =   "Green value for above color"
      Top             =   3540
      Width           =   555
   End
   Begin VB.TextBox txtB 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      ToolTipText     =   "Blue value for above color"
      Top             =   3540
      Width           =   555
   End
   Begin VB.TextBox txtR 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000C0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   5
      ToolTipText     =   "Red value for above color"
      Top             =   3540
      Width           =   555
   End
   Begin VB.CommandButton cmdChange 
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Change the Background Color"
      Top             =   120
      Width           =   675
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00C0C0C0&
      Height          =   1755
      Left            =   180
      ScaleHeight     =   1695
      ScaleWidth      =   4155
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Current color"
      Top             =   300
      Width           =   4215
      Begin VB.Label lblAdj 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "+20"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   7
         Left            =   3540
         TabIndex        =   73
         ToolTipText     =   "Click to adjust brightness"
         Top             =   1050
         Width           =   435
      End
      Begin VB.Label lblAdj 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "+15"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   6
         Left            =   3120
         TabIndex        =   72
         ToolTipText     =   "Click to adjust brightness"
         Top             =   1050
         Width           =   435
      End
      Begin VB.Label lblAdj 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "+10"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   5
         Left            =   2700
         TabIndex        =   71
         ToolTipText     =   "Click to adjust brightness"
         Top             =   1050
         Width           =   435
      End
      Begin VB.Label lblAdj 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "+5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   4
         Left            =   3540
         TabIndex        =   70
         ToolTipText     =   "Click to adjust brightness"
         Top             =   615
         Width           =   435
      End
      Begin VB.Label lblAdj 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "-5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   3
         Left            =   2700
         TabIndex        =   69
         ToolTipText     =   "Click to adjust brightness"
         Top             =   615
         Width           =   435
      End
      Begin VB.Label lblAdj 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "-10"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   2
         Left            =   3540
         TabIndex        =   68
         ToolTipText     =   "Click to adjust brightness"
         Top             =   180
         Width           =   435
      End
      Begin VB.Label lblAdj 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "-15"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   1
         Left            =   3120
         TabIndex        =   67
         ToolTipText     =   "Click to adjust brightness"
         Top             =   180
         Width           =   435
      End
      Begin VB.Label lblAdj 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "-20"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   0
         Left            =   2700
         TabIndex        =   66
         ToolTipText     =   "Click to adjust brightness"
         Top             =   180
         Width           =   435
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dk Gray"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   225
         Index           =   15
         Left            =   1800
         TabIndex        =   29
         Top             =   1320
         Width           =   645
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gray"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   225
         Index           =   14
         Left            =   1140
         TabIndex        =   28
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Navy"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   13
         Left            =   1800
         TabIndex        =   27
         Top             =   840
         Width           =   435
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Violet"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   225
         Index           =   12
         Left            =   1140
         TabIndex        =   26
         Top             =   840
         Width           =   465
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Teal"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   225
         Index           =   11
         Left            =   1800
         TabIndex        =   25
         Top             =   600
         Width           =   345
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lt Blue"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   225
         Index           =   10
         Left            =   1140
         TabIndex        =   24
         Top             =   600
         Width           =   570
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mustard"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   225
         Index           =   9
         Left            =   1800
         TabIndex        =   23
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Yellow"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   8
         Left            =   1140
         TabIndex        =   22
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Green"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   225
         Index           =   7
         Left            =   1800
         TabIndex        =   21
         Top             =   360
         Width           =   465
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lime"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   225
         Index           =   6
         Left            =   1140
         TabIndex        =   20
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brown"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   5
         Left            =   1800
         TabIndex        =   19
         Top             =   120
         Width           =   525
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   4
         Left            =   1140
         TabIndex        =   18
         Top             =   120
         Width           =   315
      End
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Caption         =   "Purple Text"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   1005
      End
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Caption         =   "Blue Text"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Caption         =   "White Text"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Caption         =   "Black Text"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1005
      End
   End
   Begin VB.Label lblBlend 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   4560
      TabIndex        =   97
      ToolTipText     =   "BlendBar: Click to make blended color the current color"
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label lblBlend 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   4560
      TabIndex        =   96
      ToolTipText     =   "BlendBar: Click to make blended color the current color"
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label lblBlend 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   4560
      TabIndex        =   95
      ToolTipText     =   "BlendBar: Click to make blended color the current color"
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label lblBlend 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   4560
      TabIndex        =   94
      ToolTipText     =   "BlendBar: Click to make blended color the current color"
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label lblBlend 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   4560
      TabIndex        =   93
      ToolTipText     =   "BlendBar: Click to make blended color the current color"
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblBlend 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   4560
      TabIndex        =   92
      ToolTipText     =   "BlendBar: Click to make blended color the current color"
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label lblBlend 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   4560
      TabIndex        =   91
      ToolTipText     =   "BlendBar: Click to make blended color the current color"
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lblBlend 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   4560
      TabIndex        =   90
      ToolTipText     =   "BlendBar: Click to make blended color the current color"
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label lblBlend 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   4560
      TabIndex        =   89
      ToolTipText     =   "BlendBar: Click to make blended color the current color"
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblBlend 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   4560
      TabIndex        =   88
      ToolTipText     =   "BlendBar: Click to make blended color the current color"
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblBlend 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   4560
      TabIndex        =   87
      ToolTipText     =   "BlendBar: Click to make blended color the current color"
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblBlend 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   4560
      TabIndex        =   86
      ToolTipText     =   "BlendBar: Click to make blended color the current color"
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label lblBlend 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   85
      ToolTipText     =   "BlendBar: Click to make blended color the current color"
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblBlend 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   84
      ToolTipText     =   "BlendBar: Click to make blended color the current color"
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblBlend 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   83
      ToolTipText     =   "BlendBar: Click to make blended color the current color"
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblInstruct 
      Alignment       =   2  'Center
      Caption         =   "Hold mouse over item for description"
      Height          =   255
      Left            =   1260
      TabIndex        =   80
      Top             =   60
      Width           =   2955
   End
   Begin VB.Image imgHappy 
      Height          =   240
      Left            =   4200
      Picture         =   "frmColorRef.frx":571E
      ToolTipText     =   "Happy, happy little program!"
      Top             =   0
      Width           =   240
   End
   Begin VB.Label lblH 
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   720
      TabIndex        =   39
      Top             =   4260
      Width           =   210
   End
   Begin VB.Label lblL 
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3000
      TabIndex        =   38
      Top             =   4260
      Width           =   195
   End
   Begin VB.Label lblS 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1860
      TabIndex        =   37
      Top             =   4260
      Width           =   225
   End
   Begin VB.Label lblHSL 
      Alignment       =   1  'Right Justify
      Caption         =   "Hue / Saturation / Luminence values are:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   300
      TabIndex        =   36
      Top             =   3960
      Width           =   3270
   End
   Begin VB.Label lblRGB 
      Alignment       =   1  'Right Justify
      Caption         =   "The RGB valuesfor the above color are:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   420
      TabIndex        =   35
      Top             =   3300
      Width           =   3360
   End
   Begin VB.Label lblVB 
      Caption         =   "&&H"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2520
      TabIndex        =   34
      Top             =   2730
      Width           =   1755
   End
   Begin VB.Label lblHTML 
      Caption         =   """       """
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2520
      TabIndex        =   33
      Top             =   2190
      Width           =   1755
   End
   Begin VB.Label lblG 
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1860
      TabIndex        =   32
      Top             =   3600
      Width           =   225
   End
   Begin VB.Label lblB 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3000
      TabIndex        =   31
      Top             =   3600
      Width           =   195
   End
   Begin VB.Label lblR 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   720
      TabIndex        =   30
      Top             =   3600
      Width           =   210
   End
   Begin VB.Label lblV 
      Alignment       =   1  'Right Justify
      Caption         =   "The VB hex code for the above color is:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   360
      TabIndex        =   13
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label lblHT 
      Alignment       =   1  'Right Justify
      Caption         =   "The HTML hex code for the above color is:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   420
      TabIndex        =   12
      Top             =   2100
      Width           =   1995
   End
End
Attribute VB_Name = "frmColorRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sColor As SelectedColor
Dim iHpos As Integer, iVpos As Integer
Dim blnUpdate As Boolean, blnHue As Boolean, blnSat As Boolean, blnLum As Boolean, blnBuddy As Boolean

'APIs for color-sampling routines

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RGBType
    R As Byte
    G As Byte
    B As Byte
    Filler As Byte
End Type

Private Type RGBLongType
    clr As Long
End Type

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Sub cmd5by5_Click()
'Start a 5x5 average sample
Dim lReturn As Long
    lReturn = SetCapture(picColor.hwnd)
    tmr5by5.Interval = 100
    pic5x5.Visible = True
End Sub

Private Sub cmdChange_Click()
'Call common color dialog (no .dll)
'This routine by Paul Mather, with minor modifications
    On Error GoTo e_Trap
    sColor = ShowColor(Me.hwnd, True, sColor.oSelectedColor)
    If Not sColor.bCanceled Then
        picColor.BackColor = sColor.oSelectedColor
        updateCodes
    End If
    Exit Sub
e_Trap:
    Exit Sub
End Sub

Private Sub cmdPick_Click()
'start a single point sampling
Dim lReturn As Long
    lReturn = SetCapture(picColor.hwnd)
    tmrPick.Interval = 50
End Sub

Private Sub Form_Load()
'Startup

    cmdChange.Picture = Me.Icon
    'get saved custom colors
    cc = GetINIString("Settings", "CustomColors", "")
    HtoU cc
    'setup screen w/ default gray
    sColor.oSelectedColor = picColor.BackColor
    updateCodes
    BlendEm
End Sub

Private Sub updateCodes()
Dim R As Long, G As Long, B As Long, HSLV As HSLCol, lColor As Long, i As Integer
    
    blnUpdate = True 'keeps the routine from being called again everytime it sets a value
    lColor = picColor.BackColor
    'the VB one is easy! ;)
    txtVB.Text = Hex$(lColor)
        
    'calculate & set the individual RGB values & scrolls
    
    R = RGBRed(lColor)
    G = RGBGreen(lColor)
    B = RGBBlue(lColor)
    txtR.Text = R
    vR = 255 - R
    txtR.BackColor = RGB(Val(txtR.Text), 0, 0)
    txtG.Text = G
    vG = 255 - G
    txtG.BackColor = RGB(0, Val(txtG.Text), 0)
    If Val(txtG.Text) > 172 Then
        txtG.ForeColor = vbBlack
    Else
        txtG.ForeColor = vbWhite
    End If
    txtB.Text = B
    vB = 255 - B
    txtB.BackColor = RGB(0, 0, Val(txtB.Text))
    
    'put together the HTML code
    
    txtHTML.Text = ZHex(R, 2) & ZHex(G, 2) & ZHex(B, 2)
    
    'Calculate & set HSL boxes & scrolls
    HSLV = RGBtoHSL(lColor)
    
    If Not blnHue Then
        txtH.Text = HSLV.Hue
        vH = 239 - HSLV.Hue
    End If
    
    If Not blnSat Then
        txtS.Text = HSLV.Sat
        vS = 240 - HSLV.Sat
    End If
    
    If Not blnLum Then
        txtL.Text = HSLV.Lum
        vL = 240 - HSLV.Lum
    End If
    
    'set the adjust brightness boxes if they're visible (not hidden by
    'pic5x5, which is acting as a container only).  If they're not visible,
    'an average sample is going on, and why add extra processing to an
    'already complicated task?
    
    If Not pic5x5.Visible Then
        For i = 0 To 7
            'the first function produces the values .2, .15, .1, .05
            'for the darken routine, the second function reverses the
            'order for brighten
            Select Case i
                Case 0 To 3
                    lblAdj(i).BackColor = Darken(lColor, (0.05 * (4 - i)))
                Case 4 To 7
                    lblAdj(i).BackColor = Brighten(lColor, (0.05 * (i - 3)))
            End Select
            lblAdj(i).ForeColor = ContrastingColor(lblAdj(i).BackColor)
        Next i
    End If
    blnUpdate = False 'now changes to the text boxes with trigger this routine
End Sub

Private Sub Form_Paint()
Dim RC As RECT, lReturn As Long, i As Integer
        RC.top = lblBlend(0).top - 2
        RC.left = lblBlend(0).left - 2
        RC.Bottom = lblBlend(14).top + lblBlend(14).Height + 2
        RC.Right = lblBlend(14).left + lblBlend(14).Width + 2
        lReturn = DrawEdge(hdc, RC, EDGE_BUMP, BF_RECT) ' Or BF_SOFT)
    For i = 0 To 1
        lReturn = GetClientRect(picAnchor(i).hwnd, RC)
        lReturn = InflateRect(RC, 2, 2)
        RC.top = RC.top + picAnchor(i).top
        RC.left = RC.left + picAnchor(i).left
        RC.Bottom = RC.Bottom + picAnchor(i).top
        RC.Right = RC.Right + picAnchor(i).left
        lReturn = DrawEdge(hdc, RC, EDGE_BUMP, BF_RECT) ' Or BF_SOFT)
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Save settings
    WriteINI "Settings", "CustomColors", Chr$(34) & cc & Chr$(34) 'save custom colors
End Sub

Private Sub imgHappy_Click()
' "About"
    MsgBox "Color Lab  v2.0" & vbNewLine & _
        "{ f r e e w a r e }" & vbNewLine & _
        "1999 B/W Software" & vbNewLine & _
        "Dan Redding" & vbNewLine & vbNewLine & _
        "Screen color pick adapted from a sample by Matt Hart" & vbNewLine & _
        "[ http://www.matthart.com ]" & vbNewLine & _
        "Color Dialog based on routines by Paul Mather" & vbNewLine & _
        "Source available at http://www.planet-source-code.com" & vbNewLine & vbNewLine & _
        "for René, who had way too many links to HTML color tables...", _
        vbInformation + vbOKOnly, "About Color Lab"
End Sub


Private Sub lblAdj_Click(Index As Integer)
'Brightness adjustment boxes
    picColor.BackColor = lblAdj(Index).BackColor
    updateCodes
End Sub

Private Sub lblBlend_Click(Index As Integer)
    picColor.BackColor = lblBlend(Index).BackColor
    updateCodes
End Sub

Private Sub picAnchor_Click(Index As Integer)
    picAnchor(Index).BackColor = picColor.BackColor
    BlendEm
End Sub

Private Sub picColor_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'ends a color sampling session (mouse events are captured to this picture box, so whereever
'you click, the event fires here
Dim lReturn As Long
    'check that we are actually sampling
    If tmrPick.Interval > 0 Or tmr5by5.Interval > 0 Then
        lReturn = ReleaseCapture
        tmrPick.Interval = 0
        tmr5by5.Interval = 0
        pic5x5.Visible = False
        updateCodes
    End If
End Sub

Private Sub tmr5by5_Timer()
'adapted from the routine in tmrPick, this samples 25 different
'pixels centering on the cursor, setting 25 small picture boxes
'to produce the 'enlarged' view.  The routine then runs through
'all 25 calculates the average color by averaging the seperate
'red, green, and blue values to set the main color window.

Static lX As Long, lY As Long
On Local Error Resume Next
Dim P As POINTAPI, H As Long, hD As Long, R As Long
Dim i As Integer, Red As Long, Blue As Long, Green As Long
Dim X1 As Long, Y1 As Long
Static ScrX As Long, ScrY As Long
    If ScrX = 0 Then
        ScrX = Screen.Width / Screen.TwipsPerPixelX
        ScrY = Screen.Height / Screen.TwipsPerPixelY
    End If
    GetCursorPos P
    If P.x = lX And P.y = lY Then Exit Sub
    lX = P.x: lY = P.y
    For i = 0 To 24
        '5x5 position relative to cursor (x & y = -2 to 2)
        X1 = (lX + (i Mod 5) - 2)
        Y1 = (lY + (i \ 5) - 2)
        P.x = X1
        P.y = Y1
        
        If X1 < 0 Or Y1 < 0 Or X1 > ScrX Or Y1 > ScrY Then
            R = 0
        Else
            'this information needs to be recalcualted
            'for each point; after all, the 5x5 square
            'could overlap 2 or more windows
            
            'which window?
            H = WindowFromPoint(X1, Y1)
            
            'get device context for that window
            hD = GetDC(H)
            
            'convert screen coordinates to local window
            ScreenToClient H, P
            
            'get color
            R = GetPixel(hD, P.x, P.y)
            If R = -1 Then
                'titlebar or other special area
                'get color by copying the point to picturebox, then checking that
                BitBlt picPoint(i).hdc, 0, 0, 1, 1, hD, P.x, P.y, vbSrcCopy
                R = picPoint(i).Point(0, 0)
            Else
                'R is the color
                picPoint(i).PSet (0, 0), R
            End If
            'Must do to prevent memory leaks
            ReleaseDC H, hD
        End If
        'set backcolor of whole picturebox to R
        picPoint(i).BackColor = R
    Next i
    
    'averaging
    
    For i = 0 To 24
        Red = Red + RGBRed(picPoint(i).BackColor)
        Blue = Blue + RGBBlue(picPoint(i).BackColor)
        Green = Green + RGBGreen(picPoint(i).BackColor)
    Next i
    
    'set main picturebox w/ average color
    picColor.BackColor = RGB(CInt(Red / 25), CInt(Green / 25), CInt(Blue / 25))
    
    updateCodes

End Sub

Private Sub tmrPick_Timer()
'This routine adapted from a project by Matt Hart

'Matt's comments follow:
' Getpixel sample by Matt Hart - vbhelp@matthart.com
' http://matthart.com
'
' This sample shows how to get the pixel color of any point
' on the screen. The GetPixel API requires CLIENT coordinates,
' so you must first get the window handle and hDC where the
' cursor is. Once you get that, you can get the pixel.
'
' However, there's one "gotcha" I found while writing this.
' Window titlebars return a "-1" for the pixel color, which
' is invalid! So, what I did to get around that was use
' BitBlt to copy a pixel from that device to the PictureBox
' control I'm using to show the colors, then use the Point
' method to check the color.

'for detailed comments, see corresponding function in tmr5x5
Static lX As Long, lY As Long
On Local Error Resume Next
Dim P As POINTAPI, H As Long, hD As Long, R As Long
    GetCursorPos P
    If P.x = lX And P.y = lY Then Exit Sub
    lX = P.x: lY = P.y
    H = WindowFromPoint(lX, lY)
    hD = GetDC(H)
    ScreenToClient H, P
    R = GetPixel(hD, P.x, P.y)
    If R = -1 Then
        BitBlt picColor.hdc, 0, 0, 1, 1, hD, P.x, P.y, vbSrcCopy
        R = picColor.Point(0, 0)
    Else
        picColor.PSet (0, 0), R
    End If
    ReleaseDC H, hD
    picColor.BackColor = R
    updateCodes
End Sub

Private Sub txtB_Change()
    'change blue value
    If blnUpdate Then Exit Sub 'updating, don't need to adjust anything else here
    'too high?
    If Val(txtR.Text) > 255 Then
        txtB.Text = "255"
    Else
        txtB.Text = Val(txtB.Text)
    End If
    'set new color
    picColor.BackColor = RGB(Val(txtR.Text), Val(txtG.Text), Val(txtB.Text))
    updateCodes
    
    'blnbuddy if txtB was changed by vB scroller.
    'without this, the two routines would trigger each other until overflow
    If Not blnBuddy Then
        blnBuddy = True
        vB.Value = 255 - Val(txtB.Text)
        blnBuddy = False
    End If
    'select if one of these two values (easy to overtype)
    If txtB.Text = "0" Or txtB.Text = "255" Then txtB_GotFocus
End Sub

Private Sub txtB_GotFocus()
'select all when get focus
    txtB.SelStart = 0
    txtB.SelLength = Len(txtB.Text)
End Sub

'For txtG routine comments, see txtB
Private Sub txtG_Change()
    If blnUpdate Then Exit Sub
    If Val(txtR.Text) > 255 Then
        txtG.Text = "255"
    Else
        txtG.Text = Val(txtG.Text)
    End If
    picColor.BackColor = RGB(Val(txtR.Text), Val(txtG.Text), Val(txtB.Text))
    updateCodes
    If Not blnBuddy Then
        blnBuddy = True
        vR.Value = 255 - Val(txtG.Text)
        blnBuddy = False
    End If
    If txtG.Text = "0" Or txtG.Text = "255" Then txtG_GotFocus
End Sub

Private Sub txtG_GotFocus()
    txtG.SelStart = 0
    txtG.SelLength = Len(txtG.Text)
End Sub

Private Sub txtH_Change()
Dim HSLV As HSLCol
    If blnUpdate Then Exit Sub 'updating, don''t need to change here
    'too high?
    If Val(txtH.Text) >= HSLMAX Then
        txtH.Text = HSLMAX - 1
    Else
        txtH.Text = Val(txtH.Text)
    End If
    'calc & set new rgb color
    HSLV.Hue = Val(txtH.Text)
    HSLV.Sat = Val(txtS.Text)
    HSLV.Lum = Val(txtL.Text)
    picColor.BackColor = HSLtoRGB(HSLV)
    
    'protect from another loop (HSL->RGB->HSL sometimes changes HSL due to rounding errors)
    blnHue = True
    updateCodes
    blnHue = False
    'Protect from infinite loop adjusting vH scroller
    If Not blnBuddy Then
        blnBuddy = True
        vH.Value = 239 - Val(txtH.Text)
        blnBuddy = False
    End If
    'select for overtyping if high or low val
    If txtH.Text = "0" Or txtH.Text = "239" Then txtH_GotFocus
End Sub

Private Sub txtH_GotFocus()
    txtH.SelStart = 0
    txtH.SelLength = Len(txtH.Text)
End Sub

Private Sub txtHTML_Change()
Dim R As Long, G As Long, B As Long
    txtHTML.Text = UCase$(txtHTML.Text) 'uppercase it
    txtHTML.SelStart = iHpos 'keep cursor where it was after uppercase
    If Len(txtHTML.Text) = 6 Then 'full code; change color
        If Not isHex(txtHTML.Text) Then
            Beep 'not valid!
            Exit Sub
        End If
        'get RGB values from hex string, the easy way
        R = Val("&H" & Mid$(txtHTML.Text, 1, 2))
        G = Val("&H" & Mid$(txtHTML.Text, 3, 2))
        B = Val("&H" & Right$(txtHTML.Text, 2))
        
        'set color and update codes
        '(unless txtHTML was changed BY the updateCodes rotuine)
        picColor.BackColor = RGB(R, G, B)
        If Not blnUpdate Then
            updateCodes
        End If
    End If
End Sub

Private Sub txtHTML_GotFocus()
    txtHTML.SelStart = 0
    txtHTML.SelLength = Len(txtHTML.Text)
    iHpos = 0 'save cursor position
End Sub

Private Sub txtHTML_KeyDown(KeyCode As Integer, Shift As Integer)
'save new cursor position
    If KeyCode = vbKeyBack Then
        iHpos = txtHTML.SelStart - 1
    Else
        iHpos = txtHTML.SelStart + 1
    End If
End Sub

'for txtL comments, see corresponding in txtH
Private Sub txtL_Change()
Dim HSLV As HSLCol
    If blnUpdate Then Exit Sub
    If Val(txtL.Text) > HSLMAX Then
        txtL.Text = HSLMAX
    Else
        txtL.Text = Val(txtL.Text)
    End If
    HSLV.Hue = Val(txtH.Text)
    HSLV.Sat = Val(txtS.Text)
    HSLV.Lum = Val(txtL.Text)
    picColor.BackColor = HSLtoRGB(HSLV)
    blnHue = True
    updateCodes
    blnHue = False
    If Not blnBuddy Then
        blnBuddy = True
        vL.Value = 240 - Val(txtL.Text)
        blnBuddy = False
    End If
    If txtL.Text = "0" Or txtL.Text = "240" Then txtL_GotFocus
End Sub

Private Sub txtL_GotFocus()
    txtL.SelStart = 0
    txtL.SelLength = Len(txtL.Text)
End Sub

'for txtR comments, see corresponding in txtB
Private Sub txtR_Change()
    If blnUpdate Then Exit Sub
    If Val(txtR.Text) > 255 Then
        txtR.Text = "255"
    Else
        txtR.Text = Val(txtR.Text)
    End If
    picColor.BackColor = RGB(Val(txtR.Text), Val(txtG.Text), Val(txtB.Text))
    updateCodes
    blnBuddy = True
    vR.Value = 255 - Val(txtR.Text)
    blnBuddy = False
    If Not blnBuddy Then
        blnBuddy = True
        vR.Value = 255 - Val(txtR.Text)
        blnBuddy = False
    End If
    If txtR.Text = "0" Or txtR.Text = "255" Then txtR_GotFocus
End Sub

Private Sub txtR_GotFocus()
    txtR.SelStart = 0
    txtR.SelLength = Len(txtR.Text)
End Sub

'for txtS comments, see corresponding in txtH
Private Sub txtS_Change()
Dim HSLV As HSLCol
    If blnUpdate Then Exit Sub
    If Val(txtS.Text) > HSLMAX Then
        txtS.Text = HSLMAX
    Else
        txtS.Text = Val(txtS.Text)
    End If
    HSLV.Hue = Val(txtH.Text)
    HSLV.Sat = Val(txtS.Text)
    HSLV.Lum = Val(txtL.Text)
    picColor.BackColor = HSLtoRGB(HSLV)
    blnSat = True
    updateCodes
    blnSat = False
    If Not blnBuddy Then
        blnBuddy = True
        vS.Value = 240 - Val(txtS.Text)
        blnBuddy = False
    End If
    If txtS.Text = "0" Or txtS.Text = "240" Then txtS_GotFocus
End Sub

Private Sub txtS_GotFocus()
    txtS.SelStart = 0
    txtS.SelLength = Len(txtS.Text)
End Sub

Private Sub txtVB_Change()
Dim VBV As Long
    txtVB.Text = UCase$(txtVB.Text)
    If Not isHex(txtVB.Text) Then
        Beep 'invalid hex code
        Exit Sub
    End If
    'adjust selection
    txtVB.SelStart = Len(txtVB.Text)
    'change color
    VBV = CLng("&H1" & txtVB.Text)
    'this avoids the negative if the 16th bit is set
    If VBV > 0 Then
        VBV = VBV - CLng("&H1" & String$(Len(txtVB.Text), "0"))
    Else
        VBV = 0
    End If
    
    'set color & update the codes
    picColor.BackColor = VBV
    If Not blnUpdate Then
        updateCodes
    End If
    If VBV = 0 Then txtVB_GotFocus
End Sub

Private Sub txtVB_GotFocus()
    txtVB.SelStart = 0
    txtVB.SelLength = Len(txtVB.Text)
    iVpos = 0
End Sub

Private Function isHex(strHex As String) As Boolean
'check that a string contains only 0-9 and A-F
Dim blnHex As Boolean, i As Integer, strChar As String * 1
    If Len(strHex) = 0 Then Exit Function
    blnHex = True
    For i = 1 To Len(strHex)
        strChar = Mid$(strHex, i, 1)
        blnHex = blnHex And ((strChar >= "0" And strChar <= "9") Or (strChar >= "A" And strChar <= "F"))
    Next i
    isHex = blnHex
End Function

'Scroll bars imitating spin buttons

Private Sub vB_Change()
    'blnBuddy keeps this event and the txt?_change events from
    'calling each other
    If Not blnBuddy Then
        blnBuddy = True
        'up is down and down is up!
        txtB.Text = 255 - vB.Value
        blnBuddy = False
    End If
End Sub

Private Sub vG_Change()
    If Not blnBuddy Then
        blnBuddy = True
        txtG.Text = 255 - vG.Value
        blnBuddy = False
    End If
End Sub

Private Sub vH_Change()
    If Not blnBuddy Then
        blnBuddy = True
        txtH.Text = 239 - vH.Value
        blnBuddy = False
    End If
End Sub

Private Sub vL_Change()
    If Not blnBuddy Then
        blnBuddy = True
        txtL.Text = 240 - vL.Value
        blnBuddy = False
    End If
End Sub

Private Sub vR_Change()
    If Not blnBuddy Then
        blnBuddy = True
        txtR.Text = 255 - vR.Value
        blnBuddy = False
    End If
End Sub

Private Sub vS_Change()
    If Not blnBuddy Then
        blnBuddy = True
        txtS.Text = 240 - vS.Value
        blnBuddy = False
    End If
End Sub

Private Sub BlendEm()
Dim i As Integer
    For i = 0 To 14
        lblBlend(i).BackColor = Blend(picAnchor(0).BackColor, picAnchor(1).BackColor, (i + 1) / 16)
    Next i
End Sub
