VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Leitor de Mentes"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   371
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   403
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Site do Autor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   105
      Top             =   5160
      Width           =   1455
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   100
      Left            =   5400
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   103
      Top             =   4680
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   99
      Left            =   4800
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   102
      Top             =   4680
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   98
      Left            =   4200
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   101
      Top             =   4680
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   97
      Left            =   3600
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   100
      Top             =   4680
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   96
      Left            =   3000
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   99
      Top             =   4680
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   95
      Left            =   2400
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   98
      Top             =   4680
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   94
      Left            =   1800
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   97
      Top             =   4680
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   93
      Left            =   1200
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   96
      Top             =   4680
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   92
      Left            =   600
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   95
      Top             =   4680
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   91
      Left            =   0
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   94
      Top             =   4680
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   90
      Left            =   5400
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   93
      Top             =   4320
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   89
      Left            =   4800
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   92
      Top             =   4320
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   88
      Left            =   4200
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   91
      Top             =   4320
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   87
      Left            =   3600
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   90
      Top             =   4320
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   86
      Left            =   3000
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   89
      Top             =   4320
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   85
      Left            =   2400
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   88
      Top             =   4320
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   84
      Left            =   1800
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   87
      Top             =   4320
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   83
      Left            =   1200
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   86
      Top             =   4320
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   82
      Left            =   600
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   85
      Top             =   4320
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   81
      Left            =   0
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   84
      Top             =   4320
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   80
      Left            =   5400
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   83
      Top             =   3960
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   79
      Left            =   4800
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   82
      Top             =   3960
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   78
      Left            =   4200
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   81
      Top             =   3960
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   77
      Left            =   3600
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   80
      Top             =   3960
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   76
      Left            =   3000
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   79
      Top             =   3960
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   75
      Left            =   2400
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   78
      Top             =   3960
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   74
      Left            =   1800
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   77
      Top             =   3960
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   73
      Left            =   1200
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   76
      Top             =   3960
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   72
      Left            =   600
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   75
      Top             =   3960
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   71
      Left            =   0
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   74
      Top             =   3960
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   70
      Left            =   5400
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   73
      Top             =   3600
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   69
      Left            =   4800
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   72
      Top             =   3600
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   68
      Left            =   4200
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   71
      Top             =   3600
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   67
      Left            =   3600
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   70
      Top             =   3600
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   66
      Left            =   3000
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   69
      Top             =   3600
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   65
      Left            =   2400
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   68
      Top             =   3600
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   64
      Left            =   1800
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   67
      Top             =   3600
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   63
      Left            =   1200
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   66
      Top             =   3600
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   62
      Left            =   600
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   65
      Top             =   3600
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   61
      Left            =   0
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   64
      Top             =   3600
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   60
      Left            =   5400
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   63
      Top             =   3240
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   59
      Left            =   4800
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   62
      Top             =   3240
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   58
      Left            =   4200
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   61
      Top             =   3240
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   57
      Left            =   3600
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   60
      Top             =   3240
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   56
      Left            =   3000
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   59
      Top             =   3240
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   55
      Left            =   2400
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   58
      Top             =   3240
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   54
      Left            =   1800
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   57
      Top             =   3240
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   53
      Left            =   1200
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   56
      Top             =   3240
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   52
      Left            =   600
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   55
      Top             =   3240
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   51
      Left            =   0
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   54
      Top             =   3240
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   50
      Left            =   5400
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   53
      Top             =   2880
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   49
      Left            =   4800
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   52
      Top             =   2880
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   48
      Left            =   4200
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   51
      Top             =   2880
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   47
      Left            =   3600
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   50
      Top             =   2880
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   46
      Left            =   3000
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   49
      Top             =   2880
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   45
      Left            =   2400
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   48
      Top             =   2880
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   44
      Left            =   1800
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   47
      Top             =   2880
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   43
      Left            =   1200
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   46
      Top             =   2880
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   42
      Left            =   600
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   45
      Top             =   2880
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   41
      Left            =   0
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   44
      Top             =   2880
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   40
      Left            =   5400
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   43
      Top             =   2520
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   39
      Left            =   4800
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   42
      Top             =   2520
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   38
      Left            =   4200
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   41
      Top             =   2520
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   37
      Left            =   3600
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   40
      Top             =   2520
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   36
      Left            =   3000
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   39
      Top             =   2520
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   35
      Left            =   2400
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   38
      Top             =   2520
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   34
      Left            =   1800
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   37
      Top             =   2520
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   33
      Left            =   1200
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   36
      Top             =   2520
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   32
      Left            =   600
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   35
      Top             =   2520
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   31
      Left            =   0
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   34
      Top             =   2520
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   30
      Left            =   5400
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   33
      Top             =   2160
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   29
      Left            =   4800
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   32
      Top             =   2160
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   28
      Left            =   4200
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   31
      Top             =   2160
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   27
      Left            =   3600
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   30
      Top             =   2160
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   26
      Left            =   3000
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   29
      Top             =   2160
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   25
      Left            =   2400
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   28
      Top             =   2160
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   24
      Left            =   1800
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   27
      Top             =   2160
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   23
      Left            =   1200
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   26
      Top             =   2160
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   22
      Left            =   600
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   25
      Top             =   2160
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   21
      Left            =   0
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   24
      Top             =   2160
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   20
      Left            =   5400
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   23
      Top             =   1800
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   19
      Left            =   4800
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   22
      Top             =   1800
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   18
      Left            =   4200
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   21
      Top             =   1800
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   17
      Left            =   3600
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   20
      Top             =   1800
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   16
      Left            =   3000
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   19
      Top             =   1800
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   15
      Left            =   2400
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   18
      Top             =   1800
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   14
      Left            =   1800
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   17
      Top             =   1800
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   13
      Left            =   1200
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   16
      Top             =   1800
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   12
      Left            =   600
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   15
      Top             =   1800
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   11
      Left            =   0
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   14
      Top             =   1800
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   10
      Left            =   5400
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   13
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   9
      Left            =   4800
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   12
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   8
      Left            =   4200
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   11
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   7
      Left            =   3600
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   10
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   6
      Left            =   3000
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   9
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   2400
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   8
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   1800
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   7
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   1200
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   6
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   600
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   5
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "De Novo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adivinhar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5160
      Width           =   2655
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   0
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   2
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox PicResp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   0
      ScaleHeight     =   239
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   399
      TabIndex        =   104
      Top             =   1440
      Width           =   6015
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Form1.frx":08CA
      Top             =   45
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   $"Form1.frx":1194
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   6015
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   408
      Y1              =   40
      Y2              =   40
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "       Leitor de Mentes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ---------------------------------------------------------------
' Por Walter Staeblein - 2003
' http://www.codex.com.br
' Código Gratuito e sem restrições, uso por conta e risco próprio
' ---------------------------------------------------------------

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Const DT_WORDBREAK = &H10
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal LpszDir As String, ByVal FsShowCmd As Long) As Long

Private Sub Command1_Click()

    PicResp.ZOrder
    
End Sub

Private Sub Command2_Click()

    fillPics
    PicResp.ZOrder (1)
    
End Sub

Private Sub Command3_Click()

    Dim ScrhDC As Long, StartDoc As Long
    
    Scr_hDC = GetDesktopWindow()
    StartDoc = ShellExecute(Scr_hDC, "Open", "http://www.codex.com.br", "", Left$(App.Path, 3), vbMaximizedFocus)
    
End Sub

Private Sub Form_Load()

    fillPics
    
End Sub

Sub fillPics()

    Dim RC1 As RECT, RC2 As RECT, RC3 As RECT, I As Integer
    
    Randomize Timer
    
    With RC1
        .Left = 1
        .Top = 3
        .Right = 31
        .Bottom = 33
    End With
    
    With RC2
        .Left = Pic(1).Width - 18
        .Top = 3
        .Right = Pic(1).Width - 1
        .Bottom = 33
    End With

    With PicResp
            .Font.Name = "Wingdings"
            .Font.Size = 150
            .ForeColor = vbBlue
    End With
    Resp = Chr(65 + Rnd * 26)

    With RC3
        .Left = (PicResp.Width - PicResp.TextWidth(Resp)) / 2
        .Top = (PicResp.Height - PicResp.TextHeight(Resp)) / 2
        .Right = .Left + PicResp.TextWidth(Resp)
        .Bottom = .Top + PicResp.TextHeight(Resp)
    End With
    PicResp.Cls
    A = DrawText(PicResp.hDC, Resp, -1, RC3, DT_WORDBREAK)

    
    For I = 1 To 100
        With Pic(I)
            Pic(I).Cls
            .Font.Name = "Arial"
            .Font.Size = 9
            A = DrawText(Pic(I).hDC, CStr(I), -1, RC1, DT_WORDBREAK)
            
            Char = Chr(65 + Rnd * 26)
            If I Mod 9 = 0 Then Char = Resp
            .Font.Name = "Wingdings"
            .Font.Size = 14
            .ForeColor = vbBlue
            A = DrawText(Pic(I).hDC, Char, -1, RC2, DT_WORDBREAK)
        End With
    Next
End Sub
