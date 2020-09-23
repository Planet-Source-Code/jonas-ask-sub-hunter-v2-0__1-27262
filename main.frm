VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Submarine Hunter v 2.0"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7455
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicBossM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   8940
      Picture         =   "main.frx":030A
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   42
      Top             =   4620
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox PicBoss 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   8940
      Picture         =   "main.frx":1F74
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   41
      Top             =   4080
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox PicSignM 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   9360
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   40
      Top             =   660
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox PicSign 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   9360
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   39
      Top             =   300
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox PicExpM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Index           =   3
      Left            =   9780
      Picture         =   "main.frx":3BDE
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   38
      Top             =   2520
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox PicExpM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Index           =   2
      Left            =   9720
      Picture         =   "main.frx":6BF0
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   37
      Top             =   2400
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox PicExpM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Index           =   1
      Left            =   9600
      Picture         =   "main.frx":9C02
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   36
      Top             =   2280
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox PicExpM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Index           =   0
      Left            =   9480
      Picture         =   "main.frx":CC14
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   35
      Top             =   2160
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox PicExp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Index           =   3
      Left            =   8580
      Picture         =   "main.frx":FC26
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   34
      Top             =   2880
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox PicExp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Index           =   2
      Left            =   8520
      Picture         =   "main.frx":12C38
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   33
      Top             =   2640
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox PicExp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Index           =   1
      Left            =   8460
      Picture         =   "main.frx":15C4A
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   32
      Top             =   2400
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox PicExp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Index           =   0
      Left            =   8400
      Picture         =   "main.frx":18C5C
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   31
      Top             =   2160
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox PicElvM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   8700
      Picture         =   "main.frx":1BC6E
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   30
      Top             =   1920
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox PicElvM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   8700
      Picture         =   "main.frx":1C010
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   29
      Top             =   1740
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox PicElvM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   8700
      Picture         =   "main.frx":1C3B2
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   28
      Top             =   1560
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox PicElv 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   8340
      Picture         =   "main.frx":1C754
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   27
      Top             =   1920
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox PicElv 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   8340
      Picture         =   "main.frx":1CAF6
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   26
      Top             =   1740
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox PicElv 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   8340
      Picture         =   "main.frx":1CE98
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   25
      Top             =   1560
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox PicPlane2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   6840
      Picture         =   "main.frx":1D23A
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   24
      Top             =   3780
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox PicPlane2M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   6840
      Picture         =   "main.frx":1D69C
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   23
      Top             =   4080
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox PicPlane1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   6360
      Picture         =   "main.frx":1DAFE
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   22
      Top             =   3780
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox PicPlane1M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   6360
      Picture         =   "main.frx":1DF60
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   21
      Top             =   4080
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox PicExit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   6360
      Picture         =   "main.frx":1E3C2
      ScaleHeight     =   300
      ScaleWidth      =   900
      TabIndex        =   18
      Top             =   5220
      Width           =   930
   End
   Begin VB.PictureBox PicCrash2M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7560
      Picture         =   "main.frx":1F214
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   17
      Top             =   5160
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox PicCrash2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7560
      Picture         =   "main.frx":1FAA6
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   16
      Top             =   4800
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox PicCrash1M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7500
      Picture         =   "main.frx":20338
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   15
      Top             =   4440
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox PicCrash1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7560
      Picture         =   "main.frx":20BCA
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   14
      Top             =   4020
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox Pic3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   5460
      Left            =   4620
      ScaleHeight     =   360
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   13
      Top             =   6180
      Width           =   6060
   End
   Begin VB.PictureBox Pic2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   5460
      Left            =   120
      ScaleHeight     =   360
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   11
      Top             =   6000
      Width           =   6060
   End
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   5460
      Left            =   8220
      ScaleHeight     =   360
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   12
      Top             =   6060
      Width           =   6060
   End
   Begin VB.PictureBox PicSub2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   7560
      Picture         =   "main.frx":2145C
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   10
      Top             =   3300
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox PicSub2M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   7560
      Picture         =   "main.frx":218D6
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   9
      Top             =   3540
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox PicSub1M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   7560
      Picture         =   "main.frx":21D50
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   8
      Top             =   3000
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox PicSub1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   7560
      Picture         =   "main.frx":221CA
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox PicShip2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7500
      Picture         =   "main.frx":22644
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.PictureBox PicShip2M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7500
      Picture         =   "main.frx":22F46
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   5
      Top             =   1860
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.PictureBox PicMineM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   90
      Left            =   7680
      Picture         =   "main.frx":23848
      ScaleHeight     =   6
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   4
      Top             =   2460
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox PicMine 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   90
      Left            =   7680
      Picture         =   "main.frx":23902
      ScaleHeight     =   6
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   3
      Top             =   2340
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox PicShipM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7380
      Picture         =   "main.frx":239BC
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   2
      Top             =   4020
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.PictureBox PicShip 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7380
      Picture         =   "main.frx":242BE
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   1
      Top             =   3720
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   6540
      Top             =   240
   End
   Begin VB.PictureBox MainBoard 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5460
      Left            =   120
      Picture         =   "main.frx":24BC0
      ScaleHeight     =   360
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   120
      Width           =   6060
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HighScores"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   6285
      TabIndex        =   20
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      Caption         =   "I made this"
      Height          =   1695
      Left            =   6300
      TabIndex        =   19
      Top             =   2040
      Width           =   1035
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Firsttime As Integer
Dim MaxSpeed As Boolean
Sub PaintBoard()
    Pic1.Cls
    Pic2.Cls
    Pic3.Cls
    MainBoard.Cls
    'HUD
    MainBoard.Print " Ammo: "
    MainBoard.Print " Score: " & P1.Score
    MainBoard.Print " Strenght: " & P1.Health
    For A = 0 To P1.Ammo - 1
        BitBlt MainBoard.hdc, (A * 7) + 38, 5, 6, 6, PicMineM.hdc, 0, 0, SRCAND
        BitBlt MainBoard.hdc, (A * 7) + 38, 5, 6, 6, PicMine.hdc, 0, 0, SRCPAINT
    Next A
    If Boss.Act Then
        txt = " Boss: "
        For A = 0 To Boss.Life - 1
           txt = txt & "I"
        Next A
        MainBoard.Print txt
    End If
    
    'The Ship
    Select Case P1.Dire
    Case 1
        BitBlt Pic2.hdc, P1.X, P1.Y, ShipBredde, ShipHoyde, PicShipM.hdc, 0, 0, SRCAND
        BitBlt Pic1.hdc, P1.X, P1.Y, ShipBredde, ShipHoyde, PicShip.hdc, 0, 0, SRCPAINT
    Case 2
        BitBlt Pic2.hdc, P1.X, P1.Y, ShipBredde, ShipHoyde, PicShip2M.hdc, 0, 0, SRCAND
        BitBlt Pic1.hdc, P1.X, P1.Y, ShipBredde, ShipHoyde, PicShip2.hdc, 0, 0, SRCPAINT
    End Select
    'The King
    If TheKing.Act Then
        BitBlt Pic2.hdc, TheKing.X, TheKing.Y, 23, 12, PicElvM(TheKing.Tag).hdc, 0, 0, SRCAND
        BitBlt Pic1.hdc, TheKing.X, TheKing.Y, 23, 12, PicElv(TheKing.Tag).hdc, 0, 0, SRCPAINT
    End If
    'Bombs
    For A = 1 To 30
        With Bombs(A)
            If .Act Then
                Pic3.Line (.X, .Y)-Step(1, 1)
            End If
        End With
    Next
    'Bossbombs
    For A = 1 To UBound(BossBombs)
        If BossBombs(A).Act Then
            Pic3.Line (BossBombs(A).X, BossBombs(A).Y)-Step(1, 2), vbRed, BF
        End If
    Next A
    
    'Mines
    For A = 1 To 30
        With Shot(A)
            If .Act Then
                BitBlt Pic3.hdc, .X, .Y, 6, 6, PicMineM.hdc, 0, 0, SRCAND
                BitBlt Pic3.hdc, .X, .Y, 6, 6, PicMineM.hdc, 0, 0, SRCAND
                BitBlt Pic1.hdc, .X, .Y, 6, 6, PicMine.hdc, 0, 0, SRCPAINT
            End If
        End With
    Next
    
    'planes
    For A = 1 To 10
        With Planes(A)
            If .Act Then
                If .Dire = 1 Then
                    BitBlt Pic3.hdc, .X, .Y, PlaneBredde, PlaneHoyde, PicPlane1M.hdc, 0, 0, SRCAND
                    BitBlt Pic3.hdc, .X, .Y, PlaneBredde, PlaneHoyde, PicPlane1M.hdc, 0, 0, SRCAND
                    BitBlt Pic1.hdc, .X, .Y, PlaneBredde, PlaneHoyde, PicPlane1.hdc, 0, 0, SRCPAINT
                Else
                    BitBlt Pic3.hdc, .X, .Y, PlaneBredde, PlaneHoyde, PicPlane2M.hdc, 0, 0, SRCAND
                    BitBlt Pic3.hdc, .X, .Y, PlaneBredde, PlaneHoyde, PicPlane2M.hdc, 0, 0, SRCAND
                    BitBlt Pic1.hdc, .X, .Y, PlaneBredde, PlaneHoyde, PicPlane2.hdc, 0, 0, SRCPAINT
                End If
            End If
        End With
    Next
    
    'Subs
    For A = 1 To 30
        With Subs(A)
            If .Act Then
                If .Dire = 1 Then
                    Select Case .Damaged
                    Case 0
                        BitBlt Pic2.hdc, .X, .Y, PicSub1M.ScaleWidth, PicSub1M.ScaleHeight, PicSub1M.hdc, 0, 0, SRCAND
                        BitBlt Pic1.hdc, .X, .Y, PicSub1M.ScaleWidth, PicSub1M.ScaleHeight, PicSub1M.hdc, 0, 0, SRCAND
                        BitBlt Pic1.hdc, .X, .Y, PicSub1.ScaleWidth, PicSub1.ScaleHeight, PicSub1.hdc, 0, 0, SRCPAINT
                    Case Else
                        BitBlt Pic3.hdc, .X, .Y, PicCrash1M.ScaleWidth, PicCrash1M.ScaleHeight, PicCrash1M.hdc, 0, 0, SRCAND
                        BitBlt Pic1.hdc, .X, .Y, PicCrash1M.ScaleWidth, PicCrash1M.ScaleHeight, PicCrash1M.hdc, 0, 0, SRCAND
                        BitBlt Pic1.hdc, .X, .Y, PicCrash1.ScaleWidth, PicCrash1.ScaleHeight, PicCrash1.hdc, 0, 0, SRCPAINT
                    End Select
                Else
                    Select Case .Damaged
                    Case 0
                        BitBlt Pic2.hdc, .X, .Y, PicSub1M.ScaleWidth, PicSub2M.ScaleHeight, PicSub2M.hdc, 0, 0, SRCAND
                        BitBlt Pic1.hdc, .X, .Y, PicSub1M.ScaleWidth, PicSub2M.ScaleHeight, PicSub2M.hdc, 0, 0, SRCAND
                        BitBlt Pic1.hdc, .X, .Y, PicSub2.ScaleWidth, PicSub2.ScaleHeight, PicSub2.hdc, 0, 0, SRCPAINT
                    Case Else
                        BitBlt Pic3.hdc, .X, .Y, PicCrash1M.ScaleWidth, PicCrash2M.ScaleHeight, PicCrash2M.hdc, 0, 0, SRCAND
                        BitBlt Pic1.hdc, .X, .Y, PicCrash1M.ScaleWidth, PicCrash2M.ScaleHeight, PicCrash2M.hdc, 0, 0, SRCAND
                        BitBlt Pic1.hdc, .X, .Y, PicCrash2.ScaleWidth, PicCrash2.ScaleHeight, PicCrash2.hdc, 0, 0, SRCPAINT
                    End Select
                End If
            End If
        End With
    Next
    'The boss
    If Boss.Act Then
        BitBlt Pic2.hdc, Boss.X, Boss.Y, BossBredde, BossHoyde, PicBossM.hdc, 0, 0, SRCAND
        BitBlt Pic1.hdc, Boss.X, Boss.Y, BossBredde, BossHoyde, PicBossM.hdc, 0, 0, SRCAND
        BitBlt Pic1.hdc, Boss.X, Boss.Y, BossBredde, BossHoyde, PicBoss.hdc, 0, 0, SRCPAINT
    End If
    'Explosions
    For A = 1 To UBound(Explo)
        If Explo(A).Act Then
            BitBlt Pic3.hdc, Explo(A).X, Explo(A).Y, 67, 60, PicExpM(Int(Explo(A).Tag / 3)).hdc, 0, 0, SRCAND
            BitBlt Pic1.hdc, Explo(A).X, Explo(A).Y, 67, 60, PicExp(Int(Explo(A).Tag / 3)).hdc, 0, 0, SRCPAINT
        End If
    Next A
    'Score signs
    For A = 1 To UBound(Signs)
        If Signs(A).Tag > 0 Then
            PicSign.Cls: PicSignM.Cls
            PicSign.Print Signs(A).Score: PicSignM.Print Signs(A).Score
            BitBlt Pic3.hdc, Signs(A).X, Signs(A).Y, 49, 21, PicSignM.hdc, 0, 0, SRCAND
            BitBlt Pic1.hdc, Signs(A).X, Signs(A).Y, 49, 21, PicSignM.hdc, 0, 0, SRCAND
            BitBlt Pic1.hdc, Signs(A).X, Signs(A).Y, 49, 21, PicSign.hdc, 0, 0, SRCPAINT
        End If
    Next A
    'Now paint put it all together
    BitBlt MainBoard.hdc, 0, 0, MainBoard.ScaleWidth, MainBoard.ScaleHeight, Pic2.hdc, 0, 0, SRCAND
    BitBlt MainBoard.hdc, 0, 0, MainBoard.ScaleWidth, MainBoard.ScaleHeight, Pic3.hdc, 0, 0, SRCAND
    BitBlt MainBoard.hdc, 0, 0, MainBoard.ScaleWidth, MainBoard.ScaleHeight, Pic1.hdc, 0, 0, SRCPAINT
End Sub

Private Sub Form_Load()
    temp = "Game and graphics by Jonas Ask"
    lblCredits.Caption = temp
    NumShots = 0
    Firsttime = True
    Randomize
    d = Int((Rnd * 2) + 1)
    P1.Dire = d
    P1.Ammo = MaxAmmo
    P1.X = 200
    P1.Y = 104
    P1.Health = 20
    LoadScore
End Sub

Private Sub Label2_Click()
    frmHigh.Show , Me
    MainPause = True
End Sub

Public Sub PicExit_Click()
    MainPause = True
    If DontClose Then Exit Sub
    UpdateScore
    If Not DontClose Then SaveScore: End
End Sub

Private Sub PicExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PicExit.BorderStyle = 0
End Sub

Private Sub PicExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PicExit.BorderStyle = 1
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    MainLoop
End Sub
Sub MainLoop()
Dim countAmmo As Integer
Dim countMake As Integer
Dim LastTick As String
    LastTick = "0"
    Do Until WW3 'Not THAT long in other words ;)
        DoEvents
        'Pause function
        Do Until MainPause = False
            DoEvents
        Loop
        
        'Framelimiter
        Do Until GetTickCount > LastTick Or MaxSpeed = True
            DoEvents
            NowTick = GetTickCount
        Loop
        LastTick = GetTickCount + 28
        
        'Stuff that does not trigger each tick:
        If countAmmo = 40 Then
            Addammo
            countAmmo = 0
        Else
            countAmmo = countAmmo + 1
        End If
        If countMake = 15 Then
            MakeSub
            MakePlane
            countMake = 0
        Else
            countMake = countMake + 1
        End If
        
        
        'Other things, pretty self-explaining :)
        DoKeys
        CheckKing
        CheckBoss
        MoveShots
        Movesubs
        MoveBoss
        MovePlanes
        DoExplo
        DropBombs
        DoSigns
        PaintBoard

    Loop
End Sub

Sub Addammo()
    If Not P1.Ammo = MaxAmmo Then P1.Ammo = P1.Ammo + 1
End Sub
Sub DoKeys()
    With P1
        If Firsttime Then Firsttime = False: Exit Sub
        If GetAsyncKeyState(vbKeyLeft) <> 0 Then
            If .Speed >= -1 Then .Speed = -1.5
            .Speed = .Speed * 1.06
            .Dire = 1
        ElseIf GetAsyncKeyState(vbKeyRight) <> 0 Then
            If .Speed <= 1 Then .Speed = 1.5
            .Speed = .Speed * 1.06
            .Dire = 2
        Else
            .Speed = .Speed / 1.1
        End If
        
        If GetAsyncKeyState(vbKeySpace) <> 0 Then
            Fire
        End If
        
        If .Speed < -6 Then .Speed = -6
        If .Speed > 6 Then .Speed = 6
        .X = .X + .Speed
        
            
        If .X < 0 Then .X = 0
        If .X > Bredde - ShipBredde Then .X = Bredde - ShipBredde
    End With
End Sub
