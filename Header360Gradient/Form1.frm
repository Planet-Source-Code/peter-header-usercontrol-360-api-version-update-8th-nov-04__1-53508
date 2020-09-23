VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Header Examples..."
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin Header.HeaderInfo HeaderInfo8 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Header Info 3.6.0."
      Top             =   360
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1296
      BorderColor     =   12582912
      Caption         =   "Header Control"
      CapAlign        =   2
      CapStyle        =   1
      ForeColorShadow =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   9112067
      GradientAngle   =   77
      GradientStart   =   15188135
      GradientFinish  =   8388608
      ImgAlign        =   1
      Image           =   "Form1.frx":0000
      ImgSize         =   32
   End
   Begin Header.HeaderInfo HeaderInfo7 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Header Info 3.6.0."
      Top             =   1440
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1296
      BorderColor     =   0
      Caption         =   "Administration Only   "
      CapAlign        =   2
      CapStyle        =   1
      ForeColorShadow =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      GradientAngle   =   245
      GradientStart   =   12632256
      GradientFinish  =   0
      Image           =   "Form1.frx":0712
      ImgSize         =   24
   End
   Begin Header.HeaderInfo HeaderInfo4 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Header Info 3.6.0."
      Top             =   6000
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1720
      BorderColor     =   128
      BorderVisible   =   0   'False
      BorderShape     =   2
      Caption         =   "Log Details"
      CapAlign        =   2
      ForeColorShadow =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   128
      GradientAngle   =   65
      GradientStart   =   128
      ImgAlign        =   4
      Image           =   "Form1.frx":0E24
      ImgSize         =   32
   End
   Begin Header.HeaderInfo HeaderInfo5 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Header Info 3.6.0."
      Top             =   5040
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1085
      BorderColor     =   16384
      BorderShape     =   0
      Caption         =   "      Recycle Options "
      CapAlign        =   2
      CapStyle        =   1
      ForeColorShadow =   32768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16384
      GradientAngle   =   45
      GradientStart   =   49152
      GradientFinish  =   16384
      ImgAlign        =   2
      Image           =   "Form1.frx":1276
      ImgSize         =   32
   End
   Begin Header.HeaderInfo HeaderInfo2 
      Height          =   495
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Header Info 3.6.0."
      Top             =   4200
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      BorderColor     =   0
      Caption         =   "Warning "
      CapAlign        =   2
      CapStyle        =   1
      ForeColorShadow =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   192
      GradientStart   =   16777215
      GradientFinish  =   8421504
      ImgAlign        =   1
      Image           =   "Form1.frx":3A28
      ImgSize         =   32
   End
   Begin Header.HeaderInfo HeaderInfo1 
      Height          =   465
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Header Info 3.6.0."
      Top             =   2520
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   820
      BorderColor     =   12582912
      Caption         =   "Header Control v3.6.0"
      CapAlign        =   2
      CapStyle        =   2
      ForeColorShadow =   8421504
      CornerSize      =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      GradientStart   =   15976092
      GradientFinish  =   16663610
      ImgAlign        =   1
      Image           =   "Form1.frx":3D42
      ImgSize         =   24
   End
   Begin Header.HeaderInfo HeaderInfo3 
      Height          =   495
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Header Info 3.6.0."
      Top             =   3360
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      BorderColor     =   0
      BorderVisible   =   0   'False
      Caption         =   "Uninstall Options"
      CapAlign        =   2
      CapStyle        =   1
      ForeColorShadow =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      GradientAngle   =   0
      GradientStart   =   8421504
      GradientFinish  =   14933984
      GradientFinishStyle=   1
      ImgAlign        =   3
      Image           =   "Form1.frx":4454
      ImgSize         =   24
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Small corners - Vertical Gradient"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Large corners - 245 Degree Gradient"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Large corners (Top Only) - 65 Degree Gradient"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   5760
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Large corners - Horizontal Gradient - Transparent Look"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Large corners - Vertical Gradient"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Square corners - Border visible"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   4800
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Large corners - 77 Degree Gradient"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
