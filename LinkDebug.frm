VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LinkTest(s.y)"
   ClientHeight    =   8955
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16530
   FillColor       =   &H00FF0000&
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   16530
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame12 
      Caption         =   "Voltage Output"
      ForeColor       =   &H000000FF&
      Height          =   3735
      Left            =   11160
      TabIndex        =   102
      Top             =   5160
      Width           =   5295
      Begin VB.TextBox Text1 
         Height          =   3255
         Left            =   120
         ScrollBars      =   2  'Vertical
         TabIndex        =   103
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "64CH Voltage"
      ForeColor       =   &H000000FF&
      Height          =   4095
      Left            =   14400
      TabIndex        =   93
      Top             =   960
      Width           =   2055
      Begin VB.CommandButton Command6 
         Caption         =   "2.ALL Channels  Output"
         Height          =   1215
         Left            =   120
         TabIndex        =   101
         Top             =   2760
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "1.SET RANGE"
         Height          =   1215
         Left            =   120
         TabIndex        =   100
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   360
         Index           =   1
         Left            =   120
         TabIndex        =   95
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   94
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label16 
         Caption         =   "RANGE X2"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   97
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "RANGE X1"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   96
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "SerialPort"
      ForeColor       =   &H000000FF&
      Height          =   4095
      Left            =   11160
      TabIndex        =   82
      Top             =   960
      Width           =   3135
      Begin VB.CommandButton Command4 
         Caption         =   "SEND"
         Height          =   495
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "OPEN"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   3480
         Width           =   1335
      End
      Begin VB.ComboBox Combo10 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1560
         TabIndex        =   92
         Top             =   2880
         Width           =   1335
      End
      Begin VB.ComboBox Combo9 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1560
         TabIndex        =   90
         Top             =   2280
         Width           =   1335
      End
      Begin VB.ComboBox Combo8 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1560
         TabIndex        =   88
         Top             =   1680
         Width           =   1335
      End
      Begin VB.ComboBox Combo7 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1560
         TabIndex        =   86
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox Combo6 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1560
         TabIndex        =   84
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Label14"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   91
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   89
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Label12"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   87
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   85
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Label10"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   83
         Top             =   480
         Width           =   975
      End
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   15720
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Frame Frame5 
      Caption         =   "SENSOR Detect Board"
      ForeColor       =   &H000000FF&
      Height          =   1575
      Left            =   120
      TabIndex        =   41
      Top             =   7320
      Width           =   10935
      Begin VB.Shape Shape9 
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   11
         Left            =   10320
         Shape           =   3  'Circle
         Top             =   960
         Width           =   375
      End
      Begin VB.Shape Shape9 
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   10
         Left            =   9480
         Shape           =   3  'Circle
         Top             =   960
         Width           =   375
      End
      Begin VB.Shape Shape9 
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   9
         Left            =   8520
         Shape           =   3  'Circle
         Top             =   960
         Width           =   375
      End
      Begin VB.Shape Shape9 
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   8
         Left            =   7560
         Shape           =   3  'Circle
         Top             =   960
         Width           =   375
      End
      Begin VB.Shape Shape9 
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   7
         Left            =   6600
         Shape           =   3  'Circle
         Top             =   960
         Width           =   375
      End
      Begin VB.Shape Shape9 
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   6
         Left            =   5640
         Shape           =   3  'Circle
         Top             =   960
         Width           =   375
      End
      Begin VB.Shape Shape9 
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   5
         Left            =   4800
         Shape           =   3  'Circle
         Top             =   960
         Width           =   375
      End
      Begin VB.Shape Shape9 
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   4
         Left            =   3840
         Shape           =   3  'Circle
         Top             =   960
         Width           =   375
      End
      Begin VB.Shape Shape9 
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   3
         Left            =   2880
         Shape           =   3  'Circle
         Top             =   960
         Width           =   375
      End
      Begin VB.Shape Shape9 
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   2
         Left            =   2040
         Shape           =   3  'Circle
         Top             =   960
         Width           =   375
      End
      Begin VB.Shape Shape9 
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   1
         Left            =   1200
         Shape           =   3  'Circle
         Top             =   960
         Width           =   375
      End
      Begin VB.Shape Shape9 
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   0
         Left            =   240
         Shape           =   3  'Circle
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   375
         Index           =   11
         Left            =   10320
         TabIndex        =   53
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   375
         Index           =   10
         Left            =   9480
         TabIndex        =   52
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   375
         Index           =   9
         Left            =   8520
         TabIndex        =   51
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   375
         Index           =   8
         Left            =   7680
         TabIndex        =   50
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   375
         Index           =   7
         Left            =   6720
         TabIndex        =   49
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   375
         Index           =   6
         Left            =   5760
         TabIndex        =   48
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   375
         Index           =   5
         Left            =   4920
         TabIndex        =   47
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   375
         Index           =   4
         Left            =   3960
         TabIndex        =   46
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   375
         Index           =   3
         Left            =   3000
         TabIndex        =   45
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   375
         Index           =   2
         Left            =   2160
         TabIndex        =   44
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   43
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   42
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "USB Switch Board"
      ForeColor       =   &H000000FF&
      Height          =   2175
      Left            =   120
      TabIndex        =   40
      Top             =   5160
      Width           =   10935
      Begin VB.Frame Frame9 
         Caption         =   "USB2"
         ForeColor       =   &H00FF0000&
         Height          =   1695
         Left            =   8280
         TabIndex        =   57
         Top             =   360
         Width           =   2415
         Begin VB.CommandButton usb2 
            Caption         =   "DISC"
            Height          =   375
            Index           =   2
            Left            =   1560
            TabIndex        =   81
            Top             =   1200
            Width           =   735
         End
         Begin VB.CommandButton usb2 
            Caption         =   "OFF"
            Height          =   375
            Index           =   1
            Left            =   840
            TabIndex        =   65
            Top             =   1200
            Width           =   615
         End
         Begin VB.CommandButton usb2 
            Caption         =   "ON"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   64
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label9 
            Caption         =   "DISC"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   1800
            TabIndex        =   77
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label9 
            Caption         =   "D2"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   1080
            TabIndex        =   76
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label9 
            Caption         =   "D1"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   360
            TabIndex        =   75
            Top             =   360
            Width           =   255
         End
         Begin VB.Shape Shape13 
            Height          =   375
            Index           =   2
            Left            =   1800
            Shape           =   3  'Circle
            Top             =   720
            Width           =   495
         End
         Begin VB.Shape Shape13 
            Height          =   375
            Index           =   1
            Left            =   1080
            Shape           =   3  'Circle
            Top             =   720
            Width           =   375
         End
         Begin VB.Shape Shape13 
            Height          =   375
            Index           =   0
            Left            =   240
            Shape           =   3  'Circle
            Top             =   720
            Width           =   375
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "USB1"
         ForeColor       =   &H00FF0000&
         Height          =   1695
         Left            =   5520
         TabIndex        =   56
         Top             =   360
         Width           =   2655
         Begin VB.CommandButton usb1 
            Caption         =   "DISC"
            Height          =   375
            Index           =   2
            Left            =   1800
            TabIndex        =   80
            Top             =   1200
            Width           =   735
         End
         Begin VB.CommandButton usb1 
            Caption         =   "OFF"
            Height          =   375
            Index           =   1
            Left            =   960
            TabIndex        =   63
            Top             =   1200
            Width           =   735
         End
         Begin VB.CommandButton usb1 
            Caption         =   "ON"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   62
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label9 
            Caption         =   "DISC"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   2040
            TabIndex        =   74
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label9 
            Caption         =   "D2"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   1200
            TabIndex        =   73
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label9 
            Caption         =   "D1"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   360
            TabIndex        =   72
            Top             =   360
            Width           =   255
         End
         Begin VB.Shape Shape12 
            Height          =   375
            Index           =   2
            Left            =   2040
            Shape           =   3  'Circle
            Top             =   720
            Width           =   375
         End
         Begin VB.Shape Shape12 
            Height          =   375
            Index           =   1
            Left            =   1080
            Shape           =   3  'Circle
            Top             =   720
            Width           =   495
         End
         Begin VB.Shape Shape12 
            Height          =   375
            Index           =   0
            Left            =   240
            Shape           =   3  'Circle
            Top             =   720
            Width           =   375
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "USART2"
         ForeColor       =   &H00FF0000&
         Height          =   1695
         Left            =   2760
         TabIndex        =   55
         Top             =   360
         Width           =   2655
         Begin VB.CommandButton usart2 
            Caption         =   "DISC"
            Height          =   375
            Index           =   2
            Left            =   1800
            TabIndex        =   79
            Top             =   1200
            Width           =   735
         End
         Begin VB.CommandButton usart2 
            Caption         =   "OFF"
            Height          =   375
            Index           =   1
            Left            =   960
            TabIndex        =   61
            Top             =   1200
            Width           =   735
         End
         Begin VB.CommandButton usart2 
            Caption         =   "ON"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   60
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label9 
            Caption         =   "DISC"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   2040
            TabIndex        =   71
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label9 
            Caption         =   "D2"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   1200
            TabIndex        =   70
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label9 
            Caption         =   "D1"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   480
            TabIndex        =   69
            Top             =   360
            Width           =   255
         End
         Begin VB.Shape Shape11 
            Height          =   375
            Index           =   3
            Left            =   2040
            Shape           =   3  'Circle
            Top             =   720
            Width           =   375
         End
         Begin VB.Shape Shape11 
            Height          =   375
            Index           =   2
            Left            =   0
            Shape           =   3  'Circle
            Top             =   -720
            Width           =   375
         End
         Begin VB.Shape Shape11 
            Height          =   375
            Index           =   1
            Left            =   1080
            Shape           =   3  'Circle
            Top             =   720
            Width           =   615
         End
         Begin VB.Shape Shape11 
            Height          =   375
            Index           =   0
            Left            =   360
            Shape           =   3  'Circle
            Top             =   720
            Width           =   375
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "USART1"
         ForeColor       =   &H00FF0000&
         Height          =   1695
         Left            =   120
         TabIndex        =   54
         Top             =   360
         Width           =   2535
         Begin VB.CommandButton usart1 
            Caption         =   "DISC"
            Height          =   360
            Index           =   2
            Left            =   1680
            TabIndex        =   78
            Top             =   1200
            Width           =   735
         End
         Begin VB.CommandButton usart1 
            Caption         =   "OFF"
            Height          =   360
            Index           =   1
            Left            =   960
            TabIndex        =   59
            Top             =   1200
            Width           =   615
         End
         Begin VB.CommandButton usart1 
            Caption         =   "ON"
            Height          =   360
            Index           =   0
            Left            =   120
            TabIndex        =   58
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label9 
            Caption         =   "DISC"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1920
            TabIndex        =   68
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label9 
            Caption         =   "D2"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   67
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label9 
            Caption         =   "D1"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   66
            Top             =   360
            Width           =   255
         End
         Begin VB.Shape Shape10 
            Height          =   375
            Index           =   2
            Left            =   1920
            Shape           =   3  'Circle
            Top             =   720
            Width           =   495
         End
         Begin VB.Shape Shape10 
            Height          =   375
            Index           =   1
            Left            =   1080
            Shape           =   3  'Circle
            Top             =   720
            Width           =   495
         End
         Begin VB.Shape Shape10 
            Height          =   375
            Index           =   0
            Left            =   240
            Shape           =   3  'Circle
            Top             =   720
            Width           =   495
         End
      End
   End
   Begin VB.CommandButton relay2 
      Caption         =   "OFF"
      Height          =   495
      Index           =   1
      Left            =   4080
      TabIndex        =   27
      Top             =   4320
      Width           =   735
   End
   Begin VB.Frame Frame3 
      Caption         =   "RELAY Board Status"
      ForeColor       =   &H000000FF&
      Height          =   1815
      Left            =   2880
      TabIndex        =   15
      Top             =   960
      Width           =   8175
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   375
         Index           =   7
         Left            =   7440
         TabIndex        =   23
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   375
         Index           =   6
         Left            =   6480
         TabIndex        =   22
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   375
         Index           =   5
         Left            =   5520
         TabIndex        =   21
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   375
         Index           =   4
         Left            =   4440
         TabIndex        =   20
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   375
         Index           =   3
         Left            =   3480
         TabIndex        =   19
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   375
         Index           =   2
         Left            =   2520
         TabIndex        =   18
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   17
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   16
         Top             =   600
         Width           =   495
      End
      Begin VB.Shape Shape8 
         Height          =   615
         Left            =   7200
         Shape           =   5  'Rounded Square
         Top             =   1080
         Width           =   615
      End
      Begin VB.Shape Shape7 
         Height          =   615
         Left            =   6240
         Shape           =   5  'Rounded Square
         Top             =   1080
         Width           =   615
      End
      Begin VB.Shape Shape6 
         Height          =   615
         Left            =   5160
         Shape           =   5  'Rounded Square
         Top             =   1080
         Width           =   735
      End
      Begin VB.Shape Shape5 
         Height          =   615
         Left            =   4200
         Shape           =   5  'Rounded Square
         Top             =   1080
         Width           =   615
      End
      Begin VB.Shape Shape4 
         Height          =   615
         Left            =   3240
         Shape           =   5  'Rounded Square
         Top             =   1080
         Width           =   615
      End
      Begin VB.Shape Shape3 
         Height          =   615
         Left            =   2160
         Shape           =   5  'Rounded Square
         Top             =   1080
         Width           =   735
      End
      Begin VB.Shape Shape2 
         Height          =   615
         Left            =   1200
         Shape           =   4  'Rounded Rectangle
         Top             =   1080
         Width           =   615
      End
      Begin VB.Shape Shape1 
         Height          =   615
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   1080
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "RELAY Board Control"
      ForeColor       =   &H000000FF&
      Height          =   2175
      Left            =   2880
      TabIndex        =   14
      Top             =   2880
      Width           =   8175
      Begin VB.CommandButton relay8 
         Caption         =   "OFF"
         Height          =   495
         Index           =   1
         Left            =   7200
         TabIndex        =   39
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton relay8 
         Caption         =   "ON"
         Height          =   495
         Index           =   0
         Left            =   7200
         TabIndex        =   38
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton relay7 
         Caption         =   "OFF"
         Height          =   495
         Index           =   1
         Left            =   6240
         TabIndex        =   37
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton relay7 
         Caption         =   "ON"
         Height          =   495
         Index           =   0
         Left            =   6240
         TabIndex        =   36
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton relay6 
         Caption         =   "OFF"
         Height          =   495
         Index           =   1
         Left            =   5160
         TabIndex        =   35
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton relay6 
         Caption         =   "ON"
         Height          =   495
         Index           =   0
         Left            =   5160
         TabIndex        =   34
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton relay5 
         Caption         =   "OFF"
         Height          =   495
         Index           =   1
         Left            =   4200
         TabIndex        =   33
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton relay5 
         Caption         =   "ON"
         Height          =   495
         Index           =   0
         Left            =   4200
         TabIndex        =   32
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton relay4 
         Caption         =   "OFF"
         Height          =   480
         Index           =   1
         Left            =   3240
         TabIndex        =   31
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton relay4 
         Caption         =   "ON"
         Height          =   480
         Index           =   0
         Left            =   3240
         TabIndex        =   30
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton relay3 
         Caption         =   "OFF"
         Height          =   495
         Index           =   1
         Left            =   2160
         TabIndex        =   29
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton relay3 
         Caption         =   "ON"
         Height          =   495
         Index           =   0
         Left            =   2160
         TabIndex        =   28
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton relay2 
         Caption         =   "ON"
         Height          =   495
         Index           =   0
         Left            =   1200
         TabIndex        =   26
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton realy1 
         Caption         =   "OFF"
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   25
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton realy1 
         Caption         =   "ON"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   600
         Width           =   735
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Frame Frame1 
      Caption         =   "SerialPort"
      ForeColor       =   &H000000FF&
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2655
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   495
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000004&
         Caption         =   "Command1"
         Height          =   495
         Left            =   120
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3480
         Width           =   1215
      End
      Begin VB.ComboBox Combo5 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1200
         TabIndex        =   11
         Top             =   2880
         Width           =   1095
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1200
         TabIndex        =   9
         Top             =   2280
         Width           =   1095
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1200
         TabIndex        =   7
         Top             =   1680
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1200
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "LinkDebug.frx":0000
         Left            =   1200
         List            =   "LinkDebug.frx":0002
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   14655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()

If (Command1.Caption = "OPEN") Then
  MSComm1.CommPort = Val(Combo1.Text)
  If MSComm1.PortOpen Then
  MSComm1.PortOpen = False
   MsgBox "COM port had been opened!", vbOKOnly + vbCritical + vbDefaultButton1, "Error"
  End If
  With MSComm1
    .Settings = Combo2.Text & "," & Mid(Combo4.Text, 1, 1) & "," & Combo3.Text & "," & Combo5.Text  '这里用"+"和用"&"的作用是一样的，都可以用来连接
    .InputLen = 0
    .InBufferSize = 1
    .RThreshold = 1
    .InputMode = comInputModeBinary
    .InBufferCount = 0
    End With
 '***************************************************************************
 
 '       串口的初始化各参数设置
 
 '**************************************************************************
  MSComm1.PortOpen = True
  Command1.Caption = "CLOSE"
  Command1.BackColor = &HFF&
Else
  MSComm1.PortOpen = False
  Command1.Caption = "OPEN"
  Command1.BackColor = &H0&
End If

End Sub

Private Sub Command3_Click()
If (Command3.Caption = "OPEN") Then
  MSComm2.CommPort = Val(Combo6.Text)
  If MSComm2.PortOpen Then
  MSComm2.PortOpen = False
   MsgBox "COM port had been opened!", vbOKOnly + vbCritical + vbDefaultButton1, "Error"
  End If
  With MSComm2
    .Settings = Combo7.Text & "," & Mid(Combo9.Text, 1, 1) & "," & Combo8.Text & "," & Combo10.Text  '这里用"+"和用"&"的作用是一样的，都可以用来连接
    .InputLen = 0
    .InBufferSize = 1
    .RThreshold = 1
    .InputMode = comInputModeBinary
    .InBufferCount = 0
    End With
 '***************************************************************************
 
 '       串口的初始化各参数设置
 
 '**************************************************************************
  MSComm2.PortOpen = True
  Command3.Caption = "CLOSE"
  Command3.BackColor = &HFF&
Else
  MSComm2.PortOpen = False
  Command3.Caption = "OPEN"
  Command3.BackColor = &H0&
End If
End Sub

Private Sub Form_Load()

Dim i As Integer

Label1.Caption = "Link Test Kit Debug"
Label2.Caption = "Port"
Label3.Caption = "Baud"
Label4.Caption = "DataBits"
Label5.Caption = "Parity"
Label6.Caption = "StopBit"

Combo2.AddItem "9600"
Combo2.AddItem "115200"
Combo2.AddItem "921600"

Combo3.AddItem ("8")

For i = 1 To 16
Combo1.AddItem CStr(i)
Next i

Combo4.AddItem "None"
Combo4.AddItem "Odd"
Combo4.AddItem "Even"
Combo5.AddItem "1"
Combo5.AddItem "2"

Command1.Caption = "OPEN"
Command1.BackColor = &HFF&
Command2.Caption = "SEND"
Command2.BackColor = &HFF&

Command3.Caption = "OPEN"
Command3.BackColor = &HFF&
Command4.Caption = "SEND"
Command4.BackColor = &HFF&


For i = 0 To 7
 Label7(i).Caption = i + 1
Next i

Shape1.BackColor = &H0&
Shape1.FillStyle = 0
Shape2.BackColor = &H0&
Shape2.FillStyle = 0
Shape3.BackColor = &H0&
Shape3.FillStyle = 0
Shape4.BackColor = &H0&
Shape4.FillStyle = 0
Shape5.BackColor = &H0&
Shape5.FillStyle = 0
Shape6.BackColor = &H0&
Shape6.FillStyle = 0
Shape7.BackColor = &H0&
Shape7.FillStyle = 0
Shape8.BackColor = &H0&
Shape8.FillStyle = 0

Shape10(0).BackColor = &H0&
Shape10(0).FillStyle = 0
Shape10(1).BackColor = &H0&
Shape10(1).FillStyle = 0
Shape10(2).BackColor = &H0&
Shape10(2).FillStyle = 0

Shape11(0).BackColor = &H0&
Shape11(0).FillStyle = 0
Shape11(1).BackColor = &H0&
Shape11(1).FillStyle = 0
Shape11(3).BackColor = &H0&
Shape11(3).FillStyle = 0

Shape12(0).BackColor = &H0&
Shape12(0).FillStyle = 0
Shape12(1).BackColor = &H0&
Shape12(1).FillStyle = 0
Shape12(2).BackColor = &H0&
Shape12(2).FillStyle = 0

Shape13(0).BackColor = &H0&
Shape13(0).FillStyle = 0
Shape13(1).BackColor = &H0&
Shape13(1).FillStyle = 0
Shape13(2).BackColor = &H0&
Shape13(2).FillStyle = 0

For i = 0 To 11
 Label8(i).Caption = i + 1
 Shape9(i).BackColor = &H0&
 Shape9(i).FillStyle = 0
Next i





End Sub

Sub InitRs232() '初始化串口副程序
   On Error Resume Next
   If MSComm1.PortOpen Then MSComm1.PortOpen = False '如果串口为打开状态则关闭它
   With MSComm1 '宣告MsComm的结构体
      .CommPort = Combo1.Text
      .Settings = Combo2.Text & "," & Mid(Combo4.Text, 1, 1) & "," & Combo3.Text & "," & Combo5.Text  '这里用"+"和用"&"的作用是一样的，都可以用来连接 '设定通讯协议 9600波特率,无奇偶校验,8位数据,一个停止位
      .InputLen = 0 '设置Input一次从接收缓冲读取
      .InBufferSize = 1 '设置缓冲区接收数据为1字节
      .RThreshold = 1 '设置接收一个字节就产生OnComm事件
      .InputMode = comInputModeBinary '设定接收模式为文字模式
      .InBufferCount = 0 '缓冲区清空
   End With
   MSComm1.PortOpen = True
End Sub





Private Sub MSComm1_OnComm()

    Dim indata As String
    Dim WChar(100) As String
    Dim bte(100) As Variant
    Call Sleep(50)
    If MSComm1.CommEvent = 2 Then
        MSComm1.RThreshold = 0
        Dim j
        For j = 1 To MSComm1.InBufferCount
            SwichVar j
            If Check3.Value = 1 Then
                Text2.Text = Text2.Text & Right("00" + Hex(out(j)), 2)
            Else
                Text2.Text = Text2.Text & Chr(out(j))
            End If
            Text2.Text = Text2.Text & " "
            rnum = rnum + 1
        Next j
        Label12.Caption = rnum
    End If
    mscSerialPort.RThreshold = 1

End Sub

Private Sub realy1_Click(Index As Integer)
    
 Dim Relay(10) As Byte
 Relay(0) = &H55
 Relay(1) = &HB
 Relay(2) = &H1
 Relay(3) = &H0
 Relay(4) = &H0
 Relay(5) = &H0
 Relay(6) = &H0
 Relay(7) = &H0
 Relay(8) = &HAA
 Relay(9) = &H1
 Relay(10) = &HB

 MSComm1.Output = Relay
 
End Sub
