VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LinkTest(simon.y)"
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
   Icon            =   "LinkDebug.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   16530
   StartUpPosition =   3  '窗口缺省
   Begin MSCommLib.MSComm MSComm3 
      Left            =   1320
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   15720
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame Frame12 
      Caption         =   "Voltage Output"
      ForeColor       =   &H000000FF&
      Height          =   3735
      Left            =   11160
      TabIndex        =   102
      Top             =   5160
      Width           =   5295
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   120
         MultiLine       =   -1  'True
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
         Caption         =   "1.SET RANGE FOR ALL CHs"
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
         Caption         =   "CLEAR"
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
         Text            =   "1"
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
         Text            =   "None"
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
         Text            =   "8"
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
         Text            =   "115200"
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
         ItemData        =   "LinkDebug.frx":25CA
         Left            =   1560
         List            =   "LinkDebug.frx":25CC
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
      Left            =   600
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
         Width           =   2535
         Begin VB.CommandButton usb2 
            Caption         =   "DISC"
            Height          =   375
            Index           =   2
            Left            =   1680
            TabIndex        =   81
            Top             =   1200
            Width           =   735
         End
         Begin VB.CommandButton usb2 
            Caption         =   "OFF"
            Height          =   375
            Index           =   1
            Left            =   960
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
      Begin VB.Shape Shape8 
         BorderColor     =   &H80000006&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   7200
         Shape           =   5  'Rounded Square
         Top             =   1080
         Width           =   615
      End
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
         Text            =   "1"
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
         Text            =   "None"
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
         Text            =   "8"
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
         Text            =   "115200"
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
         ItemData        =   "LinkDebug.frx":25CE
         Left            =   1200
         List            =   "LinkDebug.frx":25D0
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
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   14175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Pre_Relay(10) As Byte                                   '之前的状态
Dim Relay(10) As Byte                                       '当前的状态
Dim ACK_Relay() As Byte                                     'Relay的ACK

Dim USB_Board(10) As Byte
Dim Pre_USB_Board(10) As Byte
Dim ACK_USB_Board() As Byte

Dim Sensor_Status(11) As Byte
Dim CRC16(1) As Byte

Dim SetVoltageRange(12) As Byte
Dim GetAllChannelVoltage(12) As Byte
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub Cal_CRC16(dat() As Byte, crc() As Byte, length As Integer)

Dim temp As Long
Dim i As Integer

For i = 0 To length - 1
 temp = temp + dat(i)
Next i

temp = Val("&H" & Hex(temp))
crc(1) = (temp And &HFF00) / 256                             '环境会进行四舍五入
crc(0) = temp Mod 256

End Sub

'1.数据包解包
'2.获取有效数据位
'3.data - Type | Data3 | Data2 | Data1 | Data0
Public Sub DataPackageProcess(src_pack() As Byte, valid_data As Boolean, data() As Byte)
Dim i As Integer
Dim temp(8) As Byte
Dim crc(1) As Byte

Call Copy_Dat(temp, src_pack, 9)

If ((src_pack(0) <> &H55) Or (src_pack(1) <> &HB) Or (src_pack(2) <> &H1) Or (src_pack(8) <> &HAA)) Then
 valid_data = False
End If

If ((src_pack(0) = &H55) And (src_pack(1) = &HB) And (src_pack(2) = &H1) And (src_pack(8) = &HAA)) Then
 Call Cal_CRC16(temp, crc, 9)
 If ((crc(1) = src_pack(9)) And (crc(0) = src_pack(10))) Then
    valid_data = True
    For i = 3 To 7
        data(i - 3) = src_pack(i)
    Next i
 End If
End If

End Sub

Public Sub Check_Combo1_Repeat()
    Dim i, k As Integer
    Dim can As String
    Dim san As String
    For i = 0 To Combo1.ListCount - 1
        For k = i + 1 To Combo1.ListCount - 1
            can = Combo1.List(i)
            san = Combo1.List(k)
        If can = san Then Combo1.RemoveItem (k)
    Next k
    Next i
End Sub

Public Sub Check_Combo6_Repeat()
    Dim i, k As Integer
    Dim can As String
    Dim san As String
    For i = 0 To Combo6.ListCount - 1
        For k = i + 1 To Combo6.ListCount - 1
            can = Combo6.List(i)
            san = Combo6.List(k)
        If can = san Then Combo6.RemoveItem (k)
    Next k
    Next i
End Sub

Private Sub Command1_Click()

If (Command1.Caption = "OPEN") Then                                 '可以打开
 MSComm1.CommPort = Val(Mid(Combo1.Text, 4, 1))
 With MSComm1
    .Settings = Combo2.Text & "," & Mid(Combo4.Text, 1, 1) & "," & Combo3.Text & "," & Combo5.Text  '这里用"+"和用"&"的作用是一样的，都可以用来连接
    .InputLen = 0
    .InBufferSize = 1
    .RThreshold = 1
    .InputMode = comInputModeBinary
    .InBufferCount = 0
    End With
 MSComm1.PortOpen = True
 Command1.Caption = "CLOSE"
 Command1.BackColor = &HFF00&

ElseIf (Command1.Caption = "CLOSE") Then                            '可以关闭
  MSComm1.PortOpen = False
  Command1.Caption = "OPEN"
  Command1.BackColor = &HFF&
End If

End Sub

Private Sub Command3_Click()
If (Command3.Caption = "OPEN") Then                                 '可以打开
 MSComm2.CommPort = Val(Mid(Combo6.Text, 4, 1))
 With MSComm2
    .Settings = Combo7.Text & "," & Mid(Combo9.Text, 1, 1) & "," & Combo8.Text & "," & Combo10.Text  '这里用"+"和用"&"的作用是一样的，都可以用来连接
    .InputLen = 0
    .InBufferSize = 1
    .RThreshold = 1
    .InputMode = comInputModeText
    .InBufferCount = 0
    End With
 MSComm2.PortOpen = True
 Command3.Caption = "CLOSE"
 Command3.BackColor = &HFF00&

ElseIf (Command3.Caption = "CLOSE") Then                            '可以关闭
  MSComm2.PortOpen = False
  Command3.Caption = "OPEN"
  Command3.BackColor = &HFF&
End If
End Sub

Private Sub Command4_Click()
 Text1.Text = ""
End Sub

Private Sub Command5_Click()
    Dim crc(1) As Byte

    Command6.Enabled = True
    If ((Check1(0).Value = 1) And ((Check1(1).Value = 1))) Then
        MsgBox "无法实现同时选择两种量程！", vbOKOnly + vbInformation, "提示信息!" '为用户提示出错信息
    ElseIf ((Check1(0).Value = 0) And ((Check1(1).Value = 0))) Then
        MsgBox "请至少选择一种量程！", vbOKOnly + vbInformation, "提示信息!" '为用户提示出错信息
    ElseIf (Check1(0).Value = 1) Then
        SetVoltageRange(0) = &H55
        SetVoltageRange(1) = &HD
        SetVoltageRange(2) = &HA
        SetVoltageRange(3) = &H0
        SetVoltageRange(4) = &H0
        SetVoltageRange(5) = &H0
        SetVoltageRange(6) = &H0
        SetVoltageRange(7) = &H0
        SetVoltageRange(8) = &H0
        SetVoltageRange(9) = &H0
        SetVoltageRange(10) = &H0
        Call Cal_CRC16(SetVoltageRange, crc, 11)
        SetVoltageRange(11) = crc(1)
        SetVoltageRange(12) = crc(0)
    ElseIf (Check1(1).Value = 1) Then
        SetVoltageRange(0) = &H55
        SetVoltageRange(1) = &HD
        SetVoltageRange(2) = &HA
        SetVoltageRange(3) = &HFF
        SetVoltageRange(4) = &HFF
        SetVoltageRange(5) = &HFF
        SetVoltageRange(6) = &HFF
        SetVoltageRange(7) = &HFF
        SetVoltageRange(8) = &HFF
        SetVoltageRange(9) = &HFF
        SetVoltageRange(10) = &HFF
        Call Cal_CRC16(SetVoltageRange, crc, 11)
        SetVoltageRange(11) = crc(1)
        SetVoltageRange(12) = crc(0)
    End If
    MSComm2.Output = SetVoltageRange
End Sub

Private Sub Command6_Click()
    Dim crc(1) As Byte
    GetAllChannelVoltage(0) = &H55
    GetAllChannelVoltage(1) = &HD
    GetAllChannelVoltage(2) = &HB
    GetAllChannelVoltage(3) = &HFF
    GetAllChannelVoltage(4) = &HFF
    GetAllChannelVoltage(5) = &HFF
    GetAllChannelVoltage(6) = &HFF
    GetAllChannelVoltage(7) = &HFF
    GetAllChannelVoltage(8) = &HFF
    GetAllChannelVoltage(9) = &HFF
    GetAllChannelVoltage(10) = &HFF
    Call Cal_CRC16(GetAllChannelVoltage, crc, 11)
    GetAllChannelVoltage(11) = crc(1)
    GetAllChannelVoltage(12) = crc(0)

    MSComm2.Output = GetAllChannelVoltage
End Sub

Private Sub Form_Load()

Dim i As Integer

Label1.Caption = "Link Test Debug Kit V1.1"
Label2.Caption = "Port"
Label3.Caption = "Baud"
Label4.Caption = "DataBits"
Label5.Caption = "Parity"
Label6.Caption = "StopBit"

Label10.Caption = "Port"
Label11.Caption = "Baud"
Label12.Caption = "DataBits"
Label13.Caption = "Parity"
Label14.Caption = "StopBit"

Combo2.AddItem "115200"
Combo2.AddItem "921600"
Combo3.AddItem ("8")
Combo4.AddItem "None"
Combo4.AddItem "Odd"
Combo4.AddItem "Even"
Combo5.AddItem "1"

Command1.Caption = "OPEN"
Command1.BackColor = &HFF&
Command2.Caption = "SEND"
Command2.BackColor = &HFF&

Command3.Caption = "OPEN"
Command3.BackColor = &HFF&
Command4.Caption = "CLEAR"
Command4.BackColor = &HC000C0

Command2.Enabled = False
Command6.Enabled = False

For i = 0 To 10
    Relay(i) = 0
    Pre_Relay(i) = 0
Next i

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

'Init relay status
Relay(0) = &H55
Relay(1) = &HB
Relay(2) = &H1
Relay(3) = &H0                                               'Type -- RELAY
Relay(4) = &H0
Relay(5) = &H0
Relay(6) = &H0
Relay(7) = &H0
Relay(8) = &HAA
Call Copy_Dat(Pre_Relay, Relay, 9)


'Init USB_Board status
USB_Board(0) = &H55
USB_Board(1) = &HB
USB_Board(2) = &H1
USB_Board(3) = &H1                                           'Type -- RELAY
USB_Board(4) = &H0
USB_Board(5) = &H0
USB_Board(6) = &H0
USB_Board(7) = &H0
USB_Board(8) = &HAA
Call Copy_Dat(Pre_USB_Board, USB_Board, 9)
Call RecognizeCOM

'Init Sensor board status
For i = 0 To 11
    Sensor_Status(i) = 0
Next i

End Sub

Sub RecognizeCOM() '自动识别COM Port
    Dim i As Integer
    Dim item_index As Integer
    Dim j As Integer
    
    j = 0

    For i = 1 To 32 Step 1
        If MSComm1.PortOpen = True Then                 '先关闭串口
        MSComm1.PortOpen = False
        End If
        MSComm1.CommPort = i
        On Error Resume Next                            '说明当一个运行时错误发生时，控件转到紧接着发生错误的语句之后的语句，并在此继续运行。访问对象时要使用这种形式而不使用 On Error GoTo。
        MSComm1.PortOpen = True
        If Err.Number <> 8002 Then                      '无效的串口号。这样可以检测到虚拟串口，如果用Err.Number = 0的话检测不到虚拟串口
        If j = 0 Then
        j = i
        End If
    
        Combo1.AddItem "COM" & i                        '生成串口选择列表
        End If
        MSComm1.PortOpen = False
    Next i
    If j >= 1 Then
    Combo1.Text = "COM" & j                         '自动打开可用的最小串口号
    MSComm1.CommPort = j
    If Err.Number = 8005 Then                        '串口已打开,vbExclamation '
    MSComm1.PortOpen = False
    Combo1.Text = ""
    Command1.Caption = "OPEN"
    Command1.BackColor = &HFF&                      'Red
    End If
    End If
    Call Check_Combo1_Repeat
    
    j = 0
    For i = 1 To 32 Step 1
    If MSComm2.PortOpen = True Then                 '先关闭串口
    MSComm2.PortOpen = False
    End If
    MSComm2.CommPort = i
    On Error Resume Next                            '说明当一个运行时错误发生时，控件转到紧接着发生错误的语句之后的语句，并在此继续运行。访问对象时要使用这种形式而不使用 On Error GoTo。
    MSComm2.PortOpen = True
    If Err.Number <> 8002 Then                      '无效的串口号。这样可以检测到虚拟串口，如果用Err.Number = 0的话检测不到虚拟串口
    If j = 0 Then
    j = i
    End If
    Combo6.AddItem "COM" & i                         '生成串口选择列表
    End If
    MSComm2.PortOpen = False
    Next i
    If j >= 1 Then
    Combo6.Text = "COM" & j                        '自动打开可用的最小串口号
    MSComm2.CommPort = j
    If Err.Number = 8005 Then                       '串口已打开,vbExclamation '
    MSComm2.PortOpen = False
    Combo6.Text = ""
    Command3.Caption = "OPEN"
    Command3.BackColor = &HFF&                      'Red
    End If
    End If
    Call Check_Combo6_Repeat
    
End Sub

Public Sub Copy_Dat(pre() As Byte, cur() As Byte, length As Integer) '数据Copy到数组中

    Dim i As Integer
    i = length
    For i = 0 To (length - 1)
      pre(i) = cur(i)
    Next i
End Sub

Private Sub MSComm1_OnComm()

On Error GoTo MsComm_OnCommErr

    Dim RecvCount As Integer
    Dim TempBytes() As Byte
    Dim DatIndex As Integer
    Dim RecvBytes(10) As Byte
    Dim i As Integer
    Dim DataProcess_Flag As Boolean
    Dim data_valid As Boolean
    Dim data_afterprocess(4) As Byte
    
    Select Case MSComm1.CommEvent
        Case comEvReceive
            RecvCount = MSComm1.InBufferCount
            ReDim TempBytes(RecvCount - 1)
            TempBytes = MSComm1.Input
         
            For i = 0 To (RecvCount - 1)
                RecvBytes(DatIndex) = TempBytes(i)
                DatIndex = DatIndex + 1
            Next i
        
            If DatIndex >= 11 Then
                DatIndex = 0
                DataProcess_Flag = True
            End If
    
        If DataProcess_Flag = True Then
            DataProcess_Flag = False
            Call DataPackageProcess(RecvBytes, data_valid, data_afterprocess)
            If (data_valid = True) Then                              ' data valid
                Select Case data_afterprocess(0)
                    Case &H10                                        ' relay board
                        If (data_afterprocess(4) And &H1) Then
                            Shape1.FillColor = &HFF&        'red
                        Else
                            Shape1.FillColor = &H80000008   'black
                        End If
                
                        If (data_afterprocess(4) And &H2) Then
                            Shape2.FillColor = &HFF&        'red
                        Else
                            Shape2.FillColor = &H80000008   'black
                        End If
                
                        If (data_afterprocess(4) And &H4) Then
                            Shape3.FillColor = &HFF&        'red
                        Else
                            Shape3.FillColor = &H80000008   'black
                        End If
                
                        If (data_afterprocess(4) And &H8) Then
                            Shape4.FillColor = &HFF&        'red
                        Else
                            Shape4.FillColor = &H80000008   'black
                        End If
                
                        If (data_afterprocess(4) And &H10) Then
                            Shape5.FillColor = &HFF&        'red
                        Else
                            Shape5.FillColor = &H80000008   'black
                        End If
                
                        If (data_afterprocess(4) And &H20) Then
                            Shape6.FillColor = &HFF&        'red
                        Else
                            Shape6.FillColor = &H80000008   'black
                        End If
                
                        If (data_afterprocess(4) And &H40) Then
                            Shape7.FillColor = &HFF&        'red
                        Else
                            Shape7.FillColor = &H80000008   'black
                        End If
                
                        If (data_afterprocess(4) And &H80) Then
                            Shape8.FillColor = &HFF&        'red
                        Else
                            Shape8.FillColor = &H80000008   'black
                        End If

                
                    Case &H11                                        ' usb board
                        If ((data_afterprocess(4) <> 0) Or (data_afterprocess(3) <> 0)) Then
                            'uart1
                            If ((data_afterprocess(4) And &H7) = &H3) Then
                                Shape10(0).FillColor = &HFF&        'red
                                Shape10(1).FillColor = &H80000008   'black
                                Shape10(2).FillColor = &H80000008   'black
                            ElseIf ((data_afterprocess(4) And &H7) = &H1) Then
                                Shape10(0).FillColor = &H80000008   'black
                                Shape10(1).FillColor = &HFF&        'red
                                Shape10(2).FillColor = &H80000008   'black
                            ElseIf ((data_afterprocess(4) And &H7) = &H5) Then
                                Shape10(0).FillColor = &H80000008   'black
                                Shape10(1).FillColor = &H80000008   'red
                                Shape10(2).FillColor = &HFF&        'black
                            'ElseIf (((data_afterprocess(4) And &H7) <> &H3) And ((data_afterprocess(4) And &H7) <> &H1) And ((data_afterprocess(4) And &H7) <> &H5)) Then
                             Else
                                Shape10(0).FillColor = &H80000008   'black
                                Shape10(1).FillColor = &H80000008   'black
                                Shape10(2).FillColor = &H80000008   'black
                            End If
                        
                            'uart2
                            If ((data_afterprocess(4) And &H38) = &H18) Then
                                Shape11(0).FillColor = &HFF&        'red
                                Shape11(1).FillColor = &H80000008   'black
                                Shape11(3).FillColor = &H80000008   'black
                            ElseIf ((data_afterprocess(4) And &H38) = &H8) Then
                                Shape11(0).FillColor = &H80000008   'black
                                Shape11(1).FillColor = &HFF&        'red
                                Shape11(3).FillColor = &H80000008   'black
                            ElseIf ((data_afterprocess(4) And &H38) = &H28) Then
                                Shape11(0).FillColor = &H80000008   'black
                                Shape11(1).FillColor = &H80000008   'black
                                Shape11(3).FillColor = &HFF&        'red
                            'ElseIf (((data_afterprocess(4) And &H38) <> &H18) And ((data_afterprocess(4) And &H38) <> &H8) And ((data_afterprocess(4) And &H38) <> &H28)) Then
                            Else
                                Shape11(0).FillColor = &H80000008   'black
                                Shape11(1).FillColor = &H80000008   'black
                                Shape11(3).FillColor = &H80000008   'black
                            End If
                        
                            'usb1
                            If (((data_afterprocess(4) And &HC0) = &HC0) And ((data_afterprocess(3) And &H1) = &H0)) Then
                                Shape12(0).FillColor = &HFF&             'red
                                Shape12(1).FillColor = &H80000008        'black
                                Shape12(2).FillColor = &H80000008        'black
                            ElseIf (((data_afterprocess(4) And &HC0) = &H40) And ((data_afterprocess(3) And &H1) = &H0)) Then
                                Shape12(0).FillColor = &H80000008        'black
                                Shape12(1).FillColor = &HFF&             'red
                                Shape12(2).FillColor = &H80000008        'black
                            ElseIf (((data_afterprocess(4) And &HC0) = &H40) And ((data_afterprocess(3) And &H1) = &H1)) Then
                                Shape12(0).FillColor = &H80000008        'black
                                Shape12(1).FillColor = &H80000008        'black
                                Shape12(2).FillColor = &HFF&             'red
                            Else
                                Shape12(0).FillColor = &H80000008        'black
                                Shape12(1).FillColor = &H80000008        'black
                                Shape12(2).FillColor = &H80000008        'black
                            End If
                        
                            'usb2
                            If ((data_afterprocess(3) And &HE) = &H6) Then
                                Shape13(0).FillColor = &HFF&        'red
                                Shape13(1).FillColor = &H80000008   'black
                                Shape13(2).FillColor = &H80000008   'black
                            ElseIf ((data_afterprocess(3) And &HE) = &H2) Then
                                Shape13(0).FillColor = &H80000008   'black
                                Shape13(1).FillColor = &HFF&        'red
                                Shape13(2).FillColor = &H80000008   'black
                            ElseIf ((data_afterprocess(3) And &HE) = &HA) Then
                                Shape13(0).FillColor = &H80000008   'black
                                Shape13(1).FillColor = &H80000008   'black
                                Shape13(2).FillColor = &HFF&        'red
                            Else
                                Shape13(0).FillColor = &H80000008   'black
                                Shape13(1).FillColor = &H80000008   'black
                                Shape13(2).FillColor = &H80000008   'black
                            End If
                        End If
                   
                    Case &H12                                        ' sensor board
                        
                        'Sensor - Channel1~12
                        If (data_afterprocess(4) And &H1) Then
                            Shape9(0).FillColor = &HFF&        'red
                        Else
                            Shape9(0).FillColor = &H80000008   'black
                        End If
                        
                        If (data_afterprocess(4) And &H2) Then
                            Shape9(1).FillColor = &HFF&        'red
                        Else
                            Shape9(1).FillColor = &H80000008   'black
                        End If
                        
                        If (data_afterprocess(4) And &H4) Then
                            Shape9(2).FillColor = &HFF&        'red
                        Else
                            Shape9(2).FillColor = &H80000008   'black
                        End If
                        
                        If (data_afterprocess(4) And &H8) Then
                            Shape9(3).FillColor = &HFF&        'red
                        Else
                            Shape9(3).FillColor = &H80000008   'black
                        End If
                        
                        If (data_afterprocess(4) And &H10) Then
                            Shape9(4).FillColor = &HFF&        'red
                        Else
                            Shape9(4).FillColor = &H80000008   'black
                        End If
                        
                        If (data_afterprocess(4) And &H20) Then
                            Shape9(5).FillColor = &HFF&        'red
                        Else
                            Shape9(5).FillColor = &H80000008   'black
                        End If
                        
                        If (data_afterprocess(4) And &H40) Then
                            Shape9(6).FillColor = &HFF&        'red
                        Else
                            Shape9(6).FillColor = &H80000008   'black
                        End If
                        
                        If (data_afterprocess(4) And &H80) Then
                            Shape9(7).FillColor = &HFF&        'red
                        Else
                            Shape9(7).FillColor = &H80000008   'black
                        End If
                        
                        If (data_afterprocess(3) And &H1) Then
                            Shape9(8).FillColor = &HFF&        'red
                        Else
                            Shape9(8).FillColor = &H80000008   'black
                        End If
                        
                        If (data_afterprocess(3) And &H2) Then
                            Shape9(9).FillColor = &HFF&        'red
                        Else
                            Shape9(9).FillColor = &H80000008   'black
                        End If
                        
                        If (data_afterprocess(3) And &H4) Then
                            Shape9(10).FillColor = &HFF&        'red
                        Else
                            Shape9(10).FillColor = &H80000008   'black
                        End If
                        
                        If (data_afterprocess(3) And &H8) Then
                            Shape9(11).FillColor = &HFF&        'red
                        Else
                            Shape9(11).FillColor = &H80000008   'black
                        End If
                    End Select
                End If
            End If
        End Select
    Exit Sub
  
MsComm_OnCommErr:
  If Err.Number <> 0 Then '错误处理程序
    MsgBox CStr(Err.Number) + Err.Description, vbOKOnly + vbInformation, "1提示信息!" '为用户提示出错信息
  End If
  Err.Clear
End Sub

Private Sub MSComm2_OnComm()

On Error GoTo MsComm_OnCommErr2

    Dim Recv2Count As Integer
    Dim TempString As String
    
    Select Case MSComm2.CommEvent
        Case comEvReceive
            Recv2Count = MSComm2.InBufferCount
            TempString = MSComm2.Input
            Text1.Text = Text1.Text + TempString
            Text1.SelStart = Len(Text1.Text)
    End Select
  
MsComm_OnCommErr2:
  If Err.Number <> 0 Then '错误处理程序
    MsgBox CStr(Err.Number) + Err.Description, vbOKOnly + vbInformation, "1提示信息!" '为用户提示出错信息
  End If
  Err.Clear
End Sub

Private Sub realy1_Click(Index As Integer)
     
 If (Index = 0) Then
 'SOF  LEN  LINK_TEST  RELAY/USB/SENSOR  DAT3  DAT2  DAT1  DAT0  EOF  CRC1  CRC0
       
 Relay(0) = &H55
 Relay(1) = &HB
 Relay(2) = &H1
 Relay(3) = &H0                                      'Type -- RELAY
 Relay(4) = &H0
 Relay(5) = &H0
 Relay(6) = &H0
 
 Relay(7) = Pre_Relay(7) Or &H1
 Relay(8) = &HAA
 Call Cal_CRC16(Relay, CRC16, 9)
 Relay(9) = CRC16(1)
 Relay(10) = CRC16(0)
 Call Copy_Dat(Pre_Relay, Relay, 9)
 MSComm1.Output = Relay
 Call Sleep(50)
 End If
 
 
  If (Index = 1) Then
 'SOF  LEN  LINK_TEST  RELAY/USB/SENSOR  DAT3  DAT2  DAT1  DAT0  EOF  CRC1  CRC0
       
 Relay(0) = &H55
 Relay(1) = &HB
 Relay(2) = &H1
 Relay(3) = &H0                                      'Type -- RELAY
 Relay(4) = &H0
 Relay(5) = &H0
 Relay(6) = &H0
 
 Relay(7) = Pre_Relay(7) And &HFE
 Relay(8) = &HAA
 Call Cal_CRC16(Relay, CRC16, 9)
 Relay(9) = CRC16(1)
 Relay(10) = CRC16(0)
 Call Copy_Dat(Pre_Relay, Relay, 9)
 MSComm1.Output = Relay
 Call Sleep(50)
 End If
 
End Sub

Private Sub relay2_Click(Index As Integer)
 If (Index = 0) Then
 'SOF  LEN  LINK_TEST  RELAY/USB/SENSOR  DAT3  DAT2  DAT1  DAT0  EOF  CRC1  CRC0
       
 Relay(0) = &H55
 Relay(1) = &HB
 Relay(2) = &H1
 Relay(3) = &H0                                      'Type -- RELAY
 Relay(4) = &H0
 Relay(5) = &H0
 Relay(6) = &H0
 
 Relay(7) = Pre_Relay(7) Or &H2
 Relay(8) = &HAA
 
 Call Cal_CRC16(Relay, CRC16, 9)
 
 Relay(9) = CRC16(1)
 Relay(10) = CRC16(0)
 Call Copy_Dat(Pre_Relay, Relay, 9)
 MSComm1.Output = Relay
 Call Sleep(50)
 End If
 
 
  If (Index = 1) Then
 'SOF  LEN  LINK_TEST  RELAY/USB/SENSOR  DAT3  DAT2  DAT1  DAT0  EOF  CRC1  CRC0
       
 Relay(0) = &H55
 Relay(1) = &HB
 Relay(2) = &H1
 Relay(3) = &H0                                      'Type -- RELAY
 Relay(4) = &H0
 Relay(5) = &H0
 Relay(6) = &H0
 
 Relay(7) = Pre_Relay(7) And &HFD
 Relay(8) = &HAA
 
 Call Cal_CRC16(Relay, CRC16, 9)
 
 Relay(9) = CRC16(1)
 Relay(10) = CRC16(0)
 Call Copy_Dat(Pre_Relay, Relay, 9)
 MSComm1.Output = Relay
 Call Sleep(50)
 End If
End Sub

Private Sub relay3_Click(Index As Integer)
 If (Index = 0) Then
 'SOF  LEN  LINK_TEST  RELAY/USB/SENSOR  DAT3  DAT2  DAT1  DAT0  EOF  CRC1  CRC0
       
 Relay(0) = &H55
 Relay(1) = &HB
 Relay(2) = &H1
 Relay(3) = &H0                                      'Type -- RELAY
 Relay(4) = &H0
 Relay(5) = &H0
 Relay(6) = &H0
 
 Relay(7) = Pre_Relay(7) Or &H4
 Relay(8) = &HAA
 
 Call Cal_CRC16(Relay, CRC16, 9)
 
 Relay(9) = CRC16(1)
 Relay(10) = CRC16(0)
 Call Copy_Dat(Pre_Relay, Relay, 9)
 MSComm1.Output = Relay
 Call Sleep(50)
 End If
 
 
  If (Index = 1) Then
 'SOF  LEN  LINK_TEST  RELAY/USB/SENSOR  DAT3  DAT2  DAT1  DAT0  EOF  CRC1  CRC0
       
 Relay(0) = &H55
 Relay(1) = &HB
 Relay(2) = &H1
 Relay(3) = &H0                                      'Type -- RELAY
 Relay(4) = &H0
 Relay(5) = &H0
 Relay(6) = &H0
 
 Relay(7) = Pre_Relay(7) And &HFB
 Relay(8) = &HAA
 
 Call Cal_CRC16(Relay, CRC16, 9)
 
 Relay(9) = CRC16(1)
 Relay(10) = CRC16(0)
 Call Copy_Dat(Pre_Relay, Relay, 9)
 MSComm1.Output = Relay
 Call Sleep(50)
 End If
End Sub

Private Sub relay4_Click(Index As Integer)
 If (Index = 0) Then
 'SOF  LEN  LINK_TEST  RELAY/USB/SENSOR  DAT3  DAT2  DAT1  DAT0  EOF  CRC1  CRC0
       
 Relay(0) = &H55
 Relay(1) = &HB
 Relay(2) = &H1
 Relay(3) = &H0                                      'Type -- RELAY
 Relay(4) = &H0
 Relay(5) = &H0
 Relay(6) = &H0
 
 Relay(7) = Pre_Relay(7) Or &H8
 Relay(8) = &HAA
 
 Call Cal_CRC16(Relay, CRC16, 9)
 
 Relay(9) = CRC16(1)
 Relay(10) = CRC16(0)
 Call Copy_Dat(Pre_Relay, Relay, 9)
 MSComm1.Output = Relay
 Call Sleep(50)
 End If
 
 
  If (Index = 1) Then
 'SOF  LEN  LINK_TEST  RELAY/USB/SENSOR  DAT3  DAT2  DAT1  DAT0  EOF  CRC1  CRC0
       
 Relay(0) = &H55
 Relay(1) = &HB
 Relay(2) = &H1
 Relay(3) = &H0                                      'Type -- RELAY
 Relay(4) = &H0
 Relay(5) = &H0
 Relay(6) = &H0
 
 Relay(7) = Pre_Relay(7) And &HF7
 Relay(8) = &HAA
 
 Call Cal_CRC16(Relay, CRC16, 9)
 
 Relay(9) = CRC16(1)
 Relay(10) = CRC16(0)
 Call Copy_Dat(Pre_Relay, Relay, 9)
 MSComm1.Output = Relay
 Call Sleep(50)
 End If
End Sub

Private Sub relay5_Click(Index As Integer)
 If (Index = 0) Then
 'SOF  LEN  LINK_TEST  RELAY/USB/SENSOR  DAT3  DAT2  DAT1  DAT0  EOF  CRC1  CRC0
       
 Relay(0) = &H55
 Relay(1) = &HB
 Relay(2) = &H1
 Relay(3) = &H0                                      'Type -- RELAY
 Relay(4) = &H0
 Relay(5) = &H0
 Relay(6) = &H0
 
 Relay(7) = Pre_Relay(7) Or &H10
 Relay(8) = &HAA
 
 Call Cal_CRC16(Relay, CRC16, 9)
 
 Relay(9) = CRC16(1)
 Relay(10) = CRC16(0)
 Call Copy_Dat(Pre_Relay, Relay, 9)
 MSComm1.Output = Relay
 Call Sleep(50)
 End If
 
 
  If (Index = 1) Then
 'SOF  LEN  LINK_TEST  RELAY/USB/SENSOR  DAT3  DAT2  DAT1  DAT0  EOF  CRC1  CRC0
       
 Relay(0) = &H55
 Relay(1) = &HB
 Relay(2) = &H1
 Relay(3) = &H0                                      'Type -- RELAY
 Relay(4) = &H0
 Relay(5) = &H0
 Relay(6) = &H0
 
 Relay(7) = Pre_Relay(7) And &HEF
 Relay(8) = &HAA
 
 Call Cal_CRC16(Relay, CRC16, 9)
 
 Relay(9) = CRC16(1)
 Relay(10) = CRC16(0)
 Call Copy_Dat(Pre_Relay, Relay, 9)
 MSComm1.Output = Relay
 Call Sleep(50)
 End If
End Sub

Private Sub relay6_Click(Index As Integer)
 If (Index = 0) Then
 'SOF  LEN  LINK_TEST  RELAY/USB/SENSOR  DAT3  DAT2  DAT1  DAT0  EOF  CRC1  CRC0
       
 Relay(0) = &H55
 Relay(1) = &HB
 Relay(2) = &H1
 Relay(3) = &H0                                      'Type -- RELAY
 Relay(4) = &H0
 Relay(5) = &H0
 Relay(6) = &H0
 
 Relay(7) = Pre_Relay(7) Or &H20
 Relay(8) = &HAA
 
 Call Cal_CRC16(Relay, CRC16, 9)
 
 Relay(9) = CRC16(1)
 Relay(10) = CRC16(0)
 Call Copy_Dat(Pre_Relay, Relay, 9)
 MSComm1.Output = Relay
 Call Sleep(50)
 
 End If
 
 
  If (Index = 1) Then
 'SOF  LEN  LINK_TEST  RELAY/USB/SENSOR  DAT3  DAT2  DAT1  DAT0  EOF  CRC1  CRC0
       
 Relay(0) = &H55
 Relay(1) = &HB
 Relay(2) = &H1
 Relay(3) = &H0                                      'Type -- RELAY
 Relay(4) = &H0
 Relay(5) = &H0
 Relay(6) = &H0
 
 Relay(7) = Pre_Relay(7) And &HDF
 Relay(8) = &HAA
 
 Call Cal_CRC16(Relay, CRC16, 9)
 
 Relay(9) = CRC16(1)
 Relay(10) = CRC16(0)
 Call Copy_Dat(Pre_Relay, Relay, 9)
 MSComm1.Output = Relay
 Call Sleep(50)
 End If
End Sub

Private Sub relay7_Click(Index As Integer)
 If (Index = 0) Then
 'SOF  LEN  LINK_TEST  RELAY/USB/SENSOR  DAT3  DAT2  DAT1  DAT0  EOF  CRC1  CRC0
       
 Relay(0) = &H55
 Relay(1) = &HB
 Relay(2) = &H1
 Relay(3) = &H0                                      'Type -- RELAY
 Relay(4) = &H0
 Relay(5) = &H0
 Relay(6) = &H0
 
 Relay(7) = Pre_Relay(7) Or &H40
 Relay(8) = &HAA
 
 Call Cal_CRC16(Relay, CRC16, 9)
 
 Relay(9) = CRC16(1)
 Relay(10) = CRC16(0)
 Call Copy_Dat(Pre_Relay, Relay, 9)
 MSComm1.Output = Relay
 Call Sleep(50)
 End If
 
 
  If (Index = 1) Then
 'SOF  LEN  LINK_TEST  RELAY/USB/SENSOR  DAT3  DAT2  DAT1  DAT0  EOF  CRC1  CRC0
       
 Relay(0) = &H55
 Relay(1) = &HB
 Relay(2) = &H1
 Relay(3) = &H0                                      'Type -- RELAY
 Relay(4) = &H0
 Relay(5) = &H0
 Relay(6) = &H0
 
 Relay(7) = Pre_Relay(7) And &HBF
 Relay(8) = &HAA
 
 Call Cal_CRC16(Relay, CRC16, 9)
 
 Relay(9) = CRC16(1)
 Relay(10) = CRC16(0)
 Call Copy_Dat(Pre_Relay, Relay, 9)
 MSComm1.Output = Relay
 Call Sleep(50)
 End If
End Sub

Private Sub relay8_Click(Index As Integer)
 If (Index = 0) Then
 'SOF  LEN  LINK_TEST  RELAY/USB/SENSOR  DAT3  DAT2  DAT1  DAT0  EOF  CRC1  CRC0
       
 Relay(0) = &H55
 Relay(1) = &HB
 Relay(2) = &H1
 Relay(3) = &H0                                      'Type -- RELAY
 Relay(4) = &H0
 Relay(5) = &H0
 Relay(6) = &H0
 
 Relay(7) = Pre_Relay(7) Or &H80
 Relay(8) = &HAA
 
 Call Cal_CRC16(Relay, CRC16, 9)
 
 Relay(9) = CRC16(1)
 Relay(10) = CRC16(0)
 Call Copy_Dat(Pre_Relay, Relay, 9)
 MSComm1.Output = Relay
 Call Sleep(50)
 End If
 
 
  If (Index = 1) Then
 'SOF  LEN  LINK_TEST  RELAY/USB/SENSOR  DAT3  DAT2  DAT1  DAT0  EOF  CRC1  CRC0
       
 Relay(0) = &H55
 Relay(1) = &HB
 Relay(2) = &H1
 Relay(3) = &H0                                      'Type -- RELAY
 Relay(4) = &H0
 Relay(5) = &H0
 Relay(6) = &H0
 
 Relay(7) = Pre_Relay(7) And &H7F
 Relay(8) = &HAA
 
 Call Cal_CRC16(Relay, CRC16, 9)
 
 Relay(9) = CRC16(1)
 Relay(10) = CRC16(0)
 Call Copy_Dat(Pre_Relay, Relay, 9)
 MSComm1.Output = Relay
 Call Sleep(50)
 End If
End Sub

Private Sub SysInfo1_DeviceArrival(ByVal DeviceType As Long, ByVal DeviceID As Long, ByVal DeviceName As String, ByVal DeviceData As Long)
    If DeviceType = 3 Then
        Combo1.AddItem DeviceName
        Combo6.AddItem DeviceName
    End If
End Sub
Private Sub SysInfo1_DeviceRemoveComplete(ByVal DeviceType As Long, ByVal DeviceID As Long, ByVal DeviceName As String, ByVal DeviceData As Long)
    Dim i As Integer
    If DeviceType = 3 Then
        For i = 0 To 31
        If Combo1.List(i) = DeviceName Then
            Combo1.RemoveItem (i)
        End If
        If Combo6.List(i) = DeviceName Then
            Combo6.RemoveItem (i)
        End If
        Next i
    End If
End Sub

Private Sub usart1_Click(Index As Integer)
 If (Index = 0) Then
 'SOF  LEN  LINK_TEST  RELAY/USB/SENSOR  DAT3  DAT2  DAT1  DAT0  EOF  CRC1  CRC0
       
 USB_Board(0) = &H55
 USB_Board(1) = &HB
 USB_Board(2) = &H1
 USB_Board(3) = &H1                                      'Type -- RELAY
 USB_Board(4) = &H0
 USB_Board(5) = &H0
 
 USB_Board(6) = Pre_USB_Board(6)
 USB_Board(7) = (Pre_USB_Board(7) And &HF8) Or &H3
 USB_Board(8) = &HAA
 
 Call Cal_CRC16(USB_Board, CRC16, 9)
 
 USB_Board(9) = CRC16(1)
 USB_Board(10) = CRC16(0)
 Call Copy_Dat(Pre_USB_Board, USB_Board, 9)
 MSComm1.Output = USB_Board
 
 End If
 
 
  If (Index = 1) Then
 'SOF  LEN  LINK_TEST  RELAY/USB/SENSOR  DAT3  DAT2  DAT1  DAT0  EOF  CRC1  CRC0
       
 USB_Board(0) = &H55
 USB_Board(1) = &HB
 USB_Board(2) = &H1
 USB_Board(3) = &H1                                      'Type -- RELAY
 USB_Board(4) = &H0
 USB_Board(5) = &H0
 
 USB_Board(6) = Pre_USB_Board(6)
 USB_Board(7) = (Pre_USB_Board(7) And &HF8) Or &H1
 USB_Board(8) = &HAA
 
 Call Cal_CRC16(USB_Board, CRC16, 9)
 
 USB_Board(9) = CRC16(1)
 USB_Board(10) = CRC16(0)
 Call Copy_Dat(Pre_USB_Board, USB_Board, 9)
 MSComm1.Output = USB_Board
 
 End If
 
 
 If (Index = 2) Then
 'SOF  LEN  LINK_TEST  RELAY/USB/SENSOR  DAT3  DAT2  DAT1  DAT0  EOF  CRC1  CRC0
       
       
 USB_Board(0) = &H55
 USB_Board(1) = &HB
 USB_Board(2) = &H1
 USB_Board(3) = &H1                                      'Type -- RELAY
 USB_Board(4) = &H0
 USB_Board(5) = &H0
 
 USB_Board(6) = Pre_USB_Board(6)
 USB_Board(7) = (Pre_USB_Board(7) And &HF8) Or &H5
 USB_Board(8) = &HAA
 
 Call Cal_CRC16(USB_Board, CRC16, 9)
 
 USB_Board(9) = CRC16(1)
 USB_Board(10) = CRC16(0)
 Call Copy_Dat(Pre_USB_Board, USB_Board, 9)
 MSComm1.Output = USB_Board
 
 End If
 
End Sub

Private Sub usart2_Click(Index As Integer)
If (Index = 0) Then
 'SOF  LEN  LINK_TEST  RELAY/USB/SENSOR  DAT3  DAT2  DAT1  DAT0  EOF  CRC1  CRC0
       
 USB_Board(0) = &H55
 USB_Board(1) = &HB
 USB_Board(2) = &H1
 USB_Board(3) = &H1                                      'Type -- RELAY
 USB_Board(4) = &H0
 USB_Board(5) = &H0
 
 USB_Board(6) = Pre_USB_Board(6)
 USB_Board(7) = (Pre_USB_Board(7) And &HC7) Or &H18
 USB_Board(8) = &HAA
 
 Call Cal_CRC16(USB_Board, CRC16, 9)
 
 USB_Board(9) = CRC16(1)
 USB_Board(10) = CRC16(0)
 Call Copy_Dat(Pre_USB_Board, USB_Board, 9)
 MSComm1.Output = USB_Board
 
 End If
 
 
  If (Index = 1) Then
 'SOF  LEN  LINK_TEST  RELAY/USB/SENSOR  DAT3  DAT2  DAT1  DAT0  EOF  CRC1  CRC0
       
 USB_Board(0) = &H55
 USB_Board(1) = &HB
 USB_Board(2) = &H1
 USB_Board(3) = &H1                                      'Type -- RELAY
 USB_Board(4) = &H0
 USB_Board(5) = &H0
 
 USB_Board(6) = Pre_USB_Board(6)
 USB_Board(7) = (Pre_USB_Board(7) And &HC7) Or &H8
 USB_Board(8) = &HAA
 
 Call Cal_CRC16(USB_Board, CRC16, 9)
 
 USB_Board(9) = CRC16(1)
 USB_Board(10) = CRC16(0)
 Call Copy_Dat(Pre_USB_Board, USB_Board, 9)
 MSComm1.Output = USB_Board
 
 End If
 
 
 If (Index = 2) Then
 'SOF  LEN  LINK_TEST  RELAY/USB/SENSOR  DAT3  DAT2  DAT1  DAT0  EOF  CRC1  CRC0
       
       
 USB_Board(0) = &H55
 USB_Board(1) = &HB
 USB_Board(2) = &H1
 USB_Board(3) = &H1                                      'Type -- RELAY
 USB_Board(4) = &H0
 USB_Board(5) = &H0
 
 USB_Board(6) = Pre_USB_Board(6)
 USB_Board(7) = (Pre_USB_Board(7) And &HC7) Or &H28
 USB_Board(8) = &HAA
 
 Call Cal_CRC16(USB_Board, CRC16, 9)
 
 USB_Board(9) = CRC16(1)
 USB_Board(10) = CRC16(0)
 Call Copy_Dat(Pre_USB_Board, USB_Board, 9)
 MSComm1.Output = USB_Board
 
 End If
 
End Sub

Private Sub usb1_Click(Index As Integer)
If (Index = 0) Then
 'SOF  LEN  LINK_TEST  RELAY/USB/SENSOR  DAT3  DAT2  DAT1  DAT0  EOF  CRC1  CRC0
       
 USB_Board(0) = &H55
 USB_Board(1) = &HB
 USB_Board(2) = &H1
 USB_Board(3) = &H1                                      'Type -- RELAY
 USB_Board(4) = &H0
 USB_Board(5) = &H0
 
 USB_Board(6) = (Pre_USB_Board(6) And &HFE) Or &H0
 USB_Board(7) = (Pre_USB_Board(7) And &H3F) Or &HC0
 USB_Board(8) = &HAA
 
 Call Cal_CRC16(USB_Board, CRC16, 9)
 
 USB_Board(9) = CRC16(1)
 USB_Board(10) = CRC16(0)
 Call Copy_Dat(Pre_USB_Board, USB_Board, 9)
 MSComm1.Output = USB_Board
 
 End If
 
 
  If (Index = 1) Then
 'SOF  LEN  LINK_TEST  RELAY/USB/SENSOR  DAT3  DAT2  DAT1  DAT0  EOF  CRC1  CRC0
       
 USB_Board(0) = &H55
 USB_Board(1) = &HB
 USB_Board(2) = &H1
 USB_Board(3) = &H1                                      'Type -- RELAY
 USB_Board(4) = &H0
 USB_Board(5) = &H0
 
 USB_Board(6) = (Pre_USB_Board(6) And &HFE) Or &H0
 USB_Board(7) = (Pre_USB_Board(7) And &H3F) Or &H40
 USB_Board(8) = &HAA
 
 Call Cal_CRC16(USB_Board, CRC16, 9)
 
 USB_Board(9) = CRC16(1)
 USB_Board(10) = CRC16(0)
 Call Copy_Dat(Pre_USB_Board, USB_Board, 9)
 MSComm1.Output = USB_Board
 
 End If
 
 
 If (Index = 2) Then
 'SOF  LEN  LINK_TEST  RELAY/USB/SENSOR  DAT3  DAT2  DAT1  DAT0  EOF  CRC1  CRC0
       
       
 USB_Board(0) = &H55
 USB_Board(1) = &HB
 USB_Board(2) = &H1
 USB_Board(3) = &H1                                      'Type -- RELAY
 USB_Board(4) = &H0
 USB_Board(5) = &H0
 
 USB_Board(6) = (Pre_USB_Board(6) And &HFE) Or &H1
 USB_Board(7) = (Pre_USB_Board(7) And &H3F) Or &H40
 USB_Board(8) = &HAA
 
 Call Cal_CRC16(USB_Board, CRC16, 9)
 
 USB_Board(9) = CRC16(1)
 USB_Board(10) = CRC16(0)
 Call Copy_Dat(Pre_USB_Board, USB_Board, 9)
 MSComm1.Output = USB_Board
 
 End If
 
End Sub

Private Sub usb2_Click(Index As Integer)
If (Index = 0) Then
 'SOF  LEN  LINK_TEST  RELAY/USB/SENSOR  DAT3  DAT2  DAT1  DAT0  EOF  CRC1  CRC0
       
 USB_Board(0) = &H55
 USB_Board(1) = &HB
 USB_Board(2) = &H1
 USB_Board(3) = &H1                                      'Type -- RELAY
 USB_Board(4) = &H0
 USB_Board(5) = &H0
 
 USB_Board(6) = (Pre_USB_Board(6) And &HF1) Or &H6
 USB_Board(7) = Pre_USB_Board(7)
 USB_Board(8) = &HAA
 
 Call Cal_CRC16(USB_Board, CRC16, 9)
 
 USB_Board(9) = CRC16(1)
 USB_Board(10) = CRC16(0)
 Call Copy_Dat(Pre_USB_Board, USB_Board, 9)
 MSComm1.Output = USB_Board
 
 End If
 
 
  If (Index = 1) Then
 'SOF  LEN  LINK_TEST  RELAY/USB/SENSOR  DAT3  DAT2  DAT1  DAT0  EOF  CRC1  CRC0
       
 USB_Board(0) = &H55
 USB_Board(1) = &HB
 USB_Board(2) = &H1
 USB_Board(3) = &H1                                      'Type -- RELAY
 USB_Board(4) = &H0
 USB_Board(5) = &H0
 
 USB_Board(6) = (Pre_USB_Board(6) And &HF1) Or &H2
 USB_Board(7) = Pre_USB_Board(7)
 USB_Board(8) = &HAA
 
 Call Cal_CRC16(USB_Board, CRC16, 9)
 
 USB_Board(9) = CRC16(1)
 USB_Board(10) = CRC16(0)
 Call Copy_Dat(Pre_USB_Board, USB_Board, 9)
 MSComm1.Output = USB_Board
 
 End If
 
 
 If (Index = 2) Then
 'SOF  LEN  LINK_TEST  RELAY/USB/SENSOR  DAT3  DAT2  DAT1  DAT0  EOF  CRC1  CRC0
       
       
 USB_Board(0) = &H55
 USB_Board(1) = &HB
 USB_Board(2) = &H1
 USB_Board(3) = &H1                                      'Type -- RELAY
 USB_Board(4) = &H0
 USB_Board(5) = &H0
 
 USB_Board(6) = (Pre_USB_Board(6) And &HF1) Or &HA
 USB_Board(7) = Pre_USB_Board(7)
 USB_Board(8) = &HAA
 
 Call Cal_CRC16(USB_Board, CRC16, 9)
 
 USB_Board(9) = CRC16(1)
 USB_Board(10) = CRC16(0)
 Call Copy_Dat(Pre_USB_Board, USB_Board, 9)
 MSComm1.Output = USB_Board
 
 End If
End Sub
