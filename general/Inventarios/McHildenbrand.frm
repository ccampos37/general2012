VERSION 5.00
Begin VB.Form frmMcHildenbrand 
   BackColor       =   &H00C0C0C0&
   Caption         =   "McHildenbrand-Por Favor, si copian este código fuente, hagan mención al Autor"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9675
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "McHildenbrand.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   9675
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Beverages"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   5040
      TabIndex        =   47
      Top             =   3720
      Width           =   4575
      Begin VB.PictureBox picBack4 
         BackColor       =   &H00C0C0C0&
         Height          =   3735
         Left            =   120
         ScaleHeight     =   3675
         ScaleWidth      =   4275
         TabIndex        =   48
         Top             =   240
         Width           =   4335
         Begin VB.CommandButton cmdbeverage 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Small Coke"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdbeverage 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Large Coke"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdbeverage 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Small Diet"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   840
            Width           =   1335
         End
         Begin VB.CommandButton cmdbeverage 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Large Diet"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   840
            Width           =   1335
         End
         Begin VB.CommandButton cmdbeverage 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Small 7up"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   1560
            Width           =   1335
         End
         Begin VB.CommandButton cmdbeverage 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Large 7up"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   1560
            Width           =   1335
         End
         Begin VB.CommandButton cmdbeverage 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Small Coffee"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   6
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   2280
            Width           =   1335
         End
         Begin VB.CommandButton cmdbeverage 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Large Coffee"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   7
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   2280
            Width           =   1335
         End
         Begin VB.CommandButton cmdbeverage 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Small Tea"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   8
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   3000
            Width           =   1335
         End
         Begin VB.CommandButton cmdbeverage 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Large Tea"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   9
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   3000
            Width           =   1335
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   2
            X1              =   120
            X2              =   4200
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   3
            X1              =   120
            X2              =   4200
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   4
            X1              =   120
            X2              =   4200
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   7
            X1              =   120
            X2              =   4200
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   8
            X1              =   120
            X2              =   4200
            Y1              =   3600
            Y2              =   3600
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFFFF&
            Index           =   4
            X1              =   2160
            X2              =   2160
            Y1              =   120
            Y2              =   600
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00E0E0E0&
            Index           =   5
            X1              =   2160
            X2              =   2160
            Y1              =   840
            Y2              =   1320
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00E0E0E0&
            Index           =   6
            X1              =   2160
            X2              =   2160
            Y1              =   2280
            Y2              =   2760
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFFFF&
            Index           =   7
            X1              =   2160
            X2              =   2160
            Y1              =   1560
            Y2              =   2040
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFFFF&
            Index           =   8
            X1              =   2160
            X2              =   2160
            Y1              =   3000
            Y2              =   3480
         End
         Begin VB.Label lblBeverage 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   0
            Left            =   1560
            TabIndex        =   68
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblBeverage 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   1
            Left            =   3840
            TabIndex        =   67
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblBeverage 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   2
            Left            =   1560
            TabIndex        =   66
            Top             =   960
            Width           =   375
         End
         Begin VB.Label lblBeverage 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   3
            Left            =   3840
            TabIndex        =   65
            Top             =   960
            Width           =   375
         End
         Begin VB.Label lblBeverage 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   4
            Left            =   1560
            TabIndex        =   64
            Top             =   1680
            Width           =   375
         End
         Begin VB.Label lblBeverage 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   5
            Left            =   3840
            TabIndex        =   63
            Top             =   1680
            Width           =   375
         End
         Begin VB.Label lblBeverage 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   6
            Left            =   1560
            TabIndex        =   62
            Top             =   2400
            Width           =   375
         End
         Begin VB.Label lblBeverage 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   7
            Left            =   3840
            TabIndex        =   61
            Top             =   2400
            Width           =   375
         End
         Begin VB.Label lblBeverage 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   8
            Left            =   1560
            TabIndex        =   60
            Top             =   3120
            Width           =   375
         End
         Begin VB.Label lblBeverage 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   9
            Left            =   3840
            TabIndex        =   59
            Top             =   3120
            Width           =   375
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Height          =   6615
      Left            =   120
      TabIndex        =   27
      Top             =   1200
      Width           =   4815
      Begin VB.PictureBox picback3 
         BackColor       =   &H00000000&
         Height          =   6255
         Left            =   120
         ScaleHeight     =   6195
         ScaleWidth      =   4515
         TabIndex        =   28
         Top             =   240
         Width           =   4575
         Begin VB.Label lblBevSums 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   78
            Top             =   5040
            Width           =   255
         End
         Begin VB.Label lblBevSums 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   77
            Top             =   4800
            Width           =   255
         End
         Begin VB.Label lblBevSums 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   76
            Top             =   4560
            Width           =   255
         End
         Begin VB.Label lblBevSums 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   75
            Top             =   4320
            Width           =   255
         End
         Begin VB.Label lblBevSums 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   74
            Top             =   4080
            Width           =   255
         End
         Begin VB.Label lblBevSums 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   73
            Top             =   3840
            Width           =   255
         End
         Begin VB.Label lblBevSums 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   72
            Top             =   3600
            Width           =   255
         End
         Begin VB.Label lblBevSums 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   71
            Top             =   3360
            Width           =   255
         End
         Begin VB.Label lblBevSums 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   70
            Top             =   3120
            Width           =   255
         End
         Begin VB.Label lblBevitem 
            BackColor       =   &H00000000&
            Caption         =   "Large Tea"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   9
            Left            =   840
            TabIndex        =   89
            Top             =   5280
            Width           =   1815
         End
         Begin VB.Label lblBevitem 
            BackColor       =   &H00000000&
            Caption         =   "Small Tea"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   8
            Left            =   840
            TabIndex        =   88
            Top             =   5040
            Width           =   1815
         End
         Begin VB.Label lblBevitem 
            BackColor       =   &H00000000&
            Caption         =   "Large Coffee"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   7
            Left            =   840
            TabIndex        =   87
            Top             =   4800
            Width           =   1815
         End
         Begin VB.Label lblBevitem 
            BackColor       =   &H00000000&
            Caption         =   "Small Coffee"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   6
            Left            =   840
            TabIndex        =   86
            Top             =   4560
            Width           =   1815
         End
         Begin VB.Label lblBevitem 
            BackColor       =   &H00000000&
            Caption         =   "Large 7up"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   5
            Left            =   840
            TabIndex        =   85
            Top             =   4320
            Width           =   1815
         End
         Begin VB.Label lblBevitem 
            BackColor       =   &H00000000&
            Caption         =   "Small 7up"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   4
            Left            =   840
            TabIndex        =   84
            Top             =   4080
            Width           =   1815
         End
         Begin VB.Label lblBevitem 
            BackColor       =   &H00000000&
            Caption         =   "Large Diet"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   3
            Left            =   840
            TabIndex        =   83
            Top             =   3840
            Width           =   1815
         End
         Begin VB.Label lblBevitem 
            BackColor       =   &H00000000&
            Caption         =   "Small Diet"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   840
            TabIndex        =   82
            Top             =   3600
            Width           =   1815
         End
         Begin VB.Label lblBevitem 
            BackColor       =   &H00000000&
            Caption         =   "Large Coke"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   81
            Top             =   3360
            Width           =   1815
         End
         Begin VB.Label lblBevitem 
            BackColor       =   &H00000000&
            Caption         =   "Small Coke"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   80
            Top             =   3120
            Width           =   1815
         End
         Begin VB.Label lblBevSums 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   79
            Top             =   5280
            Width           =   255
         End
         Begin VB.Label lblDummy3 
            BackColor       =   &H00000000&
            Caption         =   "Beverages"
            ForeColor       =   &H0080C0FF&
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   2880
            Width           =   1335
         End
         Begin VB.Label lblAmount 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   276
            Index           =   0
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   492
         End
         Begin VB.Label lblSandType 
            BackColor       =   &H00000000&
            Caption         =   "Cheeseburger"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   252
            Index           =   0
            Left            =   840
            TabIndex        =   45
            Top             =   360
            Width           =   1932
         End
         Begin VB.Label lblAmount 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   276
            Index           =   1
            Left            =   120
            TabIndex        =   44
            Top             =   600
            Width           =   492
         End
         Begin VB.Label lblAmount 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   276
            Index           =   2
            Left            =   120
            TabIndex        =   43
            Top             =   840
            Width           =   492
         End
         Begin VB.Label lblAmount 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   276
            Index           =   3
            Left            =   120
            TabIndex        =   42
            Top             =   1080
            Width           =   492
         End
         Begin VB.Label lblSandType 
            BackColor       =   &H00000000&
            Caption         =   "Hamburger"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   252
            Index           =   1
            Left            =   840
            TabIndex        =   41
            Top             =   600
            Width           =   1932
         End
         Begin VB.Label lblSandType 
            BackColor       =   &H00000000&
            Caption         =   "Double Cheeseburger"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   252
            Index           =   2
            Left            =   840
            TabIndex        =   40
            Top             =   840
            Width           =   3252
         End
         Begin VB.Label lblSandType 
            BackColor       =   &H00000000&
            Caption         =   "Chicken Sandwich"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   372
            Index           =   3
            Left            =   840
            TabIndex        =   39
            Top             =   1080
            Width           =   2652
         End
         Begin VB.Label Lbldummy1 
            BackColor       =   &H00000000&
            Caption         =   "Sandwiches"
            ForeColor       =   &H0080C0FF&
            Height          =   252
            Left            =   120
            TabIndex        =   38
            Top             =   0
            Width           =   2892
         End
         Begin VB.Label LblDummy2 
            BackColor       =   &H00000000&
            Caption         =   "Side Dishes"
            ForeColor       =   &H0080C0FF&
            Height          =   252
            Left            =   120
            TabIndex        =   37
            Top             =   1440
            Width           =   1332
         End
         Begin VB.Label lblSidesItem 
            BackColor       =   &H00000000&
            Caption         =   "Large Fry"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   252
            Index           =   0
            Left            =   840
            TabIndex        =   36
            Top             =   1680
            Width           =   2892
         End
         Begin VB.Label lblSidesItem 
            BackColor       =   &H00000000&
            Caption         =   "Small Fry"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   252
            Index           =   1
            Left            =   840
            TabIndex        =   35
            Top             =   1920
            Width           =   2892
         End
         Begin VB.Label lblSidesItem 
            BackColor       =   &H00000000&
            Caption         =   "Sundae"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   252
            Index           =   2
            Left            =   840
            TabIndex        =   34
            Top             =   2160
            Width           =   2892
         End
         Begin VB.Label lblSidesItem 
            BackColor       =   &H00000000&
            Caption         =   "Apple Pie"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   252
            Index           =   3
            Left            =   840
            TabIndex        =   33
            Top             =   2400
            Width           =   2892
         End
         Begin VB.Label lblSidesAmount 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   252
            Index           =   0
            Left            =   240
            TabIndex        =   32
            Top             =   1680
            Width           =   252
         End
         Begin VB.Label lblSidesAmount 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   252
            Index           =   1
            Left            =   240
            TabIndex        =   31
            Top             =   1920
            Width           =   252
         End
         Begin VB.Label lblSidesAmount 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   252
            Index           =   2
            Left            =   240
            TabIndex        =   30
            Top             =   2160
            Width           =   252
         End
         Begin VB.Label lblSidesAmount 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   252
            Index           =   3
            Left            =   240
            TabIndex        =   29
            Top             =   2400
            Width           =   252
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   1212
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   4815
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00C0C000&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Print Receipt"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdEatIn 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Eat In"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   720
         Width           =   972
      End
      Begin VB.CommandButton cmdTakeout 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Take Out"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdVoid 
         BackColor       =   &H008080FF&
         Caption         =   "Void"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblOutput 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Modern"
            Size            =   20.25
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   372
         Left            =   1440
         TabIndex        =   26
         Top             =   240
         Width           =   2052
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Side Dishes"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1692
      Left            =   5040
      TabIndex        =   6
      Top             =   1920
      Width           =   4572
      Begin VB.PictureBox picback2 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1332
         Left            =   120
         ScaleHeight     =   1275
         ScaleWidth      =   4275
         TabIndex        =   7
         Top             =   240
         Width           =   4332
         Begin VB.CommandButton cmdSides 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Apple Pie"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Index           =   3
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   720
            Width           =   1332
         End
         Begin VB.CommandButton cmdSides 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Sundae"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Index           =   2
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   720
            Width           =   1332
         End
         Begin VB.CommandButton cmdSides 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Small Fry"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Index           =   1
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   120
            Width           =   1332
         End
         Begin VB.CommandButton cmdSides 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Large Fry"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   120
            Width           =   1332
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00E0E0E0&
            Index           =   3
            X1              =   2160
            X2              =   2160
            Y1              =   720
            Y2              =   1080
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFFFF&
            Index           =   2
            X1              =   2160
            X2              =   2160
            Y1              =   120
            Y2              =   480
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   0
            X1              =   120
            X2              =   4200
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   1
            X1              =   120
            X2              =   4200
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Label lblSides 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2"
            ForeColor       =   &H008080FF&
            Height          =   372
            Index           =   0
            Left            =   1560
            TabIndex        =   15
            Top             =   120
            Width           =   372
         End
         Begin VB.Label lblSides 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2"
            ForeColor       =   &H008080FF&
            Height          =   372
            Index           =   1
            Left            =   3840
            TabIndex        =   14
            Top             =   120
            Width           =   372
         End
         Begin VB.Label lblSides 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2"
            ForeColor       =   &H008080FF&
            Height          =   372
            Index           =   2
            Left            =   1560
            TabIndex        =   13
            Top             =   720
            Width           =   372
         End
         Begin VB.Label lblSides 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2"
            ForeColor       =   &H008080FF&
            Height          =   372
            Index           =   3
            Left            =   3840
            TabIndex        =   12
            Top             =   720
            Width           =   372
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sandwiches"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1932
      Left            =   5040
      TabIndex        =   0
      Top             =   0
      Width           =   4572
      Begin VB.PictureBox picback1 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1572
         Left            =   120
         ScaleHeight     =   1515
         ScaleWidth      =   4275
         TabIndex        =   1
         Top             =   240
         Width           =   4332
         Begin VB.CommandButton cmdSndWich 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Chicken Sandwich"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   492
            Index           =   3
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   840
            Width           =   1332
         End
         Begin VB.CommandButton cmdSndWich 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Double Cheeseburger"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   492
            Index           =   2
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   840
            Width           =   1332
         End
         Begin VB.CommandButton cmdSndWich 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Hamburger"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   492
            Index           =   1
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   120
            Width           =   1332
         End
         Begin VB.CommandButton cmdSndWich 
            BackColor       =   &H00C0C0C0&
            Caption         =   "CheeseBurger"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   492
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   120
            Width           =   1332
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFFFF&
            Index           =   1
            X1              =   2160
            X2              =   2160
            Y1              =   840
            Y2              =   1320
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00E0E0E0&
            Index           =   0
            X1              =   2160
            X2              =   2160
            Y1              =   120
            Y2              =   600
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   6
            X1              =   120
            X2              =   4200
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   5
            X1              =   120
            X2              =   4200
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label lblBurger 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1"
            ForeColor       =   &H00808000&
            Height          =   372
            Index           =   3
            Left            =   3840
            TabIndex        =   5
            Top             =   960
            Width           =   372
         End
         Begin VB.Label lblBurger 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1"
            ForeColor       =   &H00808000&
            Height          =   372
            Index           =   2
            Left            =   1560
            TabIndex        =   4
            Top             =   960
            Width           =   372
         End
         Begin VB.Label lblBurger 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1"
            ForeColor       =   &H00808000&
            Height          =   372
            Index           =   1
            Left            =   3840
            TabIndex        =   3
            Top             =   240
            Width           =   372
         End
         Begin VB.Label lblBurger 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1"
            ForeColor       =   &H00808000&
            Height          =   372
            Index           =   0
            Left            =   1560
            TabIndex        =   2
            Top             =   240
            Width           =   372
         End
      End
   End
End
Attribute VB_Name = "frmMcHildenbrand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''This program is copyright 1998 to Ray Hildenbrand'''''''''''''''''''''
''''''''''''''This code is provided as freeware to anyone who wants to use it. i had to develop it for a visual basic lab that i have in school
''''''''''''''primary lessons to be learned from this? who knows, who cares. pretty good example of using control arrays though



''''Misc. Const
Const Tax = 0.07
Const StoreName = "McHildenbrand"

''''Sandwich Constants
Const Hamburger = 0.88
Const CheeseBurger = 0.99
Const DoubleCheeseburger = 1.25
Const ChixSandwich = 1.99

''''Beverage COnstants
Const SmCoke = 0.87
Const LgCoke = 1.25
Const SmDiet = 0.87
Const LgDiet = 1.25
Const SmCoffee = 0.85
Const LgCoffee = 1.05
Const Sm7up = 0.87
Const Lg7up = 1.25
Const SmTea = 0.87
Const LgTea = 1.25

''''Sides Constants
Const LgFry = 1.27
Const SmFry = 0.86
Const ApplePie = 0.96
Const Sundae = 0.67

''''misc. variables
Dim TotalCost As Currency
Dim Adding As Boolean


'Ouput Message Dims
Dim OutputMessage As String
Public Sub CheckOutputMenu()
Dim i%

''''Check Label amounts (if they are set to zero then turn the "light" off"
For i = 0 To 3
If CStr(lblAmount(i).Caption) = 0 Then
    lblAmount(i).Enabled = False
    lblSandType(i).Enabled = False
    End If
       If CStr(lblSidesAmount(i).Caption) = 0 Then
          lblSidesAmount(i).Enabled = False
            lblSidesItem(i).Enabled = False
        End If
Next i

For i = 0 To 9
    If CStr(lblBevSums(i).Caption) = 0 Then
    lblBevitem(i).Enabled = False
    lblBevSums(i).Enabled = False
    End If
Next i

End Sub




Public Sub TellemAboutit(MenuItem As String)
'''' Tell the user they have not ordered any of those items yet
MsgBox "You have not placed an order for a " & MenuItem & ". Cannot complete this void.", vbExclamation, StoreName
cmdVoid.Enabled = True
End Sub





Private Sub ClearOrder()

'''' Set all of our variables and counters, labels to zero when old order is totaled
Dim counter As Integer
  For counter = 0 To 3
      lblBurger(counter).Caption = 0
      lblSides(counter).Caption = 0
      lblAmount(counter).Caption = 0
      lblSidesAmount(counter).Caption = 0
  Next counter
  
  For counter = 0 To 9
      lblBevSums(counter).Caption = 0
      lblBeverage(counter).Caption = 0
  Next counter
  
OutputMessage = ""
finalmessage = ""
lblOutput.Caption = 0
TotalCost = 0

End Sub

Private Sub cmdbeverage_Click(Index As Integer)
If cmdVoid.Enabled = False Then
      Adding = False
Else: Adding = True
End If

Dim BevType As String

If Adding Then
  Select Case Index
    Dim newvalue As Integer
    Case 0
        newvalue = CStr(lblBeverage(Index).Caption + 1)
        TotalCost = TotalCost + SmCoke
        lblBeverage(Index).Caption = newvalue
        totSmCoke = newvalue
        lblBevSums(Index).Caption = newvalue
        lblBevSums(Index).Enabled = True
        lblBevitem(Index).Enabled = True
    Case 1
        newvalue = CStr(lblBeverage(Index).Caption + 1)
        TotalCost = TotalCost + LgCoke
        lblBeverage(Index).Caption = newvalue
        totLgCoke = newvalue
        lblBevSums(Index).Caption = newvalue
        lblBevSums(Index).Enabled = True
        lblBevitem(Index).Enabled = True
    Case 2
        newvalue = CStr(lblBeverage(Index).Caption + 1)
        TotalCost = TotalCost + SmDiet
        lblBeverage(Index).Caption = newvalue
        totSmDiet = newvalue
        lblBevSums(Index).Caption = newvalue
        lblBevSums(Index).Enabled = True
        lblBevitem(Index).Enabled = True
    Case 3
        newvalue = CStr(lblBeverage(Index).Caption + 1)
        TotalCost = TotalCost + LgDiet
        lblBeverage(Index).Caption = newvalue
        totLgDiet = newvalue
        lblBevSums(Index).Caption = newvalue
        lblBevSums(Index).Enabled = True
        lblBevitem(Index).Enabled = True
         
     Case 4
        newvalue = CStr(lblBeverage(Index).Caption + 1)
        TotalCost = TotalCost + Sm7up
        lblBeverage(Index).Caption = newvalue
        totSm7up = newvalue
        lblBevSums(Index).Caption = newvalue
        lblBevSums(Index).Enabled = True
        lblBevitem(Index).Enabled = True
    Case 5
        newvalue = CStr(lblBeverage(Index).Caption + 1)
        TotalCost = TotalCost + Lg7up
        lblBeverage(Index).Caption = newvalue
        totLg7up = newvalue
        lblBevSums(Index).Caption = newvalue
        lblBevSums(Index).Enabled = True
        lblBevitem(Index).Enabled = True
         'lblAmount(Index).Caption = newvalue
    Case 6
        newvalue = CStr(lblBeverage(Index).Caption + 1)
        TotalCost = TotalCost + SmCoffee
        lblBeverage(Index).Caption = newvalue
        totSmCoffee = newvalue
        lblBevSums(Index).Caption = newvalue
        lblBevSums(Index).Enabled = True
        lblBevitem(Index).Enabled = True
    Case 7
        newvalue = CStr(lblBeverage(Index).Caption + 1)
        TotalCost = TotalCost + LgCoffee
        lblBeverage(Index).Caption = newvalue
        totLgCoffee = newvalue
        lblBevSums(Index).Caption = newvalue
        lblBevSums(Index).Enabled = True
        lblBevitem(Index).Enabled = True
    Case 8
        newvalue = CStr(lblBeverage(Index).Caption + 1)
        TotalCost = TotalCost + SmTea
        lblBeverage(Index).Caption = newvalue
        totSmTea = newvalue
        lblBevSums(Index).Caption = newvalue
        lblBevSums(Index).Enabled = True
        lblBevitem(Index).Enabled = True
    Case 9
        newvalue = CStr(lblBeverage(Index).Caption + 1)
        TotalCost = TotalCost + LgTea
        lblBeverage(Index).Caption = newvalue
        totLgTea = newvalue
        lblBevSums(Index).Caption = newvalue
        lblBevSums(Index).Enabled = True
        lblBevitem(Index).Enabled = True
         
End Select

Else

    Select Case Index

    Case 0
         If lblBeverage(Index).Caption = 0 Then
             SideType = "Small Coke"
             lblBevSums(Index).Enabled = False
             lblBevitem(Index).Enabled = False
             TellemAboutit (SideType)
             Exit Sub
         End If
         
         newvalue = CStr(lblBeverage(Index).Caption - 1)
         If newvalue <= 0 Then newvalue = 0
            TotalCost = TotalCost - SmCoke
            lblBeverage(Index).Caption = newvalue
            totSmCoke = newvalue
            lblBevSums(Index).Caption = newvalue
       
    Case 1
         If lblBeverage(Index).Caption = 0 Then
            SideType = "Large Coke"
            lblBevSums(Index).Enabled = False
            lblBevitem(Index).Enabled = False
            TellemAboutit (SideType)
            Exit Sub
         End If
         
            newvalue = CStr(lblBeverage(Index).Caption - 1)
         If newvalue <= 0 Then newvalue = 0
            TotalCost = TotalCost - LgCoke
            lblBeverage(Index).Caption = newvalue
            totLgCoke = newvalue
            lblBevSums(Index).Caption = newvalue
    Case 2
    
        If lblBeverage(Index).Caption = 0 Then
            SideType = "Small Diet"
            lblBevSums(Index).Enabled = False
            lblBevitem(Index).Enabled = False
            TellemAboutit (SideType)
            Exit Sub
        End If
        
            newvalue = CStr(lblBeverage(Index).Caption - 1)
        If newvalue <= 0 Then newvalue = 0
            TotalCost = TotalCost - SmDiet
            lblBeverage(Index).Caption = newvalue
            totSmDiet = newvalue
            lblBevSums(Index).Caption = newvalue
    Case 3
    
         If lblBeverage(Index).Caption = 0 Then
            SideType = "Large Diet"
            lblBevSums(Index).Enabled = False
            lblBevitem(Index).Enabled = False
            TellemAboutit (SideType)
            Exit Sub
        End If
        
            newvalue = CStr(lblBeverage(Index).Caption - 1)
            If newvalue <= 0 Then newvalue = 0
            TotalCost = TotalCost - LgDiet
            lblBeverage(Index).Caption = newvalue
            totLgDiet = newvalue
            lblBevSums(Index).Caption = newvalue
         
     Case 4
     
         If lblBeverage(Index).Caption = 0 Then
            SideType = "Small 7up"
            lblBevSums(Index).Enabled = False
            lblBevitem(Index).Enabled = False
            TellemAboutit (SideType)
            Exit Sub
         End If
            newvalue = CStr(lblBeverage(Index).Caption - 1)
            If newvalue <= 0 Then newvalue = 0
            TotalCost = TotalCost - Sm7up
            lblBeverage(Index).Caption = newvalue
            totSm7up = newvalue
            lblBevSums(Index).Caption = newvalue
    Case 5
    
        If lblBeverage(Index).Caption = 0 Then
            SideType = "Large 7up"
            lblBevSums(Index).Enabled = False
            lblBevitem(Index).Enabled = False
            TellemAboutit (SideType)
            Exit Sub
        End If
            newvalue = CStr(lblBeverage(Index).Caption - 1)
            If newvalue <= 0 Then newvalue = 0
            TotalCost = TotalCost - Lg7up
            lblBeverage(Index).Caption = newvalue
            totLg7up = newvalue
            lblBevSums(Index).Caption = newvalue
    Case 6
        
         If lblBeverage(Index).Caption = 0 Then
            SideType = "Small Coffee"
            lblBevSums(Index).Enabled = False
            lblBevitem(Index).Enabled = False
            TellemAboutit (SideType)
            Exit Sub
        End If
            newvalue = CStr(lblBeverage(Index).Caption - 1)
         If newvalue <= 0 Then newvalue = 0
            TotalCost = TotalCost - SmCoffee
            lblBeverage(Index).Caption = newvalue
            totSmCoffee = newvalue
            lblBevSums(Index).Caption = newvalue
    Case 7
         If lblBeverage(Index).Caption = 0 Then
            SideType = "Large Coffee"
            lblBevSums(Index).Enabled = False
            lblBevitem(Index).Enabled = False
            TellemAboutit (SideType)
            Exit Sub
        End If
            newvalue = CStr(lblBeverage(Index).Caption - 1)
        If newvalue <= 0 Then newvalue = 0
            TotalCost = TotalCost - LgCoffee
            lblBeverage(Index).Caption = newvalue
            totLgCoffee = newvalue
            lblBevSums(Index).Caption = newvalue
            'lblAmount(Index).Caption = newvalue
    Case 8
         If lblBeverage(Index).Caption = 0 Then
            SideType = "Small Tea"
            lblBevSums(Index).Enabled = False
            lblBevitem(Index).Enabled = False
            TellemAboutit (SideType)
            Exit Sub
         End If
            newvalue = CStr(lblBeverage(Index).Caption - 1)
         If newvalue <= 0 Then newvalue = 0
            TotalCost = TotalCost - SmTea
            lblBeverage(Index).Caption = newvalue
            totSmTea = newvalue
            lblBevSums(Index).Caption = newvalue
    Case 9
    
        If lblBeverage(Index).Caption = 0 Then
            SideType = "Large Tea"
            lblBevSums(Index).Enabled = False
            lblBevitem(Index).Enabled = False
            TellemAboutit (SideType)
            Exit Sub
         End If
           newvalue = CStr(lblBeverage(Index).Caption - 1)
         If newvalue <= 0 Then newvalue = 0
            TotalCost = TotalCost - LgTea
            lblBeverage(Index).Caption = newvalue
            totLgTea = newvalue
            lblBevSums(Index).Caption = newvalue
         
  End Select
End If


Adding = True
cmdVoid.Enabled = True
lblOutput.Caption = TotalCost
CheckOutputMenu
End Sub

Private Sub cmdCancel_Click()
Dim response
    response = MsgBox("This will delete the entire order. Are you sure that you want to do this?", vbYesNo, StoreName)
If response = vbYes Then
    ClearOrder
    CheckOutputMenu
Else
    Exit Sub
End If
End Sub



Private Sub cmdExit_Click()
Dim response
    response = MsgBox("This will end your session. Are you sure that you want to do this?", vbYesNo, StoreName)
If response = vbYes Then
    Unload Me
Else
    Exit Sub
End If
End Sub

Private Sub cmdSides_Click(Index As Integer)
If cmdVoid.Enabled = False Then
      Adding = False
Else: Adding = True
End If


Dim SideType As String
If Adding Then
  Select Case Index
    Dim newvalue As Integer
    Case 0
        newvalue = CStr(lblSides(Index).Caption + 1)
        TotalCost = TotalCost + LgFry
        lblSides(Index).Caption = newvalue
        lblSidesAmount(Index).Caption = newvalue
        lblSidesAmount(Index).Enabled = True
        lblSidesItem(Index).Enabled = True
    Case 1
        newvalue = CStr(lblSides(Index).Caption + 1)
        TotalCost = TotalCost + SmFry
        lblSides(Index).Caption = newvalue
        lblSidesAmount(Index).Caption = newvalue
        lblSidesAmount(Index).Enabled = True
        lblSidesItem(Index).Enabled = True
    Case 2
        newvalue = CStr(lblSides(Index).Caption + 1)
        TotalCost = TotalCost + Sundae
        lblSides(Index).Caption = newvalue
        lblSidesAmount(Index).Caption = newvalue
        lblSidesAmount(Index).Enabled = True
        lblSidesItem(Index).Enabled = True
    Case 3
        newvalue = CStr(lblSides(Index).Caption + 1)
        TotalCost = TotalCost + ApplePie
        lblSides(Index).Caption = newvalue
        lblSidesAmount(Index).Caption = newvalue
        lblSidesAmount(Index).Enabled = True
        lblSidesItem(Index).Enabled = True
    End Select
        
Else
Select Case Index
   
    Case 0
        If lblSides(Index) = 0 Then
            SideType = "Large Fry"
            lblSidesAmount(Index).Enabled = False
            lblSidesItem(Index).Enabled = False
            TellemAboutit (SideType)
            Exit Sub
        End If
        newvalue = CStr(lblSides(Index).Caption - 1)
        If newvalue < 0 Then newvalue = 0
              TotalCost = TotalCost - LgFry
              lblSides(Index).Caption = newvalue
              lblSidesAmount(Index).Caption = newvalue
        
         Case 1
          If lblSides(Index) = 0 Then
            SideType = "Small Fry"
            lblSidesAmount(Index).Enabled = False
            lblSidesItem(Index).Enabled = False
            TellemAboutit (SideType)
            Exit Sub
          End If
            newvalue = CStr(lblSides(Index).Caption - 1)
        If newvalue < 0 Then newvalue = 0
            TotalCost = TotalCost - SmFry
            lblSides(Index).Caption = newvalue
            lblSidesAmount(Index).Caption = newvalue
        
    Case 2
        If lblSides(Index) = 0 Then
            SideType = "Sundae"
            lblSidesAmount(Index).Enabled = False
            lblSidesItem(Index).Enabled = False
            TellemAboutit (SideType)
            Exit Sub
        End If
            newvalue = CStr(lblSides(Index).Caption - 1)
        If newvalue < 0 Then newvalue = 0
            TotalCost = TotalCost - Sundae
            lblSides(Index).Caption = newvalue
            lblSidesAmount(Index).Caption = newvalue
        
    Case 3
         If lblSides(Index) = 0 Then
            SideType = "Apple Pie"
            lblSidesAmount(Index).Enabled = False
            lblSidesItem(Index).Enabled = False
         
            TellemAboutit (SideType)
            Exit Sub
         End If
            newvalue = CStr(lblSides(Index).Caption - 1)
         If newvalue < 0 Then newvalue = 0
            TotalCost = TotalCost - ApplePie
            lblSides(Index).Caption = newvalue
            lblSidesAmount(Index).Caption = newvalue
       
        
    End Select

End If

Adding = True
cmdVoid.Enabled = True
lblOutput.Caption = TotalCost
CheckOutputMenu
End Sub

Private Sub cmdSndWich_Click(Index As Integer)
If cmdVoid.Enabled = False Then
      Adding = False
Else: Adding = True
End If

Dim SandType As String

If Adding Then
  Select Case Index
    Dim newvalue As Integer
    Case 0
        newvalue = CStr(lblBurger(Index).Caption + 1)
        TotalCost = TotalCost + CheeseBurger
        lblBurger(Index).Caption = newvalue
        totCheeseBurger = newvalue
        lblAmount(Index).Caption = newvalue
        lblAmount(Index).Enabled = True
        lblSandType(Index).Enabled = True
        
    Case 1
        newvalue = CStr(lblBurger(Index).Caption + 1)
        TotalCost = TotalCost + Hamburger
        lblBurger(Index).Caption = newvalue
        totHamburger = newvalue
        lblAmount(Index).Caption = newvalue
        lblAmount(Index).Enabled = True
        lblSandType(Index).Enabled = True
    Case 2
        newvalue = CStr(lblBurger(Index).Caption + 1)
        TotalCost = TotalCost + DoubleCheeseburger
        lblBurger(Index).Caption = newvalue
        totDoubleCheeseburger = newvalue
        lblAmount(Index).Caption = newvalue
        lblAmount(Index).Enabled = True
        lblSandType(Index).Enabled = True
    Case 3
        newvalue = CStr(lblBurger(Index).Caption + 1)
        TotalCost = TotalCost + ChixSandwich
        lblBurger(Index).Caption = newvalue
        totChixSandwich = newvalue
        lblAmount(Index).Caption = newvalue
        lblAmount(Index).Enabled = True
        lblSandType(Index).Enabled = True
    End Select
        
Else
Select Case Index
   
Case 0
        If lblBurger(Index) = 0 Then
            SandType = "Cheeseburger"
            lblAmount(Index).Enabled = False
            lblSandType(Index).Enabled = False
            TellemAboutit (SandType)
            Exit Sub
        End If
          newvalue = CStr(lblBurger(Index).Caption - 1)
          If newvalue < 0 Then newvalue = 0
            TotalCost = TotalCost - CheeseBurger
            lblBurger(Index).Caption = newvalue
            totCheeseBurger = newvalue
            lblAmount(Index).Caption = totCheeseBurger
        
Case 1
          If lblBurger(Index) = 0 Then
            SandType = "Hamburger"
            lblAmount(Index).Enabled = False
            lblSandType(Index).Enabled = False
            TellemAboutit (SandType)
            Exit Sub
          End If
            newvalue = CStr(lblBurger(Index).Caption - 1)
          If newvalue < 0 Then newvalue = 0
            TotalCost = TotalCost - Hamburger
            totHamburgers = newvalue
            lblBurger(Index).Caption = newvalue
            lblAmount(Index).Caption = totHamburgers
Case 2
        If lblBurger(Index) = 0 Then
            SandType = "Double Cheeseburger"
            lblAmount(Index).Enabled = False
            lblSandType(Index).Enabled = False
            TellemAboutit (SandType)
            Exit Sub
        End If
           newvalue = CStr(lblBurger(Index).Caption - 1)
        If newvalue < 0 Then newvalue = 0
           TotalCost = TotalCost - DoubleCheeseburger
           lblBurger(Index).Caption = newvalue
           totDoubleCheeseburgers = newvalue
           lblAmount(Index).Caption = totDoubleCheeseburgers
        
Case 3
         If lblBurger(Index) = 0 Then
            SandType = "Chicken Sandwich"
            lblAmount(Index).Enabled = False
            lblSandType(Index).Enabled = False
            TellemAboutit (SandType)
            Exit Sub
         End If
            newvalue = CStr(lblBurger(Index).Caption - 1)
         If newvalue < 0 Then newvalue = 0
              TotalCost = TotalCost - ChixSandwich
              lblBurger(Index).Caption = newvalue
              totChixSandwich = newvalue
              lblAmount(Index).Caption = totChixSandwich
        
        
    End Select

End If

Adding = True
cmdVoid.Enabled = True
lblOutput.Caption = TotalCost
CheckOutputMenu
End Sub

Private Sub cmdVoid_Click()
cmdVoid.Enabled = False

End Sub

Private Sub CmdPrint_Click()
Dim myprinter As Printer
Dim moneydue As Currency
first = "Thanks for shopping at McHildenbrands. Here is a summary of your order" & vbCrLf & vbCrLf

For counter1 = 0 To 3
    newvalue = CStr(lblAmount(counter1).Caption)
  If newvalue > 0 Then
        Dim tempstring As String
        OutputMessage = OutputMessage & newvalue & "    " & lblSandType(counter1).Caption & vbCrLf
  End If
  
Next counter1

OutputMessage = OutputMessage & vbCrLf

For counter1 = 0 To 3
    SidesTotal = CStr(lblSidesAmount(counter1).Caption)
    
    If SidesTotal > 0 Then
        OutputMessage = OutputMessage & SidesTotal & "    " & lblSidesItem(counter1).Caption & vbCrLf
    End If
    
Next counter1

For counter1 = 0 To 9
    bevstotal = CStr(lblBevSums(counter1).Caption)
    
    If bevstotal > 0 Then
        OutputMessage = OutputMessage & bevstotal & "    " & lblBevitem(counter1).Caption & vbCrLf
    End If
Next counter1

OutputMessage = OutputMessage + vbCrLf
last = "_____________" & vbCrLf & "$   " & TotalCost & vbCrLf & "x    " & Tax & "  Tax" & vbCrLf & "_____________" & vbCrLf & vbCrLf
moneydue = TotalCost * Tax
moneydue = Int(moneydue * 100) / 100
moneydue = CStr(lblOutput.Caption) + moneydue
part = "$   " & moneydue
finalmessage = first & OutputMessage & last & part
Printer.Print finalmessage
Printer.EndDoc
MsgBox "Press either  the Take Out or Eat In button to clear the order.", vbInformation, StoreName
OutputMessage = ""
finalmessage = ""
End Sub

Private Sub CmdTakeOut_Click()
Dim counter1 As Integer
Dim newvalue As Integer
Dim SidesTotal As Integer
Dim bevstotal As Integer
Dim first As String
Dim last As String
Dim finalmessage As String

If TotalCost = 0 Then
   MsgBox "You will have to enter an order first", vbInformation, StoreName
   Exit Sub
End If
first = "Here is a summary of your order" & vbCrLf & vbCrLf
For counter1 = 0 To 3
    newvalue = CStr(lblAmount(counter1).Caption)
    If newvalue > 0 Then
        Dim tempstring As String
        
        OutputMessage = OutputMessage & newvalue & "    " & lblSandType(counter1).Caption & vbCrLf
    End If
Next counter1
OutputMessage = OutputMessage & vbCrLf
For counter1 = 0 To 3
    SidesTotal = CStr(lblSidesAmount(counter1).Caption)
    
    If SidesTotal > 0 Then
        
        
        OutputMessage = OutputMessage & SidesTotal & "    " & lblSidesItem(counter1).Caption & vbCrLf
    End If
Next counter1

For counter1 = 0 To 9
    bevstotal = CStr(lblBevSums(counter1).Caption)
    
    If bevstotal > 0 Then
        
        
        OutputMessage = OutputMessage & bevstotal & "    " & lblBevitem(counter1).Caption & vbCrLf
    End If
Next counter1

OutputMessage = OutputMessage + vbCrLf
last = "_____________" & vbCrLf & "$   " & TotalCost & vbCrLf & "x    " & Tax & "  Tax" & vbCrLf & "_____________" & vbCrLf & vbCrLf
Dim moneydue As Currency
moneydue = TotalCost * Tax
moneydue = Int(moneydue * 100) / 100
moneydue = CStr(lblOutput.Caption) + moneydue

part = "$   " & moneydue
finalmessage = first & OutputMessage & last & part
MsgBox finalmessage
ClearOrder
CheckOutputMenu
End Sub

Private Sub CmdEatIn_Click()


Dim counter1 As Integer
Dim newvalue As Integer
Dim SidesTotal As Integer
Dim first As String
Dim last As String
Dim finalmessage As String


If TotalCost = 0 Then
   MsgBox "You will have to enter an order first", vbInformation, StoreName
   Exit Sub
End If
first = "Here is a summary of your order" & vbCrLf & vbCrLf
For counter1 = 0 To 3
    newvalue = CStr(lblAmount(counter1).Caption)
    If newvalue > 0 Then
        Dim tempstring As String
        
        OutputMessage = OutputMessage & newvalue & "    " & lblSandType(counter1).Caption & vbCrLf
    End If
Next counter1
OutputMessage = OutputMessage & vbCrLf
For counter1 = 0 To 3
    SidesTotal = CStr(lblSidesAmount(counter1).Caption)
    
    If SidesTotal > 0 Then
        
        
        OutputMessage = OutputMessage & SidesTotal & "    " & lblSidesItem(counter1).Caption & vbCrLf
    End If
Next counter1


For counter1 = 0 To 9
    bevstotal = CStr(lblBevSums(counter1).Caption)
    
    If bevstotal > 0 Then
        
        
        OutputMessage = OutputMessage & bevstotal & "    " & lblBevitem(counter1).Caption & vbCrLf
    End If
Next counter1
OutputMessage = OutputMessage + vbCrLf
last = "_____________" & vbCrLf & "$   " & TotalCost & vbCrLf & "x    " & Tax & "  Tax" & vbCrLf & "_____________" & vbCrLf & vbCrLf
Dim moneydue As Currency
moneydue = TotalCost * Tax
moneydue = Int(moneydue * 100) / 100
moneydue = CStr(lblOutput.Caption) + moneydue

part = "$   " & moneydue
finalmessage = first & OutputMessage & last & part
MsgBox finalmessage
ClearOrder
CheckOutputMenu
End Sub

Private Sub Form_Load()
ClearOrder
CheckOutputMenu
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

