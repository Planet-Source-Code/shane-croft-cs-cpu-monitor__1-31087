VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CS Cpu Monitor"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7245
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   7245
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Window Startup Positions"
      Height          =   1815
      Left            =   4035
      TabIndex        =   18
      Top             =   1575
      Width           =   3135
      Begin VB.OptionButton Option17 
         Caption         =   "Lower Center"
         Height          =   225
         Left            =   1560
         TabIndex        =   25
         Top             =   1440
         Width           =   1455
      End
      Begin VB.OptionButton Option16 
         Caption         =   "Lower Left"
         Height          =   225
         Left            =   1560
         TabIndex        =   24
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton Option15 
         Caption         =   "Lower Right"
         Height          =   225
         Left            =   1560
         TabIndex        =   23
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Option14 
         Caption         =   "Upper Center"
         Height          =   225
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Option13 
         Caption         =   "Upper Left"
         Height          =   225
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton Option12 
         Caption         =   "Upper Right"
         Height          =   225
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Option11 
         Caption         =   "Center"
         Height          =   225
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Misc. Options"
      Height          =   1455
      Left            =   4035
      TabIndex        =   6
      Top             =   15
      Width           =   3135
      Begin VB.CheckBox Check3 
         Caption         =   "Start Program Minimized"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   2655
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Launch At Windows Startup"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   2655
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Always On Top"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Value           =   1  'Checked
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Update Interval"
      Height          =   1455
      Left            =   2235
      TabIndex        =   5
      Top             =   2175
      Width           =   1575
      Begin VB.OptionButton Option5 
         Caption         =   "Low"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         ToolTipText     =   "Update Every 2 Sec."
         Top             =   1080
         Width           =   855
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Normal"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         ToolTipText     =   "Update Every 1 Sec."
         Top             =   720
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "High"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         ToolTipText     =   "Update Every .5 Sec."
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Show Options"
      Height          =   2055
      Left            =   75
      TabIndex        =   4
      Top             =   15
      Width           =   3735
      Begin VB.OptionButton Option10 
         Caption         =   "Show Numbers Only"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1680
         Width           =   3495
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Show Little Bar With Graph"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   3255
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Show Graph"
         Height          =   225
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   3015
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Show Little Bar"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   2895
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Show Large Bar"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Systray Options"
      Height          =   1455
      Left            =   75
      TabIndex        =   1
      Top             =   2175
      Width           =   1935
      Begin VB.OptionButton Option2 
         Caption         =   "Numbers"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Color"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   1560
         Picture         =   "FrmMain.frx":030A
         Top             =   960
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   1560
         Picture         =   "FrmMain.frx":0894
         Top             =   360
         Width           =   240
      End
   End
   Begin VB.PictureBox Pic1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3315
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   615
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3960
      Top             =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Crofts Software"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   4080
      MouseIcon       =   "FrmMain.frx":0E1E
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   3480
      Width           =   3015
   End
   Begin ComctlLib.ImageList IL2 
      Left            =   3960
      Top             =   15
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   101
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":1128
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":1302
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":14DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":16B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":1890
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":1A6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":1C44
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":1E1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":1FF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":21D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":23AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":2586
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":2760
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":293A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":2B14
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":2CEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":2EC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":30A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":327C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":3456
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":3630
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":380A
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":39E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":3BBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":3D98
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":3F72
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":414C
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":4326
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":4500
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":46DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":48B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":4A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":4C68
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":4E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":501C
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":51F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":53D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":55AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":5784
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":595E
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":5B38
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":5D12
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":5EEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":60C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":62A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":647A
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":6654
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":682E
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":6A08
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":6BE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":6DBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":6F96
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":7170
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":734A
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":7524
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":76FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":78D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":7AB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":7C8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":7E66
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":8040
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":821A
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":83F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":85CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":87A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":8982
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":8B5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":8D36
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":8F10
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":90EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":92C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":949E
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":9678
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":9852
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":9A2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":9C06
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":9DE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":9FBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":A194
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":A36E
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":A548
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":A722
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":A8FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":AAD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":ACB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":AE8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":B064
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":B23E
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":B418
            Key             =   ""
         EndProperty
         BeginProperty ListImage90 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":B5F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage91 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":B7CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage92 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":B9A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage93 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":BB80
            Key             =   ""
         EndProperty
         BeginProperty ListImage94 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":BD5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage95 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":BF34
            Key             =   ""
         EndProperty
         BeginProperty ListImage96 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":C10E
            Key             =   ""
         EndProperty
         BeginProperty ListImage97 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":C2E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage98 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":C4C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage99 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":C69C
            Key             =   ""
         EndProperty
         BeginProperty ListImage100 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":C876
            Key             =   ""
         EndProperty
         BeginProperty ListImage101 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":CA50
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList IL1 
      Left            =   3360
      Top             =   15
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   20
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":CC2A
            Key             =   "A1"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":CE54
            Key             =   "A2"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":D07E
            Key             =   "A3"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":D2A8
            Key             =   "A4"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":D4D2
            Key             =   "A5"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":D6FC
            Key             =   "A6"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":D926
            Key             =   "A7"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":DB50
            Key             =   "A8"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":DD7A
            Key             =   "A9"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":DFA4
            Key             =   "A10"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":E1CE
            Key             =   "A11"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":E3F8
            Key             =   "A12"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":E622
            Key             =   "A13"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":E84C
            Key             =   "A14"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":EA76
            Key             =   "A15"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":ECA0
            Key             =   "A16"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":EECA
            Key             =   "A17"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":F0F4
            Key             =   "A18"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":F31E
            Key             =   "A19"
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":F548
            Key             =   "A20"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu MenuShow 
         Caption         =   "Show"
         Begin VB.Menu MenuShowLarge 
            Caption         =   "Show Large Bar"
         End
         Begin VB.Menu MenuLittle 
            Caption         =   "Show Little Bar"
         End
         Begin VB.Menu MenuGraph 
            Caption         =   "Show Graph"
         End
         Begin VB.Menu MenuLittleWithGraph 
            Caption         =   "Show Little Bar With Graph"
         End
         Begin VB.Menu MenuNumbers 
            Caption         =   "Show Numbers Only"
         End
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuSettings 
         Caption         =   "Settings"
      End
      Begin VB.Menu menuabout 
         Caption         =   "About"
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu menuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long


Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long


Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long


Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long


Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    Const ERROR_SUCCESS = 0&
    Const REG_SZ = 1 ' Unicode nul terminated String
    Const REG_DWORD = 4 ' 32-bit number


Public Enum HKeyTypes
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
End Enum

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
Private Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal X&, ByVal Y&, ByVal Flags&) As Long

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

Private Const HKEY_DYN_DATA = &H80000006

Private Const DFC_BUTTON = 4
Private Const DFCS_BUTTON3STATE = &H10

Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_SHOWWINDOW = &H40
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTTOPMOST = -2

Private Const ILD_TRANSPARENT = &H1

Private Type NOTIFYICONDATA
    cbSize As Long
    mhWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function ClipCursor Lib "User32" _
    (lpRect As Any) As Long

Private Declare Function OSGetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function OSGetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function OSGetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function OSWritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function OSWritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Declare Function OSGetProfileInt Lib "kernel32" Alias "GetProfileIntA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Long) As Long
Private Declare Function OSGetProfileSection Lib "kernel32" Alias "GetProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function OSGetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long

Private Declare Function OSWriteProfileSection Lib "kernel32" Alias "WriteProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String) As Long
Private Declare Function OSWriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long

Private Const nBUFSIZEINI = 1024
Private Const nBUFSIZEINIALL = 4096
Private FilePathName As String

Dim TheForm As NOTIFYICONDATA
Dim Status As Long
Private QueryObject As Object
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Sub Check2_Click()
If Check2.Value = 1 Then
Call AddToRun("CS CPU Monitor", App.Path & "\" & App.EXEName & ".exe")
End If
If Check2.Value = 0 Then
Call RemoveFromRun("CS CPU Monitor")
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
'If UnloadMode = 0 Then
Cancel = True
'End If
Me.Hide
End Sub

Private Sub Label1_Click()
On Error Resume Next
Call ShellExecute(hwnd, "Open", "http://www.croftssoftware.com", "", App.Path, 1)
End Sub

Private Sub menuabout_Click()
frmAbout.Show
End Sub

Private Sub menuExit_Click()
On Error Resume Next
Dim fFile As Integer
fFile = FreeFile
'save settings
Open App.Path & "\Settings.inf" For Output As fFile
Print #fFile, "[settings]"
Print #fFile, "scheck1=" & Me.Check1.Value
Print #fFile, "scheck2=" & Me.Check2.Value
Print #fFile, "scheck3=" & Me.Check3.Value
Print #fFile, "soption1=" & Me.Option1.Value
Print #fFile, "soption2=" & Me.Option2.Value
Print #fFile, "soption3=" & Me.Option3.Value
Print #fFile, "soption4=" & Me.Option4.Value
Print #fFile, "soption5=" & Me.Option5.Value
Print #fFile, "soption6=" & Me.Option6.Value
Print #fFile, "soption7=" & Me.Option7.Value
Print #fFile, "soption8=" & Me.Option8.Value
Print #fFile, "soption9=" & Me.Option9.Value
Print #fFile, "soption10=" & Me.Option10.Value
Print #fFile, "soption11=" & Me.Option11.Value
Print #fFile, "soption12=" & Me.Option12.Value
Print #fFile, "soption13=" & Me.Option13.Value
Print #fFile, "soption14=" & Me.Option14.Value
Print #fFile, "soption15=" & Me.Option15.Value
Print #fFile, "soption16=" & Me.Option16.Value
Print #fFile, "soption17=" & Me.Option17.Value
Close fFile
DoEvents

Shell_NotifyIcon NIM_DELETE, TheForm
    'stop the timer
    tmrRefresh.Enabled = False
    'clean up
    QueryObject.Terminate
    Set QueryObject = Nothing
    DoEvents
End
End Sub

Private Sub MenuGraph_Click()
Option8.Value = True
Call Option8_Click
Me.MenuShowLarge.Checked = False
Me.MenuLittle.Checked = False
Me.MenuGraph.Checked = True
Me.MenuLittleWithGraph.Checked = False
Me.MenuNumbers.Checked = False
End Sub

Private Sub MenuLittle_Click()
Option7.Value = True
Call Option7_Click
Me.MenuShowLarge.Checked = False
Me.MenuLittle.Checked = True
Me.MenuGraph.Checked = False
Me.MenuLittleWithGraph.Checked = False
Me.MenuNumbers.Checked = False
End Sub

Private Sub MenuLittleWithGraph_Click()
Option9.Value = True
Call Option9_Click
Me.MenuShowLarge.Checked = False
Me.MenuLittle.Checked = False
Me.MenuGraph.Checked = False
Me.MenuLittleWithGraph.Checked = True
Me.MenuNumbers.Checked = False
End Sub

Private Sub MenuNumbers_Click()
Option10.Value = True
Call Option10_Click
Me.MenuShowLarge.Checked = False
Me.MenuLittle.Checked = False
Me.MenuGraph.Checked = False
Me.MenuLittleWithGraph.Checked = False
Me.MenuNumbers.Checked = True
End Sub

Private Sub MenuSettings_Click()
Me.Show
End Sub
Private Sub MenuShowLarge_Click()
Option6.Value = True
Call Option6_Click
Me.MenuShowLarge.Checked = True
Me.MenuLittle.Checked = False
Me.MenuGraph.Checked = False
Me.MenuLittleWithGraph.Checked = False
Me.MenuNumbers.Checked = False
End Sub

Private Sub Option10_Click()
On Error Resume Next
If Option10.Value = True Then
Unload FrmLittleBar
Unload FrmBigBar
Unload FrmGraph
Unload FrmLittleBarGraph
DoEvents
FrmNumber.Show
Me.MenuShowLarge.Checked = False
Me.MenuLittle.Checked = False
Me.MenuGraph.Checked = False
Me.MenuLittleWithGraph.Checked = False
Me.MenuNumbers.Checked = True
End If
End Sub

Private Sub Option3_Click()
If FrmMain.Option3.Value = True Then
Me.tmrRefresh.Interval = 500
FrmGraph.Timer1.Interval = 500
FrmLittleBar.Timer1.Interval = 500
FrmLittleBarGraph.Timer1.Interval = 500
FrmNumber.Timer1.Interval = 500
FrmBigBar.Timer1.Interval = 500
End If
End Sub

Private Sub Option4_Click()
If FrmMain.Option4.Value = True Then
Me.tmrRefresh.Interval = 1000
FrmGraph.Timer1.Interval = 1000
FrmLittleBar.Timer1.Interval = 1000
FrmLittleBarGraph.Timer1.Interval = 1000
FrmNumber.Timer1.Interval = 1000
FrmBigBar.Timer1.Interval = 1000
End If
End Sub

Private Sub Option5_Click()
If FrmMain.Option5.Value = True Then
Me.tmrRefresh.Interval = 2000
FrmGraph.Timer1.Interval = 2000
FrmLittleBar.Timer1.Interval = 2000
FrmLittleBarGraph.Timer1.Interval = 2000
FrmNumber.Timer1.Interval = 2000
FrmBigBar.Timer1.Interval = 2000
End If
End Sub

Private Sub Option6_Click()
On Error Resume Next
If Option6.Value = True Then
Unload FrmGraph
Unload FrmLittleBar
Unload FrmLittleBarGraph
Unload FrmNumber
DoEvents
FrmBigBar.Show
Me.MenuShowLarge.Checked = True
Me.MenuLittle.Checked = False
Me.MenuGraph.Checked = False
Me.MenuLittleWithGraph.Checked = False
Me.MenuNumbers.Checked = False
End If
End Sub

Private Sub Option7_Click()
On Error Resume Next
If Option7.Value = True Then
Unload FrmGraph
Unload FrmBigBar
Unload FrmLittleBarGraph
Unload FrmNumber
DoEvents
FrmLittleBar.Show
Me.MenuShowLarge.Checked = False
Me.MenuLittle.Checked = True
Me.MenuGraph.Checked = False
Me.MenuLittleWithGraph.Checked = False
Me.MenuNumbers.Checked = False
End If
End Sub

Private Sub Option8_Click()
On Error Resume Next
If Option8.Value = True Then
Unload FrmLittleBar
Unload FrmBigBar
Unload FrmLittleBarGraph
Unload FrmNumber
DoEvents
FrmGraph.Show
Me.MenuShowLarge.Checked = False
Me.MenuLittle.Checked = False
Me.MenuGraph.Checked = True
Me.MenuLittleWithGraph.Checked = False
Me.MenuNumbers.Checked = False
End If
End Sub

Private Sub Option9_Click()
On Error Resume Next
If Option9.Value = True Then
Unload FrmLittleBar
Unload FrmBigBar
Unload FrmGraph
Unload FrmNumber
DoEvents
FrmLittleBarGraph.Show
Me.MenuShowLarge.Checked = False
Me.MenuLittle.Checked = False
Me.MenuGraph.Checked = False
Me.MenuLittleWithGraph.Checked = True
Me.MenuNumbers.Checked = False
End If
End Sub

Private Sub tmrRefresh_Timer()
    Dim Ret As Long
    Dim Which As Long
    'query the CPU usage
    Ret = QueryObject.Query
    If Ret = -1 Then
        tmrRefresh.Enabled = False
        MsgBox "Error while retrieving CPU usage"
    Else
    Status = CStr(Ret)
    
    If Option1.Value = True Then
    Select Case Status
    
        Case Is <= 5
          Which = 1
        Case 6 To 10
          Which = 2
        Case 11 To 15
          Which = 3
        Case 16 To 20
          Which = 4
        Case 21 To 25
          Which = 5
        Case 26 To 30
          Which = 6
        Case 31 To 35
          Which = 7
        Case 36 To 40
          Which = 8
        Case 41 To 45
          Which = 9
        Case 46 To 50
          Which = 10
        Case 51 To 55
          Which = 11
        Case 56 To 60
          Which = 12
        Case 61 To 65
          Which = 13
        Case 66 To 70
          Which = 14
        Case 71 To 75
          Which = 15
        Case 76 To 80
          Which = 16
        Case 81 To 85
          Which = 17
        Case 86 To 90
          Which = 18
        Case 91 To 95
          Which = 19
        Case 96 To 100
          Which = 20
      
    End Select
    
    Pic1.Picture = IL1.ListImages(Which).ExtractIcon
    DoEvents
    ModifyIcon
    DoEvents
    End If
    
    If Option2.Value = True Then
    Pic1.Picture = IL2.ListImages(Status + 1).ExtractIcon
    DoEvents
    ModifyIcon
    DoEvents
    End If
    End If
End Sub

Public Function SysTray()
TheForm.cbSize = Len(TheForm)
    
    TheForm.mhWnd = Pic1.hwnd
    TheForm.hIcon = Pic1.Picture
    TheForm.uId = 1&
    
    TheForm.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    
    TheForm.ucallbackMessage = WM_MOUSEMOVE
    
    'TheForm.szTip = "Michael's Helper" & Chr$(0)
    
    Shell_NotifyIcon NIM_ADD, TheForm
End Function
Function ModifyIcon()
TheForm.cbSize = Len(TheForm)
    
    TheForm.mhWnd = Pic1.hwnd
    TheForm.hIcon = Pic1.Picture
    TheForm.uId = 1&
    
    TheForm.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    
    TheForm.ucallbackMessage = WM_MOUSEMOVE
    
    TheForm.szTip = "CPU Usage: " & Status & "%" & Chr$(0)
    
    Shell_NotifyIcon NIM_MODIFY, TheForm
End Function
Private Sub Form_Load()
On Error Resume Next

Dim AppDir As String
AppDir = App.Path

DoEvents

FilePathName = AppDir + "\Settings.inf"
scheck1 = GetPrivateProfileString("settings", "scheck1", "", FilePathName)
scheck2 = GetPrivateProfileString("settings", "scheck2", "", FilePathName)
scheck3 = GetPrivateProfileString("settings", "scheck3", "", FilePathName)
soption1 = GetPrivateProfileString("settings", "soption1", "", FilePathName)
soption2 = GetPrivateProfileString("settings", "soption2", "", FilePathName)
soption3 = GetPrivateProfileString("settings", "soption3", "", FilePathName)
soption4 = GetPrivateProfileString("settings", "soption4", "", FilePathName)
soption5 = GetPrivateProfileString("settings", "soption5", "", FilePathName)
soption6 = GetPrivateProfileString("settings", "soption6", "", FilePathName)
soption7 = GetPrivateProfileString("settings", "soption7", "", FilePathName)
soption8 = GetPrivateProfileString("settings", "soption8", "", FilePathName)
soption9 = GetPrivateProfileString("settings", "soption9", "", FilePathName)
soption10 = GetPrivateProfileString("settings", "soption10", "", FilePathName)
soption11 = GetPrivateProfileString("settings", "soption11", "", FilePathName)
soption12 = GetPrivateProfileString("settings", "soption12", "", FilePathName)
soption13 = GetPrivateProfileString("settings", "soption13", "", FilePathName)
soption14 = GetPrivateProfileString("settings", "soption14", "", FilePathName)
soption15 = GetPrivateProfileString("settings", "soption15", "", FilePathName)
soption16 = GetPrivateProfileString("settings", "soption16", "", FilePathName)
soption17 = GetPrivateProfileString("settings", "soption17", "", FilePathName)

DoEvents

Me.Check1.Value = scheck1
Me.Check2.Value = scheck2
Me.Check3.Value = scheck3
Me.Option1.Value = soption1
Me.Option2.Value = soption2
Me.Option3.Value = soption3
Me.Option4.Value = soption4
Me.Option5.Value = soption5
Me.Option11.Value = soption11
Me.Option12.Value = soption12
Me.Option13.Value = soption13
Me.Option14.Value = soption14
Me.Option15.Value = soption15
Me.Option16.Value = soption16
Me.Option17.Value = soption17
Me.Option6.Value = soption6
Me.Option7.Value = soption7
Me.Option8.Value = soption8
Me.Option9.Value = soption9
Me.Option10.Value = soption10
DoEvents

Me.Caption = "CS Cpu Monitor v" & App.Major & "." & App.Minor & "." & App.Revision
DoEvents
If FrmMain.Option3.Value = True Then
FrmMain.tmrRefresh.Interval = 500
FrmGraph.Timer1.Interval = 500
FrmLittleBar.Timer1.Interval = 500
FrmLittleBarGraph.Timer1.Interval = 500
FrmNumber.Timer1.Interval = 500
FrmBigBar.Timer1.Interval = 500
End If

If FrmMain.Option4.Value = True Then
FrmMain.tmrRefresh.Interval = 1000
FrmGraph.Timer1.Interval = 1000
FrmLittleBar.Timer1.Interval = 1000
FrmLittleBarGraph.Timer1.Interval = 1000
FrmNumber.Timer1.Interval = 1000
FrmBigBar.Timer1.Interval = 1000
End If

If FrmMain.Option5.Value = True Then
FrmMain.tmrRefresh.Interval = 2000
FrmGraph.Timer1.Interval = 2000
FrmLittleBar.Timer1.Interval = 2000
FrmLittleBarGraph.Timer1.Interval = 2000
FrmNumber.Timer1.Interval = 2000
FrmBigBar.Timer1.Interval = 2000
End If
DoEvents
SysTray
    'set the Priority of this process to 'High'
    'this makes sure our program gets updated, even when
    'another process is consuming lots of CPU cycles
    SetThreadPriority GetCurrentThread, THREAD_BASE_PRIORITY_MAX
    SetPriorityClass GetCurrentProcess, HIGH_PRIORITY_CLASS
    'Initialize our Query object
    If IsWinNT Then
        Set QueryObject = New clsCPUUsageNT
    Else
        Set QueryObject = New clsCPUUsage
    End If
    'Initializing is necesarry for the correct values to be retrieved
    QueryObject.Initialize
    'start the timer
    tmrRefresh.Enabled = True
    'don't wait for the first interval to elapse
    tmrRefresh_Timer


If Check3.Value = 1 Then
Unload FrmGraph
Unload FrmLittleBar
Unload FrmLittleBarGraph
Unload FrmNumber
Unload FrmBigBar
DoEvents
End If
Me.Hide
End Sub
Public Sub pic1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static Rec As Boolean, Msg As Long
    Msg = X / Screen.TwipsPerPixelX
    If Rec = False Then
        Rec = True
        Select Case Msg
            Case WM_LBUTTONDBLCLK:
                'PopupMenu mnufile
            Case WM_LBUTTONDOWN:

            Case WM_LBUTTONUP:
                PopupMenu mnufile
            Case WM_RBUTTONDBLCLK:
                'PopupMenu mnufile
            Case WM_RBUTTONDOWN:
                
            Case WM_RBUTTONUP:
                PopupMenu mnufile
        End Select
        Rec = False
    End If

End Sub
Public Sub AddToRun(ProgramName As String, FileToRun As String)
    'Add a program to the 'Run at Startup' r
    '     egistry keys
    Call SaveString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", ProgramName, FileToRun)
End Sub


Public Sub RemoveFromRun(ProgramName As String)
    'Remove a program from the 'Run at Start
    '     up' registry keys
    Call DeleteValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", ProgramName)
End Sub
Public Sub SaveString(hKey As HKeyTypes, strPath As String, strValue As String, strData As String)
    'EXAMPLE:
    '
    'Call savestring(HKEY_CURRENT_USER, "Sof
    '     tware\VBW\Registry", "String", text1.tex
    '     t)
    '
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strData, Len(strData))
    r = RegCloseKey(keyhand)
End Sub


Public Function DeleteValue(ByVal hKey As HKeyTypes, ByVal strPath As String, ByVal strValue As String)
    'EXAMPLE:
    '
    'Call DeleteValue(HKEY_CURRENT_USER, "So
    '     ftware\VBW\Registry", "Dword")
    '
    Dim keyhand As Long
    r = RegOpenKey(hKey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
End Function


Public Function DeleteKey(ByVal hKey As HKeyTypes, ByVal strPath As String)
    'EXAMPLE:
    '
    'Call DeleteKey(HKEY_CURRENT_USER, "Soft
    '     ware\VBW\Registry")
    '
    Dim keyhand As Long
    r = RegDeleteKey(hKey, strPath)
End Function
Private Function GetPrivateProfileString(ByVal szSection As String, ByVal szEntry As Variant, ByVal szDefault As String, ByVal szFileName As String) As String
   ' *** Get an entry in the inifile ***

   Dim szTmp                     As String
   Dim nRet                      As Long

   If (IsNull(szEntry)) Then
      ' *** Get names of all entries in the named Section ***
      szTmp = String$(nBUFSIZEINIALL, 0)
      nRet = OSGetPrivateProfileString(szSection, 0&, szDefault, szTmp, nBUFSIZEINIALL, szFileName)
   Else
      ' *** Get the value of the named Entry ***
      szTmp = String$(nBUFSIZEINI, 0)
      nRet = OSGetPrivateProfileString(szSection, CStr(szEntry), szDefault, szTmp, nBUFSIZEINI, szFileName)
   End If
   GetPrivateProfileString = Left$(szTmp, nRet)

End Function
Private Function GetProfileString(ByVal szSection As String, ByVal szEntry As Variant, ByVal szDefault As String) As String
   ' *** Get an entry in the WIN inifile ***

   Dim szTmp                    As String
   Dim nRet                     As Long

   If (IsNull(szEntry)) Then
      ' *** Get names of all entries in the named Section ***
      szTmp = String$(nBUFSIZEINIALL, 0)
      nRet = OSGetProfileString(szSection, 0&, szDefault, szTmp, nBUFSIZEINIALL)
   Else
      ' *** Get the value of the named Entry ***
      szTmp = String$(nBUFSIZEINI, 0)
      nRet = OSGetProfileString(szSection, CStr(szEntry), szDefault, szTmp, nBUFSIZEINI)
   End If
   GetProfileString = Left$(szTmp, nRet)

End Function

