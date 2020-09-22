VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Service Centre Entry"
   ClientHeight    =   6255
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10935
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   1190
      Left            =   240
      ScaleHeight     =   1155
      ScaleWidth      =   3090
      TabIndex        =   60
      Top             =   240
      Width           =   3120
      Begin VB.Image Image1 
         Height          =   960
         Left            =   240
         Picture         =   "Form1.frx":0442
         Top             =   120
         Width           =   2595
      End
   End
   Begin VB.TextBox Text15 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   285
      Left            =   5040
      MaxLength       =   50
      TabIndex        =   59
      Top             =   2760
      Width           =   3195
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   285
      Left            =   6960
      MaxLength       =   10
      TabIndex        =   44
      Top             =   5400
      Width           =   1275
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   285
      Left            =   6960
      MaxLength       =   10
      TabIndex        =   43
      Top             =   4920
      Width           =   1275
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   285
      Left            =   6960
      MaxLength       =   10
      TabIndex        =   42
      Top             =   4560
      Width           =   1275
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   285
      Left            =   6960
      MaxLength       =   10
      TabIndex        =   41
      Top             =   4200
      Width           =   1275
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   285
      Left            =   6960
      MaxLength       =   10
      TabIndex        =   40
      Top             =   3840
      Width           =   1275
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   285
      Left            =   6960
      MaxLength       =   10
      TabIndex        =   39
      Top             =   3480
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   285
      Left            =   6960
      MaxLength       =   10
      TabIndex        =   38
      Top             =   3120
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4500
      Left            =   8280
      TabIndex        =   35
      Top             =   1320
      Visible         =   0   'False
      Width           =   2295
      Begin VB.CommandButton Command13 
         BackColor       =   &H0000C000&
         Caption         =   "Delete Vehicle Type"
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
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   4130
         Width           =   2295
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   4125
         ItemData        =   "Form1.frx":8684
         Left            =   0
         List            =   "Form1.frx":8686
         TabIndex        =   36
         Top             =   0
         Width           =   2295
      End
   End
   Begin VB.TextBox Text14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3840
      TabIndex        =   33
      Top             =   1590
      Width           =   1275
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   375
      Left            =   7800
      TabIndex        =   32
      Top             =   7800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      ForeColor       =   &H80000008&
      Height          =   4415
      Left            =   8520
      ScaleHeight     =   4380
      ScaleWidth      =   2145
      TabIndex        =   22
      Top             =   1545
      Width           =   2175
      Begin VB.CommandButton Command12 
         BackColor       =   &H00808080&
         Caption         =   "Move to Billing Section"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3120
         Width           =   1815
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00808080&
         Caption         =   "Reset Numbers"
         Enabled         =   0   'False
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2160
         Width           =   1815
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00808080&
         Caption         =   "Print Bill"
         Enabled         =   0   'False
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00808080&
         Caption         =   "Save"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H000000FF&
         Caption         =   "Delete"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   3840
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00808080&
         Caption         =   "Exit"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00808080&
         Caption         =   "Blank"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00808080&
         Caption         =   "Edit"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1080
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00808080&
      Caption         =   "Go ->"
      Height          =   300
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   480
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      ItemData        =   "Form1.frx":8688
      Left            =   7320
      List            =   "Form1.frx":8695
      TabIndex        =   20
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox Text13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   285
      Left            =   4800
      TabIndex        =   17
      Top             =   480
      Width           =   1755
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00808080&
      Caption         =   "..."
      Height          =   255
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Click for Present Time"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   285
      Left            =   8280
      TabIndex        =   15
      Top             =   960
      Width           =   2190
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   285
      Left            =   4800
      TabIndex        =   0
      Top             =   960
      Width           =   1755
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      TabIndex        =   11
      Top             =   1590
      Width           =   1275
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   4920
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   285
      Left            =   6960
      TabIndex        =   7
      Top             =   2040
      Width           =   1275
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   285
      Left            =   4560
      TabIndex        =   6
      Top             =   2040
      Width           =   1155
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   3540
      Left            =   1320
      TabIndex        =   3
      Top             =   2400
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   3540
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Restore Deleted"
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
      Left            =   11900
      TabIndex        =   1
      Top             =   5880
      Width           =   1695
   End
   Begin MSComctlLib.ListView List 
      Height          =   1695
      Left            =   1560
      TabIndex        =   31
      Top             =   7320
      Visible         =   0   'False
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   2990
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   8421376
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Sl."
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   6703
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Amount"
         Object.Width           =   1834
      EndProperty
   End
   Begin RichTextLib.RichTextBox RT1 
      Height          =   285
      Left            =   3960
      TabIndex        =   45
      Top             =   3120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      _Version        =   393217
      BackColor       =   12632256
      BorderStyle     =   0
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Form1.frx":86BE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RT2 
      Height          =   285
      Left            =   3960
      TabIndex        =   46
      Top             =   3480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      _Version        =   393217
      BackColor       =   12632256
      BorderStyle     =   0
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Form1.frx":874D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox ReT3 
      Height          =   285
      Left            =   3960
      TabIndex        =   47
      Top             =   3840
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      _Version        =   393217
      BackColor       =   12632256
      BorderStyle     =   0
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Form1.frx":87DC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox ReT4 
      Height          =   285
      Left            =   3960
      TabIndex        =   48
      Top             =   4200
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      _Version        =   393217
      BackColor       =   12632256
      BorderStyle     =   0
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Form1.frx":8867
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox ReT5 
      Height          =   285
      Left            =   3960
      TabIndex        =   49
      Top             =   4560
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      _Version        =   393217
      BackColor       =   12632256
      BorderStyle     =   0
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Form1.frx":88F6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox ReT6 
      Height          =   285
      Left            =   3960
      TabIndex        =   50
      Top             =   4920
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      _Version        =   393217
      BackColor       =   12632256
      BorderStyle     =   0
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Form1.frx":8982
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RT7 
      Height          =   285
      Left            =   3960
      TabIndex        =   51
      Top             =   5400
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      _Version        =   393217
      BackColor       =   12632256
      BorderStyle     =   0
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Form1.frx":8A0F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   285
      Left            =   3600
      TabIndex        =   58
      Top             =   2760
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393217
      BackColor       =   12632256
      BorderStyle     =   0
      Appearance      =   0
      TextRTF         =   $"Form1.frx":8ABA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "6"
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
      Index           =   5
      Left            =   3600
      TabIndex        =   57
      Top             =   4920
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
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
      Index           =   4
      Left            =   3600
      TabIndex        =   56
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
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
      Index           =   3
      Left            =   3600
      TabIndex        =   55
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
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
      Index           =   2
      Left            =   3600
      TabIndex        =   54
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
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
      Index           =   1
      Left            =   3600
      TabIndex        =   53
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
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
      Index           =   0
      Left            =   3600
      TabIndex        =   52
      Top             =   3120
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Height          =   1215
      Left            =   3480
      Top             =   240
      Width           =   7215
   End
   Begin VB.Shape Shape3 
      Height          =   6255
      Left            =   0
      Top             =   0
      Width           =   10935
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Height          =   1215
      Left            =   240
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   3120
      TabIndex        =   34
      Top             =   1605
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "By"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   6840
      TabIndex        =   19
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Searching"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   3720
      TabIndex        =   18
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type of Vehicle "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   6840
      TabIndex        =   14
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   3720
      TabIndex        =   13
      Top             =   975
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Job Card No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   360
      TabIndex        =   12
      Top             =   1605
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   390
      Index           =   4
      Left            =   240
      Top             =   1545
      Width           =   8175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Service Centre Entry   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   10
      Left            =   240
      TabIndex        =   10
      Top             =   1545
      Width           =   8175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Entry Time"
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
      Index           =   9
      Left            =   3600
      TabIndex        =   9
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   6120
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   3
      Left            =   3360
      Top             =   1920
      Width           =   5055
   End
   Begin VB.Shape Shape1 
      Height          =   3540
      Index           =   2
      Left            =   3360
      Top             =   2400
      Width           =   5055
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   1
      Left            =   1320
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   0
      Left            =   240
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Job Card"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   285
      TabIndex        =   5
      Top             =   1995
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   1320
      TabIndex        =   4
      Top             =   1995
      Width           =   2055
   End
   Begin VB.Menu MnuFile 
      Caption         =   "File"
      Begin VB.Menu MnuAddVehicle 
         Caption         =   "Add Vehicle Type"
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////
'////                                                                    /////
'////     Developer: Shyam Singh Chandel                                 /////
'////     Jr. Technician (United News of India, Shillong)                /////
'////     URL http://tech.groups.yahoo.com/group/ssc_visual_basic/       /////
'////                                                                    /////
'/////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////

'developer: Shyam Singh Chandel < shyamschandel@rediffmail.com >
Dim M As ListViewPrinter


Private Sub Command1_Click()

LOADJOBCARDNO
Text8 = Format(Time, "hh:mm AM/PM")
Dim RS As New Recordset
 RS.Open "SELECT * FROM USSC", Connect, adOpenStatic, adLockOptimistic, adCmdText
 RS.AddNew
          RS!ServiceNo = Text10.Text
          RS!typeofVehicle = Text12.Text
          RS!NOC = Text15.Text
          RS!VNo = Text11.Text
          RS!EntryTime = Text8.Text
          RS!FullServiceAmt = Text1.Text
          RS!HalfServiceAmt = Text2.Text
          RS!UnderSideAmt = Text3.Text
          RS!EngineWashingAmt = Text4.Text
          RS!WaterSprayAmt = Text5.Text
          RS!GelingWashAmt = Text6.Text
          RS!UnderSide = ReT3.Text
          RS!EngineWashing = ReT4.Text
          RS!WaterSpray = ReT5.Text
          RS!GelingWash = ReT6.Text
          RS!amount = Text7.Text
          RS!PrintStatus = "NOT DONE"
   RS.Update
  SaveJOBCARDNO
  MsgBox "Record has been saved"
   RS.Close
   BlankFields
   List1.Clear
   List2.Clear
  load
End Sub


Private Sub Command10_Click()
PrintingBILL
End Sub

Private Sub Command11_Click()
Dim RSc As New Recordset
RSc.Open "SELECT * FROM NUMBERS", Connect, adOpenStatic, adLockOptimistic, adCmdText
        RSc.AddNew
              RSc!BillNo = 0
              RSc!JobCard = 0
              RSc!RecNo = 0
        RSc.Update
        RSc.Close
        MsgBox "done"
End Sub

Private Sub Command12_Click()
Dim RS As New Recordset
        RS.Open "SELECT * FROM USSC where ServiceNo='" & Text10.Text & "' AND VNo='" & Text11.Text & "'", Connect, adOpenStatic, adLockOptimistic, adCmdText
                currentMode = EditMode
                   RS!PrintStatus = "DONE"
                   RS!PrintBILL = "NOT DONE"
                RS.Update
                MsgBox "RECORD MOVED TO BILLING SECTION."
                load
                

End Sub



Private Sub Command13_Click()
On Error Resume Next
Dim RSJO As New Recordset
 RSJO.Open "SELECT * FROM Vehicle", Connect, adOpenStatic, adLockOptimistic, adCmdText
      RSJO.Delete
  RSJO.Close
  loadVehicle
End Sub

Private Sub Command2_Click()
ins = MsgBox("Are you sure to delete record", vbQuestion + vbYesNo)
If ins = vbYes Then
Dim RS As New Recordset
RS.Open "SELECT * FROM USSC where ServiceNo='" & Text10.Text & "'", Connect, adOpenStatic, adLockOptimistic
   RS.Delete
  MsgBox "Record has been deleted"
  Command5_Click
  RS.Close
List1.Clear
List2.Clear
load
Else
Exit Sub
End If

End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
BlankFields
End Sub

Private Sub Command6_Click()
        Dim RS As New Recordset
        RS.Open "SELECT * FROM USSC WHERE ServiceNo='" & Text10.Text & "' AND VNo='" & Text11.Text & "'", Connect, adOpenStatic, adLockOptimistic, adCmdText
        currentMode = EditMode
              RS!typeofVehicle = Text12.Text
              RS!FullServiceAmt = Text1.Text
              RS!NOC = Text15.Text
              RS!HalfServiceAmt = Text2.Text
              RS!UnderSideAmt = Text3.Text
              RS!EngineWashingAmt = Text4.Text
              RS!WaterSprayAmt = Text5.Text
              RS!GelingWashAmt = Text6.Text
              RS!UnderSide = ReT3.Text
              RS!EngineWashing = ReT4.Text
              RS!WaterSpray = ReT5.Text
              RS!GelingWash = ReT6.Text
              RS!amount = Text7.Text
        RS.Update
        MsgBox "RECORD HAS BEEN EDITED"
        RS.Close
End Sub

Private Sub Command7_Click()
Text8 = Format(Time, "hh:mm AM/PM")
End Sub

Private Sub Command8_Click()
        On Error Resume Next
        List1.Clear
        List2.Clear
        Dim SQL As String
        If Combo1.Text = "Job Card" Then
        SQL = "SELECT * FROM [" & "USSC" & "]  where [" & "Serviceno" & "]" & "LIKE '%" & Text13.Text & "%' and PrintStatus='" & "NOT DONE" & "'"        ' AND PRINT='" & "DONE" & "'"   'Initial SQL Statement  modified by Shyam Singh Chandel on 19-11-2006 at 8.00 AM
        ElseIf Combo1.Text = "Vehicle No" Then
        SQL = "SELECT * FROM [" & "USSC" & "]  where [" & "VNo" & "]" & "LIKE '%" & Text13.Text & "%' and PrintStatus='" & "NOT DONE" & "'"        ' AND PRINT='" & "DONE" & "'"   'Initial SQL Statement  modified by Shyam Singh Chandel on 19-11-2006 at 8.00 AM
        ElseIf Combo1.Text = "Customer Name" Then
        SQL = "SELECT * FROM [" & "USSC" & "]  where [" & "NOC" & "]" & "LIKE '%" & Text13.Text & "%' and PrintStatus='" & "NOT DONE" & "'"        ' AND PRINT='" & "DONE" & "'"   'Initial SQL Statement  modified by Shyam Singh Chandel on 19-11-2006 at 8.00 AM
        End If
        Dim RS As New Recordset
         RS.Open SQL, Connect, adOpenStatic, adLockOptimistic, adCmdText
            Do While Not RS.EOF
              List1.AddItem RS!ServiceNo
              List2.AddItem RS!VNo
            RS.MoveNext
            Loop
         RS.Close
End Sub

Private Sub Command9_Click()
  List.ListItems.Clear
  Set itmX = List.ListItems.Add(, "A", "1")
  itmX.SubItems(1) = RT1.Text
  itmX.SubItems(2) = Text1.Text
  
  Set itmX = List.ListItems.Add(, "B", "2")
  itmX.SubItems(1) = RT2.Text
  itmX.SubItems(2) = Text2.Text
  
  Set itmX = List.ListItems.Add(, "C", "3")
  itmX.SubItems(1) = ReT3.Text
  itmX.SubItems(2) = Text3.Text
  
  Set itmX = List.ListItems.Add(, "D", "4")
  itmX.SubItems(1) = ReT4.Text
  itmX.SubItems(2) = Text4.Text
 
  Set itmX = List.ListItems.Add(, "E", "5")
  itmX.SubItems(1) = ReT5.Text
  itmX.SubItems(2) = Text5.Text
  
  Set itmX = List.ListItems.Add(, "F", "6")
  itmX.SubItems(1) = ReT6.Text
  itmX.SubItems(2) = Text6.Text
  
  Set itmX = List.ListItems.Add(, "G", "")
  itmX.SubItems(1) = "                              Total: "
  itmX.SubItems(2) = Text7.Text
  
End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
'On Error Resume Next
Dim itemX As ListItem
Dim clmx As ColumnHeader
Dim i As Integer
Set M = New ListViewPrinter
Set M.ListViewName = List
M.DrawHorizontalLines = False
M.DrawVerticalLines = False
M.DrawBorder = False
M.BorderDistance = 2
M.PosX = 2380    'Value in Twips
M.PosY = 3430  'Value in Twips
M.HasPicture = True
Text8 = Format(Time, "hh:mm AM/PM")
load
LOADJOBCARDNO
LOADBILLNO
loadVehicle
ReT3.LoadFile App.Path & "\RT3.dll"
ReT4.LoadFile App.Path & "\RT4.TXT"
ReT5.LoadFile App.Path & "\RT5.TXT"
ReT6.LoadFile App.Path & "\RT6.TXT"
Combo1.ListIndex = 0

End Sub
Private Sub load()
On Error Resume Next
List1.Clear
List2.Clear

Dim RS As New Recordset
 RS.Open "SELECT * FROM USSC where printstatus='" & "NOT DONE" & "'", Connect, adOpenStatic, adLockOptimistic, adCmdText
   
   Do While Not RS.EOF
      List1.AddItem RS!ServiceNo
      List2.AddItem RS!VNo
    RS.MoveNext
    Loop
   RS.Close
   
   
End Sub
Private Sub loadVehicle()
On Error Resume Next
List3.Clear
Dim RSJO As New Recordset
 RSJO.Open "SELECT * FROM Vehicle", Connect, adOpenStatic, adLockOptimistic, adCmdText
        Do While Not RSJO.EOF
        
        List3.AddItem RSJO!typeofVehicle
    
    RSJO.MoveNext
    Loop
  RSJO.Close
  
End Sub
Private Sub LOADJOBCARDNO()
On Error Resume Next
Dim RSc As New Recordset
 RSc.Open "SELECT * FROM NUMBERS", Connect, adOpenStatic, adLockOptimistic, adCmdText
        SR = RSc!JobCard
        SRPLUS = SR + 1
        Text10 = "SER " & Format(SRPLUS, "000000")
  RSc.Close
  
End Sub
Private Sub LOADBILLNO()
On Error Resume Next
Dim RSc As New Recordset
RSc.Open "SELECT * FROM NUMBERS", Connect, adOpenStatic, adLockOptimistic, adCmdText
        NR = RSc!BillNo
        NRPLUS = NR + 1
        Text14 = "SEB " & Format(NRPLUS, "000000")
  RSc.Close
  
End Sub
Private Sub SaveJOBCARDNO()
On Error Resume Next
Dim RSc As New Recordset
 RSc.Open "SELECT JobCard FROM NUMBERS", Connect, adOpenStatic, adLockOptimistic, adCmdText
         currentMode = EditMode
              RSc!JobCard = SRPLUS
        RSc.Update
  RSc.Close

End Sub
Private Sub SaveBillNO()
On Error Resume Next
Dim RSc As New Recordset
 RSc.Open "SELECT JobCard FROM NUMBERS", Connect, adOpenStatic, adLockOptimistic, adCmdText
         currentMode = EditMode
              RSc!BillNo = NRPLUS
        RSc.Update
  RSc.Close

End Sub


Private Sub List1_Click()
On Error Resume Next
List2.ListIndex = List1.ListIndex
Dim RS As New Recordset
 RS.Open "SELECT * FROM USSC where Serviceno='" & List1.Text & "' and VNo='" & List2.Text & "'", Connect, adOpenStatic, adLockOptimistic, adCmdText
    Text10 = RS!ServiceNo
    Text12 = RS!typeofVehicle
    Text11 = RS!VNo
    Text15 = RS!NOC
    Text8 = RS!EntryTime
    Text1 = RS!FullServiceAmt
    Text2 = RS!HalfServiceAmt
    Text3 = RS!UnderSideAmt
    Text4 = RS!EngineWashingAmt
    Text5 = RS!WaterSprayAmt
    Text6 = RS!GelingWashAmt
    ReT3 = RS!UnderSide
    ReT4 = RS!EngineWashing
    ReT5 = RS!WaterSpray
    ReT6 = RS!GelingWash
    Text7 = RS!amount
    Command9_Click
 RS.Close
  
End Sub

Sub BlankFields()
  Text1 = ""
  Text2 = ""
  Text3 = ""
  Text4 = ""
  Text5 = ""
  Text6 = ""
  Text8 = ""
  Text9 = ""
  Text10 = ""
  Text11 = ""
  Text12 = ""
  Text13 = ""
  Text15 = ""
End Sub


Private Sub List2_Click()
List1.ListIndex = List2.ListIndex
List1_Click
End Sub

Private Sub List3_Click()
Text12.Text = List3.Text
End Sub

Private Sub List3_DblClick()
Frame1.Visible = False
Text15.SetFocus
End Sub

Private Sub List3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text15.SetFocus
Frame1.Visible = False
End If
End Sub


Private Sub MnuAddVehicle_Click()
FrmAddVehicle.Show
End Sub

Private Sub ReT3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Open App.Path & "\RT3.dll" For Output As #1
Print #1, ReT3.Text
Close #1
End If
End Sub

Private Sub RT4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Open App.Path & "\RT4.TXT" For Output As #2
Print #2, RT4.Text
Close #2
End If
End Sub

Private Sub RT5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Open App.Path & "\RT5.TXT" For Output As #3
Print #3, RT5.Text
Close #3
End If
End Sub

Private Sub RT6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Open App.Path & "\RT6.TXT" For Output As #4
Print #4, RT6.Text
Close #4
End If
End Sub



Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text1.Text = Format(Text1.Text, "0.00")
Text2.SetFocus
End If
End Sub


Private Sub Text11_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Frame1.Visible = True
List3.SetFocus
End If
End Sub

Private Sub Text12_Click()
Frame1.Visible = True
End Sub

Private Sub Text13_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Command8_Click
End If

End Sub

Private Sub Text15_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text1.SetFocus
End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text2.Text = Format(Text2.Text, "0.00")
Text3.SetFocus
End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text3.Text = Format(Text3.Text, "0.00")
Text4.SetFocus
End If
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text4.Text = Format(Text4.Text, "0.00")
Text5.SetFocus
End If
End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text5.Text = Format(Text5.Text, "0.00")
Text6.SetFocus
End If
End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text6.Text = Format(Text6.Text, "0.00")
Command1_Click
End If
End Sub

Private Sub Timer1_Timer()
Text7.Text = Val(Text1.Text) + Val(Text2.Text) + Val(Text3.Text) + Val(Text4.Text) + Val(Text5.Text) + Val(Text6.Text)
Text7.Text = Format(Text7.Text, "0.00")
End Sub

Private Sub PrintingBILL()
Dim d, ts, bs, cs, ds, es, fs, gs

Dim Page As Integer
Dim sngTotalPage As Single
M.NumOfRowsPerPage = 30
M.RowHeight = 12 * 30
If List.ListItems.Count = 0 Then
MsgBox "NO ITEM SELECTED FOR PRINT"
Exit Sub
Else
sngTotalPage = List.ListItems.Count / M.NumOfRowsPerPage
If sngTotalPage - Int(sngTotalPage) <> 0 Then sngTotalPage = Int(sngTotalPage) + 1
Me.ScaleMode = vbPixels 'this must be done, the container [LEDGER in this case] must be in vbpixels scalemode
Printer.ScaleMode = vbTwips
Printer.PaperSize = vbPRPSA4             ' vbPRPSA5
Printer.Orientation = vbPRORPortrait      'vbPRORLandscape
Printer.Font = List.Font.Name
Printer.FontSize = List.Font.Size
While Not M.LastRowPrinted
        Page = Page + 1
        M.SetRows
        Printer.CurrentX = 1900
        Printer.CurrentY = 60: Printer.FontSize = 12: Printer.FontUnderline = True: Printer.FontBold = True: Printer.FontName = "Times New Roman"
        Printer.Print "BILL/CASH MEMO"  ' - " & Text17.Text
        
        Printer.CurrentX = 980
        Printer.CurrentY = 530: Printer.FontItalic = True: Printer.FontUnderline = False: Printer.FontSize = 22: Printer.FontBold = True: Printer.FontName = "Times New Roman"
        Printer.Print "SERVICING CENTRE"
       
        Printer.FontSize = 12: Printer.FontItalic = False: Printer.FontUnderline = False: Printer.FontBold = False: Printer.FontName = "Arial"
        M.PrintHead Printer
        M.PrintBody Printer
        
        Printer.CurrentX = 1800
        Printer.CurrentY = 1030: Printer.FontBold = True: Printer.FontSize = 12: Printer.FontName = "Times New Roman"
        Printer.Print "SERVELL COMPLEX"
        
        Printer.CurrentX = 380
        Printer.CurrentY = 1330: Printer.FontBold = False: Printer.FontSize = 12: Printer.FontName = "Times New Roman"
        Printer.Print "Jingkieng Nongrim Hills, Shillong-793003(Meghalaya)"
        
        Printer.CurrentX = 1980
        Printer.CurrentY = 1580: Printer.FontSize = 12: Printer.FontName = "Times New Roman"
        Printer.Print "Bill No." & Text14.Text
        
        
        Printer.CurrentX = 3380
        Printer.CurrentY = 1580: Printer.FontBold = False: Printer.FontSize = 10: Printer.FontName = "Times New Roman"
        Printer.Print "TEL: 0364-2520730"
        
        Printer.CurrentX = 4390
        Printer.CurrentY = 230: Printer.FontBold = False: Printer.FontSize = 10: Printer.FontName = "Times New Roman"
        Printer.Print " "
        
        Printer.CurrentX = 4390
        Printer.CurrentY = 400: Printer.FontBold = False: Printer.FontSize = 10: Printer.FontName = "Times New Roman"
        Printer.Print " "

        Printer.CurrentX = 380
        Printer.CurrentY = 1930: Printer.FontSize = 12: Printer.FontBold = True: Printer.FontName = "Times New Roman"
        Printer.Print "Vehicle No. :- " & Text11.Text
        
        Printer.CurrentX = 380
        Printer.CurrentY = 2230: Printer.FontSize = 12: Printer.FontBold = False: Printer.FontName = "Times New Roman"
        Printer.Print "Type of Vehicle :- " & Text12.Text
        
        Printer.CurrentX = 380
        Printer.CurrentY = 2630: Printer.FontSize = 12: Printer.FontBold = False: Printer.FontName = "Times New Roman"
        Printer.Print "Entry Time :- " & Text8.Text

        Printer.CurrentX = 3880
        Printer.CurrentY = 2630: Printer.FontSize = 12: Printer.FontBold = False: Printer.FontName = "Times New Roman"
        Printer.Print "Exit Time :- " & Text9.Text
        
        Printer.CurrentX = 380
        Printer.CurrentY = 3030: Printer.FontSize = 10: Printer.FontBold = False: Printer.FontName = "Times New Roman"
        Printer.Print "____________________________________________________________"
        
        Printer.CurrentX = 380
        Printer.CurrentY = 7880: Printer.FontSize = 12: Printer.FontBold = True: Printer.FontName = "Times New Roman"
        Printer.Print "Date: " & Format(Date, "dd-mm-yy")
        
        Printer.FontSize = 8: Printer.FontName = "Times New Roman"
        Printer.CurrentX = 380
        Printer.CurrentY = 5800: Printer.FontBold = True: Printer.FontSize = 10: Printer.FontName = "Times New Roman"
        Printer.Print "____________________________________________________________"
        
        Printer.CurrentX = 380
        Printer.CurrentY = 6150: Printer.FontBold = True: Printer.FontSize = 10: Printer.FontName = "Times New Roman"
        Printer.Print "____________________________________________________________"

        
        
        Text7.SelStart = 0
        gs = Text6.MaxLength - Len(Text6.Text)
        Printer.CurrentX = 4780
        Printer.CurrentY = 6030: Printer.FontSize = 12: Printer.FontBold = True: Printer.FontName = "Times New Roman"
        Printer.Print Space(gs) & Text7.Text
        
        Printer.CurrentX = 380
        Printer.CurrentY = 7300: Printer.FontSize = 12: Printer.FontName = "Times New Roman"
        Printer.Print "Receive with thanks " & Chr(13) & "Rs. " & Text7.Text & " From " & Chr(13) & "Mr. " & Text15.Text

        Printer.CurrentX = 5000
        Printer.CurrentY = 7880: Printer.FontBold = False: Printer.FontItalic = True: Printer.FontSize = 12: Printer.FontName = "Times New Roman"
        Printer.Print "For"
        
        Printer.CurrentX = 5500
        Printer.CurrentY = 7880: Printer.FontBold = False: Printer.FontItalic = False: Printer.FontSize = 12: Printer.FontName = "Times New Roman"
        Printer.Print "A.P.P.H"

        Printer.CurrentX = 500
        Printer.CurrentY = 14500: Printer.FontBold = True
        Printer.NewPage
        
Wend
        Printer.EndDoc
        M.LastRowPrinted = False
        Me.ScaleMode = vbTwips
        Dim SQL As String
        Dim RS As New Recordset
        SQL = "SELECT * FROM USSC WHERE ServiceNo='" & Text10.Text & "' AND VNo='" & Text11.Text & "'"
        RS.Open SQL, Connect, adOpenStatic, adLockOptimistic, adCmdText
        If RS.EOF = False Then
                currentMode = EditMode
                   RS!PrintBILL = "DONE"
                RS.Update
                load
                MsgBox "Done"
                End If
End If
        
End Sub

Private Sub PrintJOB()
Dim d, ts, bs, cs, ds, es, fs, gs

Dim Page As Integer
Dim sngTotalPage As Single
M.NumOfRowsPerPage = 30
M.RowHeight = 12 * 30
If List.ListItems.Count = 0 Then
MsgBox "NO ITEM SELECTED FOR PRINT"
Exit Sub
Else
sngTotalPage = List.ListItems.Count / M.NumOfRowsPerPage
If sngTotalPage - Int(sngTotalPage) <> 0 Then sngTotalPage = Int(sngTotalPage) + 1
Me.ScaleMode = vbPixels 'this must be done, the container [LEDGER in this case] must be in vbpixels scalemode
Printer.ScaleMode = vbTwips
Printer.PaperSize = vbPRPSA4             ' vbPRPSA5
Printer.Orientation = vbPRORPortrait      'vbPRORLandscape
Printer.Font = List.Font.Name
Printer.FontSize = List.Font.Size
While Not M.LastRowPrinted
        Page = Page + 1
        M.SetRows
        Printer.CurrentX = 3500
        Printer.CurrentY = 60: Printer.FontSize = 12: Printer.FontUnderline = True: Printer.FontBold = True: Printer.FontName = "Times New Roman"
        Printer.Print "JOB CARD NO.:- " & Text10.Text
        
        Printer.CurrentX = 2980
        Printer.CurrentY = 530: Printer.FontItalic = True: Printer.FontUnderline = False: Printer.FontSize = 22: Printer.FontBold = True: Printer.FontName = "Times New Roman"
        Printer.Print "SERVICING CENTRE"
       
        Printer.FontSize = 12: Printer.FontItalic = False: Printer.FontUnderline = False: Printer.FontBold = False: Printer.FontName = "Arial"
        M.PrintHead Printer
        M.PrintBody Printer
        
        Printer.CurrentX = 3800
        Printer.CurrentY = 1030: Printer.FontBold = True: Printer.FontSize = 12: Printer.FontName = "Times New Roman"
        Printer.Print "SERVELL COMPLEX"
        
        Printer.CurrentX = 2380
        Printer.CurrentY = 1330: Printer.FontBold = False: Printer.FontSize = 12: Printer.FontName = "Times New Roman"
        Printer.Print "Jingkieng Nongrim Hills, Shillong-793003(Meghalaya)"
        
        
        Printer.CurrentX = 6390
        Printer.CurrentY = 230: Printer.FontBold = False: Printer.FontSize = 10: Printer.FontName = "Times New Roman"
        Printer.Print " "
        
        Printer.CurrentX = 6390
        Printer.CurrentY = 400: Printer.FontBold = False: Printer.FontSize = 10: Printer.FontName = "Times New Roman"
        Printer.Print " "

        Printer.CurrentX = 2380
        Printer.CurrentY = 1930: Printer.FontSize = 12: Printer.FontBold = True: Printer.FontName = "Times New Roman"
        Printer.Print "Vehicle No. :- " & Text11.Text
        
        Printer.CurrentX = 2380
        Printer.CurrentY = 2230: Printer.FontSize = 12: Printer.FontBold = False: Printer.FontName = "Times New Roman"
        Printer.Print "Type of Vehicle :- " & Text12.Text
        
        Printer.CurrentX = 2380
        Printer.CurrentY = 2630: Printer.FontSize = 12: Printer.FontBold = False: Printer.FontName = "Times New Roman"
        Printer.Print "Entry Time :- " & Text8.Text

        Printer.CurrentX = 5880
        Printer.CurrentY = 2630: Printer.FontSize = 12: Printer.FontBold = False: Printer.FontName = "Times New Roman"
        Printer.Print "Exit Time :- " & Text9.Text
        
        Printer.CurrentX = 2380
        Printer.CurrentY = 3030: Printer.FontSize = 10: Printer.FontBold = False: Printer.FontName = "Times New Roman"
        Printer.Print "____________________________________________________________"
        
        Printer.CurrentX = 2380
        Printer.CurrentY = 7880: Printer.FontSize = 12: Printer.FontBold = True: Printer.FontName = "Times New Roman"
        Printer.Print "Date: " & Format(Date, "dd-mm-yy")
        
        Printer.FontSize = 8: Printer.FontName = "Times New Roman"
        Printer.CurrentX = 2380
        Printer.CurrentY = 5700: Printer.FontBold = True: Printer.FontSize = 10: Printer.FontName = "Times New Roman"
        Printer.Print "____________________________________________________________"
        
        Printer.CurrentX = 2380
        Printer.CurrentY = 6150: Printer.FontBold = True: Printer.FontSize = 10: Printer.FontName = "Times New Roman"
        Printer.Print "____________________________________________________________"

        
        Printer.CurrentX = 7000
        Printer.CurrentY = 7880: Printer.FontBold = False: Printer.FontItalic = True: Printer.FontSize = 12: Printer.FontName = "Times New Roman"
        Printer.Print "For"
        
        Printer.CurrentX = 7500
        Printer.CurrentY = 7880: Printer.FontBold = False: Printer.FontItalic = False: Printer.FontSize = 12: Printer.FontName = "Times New Roman"
        Printer.Print "A.P.P.H"

        Printer.CurrentX = 2500
        Printer.CurrentY = 14500: Printer.FontBold = True
        Printer.NewPage
        
Wend
        Printer.EndDoc
        M.LastRowPrinted = False
        Me.ScaleMode = vbTwips
        Dim RS As New Recordset
       
        RS.Open "SELECT * FROM USSC where ServiceNo='" & Text10.Text & "' AND VNo='" & Text11.Text & "'", Connect, adOpenStatic, adLockOptimistic, adCmdText
        
                currentMode = EditMode
                   RS!PrintStatus = "DONE"
                   RS!PrintBILL = "NOT DONE"
                RS.Update
                MsgBox "RECORD MOVED TO BILLING SECTION."
                load
                
                End If
        
End Sub


