VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "untitled - CozIcon"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10890
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":10C2
   ScaleHeight     =   529
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   726
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picExtra 
      Height          =   1335
      Left            =   0
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   56
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1590
      Width           =   900
      Begin VB.PictureBox picToolSel 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   900
         Index           =   10
         Left            =   240
         MouseIcon       =   "frmMain.frx":2184
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":2E4E
         ScaleHeight     =   60
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   300
         Begin VB.Shape shpToolSel 
            BorderColor     =   &H000000FF&
            Height          =   255
            Index           =   10
            Left            =   0
            Top             =   360
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin VB.PictureBox picToolSel 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   600
         Index           =   1
         Left            =   240
         MouseIcon       =   "frmMain.frx":2EBA
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":3B84
         ScaleHeight     =   40
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   29
         Top             =   0
         Visible         =   0   'False
         Width           =   300
         Begin VB.Shape shpToolSel 
            BorderColor     =   &H000000FF&
            Height          =   255
            Index           =   1
            Left            =   0
            Top             =   360
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin VB.PictureBox picToolSel 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1200
         Index           =   5
         Left            =   240
         MouseIcon       =   "frmMain.frx":3C31
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":48FB
         ScaleHeight     =   80
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   300
         Begin VB.Shape shpToolSel 
            BorderColor     =   &H000000FF&
            Height          =   255
            Index           =   5
            Left            =   0
            Top             =   360
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin VB.PictureBox picToolSel 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1200
         Index           =   3
         Left            =   120
         MouseIcon       =   "frmMain.frx":4AD4
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":579E
         ScaleHeight     =   80
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   40
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   8
         Visible         =   0   'False
         Width           =   600
         Begin VB.Shape shpToolSel 
            BorderColor     =   &H000000FF&
            Height          =   255
            Index           =   3
            Left            =   0
            Top             =   360
            Visible         =   0   'False
            Width           =   255
         End
      End
   End
   Begin VB.PictureBox picUndo 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   0
      Left            =   360
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   120
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6960
      Visible         =   0   'False
      Width           =   480
   End
   Begin CozIcon.winConnect wc 
      Left            =   0
      Top             =   7320
      _ExtentX        =   1535
      _ExtentY        =   661
   End
   Begin VB.FileListBox flbPlugin 
      Height          =   285
      Left            =   240
      Pattern         =   "*.exe"
      TabIndex        =   18
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComctlLib.ImageList ilToolBar 
      Left            =   1080
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5AE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5FE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":64EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":69EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6EF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":73F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":78FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7DFE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbTools 
      Height          =   390
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   688
      ButtonWidth     =   714
      ButtonHeight    =   688
      Appearance      =   1
      Style           =   1
      ImageList       =   "ilToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Select"
            Object.Tag             =   "11"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Text"
            Object.Tag             =   "12"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pen Tool"
            Object.Tag             =   "1"
            ImageIndex      =   6
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Line Tool"
            Object.Tag             =   "10"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eye Dropper"
            Object.Tag             =   "9"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Fill Color"
            Object.Tag             =   "2"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Box"
            Object.Tag             =   "3"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Circle"
            Object.Tag             =   "5"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   1080
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Flags           =   7
   End
   Begin VB.PictureBox picTrans 
      BackColor       =   &H00808000&
      Height          =   180
      Left            =   120
      MouseIcon       =   "frmMain.frx":8302
      MousePointer    =   99  'Custom
      ScaleHeight     =   120
      ScaleWidth      =   120
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Transparent Color"
      Top             =   3480
      Width           =   180
   End
   Begin VB.PictureBox picClr 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   60
      MouseIcon       =   "frmMain.frx":8FCC
      MousePointer    =   99  'Custom
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Left Button Color"
      Top             =   3000
      Width           =   480
   End
   Begin VB.PictureBox picClr 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   300
      MouseIcon       =   "frmMain.frx":9C96
      MousePointer    =   99  'Custom
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Right Button Color"
      Top             =   3240
      Width           =   480
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00808080&
      Height          =   6780
      Left            =   960
      ScaleHeight     =   448
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   656
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   9900
      Begin MSComctlLib.ImageList ilIcons 
         Left            =   960
         Top             =   5880
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":A960
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":B23C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":BF18
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.HScrollBar hsEdit 
         Height          =   255
         LargeChange     =   10
         Left            =   0
         Max             =   0
         SmallChange     =   5
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   6465
         Width           =   9495
      End
      Begin VB.PictureBox picBrushes 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   6960
         ScaleHeight     =   177
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   169
         TabIndex        =   53
         Top             =   5280
         Width           =   2535
         Begin CozIcon.CustomBrushes cbBox 
            Height          =   2055
            Left            =   120
            TabIndex        =   56
            Top             =   420
            Visible         =   0   'False
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   3625
         End
         Begin CozIcon.CustomBrushes cbPens 
            Height          =   2055
            Left            =   120
            TabIndex        =   55
            Top             =   420
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   3625
         End
         Begin MSComctlLib.TabStrip tabBrushes 
            Height          =   2535
            Left            =   0
            TabIndex        =   54
            Top             =   0
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   4471
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   2
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Pens"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Boxes"
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   168
            Y1              =   0
            Y2              =   0
         End
      End
      Begin VB.PictureBox picLayers 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   5175
         Left            =   6960
         ScaleHeight     =   345
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   169
         TabIndex        =   32
         Top             =   0
         Width           =   2535
         Begin MSComctlLib.ImageList ilLyr 
            Left            =   1800
            Top             =   360
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   8
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":CC74
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":D010
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":D3AC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":D748
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":DAE4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":DE80
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":E21C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":E5B8
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.Toolbar tbLyr 
            Height          =   330
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            ImageList       =   "ilLyr"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   7
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Clear Layer"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Toogle Layer Visibility"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Move Layer Up"
                  ImageIndex      =   4
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Move Layer Down"
                  ImageIndex      =   5
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Merge Layer Up"
                  ImageIndex      =   7
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Merge Layer Down"
                  ImageIndex      =   6
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Flatten Image"
                  ImageIndex      =   8
               EndProperty
            EndProperty
         End
         Begin VB.VScrollBar vsLyr 
            Height          =   2475
            LargeChange     =   120
            Left            =   2280
            Max             =   0
            SmallChange     =   60
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   360
            Width           =   255
         End
         Begin VB.PictureBox picLayersBack 
            BackColor       =   &H0080C0FF&
            BorderStyle     =   0  'None
            Height          =   4335
            Left            =   0
            ScaleHeight     =   289
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   153
            TabIndex        =   33
            Top             =   360
            Width           =   2295
            Begin VB.PictureBox picLyr 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00808000&
               BorderStyle     =   0  'None
               Height          =   480
               Index           =   9
               Left            =   120
               MouseIcon       =   "frmMain.frx":E954
               MousePointer    =   99  'Custom
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   51
               TabStop         =   0   'False
               Tag             =   "visible"
               Top             =   5520
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.PictureBox picLyr 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00808000&
               BorderStyle     =   0  'None
               Height          =   480
               Index           =   8
               Left            =   120
               MouseIcon       =   "frmMain.frx":F61E
               MousePointer    =   99  'Custom
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   49
               TabStop         =   0   'False
               Tag             =   "visible"
               Top             =   4920
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.PictureBox picLyr 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00808000&
               BorderStyle     =   0  'None
               Height          =   480
               Index           =   7
               Left            =   120
               MouseIcon       =   "frmMain.frx":102E8
               MousePointer    =   99  'Custom
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   48
               TabStop         =   0   'False
               Tag             =   "visible"
               Top             =   4320
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.PictureBox picLyr 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00808000&
               BorderStyle     =   0  'None
               Height          =   480
               Index           =   6
               Left            =   120
               MouseIcon       =   "frmMain.frx":10FB2
               MousePointer    =   99  'Custom
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   47
               TabStop         =   0   'False
               Tag             =   "visible"
               Top             =   4560
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.PictureBox picLyr 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00808000&
               BorderStyle     =   0  'None
               Height          =   480
               Index           =   5
               Left            =   120
               MouseIcon       =   "frmMain.frx":11C7C
               MousePointer    =   99  'Custom
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   46
               TabStop         =   0   'False
               Tag             =   "visible"
               Top             =   4440
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.PictureBox picLyr 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00808000&
               BorderStyle     =   0  'None
               Height          =   480
               Index           =   4
               Left            =   120
               MouseIcon       =   "frmMain.frx":12946
               MousePointer    =   99  'Custom
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   44
               TabStop         =   0   'False
               Tag             =   "visible"
               Top             =   3480
               Width           =   480
            End
            Begin VB.PictureBox picLyr 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00808000&
               BorderStyle     =   0  'None
               Height          =   480
               Index           =   3
               Left            =   120
               MouseIcon       =   "frmMain.frx":13610
               MousePointer    =   99  'Custom
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   42
               TabStop         =   0   'False
               Tag             =   "visible"
               Top             =   2640
               Width           =   480
            End
            Begin VB.PictureBox picLyr 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00808000&
               BorderStyle     =   0  'None
               Height          =   480
               Index           =   2
               Left            =   120
               MouseIcon       =   "frmMain.frx":142DA
               MousePointer    =   99  'Custom
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   40
               TabStop         =   0   'False
               Tag             =   "visible"
               Top             =   1800
               Width           =   480
            End
            Begin VB.PictureBox picLyr 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00808000&
               BorderStyle     =   0  'None
               Height          =   480
               Index           =   1
               Left            =   120
               MouseIcon       =   "frmMain.frx":14FA4
               MousePointer    =   99  'Custom
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   38
               TabStop         =   0   'False
               Tag             =   "visible"
               Top             =   960
               Width           =   480
            End
            Begin VB.PictureBox picLyr 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00808000&
               BorderStyle     =   0  'None
               Height          =   480
               Index           =   0
               Left            =   120
               MouseIcon       =   "frmMain.frx":15C6E
               MousePointer    =   99  'Custom
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   35
               TabStop         =   0   'False
               Tag             =   "visible"
               Top             =   120
               Width           =   480
            End
            Begin VB.Label lblLyr 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Undo Layer 5"
               Height          =   195
               Index           =   9
               Left            =   720
               TabIndex        =   52
               Top             =   5640
               Visible         =   0   'False
               Width           =   960
            End
            Begin VB.Label lblLyr 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Undo Layer 4"
               Height          =   195
               Index           =   8
               Left            =   720
               TabIndex        =   50
               Top             =   5040
               Visible         =   0   'False
               Width           =   960
            End
            Begin VB.Label lblLyr 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Layer 4"
               Height          =   195
               Index           =   4
               Left            =   960
               TabIndex        =   45
               Top             =   3600
               Width           =   525
            End
            Begin VB.Label lblLyr 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Layer 3"
               Height          =   195
               Index           =   3
               Left            =   960
               TabIndex        =   43
               Top             =   2760
               Width           =   525
            End
            Begin VB.Label lblLyr 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Layer 2"
               Height          =   195
               Index           =   2
               Left            =   960
               TabIndex        =   41
               Top             =   1920
               Width           =   525
            End
            Begin VB.Label lblLyr 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Layer 1"
               Height          =   195
               Index           =   1
               Left            =   960
               TabIndex        =   39
               Top             =   1080
               Width           =   525
            End
            Begin VB.Shape shpLyrSel 
               BorderColor     =   &H00FF8080&
               Height          =   870
               Left            =   45
               Top             =   45
               Width           =   2175
            End
            Begin VB.Label lblLyr 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "BackGround"
               Height          =   195
               Index           =   0
               Left            =   960
               TabIndex        =   36
               Top             =   240
               Width           =   900
            End
         End
      End
      Begin VB.OptionButton optEdit 
         Height          =   375
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   5760
         Width           =   375
      End
      Begin VB.VScrollBar vsEdit 
         Height          =   6375
         LargeChange     =   10
         Left            =   9600
         Max             =   0
         SmallChange     =   5
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox picEdit 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808000&
         BorderStyle     =   0  'None
         Height          =   5760
         Left            =   120
         MousePointer    =   2  'Cross
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   384
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   440
         TabIndex        =   0
         Top             =   120
         Width           =   6600
         Begin VB.Shape shpSel 
            BorderStyle     =   3  'Dot
            Height          =   735
            Left            =   1080
            Top             =   3120
            Visible         =   0   'False
            Width           =   735
         End
      End
   End
   Begin VB.PictureBox picIconsBack 
      BackColor       =   &H00808080&
      Height          =   2655
      Left            =   0
      ScaleHeight     =   173
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   56
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3840
      Width           =   900
      Begin VB.VScrollBar vsIcons 
         Height          =   495
         Left            =   0
         SmallChange     =   34
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2100
         Width           =   855
      End
      Begin VB.PictureBox picIconsMove 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   0
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   56
         TabIndex        =   4
         Top             =   120
         Width           =   840
         Begin VB.PictureBox picIcon 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00808000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   0
            Left            =   180
            MouseIcon       =   "frmMain.frx":16938
            MousePointer    =   99  'Custom
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   5
            TabStop         =   0   'False
            Tag             =   "untitled"
            Top             =   0
            Width           =   480
         End
         Begin VB.Label txtIcon 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   420
            Index           =   0
            Left            =   90
            TabIndex        =   22
            Top             =   540
            Width           =   675
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   960
      TabIndex        =   7
      Top             =   6840
      Width           =   9900
      Begin VB.PictureBox picClrBack 
         BorderStyle     =   0  'None
         Height          =   660
         Index           =   0
         Left            =   120
         ScaleHeight     =   44
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   497
         TabIndex        =   19
         Top             =   160
         Width           =   7455
         Begin VB.PictureBox picClrSel 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   1
            Left            =   3480
            ScaleHeight     =   18
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   255
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   0
            Width           =   3855
         End
         Begin VB.PictureBox picClrSel 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   2
            Left            =   3480
            ScaleHeight     =   18
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   255
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   330
            Width           =   3855
         End
         Begin CozIcon.ColorPicker cp 
            Height          =   495
            Left            =   0
            TabIndex        =   57
            Top             =   60
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   873
            BackColor       =   14737632
         End
      End
      Begin VB.PictureBox picClrTextContainer 
         BorderStyle     =   0  'None
         Height          =   680
         Left            =   7560
         ScaleHeight     =   45
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   97
         TabIndex        =   12
         Top             =   120
         Width           =   1455
         Begin VB.TextBox txtClr 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   24
            Text            =   "000000"
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtClr 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            Height          =   285
            Index           =   2
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   15
            Text            =   "0"
            Top             =   60
            Width           =   495
         End
         Begin VB.TextBox txtClr 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Index           =   1
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   14
            Text            =   "0"
            Top             =   60
            Width           =   495
         End
         Begin VB.TextBox txtClr 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
            Height          =   285
            Index           =   0
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   13
            Text            =   "0"
            Top             =   60
            Width           =   495
         End
      End
   End
   Begin VB.Label lblSwitch 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "q"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   540
      MouseIcon       =   "frmMain.frx":17602
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   2970
      Width           =   300
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New"
         Begin VB.Menu mnuFileNew48 
            Caption         =   "48x48 Icon"
            Shortcut        =   ^K
         End
         Begin VB.Menu mnuFileNew32 
            Caption         =   "32x32 Icon"
            Shortcut        =   ^N
         End
         Begin VB.Menu mnuFileNew16 
            Caption         =   "16x16 Icon"
            Shortcut        =   ^M
         End
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuLinei43i566 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRecent 
         Caption         =   "Recent Files"
         Begin VB.Menu mnuFileRecentFiles 
            Caption         =   "(empty)"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu mnuLine39k3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save As..."
      End
      Begin VB.Menu mnuFileSaveAllAs 
         Caption         =   "Save All As..."
      End
      Begin VB.Menu mnuLine34445 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "Close"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuLinei821 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuLinej4049d 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuLinek431 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuViewZoom 
         Caption         =   "Zoom"
         Begin VB.Menu mnuViewZoomSize 
            Caption         =   "100%"
            Index           =   1
            Shortcut        =   {F5}
         End
         Begin VB.Menu mnuViewZoomSize 
            Caption         =   "500%"
            Index           =   5
            Shortcut        =   {F6}
         End
         Begin VB.Menu mnuViewZoomSize 
            Caption         =   "1000%"
            Index           =   10
            Shortcut        =   {F7}
         End
         Begin VB.Menu mnuViewZoomSize 
            Caption         =   "2000%"
            Checked         =   -1  'True
            Index           =   20
            Shortcut        =   {F8}
         End
      End
      Begin VB.Menu mnuLine45r4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewGrid 
         Caption         =   "Grid"
         Checked         =   -1  'True
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuLine44r78 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewBrushes 
         Caption         =   "Brushes"
         Checked         =   -1  'True
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuViewLayers 
         Caption         =   "Layers"
         Checked         =   -1  'True
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu mnuImage 
      Caption         =   "Image"
      Begin VB.Menu mnuImageFlipHo 
         Caption         =   "Flip Horizontal"
      End
      Begin VB.Menu mnuImageFlipVert 
         Caption         =   "Flip Vertical"
      End
      Begin VB.Menu mnuImageRotate 
         Caption         =   "Rotate Image"
      End
      Begin VB.Menu mnuLine33kk3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImageOriginal 
         Caption         =   "Original"
      End
      Begin VB.Menu mnuImagePal 
         Caption         =   "Palette"
      End
      Begin VB.Menu mnuLineii442 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImageCopyTo 
         Caption         =   "Copy to"
         Begin VB.Menu mnuImageCopyNew 
            Caption         =   "New"
         End
         Begin VB.Menu mnuImageCopyLayer 
            Caption         =   "Layer"
            Visible         =   0   'False
            Begin VB.Menu mnuImageCopyToLayer 
               Caption         =   "2"
               Index           =   2
            End
         End
      End
      Begin VB.Menu mnuImageClear 
         Caption         =   "Clear"
      End
   End
   Begin VB.Menu mnuFilters 
      Caption         =   "Filters"
      Begin VB.Menu mnuFiltersRun 
         Caption         =   "(empty)"
         Index           =   0
      End
   End
   Begin VB.Menu mnuPlugins 
      Caption         =   "Plugins"
      Begin VB.Menu mnuPluginsRun 
         Caption         =   "(empty)"
         Index           =   0
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpTest 
         Caption         =   "Test Icon"
      End
      Begin VB.Menu mnuLinei392 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'

'Supports 16x16 and 32x32 Full Transparency True Color Icons. Has most of the basic image editor type settings: pen, fill, eclipse, rect, text, selections, etc.
'But also includes gradient rectangles and gradient eclipses. Also has mild support for custom brushes and my own type of plugins.
'Drag and Drop is supported as well is bmp, gif or jpeg to icon conversion. Can also load language packs for our non-english reading friends.

'Included is the Main Project, some sample Plugins Projects, some custom brush files, a read me file about custom brushes and a read me concerning the language packs(with two example packs: one for english and the other for spanish).
'Some subsets of the Main Project include dynamically loading custom language packs, winconnect(a custom activex for communication through different windows) and hard coding icons (not using savepicture via Visual Basic).


Option Explicit
Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal fuFillType As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Enum UDE_TOOLS 'these are our tools
 tDraw = 1
 tFill = 2
 tBox = 3
 tBoxFill = 4
 tCircle = 5
 tCircleFill = 6
 tBoxGrad = 7
 tCircleGrad = 8
 tEye = 9
 tLine = 10
 tSelect = 11
 tText = 12
 tBoxGradNS = 13
 tBoxFillX = 14
 tCircleFillX = 15
 tCustom = 16
 tBoxGradNESW = 17
 tBoxGradNWSE = 18
 tLineX2 = 19
 tLineX3 = 20
End Enum

Private Enum UDE_GRADDIR 'something quick to refer to when painting Gradient Boxes
 gNS = 0
 gEW = 1
 gNESW = 2
 gNWSE = 3
End Enum

'x1 and y1 are original mousedown position, xz1 and yz1 are used for the selection tool
Dim x1 As Integer, xz1 As Integer
Dim y1 As Integer, yz1 As Integer
Dim X2 As Integer 'mouseup postion
Dim Y2 As Integer 'mouseup postion
Dim mButton As Integer 'current button down
Dim SelTool As UDE_TOOLS 'current tool
Dim KeyDown As Integer 'current keydown
Dim mColor(4) As Long 'colors I originally planned on using the mouse wheel click so it's available
Dim mZoom As Integer 'current zoom (ie. IconWidth * mZoom)
Dim mDrawGrid As Boolean 'should we draw the grid?
'current icon being edited, Custom Tool Data,    Type of custom tool
Public mCurrIcon As Integer
Dim mCustomTool As String, tCustomType As Integer
'Are we moving the selection or drawing it?, Last Select position
Dim MovingSel As Boolean, OldSel(3) As Integer
'are we drawing anything?, stores wether an icon was edited.
Dim drawing As Boolean, FileChanged(99) As Boolean
Dim mOpacity As Integer 'opacity of brush
Public mLayer As Integer 'current layer
Dim mIcons() As Picture
Dim mLPics() As Picture

Private Sub AddRecent(ByVal s As String)
'adds 's' to the recent files menu, only 10 items are allowed in the menu
Dim X As Integer
 For X = 0 To mnuFileRecentFiles.Count - 1
   If s = mnuFileRecentFiles(X).Caption Then Exit Sub
 Next X

If mnuFileRecentFiles.Count >= 10 Then
 Dim arr(9) As String
 For X = mnuFileRecentFiles.Count - 1 To 0 Step -1
  arr(X) = mnuFileRecentFiles(X).Caption
 Next X
 For X = 0 To mnuFileRecentFiles.Count - 2
   mnuFileRecentFiles(X).Caption = arr(X + 1)
 Next X
  mnuFileRecentFiles(mnuFileRecentFiles.Count - 1).Caption = s
Else
 X = mnuFileRecentFiles.Count
 If X = 1 And mnuFileRecentFiles(0).Enabled = False Then
  mnuFileRecentFiles(0).Caption = s
  mnuFileRecentFiles(0).Visible = True
  mnuFileRecentFiles(0).Enabled = True
 Else
  Call Load(mnuFileRecentFiles(X))
  mnuFileRecentFiles(X).Caption = s
  mnuFileRecentFiles(X).Visible = True
  mnuFileRecentFiles(X).Enabled = True
 End If
End If
End Sub

Private Function BlendColor(ByVal Clr1 As Long, ByVal Clr2 As Long, ByVal Amount As Single) As Long
'used for pen opacity, not implemented
'mOpacity = 50
'IIf(picIcon(mCurrIcon).Point(Int(x1 / mZoom), Int(y1 / mZoom)) = picTrans.BackColor, mColor(mButton), BlendColor(picIcon(mCurrIcon).Point(Int(x1 / mZoom), Int(y1 / mZoom)), mColor(mButton), mOpacity))
If Amount = 100 Then BlendColor = Clr1: Exit Function
If Amount < 1 Then BlendColor = Clr2: Exit Function

Dim r(2) As Single, g(2) As Single, b(2) As Single

Amount = 100 / Amount

r(0) = GetRGB(Clr1).Red
r(1) = GetRGB(Clr2).Red
g(0) = GetRGB(Clr1).Green
g(1) = GetRGB(Clr2).Green
b(0) = GetRGB(Clr1).Blue
b(1) = GetRGB(Clr2).Blue

r(2) = (r(1) - r(0)) / Amount
g(2) = (g(1) - g(0)) / Amount
b(2) = (b(1) - b(0)) / Amount

r(0) = r(0) + r(2)
If r(0) < 0 Then r(0) = 0
If r(0) > 255 Then r(0) = 255

g(0) = g(0) + g(2)
If g(0) < 0 Then g(0) = 0
If g(0) > 255 Then g(0) = 255

b(0) = b(0) + b(2)
If b(0) < 0 Then b(0) = 0
If b(0) > 255 Then b(0) = 255

BlendColor = RGB(r(0), g(0), b(0))
If BlendColor = picTrans.BackColor Then BlendColor = BlendColor + 1
End Function

Private Sub DrawGrid()
'draws a grid on the Edit screen according to the currect zoom
If picIcon(mCurrIcon).ScaleWidth > 32 And mnuViewZoomSize(20).Checked = True Then Call mnuViewZoomSize_Click(10): Exit Sub
If picEdit.Height = picIcon(mCurrIcon).Width Or mDrawGrid = False Then Exit Sub
Dim X As Integer
Dim Y As Integer

For Y = 0 To picIcon(mCurrIcon).Height
 For X = 0 To picIcon(mCurrIcon).Width
  picEdit.Line (Int(X * mZoom), Int(Y * mZoom))-(Int(X * mZoom), picEdit.ScaleHeight), RGB(190, 190, 190)
  picEdit.Line (Int(X * mZoom), Int(Y * mZoom))-(picEdit.ScaleWidth, Int(Y * mZoom)), RGB(190, 190, 190)
 Next X
Next Y

  picEdit.Refresh
  picEdit.Line (picEdit.ScaleWidth - 1, 0)-(picEdit.ScaleWidth - 1, picEdit.ScaleHeight), RGB(190, 190, 190)
  picEdit.Line (0, picEdit.ScaleHeight - 1)-(picEdit.ScaleWidth - 1, picEdit.ScaleHeight - 1), RGB(190, 190, 190)
  'picedit.Line (Int(X * mzoom), Int(Y * mzoom))-(picedit.ScaleWidth, Int(Y * mzoom)), RGB(190, 190, 190)
End Sub

Private Sub DrawLayers(ByVal Icon As Integer, Optional icStop As Integer = -1)
Dim i As Integer, c As Long
Dim X As Integer, Y As Integer
Set picIcon(Icon).Picture = LoadPicture()
If icStop = -1 Then icStop = 4
For i = 0 To icStop
 If picLyr(i).Tag <> "" And picLyr(i).Tag <> "invisible" Then
  For Y = 0 To picLyr(i).ScaleHeight - 1
   For X = 0 To picLyr(i).ScaleWidth - 1
    c = picLyr(i).Point(X, Y)
    If i = 0 Then
     picIcon(mCurrIcon).PSet (X, Y), c
    Else
     If c <> picTrans.BackColor Then picIcon(mCurrIcon).PSet (X, Y), c
    End If
   Next X
  Next Y
 End If
Next i
End Sub

Private Sub SetAndDrawLayers()
  Call DrawLayers(mCurrIcon)
  Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
  picEdit.Cls
  Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
  Call DrawGrid
  picEdit.Refresh
  Set picEdit.Picture = picEdit.Image
End Sub

Private Sub MergeLayers(ByVal Icon As Integer, ByVal icStart As Integer, ByVal icStop As Integer, Optional mDown As Boolean = False)
Dim i As Integer, c As Long
Dim X As Integer, Y As Integer
Set picIcon(Icon).Picture = LoadPicture()
If mDown = False Then
For i = icStart To icStop
  For Y = 0 To picLyr(i).ScaleHeight - 1
   For X = 0 To picLyr(i).ScaleWidth - 1
    c = picLyr(i).Point(X, Y)
    If i = 0 Then
     picLyr(icStart).PSet (X, Y), c
    Else
     If c <> picTrans.BackColor Then picLyr(icStart).PSet (X, Y), c
    End If
   Next X
  Next Y
  If i <> icStart Then Set picLyr(i).Picture = LoadPicture()
Next i
Else
For i = icStart To icStop
  For Y = 0 To picLyr(i).ScaleHeight - 1
   For X = 0 To picLyr(i).ScaleWidth - 1
    c = picLyr(i).Point(X, Y)
    If c <> picTrans.BackColor Then picLyr(icStop).PSet (X, Y), c
   Next X
  Next Y
  If i <> icStop Then Set picLyr(i).Picture = LoadPicture()
Next i
End If
End Sub

Private Function ExtractFileName(ByVal File As String) As String
'returns a file's name sans Directory and Extension
Dim i As Integer
i = InStrRev(File, "\")
If i = 0 Then ExtractFileName = File: Exit Function
ExtractFileName = Mid(File, i + 1)
i = InStrRev(ExtractFileName, ".")
ExtractFileName = Left(ExtractFileName, i - 1)
End Function

Private Sub FillRegion(ByVal X As Integer, ByVal Y As Integer)
'for paint bucket tool
  Dim a As Integer, b As Integer
  a = picIcon(mCurrIcon).FillStyle
  b = picIcon(mCurrIcon).FillColor
  'picIcon(mCurrIcon).FillStyle = 0
  picLyr(mLayer).FillStyle = 0
  'picIcon(mCurrIcon).FillColor = mColor(mButton)
  picLyr(mLayer).FillColor = mColor(mButton)
'  Call ExtFloodFill(picIcon(mCurrIcon).hdc, Int(x / mZoom), Int(y / mZoom), picIcon(mCurrIcon).Point(Int(x / mZoom), Int(y / mZoom)), 1)
  Call ExtFloodFill(picLyr(mLayer).hdc, Int(X / mZoom), Int(Y / mZoom), picLyr(mLayer).Point(Int(X / mZoom), Int(Y / mZoom)), 1)
  'picIcon(mCurrIcon).FillStyle = a
  'picIcon(mCurrIcon).FillColor = b
  picLyr(mLayer).FillStyle = a
  picLyr(mLayer).FillColor = b
End Sub

Private Sub GradientClr(ByRef PicBox As PictureBox, ByVal c1 As Long, ByVal c2 As Long)
'paints a gradient to the referred picture box
'mainly used for displaying the color selections
Dim r(2) As Single, g(2) As Single, b(2) As Single
Dim i As Integer, ix As Integer

 r(0) = GetRGB(c1).Red
 g(0) = GetRGB(c1).Green
 b(0) = GetRGB(c1).Blue

 r(1) = GetRGB(c2).Red
 g(1) = GetRGB(c2).Green
 b(1) = GetRGB(c2).Blue

i = PicBox.ScaleWidth
If i > 255 Then i = 255

 r(2) = (r(1) - r(0)) / i
 g(2) = (g(1) - g(0)) / i
 b(2) = (b(1) - b(0)) / i

For ix = 0 To PicBox.ScaleWidth

 If r(0) < 0 Then r(0) = 0
 If r(0) > 255 Then r(0) = 255
 If g(0) < 0 Then g(0) = 0
 If g(0) > 255 Then g(0) = 255
 If b(0) < 0 Then b(0) = 0
 If b(0) > 255 Then b(0) = 255

 PicBox.Line (ix, 0)-(ix, PicBox.ScaleHeight), RGB(r(0), g(0), b(0)), BF
 r(0) = r(0) + r(2)
 g(0) = g(0) + g(2)
 b(0) = b(0) + b(2)
 
Next ix
End Sub

Public Sub IconFlip(Picture1 As PictureBox)
   'flip vertical
   Dim pX As Integer, pY As Integer
   pX = Picture1.ScaleWidth
   pY = Picture1.ScaleHeight
   Call StretchBlt(Picture1.hdc, 0, pY, pX, -pY - 1, Picture1.hdc, 0, 0, pX, pY, vbSrcCopy)
End Sub

Public Sub IconMirror(Picture1 As PictureBox)
   'flip horizontal
   Dim pX As Integer, pY As Integer
   pX = Picture1.ScaleWidth - 1
   pY = Picture1.ScaleHeight - 1
   Set picTemp.Picture = picTemp.Image
   Call StretchBlt(Picture1.hdc, pX, 0, -pX, pY, Picture1.hdc, 0, 0, pX, pY, vbSrcCopy)
End Sub

Private Sub ImageRotate(picSource As PictureBox, picDestination As PictureBox, sngRotateAngle As Single, blnClockWise As Boolean)
'Rotates image picSource sngRotateAngle degree and save the result in picDestination
'borrowed from Nubee Paint

  Const conPi = 3.14159265358979
  
  Dim a As Single                                            'angle of R and dXd
  Dim intMaxXY As Single 'maximum width or height of picDestination
  Dim dXs As Long         'relative coordinate where the pixel color information
  Dim dYs As Long         '                     will be retrieved from picSource
  Dim dXd As Long         'relative coordinate where the pixel color information
  Dim dYd As Long         '                    will be written to picDestination
  Dim lngAdjustX As Long                 'to adjust the new pixel coordinates to
  Dim lngAdjustY As Long                 ' make sure the whole part of the image
                                         '  is shown (currently only for 90° and
                                         '                        270° rotation)
  Dim lngColor(3) As Long                              'pixel colors information
  Dim r As Integer                               'length of line (0,0)-(dXd,dYd)
  Dim Xs As Integer           'base coordinate where the pixel color information
  Dim Ys As Integer           '                 will be retrieved from picSource
  Dim Xd As Integer           'base coordinate where the pixel color information
  Dim Yd As Integer           '                will be written to picDestination
                              
  On Error GoTo ErrorHandler
  
  If blnClockWise Then
    sngRotateAngle = 360 - sngRotateAngle
  End If
  Xs = picSource.ScaleWidth / 2
  Ys = picSource.ScaleHeight / 2
  Xd = picDestination.ScaleWidth / 2
  Yd = picDestination.ScaleHeight / 2
  intMaxXY = IIf(picDestination.ScaleWidth > picDestination.ScaleHeight, _
                    picDestination.ScaleWidth / 2, _
                    picDestination.ScaleHeight / 2)
  If (sngRotateAngle = 90) Or (sngRotateAngle = 270) Then
    lngAdjustX = ((picDestination.ScaleHeight - _
                   picDestination.ScaleWidth) / 2) - 2
    lngAdjustY = ((picDestination.ScaleWidth - _
                   picDestination.ScaleHeight) / 2)
  Else
    lngAdjustX = 0
    lngAdjustY = 0
  End If
  sngRotateAngle = sngRotateAngle * (conPi / 180)             'convert to radian
  'Write each pixels to picDestination with transformed coordinates
  '  to make rotation effect
  picDestination.DrawMode = vbCopyPen
  For dXd = 0 To intMaxXY
    DoEvents
    For dYd = 0 To intMaxXY
      DoEvents
      If dXd = 0 Then
        a = conPi / 2
      Else
        a = Atn(dYd / dXd)
      End If
      r = Sqr((dXd * dXd) + (dYd * dYd))
      dXs = r * Cos(a + sngRotateAngle)
      dYs = r * Sin(a + sngRotateAngle)
      'Get pixel colors information from picSource
      lngColor(0) = picSource.Point(Xs + dXs, Ys + dYs)
      lngColor(1) = picSource.Point(Xs - dXs, Ys - dYs)
      lngColor(2) = picSource.Point(Xs + dYs, Ys - dXs)
      lngColor(3) = picSource.Point(Xs - dYs, Ys + dXs)
      'Set pixel colors information to picDestination
      If lngColor(0) <> -1 Then
        picDestination.PSet (Xd + dXd + lngAdjustX, Yd + dYd + lngAdjustY), lngColor(0)
      End If
      If lngColor(1) <> -1 Then
        picDestination.PSet (Xd - dXd + lngAdjustX, Yd - dYd + lngAdjustY), lngColor(1)
      End If
      If lngColor(2) <> -1 Then
        picDestination.PSet (Xd + dYd + lngAdjustX, Yd - dXd + lngAdjustY), lngColor(2)
      End If
      If lngColor(3) <> -1 Then
        picDestination.PSet (Xd - dYd + lngAdjustX, Yd + dXd + lngAdjustY), lngColor(3)
      End If
    Next
    picDestination.Refresh
  Next
  picDestination.Refresh
  Exit Sub

ErrorHandler:
End Sub

Public Sub LoadMultiIcons(ByVal lst As Collection)
'used by frmMIcons to load multiple icons
Dim i As Integer
For i = 1 To lst.Count
  Call NewIcon(Temp_Icon.ifdIconData(lst(i)).idWidth)
  'Else Call NewIcon(32)
    
 'Set picUndo(mCurrIcon).Picture = Temp_Icon.ifdIcon(lst(i))  'LoadPicture(cd.FileName)
 Set picIcon(mCurrIcon).Picture = Temp_Icon.ifdIcon(lst(i))  'LoadPicture(cd.FileName)
 Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
 picLyr(0).Width = picIcon(mCurrIcon).Width
 picLyr(0).Height = picIcon(mCurrIcon).Height
 Set picLyr(0).Picture = picIcon(mCurrIcon).Picture
   
 picIcon(mCurrIcon).Refresh
 picIcon(mCurrIcon).Tag = cd.Filename & i
 txtIcon(mCurrIcon).Caption = Truncate(LCase(ExtractFileName(cd.Filename)) & i) & vbCrLf & picIcon(mCurrIcon).Width & "x" & picIcon(mCurrIcon).Height
   
  Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
  Call DrawGrid
    
 picEdit.Refresh
 Set picEdit.Picture = picEdit.Image
 ReDim Preserve mIcons(mCurrIcon)
 Set mIcons(mCurrIcon) = Temp_Icon.ifdIcon(lst(i))
Next i

End Sub

Private Sub LoadPrefs()
'load preferences
Dim s As String
s = GetFromINI("Main", "Grid", App.Path & "\prefs.ini")
If s = "" Or s = "1" Then mDrawGrid = True Else mDrawGrid = False
mnuViewGrid.Checked = mDrawGrid

s = GetFromINI("Main", "Backcolor", App.Path & "\prefs.ini")
If s <> "" And IsNumeric(s) = True Then picBack.BackColor = CLng(s)

s = GetFromINI("Main", "Layers", App.Path & "\prefs.ini")
If s = "" Or s = "1" Then mnuViewLayers.Checked = True Else mnuViewLayers.Checked = False
picLayers.Visible = mnuViewLayers.Checked

s = GetFromINI("Main", "Brushes", App.Path & "\prefs.ini")
If s = "" Or s = "1" Then mnuViewBrushes.Checked = True Else mnuViewBrushes.Checked = False
picBrushes.Visible = mnuViewBrushes.Checked

Call picBack_Resize

s = GetFromINI("Main", "Zoom", App.Path & "\prefs.ini")
If s = "" Or IsNumeric(s) = False Then
 Call mnuViewZoomSize_Click(10)
Else
 Call mnuViewZoomSize_Click(CInt(s))
End If
s = GetFromINI("Main", "RecentCount", App.Path & "\prefs.ini")
If IsNumeric(s) = True Then
 Dim X As Integer, Y As Integer
 Y = CInt(s)
 If Y <> 0 Then
  s = GetFromINI("Main", "Recent" & 0, App.Path & "\prefs.ini")
  If s <> "" Then
   mnuFileRecentFiles(0).Caption = s
   mnuFileRecentFiles(0).Enabled = True
  End If
  For X = 1 To Y - 1
   s = GetFromINI("Main", "Recent" & X, App.Path & "\prefs.ini")
   If s <> "" Then
    Call Load(mnuFileRecentFiles(X))
    mnuFileRecentFiles(X).Caption = s
    mnuFileRecentFiles(X).Enabled = True
   End If
  Next X
 End If
End If

s = GetFromINI("Main", "Window", App.Path & "\prefs.ini")
If s = "" Or IsNumeric(s) = False Then
 Me.WindowState = vbNormal
Else
 Me.WindowState = CInt(s)
End If

If FileExist(App.Path & "\lan.lpk") = True Then Call LoadLanguage(App.Path & "\lan.lpk")

End Sub

Private Sub NewIcon(Optional ByVal Size As Integer = 32)
'completes the steps to create a new fresh icon
Dim i As Integer, X As Integer, p As Integer
p = -1

For i = 0 To picIcon.Count - 1
 If picIcon(i).Tag = "" Then
  If p <> -1 Then p = i
 Else
  X = X + ((picIcon(i).Height + txtIcon(i).Height) + 8)
 End If
Next i
If p <> -1 Then GoTo 1
p = picIcon.Count
Call Load(picIcon(p))
Call Load(txtIcon(p))

1
With picIcon(p)
 .Tag = "untitled"
 Set .Picture = LoadPicture()
 .Top = X
 .Left = 6 + ((48 - Size) / 2)
 .Height = Size
 .Width = Size
 .Visible = True
End With

txtIcon(p).Top = X + picIcon(p).Height + 4
txtIcon(p).Visible = True
txtIcon(p).Caption = "untitled" & vbCrLf & picIcon(p).Width & "x" & picIcon(p).Height

picEdit.Height = Size * mZoom
picEdit.Width = Size * mZoom

Set picEdit.Picture = LoadPicture()
picEdit.Visible = True

Dim b As Integer
 b = mCurrIcon
 For i = 0 To picLyr.Count - 1
  Set picLyr(i).Picture = picLyr(i).Image
  Set mLPics(b, i) = picLyr(i).Picture
 Next i
 
 For i = 0 To picLyr.Count - 1
  Set picLyr(i).Picture = LoadPicture()
 Next i


mCurrIcon = p

Call DrawGrid
picEdit.Refresh
Set picEdit.Picture = picEdit.Image
picIconsMove.Height = X + (((picIcon(p).Height + txtIcon(p).Height) + 8) * 2)
End Sub

Private Sub NewFilter(ByVal f As String)
'really just creates a menu item for the filter 'f'
Dim i As Integer
For i = 0 To mnuFiltersRun.Count - 1
 If mnuFiltersRun(i).Tag = "" Then GoTo 1
Next i

Call Load(mnuFiltersRun(i))

1
 mnuFiltersRun(i).Tag = f
 mnuFiltersRun(i).Caption = f
 mnuFiltersRun(i).Visible = True
End Sub

Private Sub NewPlugin(ByVal f As String)
'really just creates a menu item for the plugin 'f'
Dim i As Integer
For i = 0 To mnuPluginsRun.Count - 1
 If mnuPluginsRun(i).Tag = "" Then GoTo 1
Next i

Call Load(mnuPluginsRun(i))

1
 mnuPluginsRun(i).Tag = f
 mnuPluginsRun(i).Caption = f
 mnuPluginsRun(i).Visible = True
End Sub

Private Sub SavePrefs()
'saves preferences
Dim s As String
Call WriteToINI("Main", "Grid", IIf(mDrawGrid = True, 1, 0), App.Path & "\prefs.ini")
Call WriteToINI("Main", "Layers", IIf(mnuViewLayers.Checked = True, 1, 0), App.Path & "\prefs.ini")
Call WriteToINI("Main", "Brushes", IIf(mnuViewBrushes.Checked = True, 1, 0), App.Path & "\prefs.ini")
Call WriteToINI("Main", "Zoom", CStr(mZoom), App.Path & "\prefs.ini")
Call WriteToINI("Main", "Backcolor", CStr(picBack.BackColor), App.Path & "\prefs.ini")
Call WriteToINI("Main", "Window", CStr(Me.WindowState), App.Path & "\prefs.ini")
Call WriteToINI("Main", "RecentCount", CStr(mnuFileRecentFiles.Count), App.Path & "\prefs.ini")

Dim X As Integer
 For X = 0 To mnuFileRecentFiles.Count - 1
  If mnuFileRecentFiles(X).Caption = "(empty)" Then Exit For
  Call WriteToINI("Main", "Recent" & X, mnuFileRecentFiles(X).Caption, App.Path & "\prefs.ini")
 Next X
End Sub

Public Sub SetFont(ByVal X As Integer, ByVal Y As Integer, ByVal FontName As String, ByVal FontSize As Integer, ByVal Text As String)
'accessed by frmText to draw text to the current icon
If FontName = "" Then Exit Sub
With picLyr(mLayer)
 .CurrentX = X - 1
 .CurrentY = Y - 4
 .ForeColor = mColor(1)
 .Font = FontName
 .FontSize = FontSize
End With
  Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
  Set picLyr(mLayer).Picture = picLyr(mLayer).Image: Set picLyr(mLayer + 5).Picture = picLyr(mLayer).Picture
  Dim arr() As String, v As Variant
  arr() = Split(Text, vbCrLf)
  For Each v In arr()
   picLyr(mLayer).Print v
   picLyr(mLayer).CurrentY = picIcon(mCurrIcon).CurrentY + (FontSize / 5) - 5
   picLyr(mLayer).CurrentX = X - 1
  Next v

Call SetAndDrawLayers
End Sub

Private Sub ShowRGBVal(ByVal Index As Integer)
'displays RGB and Hex Values to the color textboxes
Dim l As Long
l = mColor(Index)
txtClr(0).Text = GetRGB(l).Red
txtClr(1).Text = GetRGB(l).Green
txtClr(2).Text = GetRGB(l).Blue
txtClr(3).Text = IIf(Len(Hex(txtClr(0).Text)) = 1, "0", "") & Hex(txtClr(0).Text) & IIf(Len(Hex(txtClr(1).Text)) = 1, "0", "") & Hex(txtClr(1).Text) & IIf(Len(Hex(txtClr(2).Text)) = 1, "0", "") & Hex(txtClr(2).Text)
End Sub

Private Sub Tool_Box_Move(ByVal X As Integer, ByVal Y As Integer, Optional OptLine As Integer, Optional HideBox As Boolean)
'displays a box outline when drawing a box or gradient box
'OptLine is the center line's orientation
On Error Resume Next
If x1 = -1 Then Exit Sub
picEdit.Cls

X2 = Int(X / mZoom) * mZoom + mZoom - 1
Y2 = Int(Y / mZoom) * mZoom + mZoom - 1
picEdit.DrawMode = vbInvert
If HideBox = False Then
 If KeyDown = 17 Then
  Y2 = y1 + (X2 - x1)
  picEdit.Line (x1, y1)-(X2, Y2), , B
 Else
  picEdit.Line (x1, y1)-(X2, Y2), , B
 End If
End If
Select Case OptLine
 Case 1
  picEdit.Line (x1, y1 + ((Y2 - y1) / 2))-(X2, y1 + ((Y2 - y1) / 2)), 0
 Case 2
  picEdit.Line (x1 + ((X2 - x1) / 2), y1)-(x1 + ((X2 - x1) / 2), Y2), 0
 Case 3
  picEdit.Line (X2, y1)-(x1, Y2), 0
 Case 4
  picEdit.Line (x1, y1)-(X2, Y2), 0
End Select
picEdit.DrawMode = vbCopyPen
End Sub

Private Sub Tool_Box_Up(ByVal X As Integer, ByVal Y As Integer, Optional ByVal Fill As Boolean = False, Optional OutLine As Boolean)
x1 = Int(x1 / mZoom)
X2 = Int(X2 / mZoom)
y1 = Int(y1 / mZoom)
Y2 = Int(Y2 / mZoom)

If Fill = True Then
 picIcon(mCurrIcon).Line (x1, y1)-(X2, Y2), mColor(mButton), BF
 picLyr(mLayer).Line (x1, y1)-(X2, Y2), mColor(mButton), BF
Else
 picIcon(mCurrIcon).Line (x1, y1)-(X2, Y2), mColor(mButton), B
 picLyr(mLayer).Line (x1, y1)-(X2, Y2), mColor(mButton), B
End If

Dim b As Integer
b = mButton
If b = 1 Then b = 2 Else b = 1
If OutLine = True Then
 picIcon(mCurrIcon).Line (x1, y1)-(X2, Y2), mColor(b), B
 picLyr(mLayer).Line (x1, y1)-(X2, Y2), mColor(b), B
End If
End Sub

Private Sub Tool_BoxGrad_Up(ByVal X As Single, ByVal Y As Single, Optional Direction As UDE_GRADDIR)
On Error Resume Next
Dim r(2) As Integer
Dim g(2) As Integer
Dim b(2) As Integer

Dim c1 As Long, c2 As Long, ix As Long
If mButton = 1 Then
 c1 = mColor(1)
 c2 = mColor(2)
Else
 c2 = mColor(1)
 c1 = mColor(2)
End If

 r(0) = GetRGB(c1).Red
 g(0) = GetRGB(c1).Green
 b(0) = GetRGB(c1).Blue

 r(1) = GetRGB(c2).Red
 g(1) = GetRGB(c2).Green
 b(1) = GetRGB(c2).Blue
 Dim hOff As Integer

Select Case Direction
 Case gEW

    r(2) = (r(1) - r(0)) / Int((X / mZoom) - Int(x1 / mZoom) - 1)
    g(2) = (g(1) - g(0)) / Int((X / mZoom) - Int(x1 / mZoom) - 1)
    b(2) = (b(1) - b(0)) / Int((X / mZoom) - Int(x1 / mZoom) - 1)

    For ix = Int(x1 / mZoom) To Int(X2 / mZoom)
    
     If r(0) < 0 Then r(0) = 0
     If r(0) > 255 Then r(0) = 255
     If g(0) < 0 Then g(0) = 0
     If g(0) > 255 Then g(0) = 255
     If b(0) < 0 Then b(0) = 0
     If b(0) > 255 Then b(0) = 255
    
       picIcon(mCurrIcon).Line (ix, Int(y1 / mZoom))-(ix, Int(Y2 / mZoom)), RGB(r(0), g(0), b(0)), BF
       picLyr(mLayer).Line (ix, Int(y1 / mZoom))-(ix, Int(Y2 / mZoom)), RGB(r(0), g(0), b(0)), BF
     
     r(0) = r(0) + r(2)
     g(0) = g(0) + g(2)
     b(0) = b(0) + b(2)
    
    Next ix
 Case gNS
    r(2) = (r(1) - r(0)) / Int((Y / mZoom) - Int(y1 / mZoom) - 1)
    g(2) = (g(1) - g(0)) / Int((Y / mZoom) - Int(y1 / mZoom) - 1)
    b(2) = (b(1) - b(0)) / Int((Y / mZoom) - Int(y1 / mZoom) - 1)
 
    For ix = Int(y1 / mZoom) To Int(Y2 / mZoom)
    
     If r(0) < 0 Then r(0) = 0
     If r(0) > 255 Then r(0) = 255
     If g(0) < 0 Then g(0) = 0
     If g(0) > 255 Then g(0) = 255
     If b(0) < 0 Then b(0) = 0
     If b(0) > 255 Then b(0) = 255
    
       picIcon(mCurrIcon).Line (Int(x1 / mZoom), ix)-(Int(X2 / mZoom), ix), RGB(r(0), g(0), b(0)), BF
       picLyr(mLayer).Line (Int(x1 / mZoom), ix)-(Int(X2 / mZoom), ix), RGB(r(0), g(0), b(0)), BF

     r(0) = r(0) + r(2)
     g(0) = g(0) + g(2)
     b(0) = b(0) + b(2)
    
    Next ix
 Case gNESW

   Set picTemp.Picture = LoadPicture()
    picTemp.Height = Int(Y2 / mZoom) - Int(y1 / mZoom)
    hOff = picTemp.Height / 2
    picTemp.Width = Int(X2 / mZoom) - Int(x1 / mZoom)
'    Debug.Print (picTemp.Width + hOff)
    r(2) = (r(1) - r(0)) / ((picTemp.Width + (hOff * 2))) ' * IIf(mButton = 1, 1, 1.3))
    g(2) = (g(1) - g(0)) / ((picTemp.Width + (hOff * 2))) ' * IIf(mButton = 1, 1, 1.3))
    b(2) = (b(1) - b(0)) / ((picTemp.Width + (hOff * 2))) ' * IIf(mButton = 1, 1, 1.3))
    
    For ix = 0 - hOff To picTemp.Width + hOff
    
     If r(0) < 0 Then r(0) = 0
     If r(0) > 255 Then r(0) = 255
     If g(0) < 0 Then g(0) = 0
     If g(0) > 255 Then g(0) = 255
     If b(0) < 0 Then b(0) = 0
     If b(0) > 255 Then b(0) = 255
    
       picTemp.Line (ix + hOff, 0)-(ix - hOff, picTemp.Height), RGB(r(0), g(0), b(0))
    
     r(0) = r(0) + r(2)
     g(0) = g(0) + g(2)
     b(0) = b(0) + b(2)
    
    Next ix
    Set picTemp.Picture = picTemp.Image
    picIcon(mCurrIcon).PaintPicture picTemp.Picture, Int(x1 / mZoom), Int(y1 / mZoom), Int(X2 / mZoom) - Int(x1 / mZoom) + 1, Int(Y2 / mZoom) - Int(y1 / mZoom) + 1, 0, 0, picTemp.Width, picTemp.Height, vbSrcCopy
    picLyr(mLayer).PaintPicture picTemp.Picture, Int(x1 / mZoom), Int(y1 / mZoom), Int(X2 / mZoom) - Int(x1 / mZoom) + 1, Int(Y2 / mZoom) - Int(y1 / mZoom) + 1, 0, 0, picTemp.Width, picTemp.Height, vbSrcCopy
 Case gNWSE

   Set picTemp.Picture = LoadPicture()
    picTemp.Height = Int(Y2 / mZoom) - Int(y1 / mZoom)
    hOff = picTemp.Height / 2
    picTemp.Width = Int(X2 / mZoom) - Int(x1 / mZoom)

    r(2) = (r(1) - r(0)) / ((picTemp.Width + (hOff * 2)))
    g(2) = (g(1) - g(0)) / ((picTemp.Width + (hOff * 2)))
    b(2) = (b(1) - b(0)) / ((picTemp.Width + (hOff * 2)))

    For ix = picTemp.Width + hOff To 0 - hOff Step -1
    
     If r(0) < 0 Then r(0) = 0
     If r(0) > 255 Then r(0) = 255
     If g(0) < 0 Then g(0) = 0
     If g(0) > 255 Then g(0) = 255
     If b(0) < 0 Then b(0) = 0
     If b(0) > 255 Then b(0) = 255
    
       picTemp.Line (ix - hOff, 0)-(ix + hOff, picTemp.Height), RGB(r(0), g(0), b(0))
    
     r(0) = r(0) + r(2)
     g(0) = g(0) + g(2)
     b(0) = b(0) + b(2)
    
    Next ix
    Set picTemp.Picture = picTemp.Image
    picIcon(mCurrIcon).PaintPicture picTemp.Picture, Int(x1 / mZoom), Int(y1 / mZoom), Int(X2 / mZoom) - Int(x1 / mZoom) + 1, Int(Y2 / mZoom) - Int(y1 / mZoom) + 1, 0, 0, picTemp.Width, picTemp.Height, vbSrcCopy
    picLyr(mLayer).PaintPicture picTemp.Picture, Int(x1 / mZoom), Int(y1 / mZoom), Int(X2 / mZoom) - Int(x1 / mZoom) + 1, Int(Y2 / mZoom) - Int(y1 / mZoom) + 1, 0, 0, picTemp.Width, picTemp.Height, vbSrcCopy
End Select
End Sub

Private Sub Tool_CircleGrad_Up(ByVal X As Single, ByVal Y As Single)
On Error Resume Next
Dim r(2) As Integer
Dim g(2) As Integer
Dim b(2) As Integer
x1 = Int(x1 / mZoom)
X2 = Int(X2 / mZoom)
y1 = Int(y1 / mZoom)
Y2 = Int(Y2 / mZoom)
Dim ra As Single, a As Single
Dim c1 As Long, c2 As Long, ix As Single, xc As Single, yc As Single
If Abs(X2 - x1) > Abs(Y2 - y1) Then ra = Abs(X2 - x1) / 2 Else ra = Abs(Y2 - y1) / 2
If KeyDown = 17 Then a = 1 Else a = Abs(Abs(y1 - Y2) / Abs(X2 - x1))

xc = x1 + (X2 - x1) / 2
yc = y1 + (Y2 - y1) / 2

'Debug.Print KeyDown

If mButton = 1 Then
 c1 = mColor(1)
 c2 = mColor(2)
Else
 c2 = mColor(1)
 c1 = mColor(2)
End If

 r(0) = GetRGB(c1).Red
 g(0) = GetRGB(c1).Green
 b(0) = GetRGB(c1).Blue

 r(1) = GetRGB(c2).Red
 g(1) = GetRGB(c2).Green
 b(1) = GetRGB(c2).Blue

 r(2) = (r(1) - r(0)) / Int(ra - 1)
 g(2) = (g(1) - g(0)) / Int(ra - 1)
 b(2) = (b(1) - b(0)) / Int(ra - 1)

Dim H As Long, k As Long
H = picIcon(mCurrIcon).FillStyle
k = picIcon(mCurrIcon).FillColor
picIcon(mCurrIcon).FillStyle = 0
picLyr(mLayer).FillStyle = 0
For ix = ra To 0 Step -1

 If r(0) < 0 Then r(0) = 0
 If r(0) > 255 Then r(0) = 255
 If g(0) < 0 Then g(0) = 0
 If g(0) > 255 Then g(0) = 255
 If b(0) < 0 Then b(0) = 0
 If b(0) > 255 Then b(0) = 255
   picIcon(mCurrIcon).FillColor = RGB(r(0), g(0), b(0))
   picLyr(mLayer).FillColor = RGB(r(0), g(0), b(0))
   picIcon(mCurrIcon).Circle (xc, yc), ix, RGB(r(0), g(0), b(0)), , , a
   picLyr(mLayer).Circle (xc, yc), ix, RGB(r(0), g(0), b(0)), , , a
 
 r(0) = r(0) + r(2)
 g(0) = g(0) + g(2)
 b(0) = b(0) + b(2)

Next ix
picIcon(mCurrIcon).FillColor = k
picLyr(mLayer).FillColor = k
picIcon(mCurrIcon).FillStyle = H
picLyr(mLayer).FillStyle = H
End Sub

Private Sub Tool_Circle_Move(ByVal X As Integer, ByVal Y As Integer)
On Error Resume Next
If x1 = -1 Then Exit Sub
picEdit.Cls
X2 = Int(X / mZoom) * mZoom + mZoom - 1
Y2 = Int(Y / mZoom) * mZoom + mZoom - 1


picEdit.DrawMode = vbInvert
Dim r As Integer, a As Single, xc As Single, yc As Single
If Abs(X2 - x1) > Abs(Y2 - y1) Then
 If Abs(X2 - x1) <> 0 Then r = Abs(X2 - x1) / 2 Else Exit Sub
Else
 If Abs(Y2 - y1) <> 0 Then r = Abs(Y2 - y1) / 2 Else Exit Sub
End If

If KeyDown = 17 Then
 a = 1
Else
 If Abs(y1 - Y2) <> 0 And Abs(X2 - x1) <> 0 Then a = Abs(Abs(y1 - Y2) / Abs(X2 - x1)): picEdit.Line (x1, y1)-(X2, Y2), , B Else Exit Sub
End If

xc = x1 + (X2 - x1) / 2
yc = y1 + (Y2 - y1) / 2
picEdit.Circle (xc, yc), r, , , , a
picEdit.DrawMode = vbCopyPen

End Sub

Private Sub Tool_Circle_Up(ByVal X As Integer, ByVal Y As Integer, Optional ByVal Fill As Boolean = False, Optional OutLine As Boolean)
x1 = Int(x1 / mZoom)
X2 = Int(X2 / mZoom)
y1 = Int(y1 / mZoom)
Y2 = Int(Y2 / mZoom)
Dim r As Single, a As Single
Dim xc As Single, yc As Single

If Abs(X2 - x1) > Abs(Y2 - y1) Then
 If Abs(X2 - x1) <> 0 Then r = Abs(X2 - x1) / 2 Else Exit Sub
Else
 If Abs(Y2 - y1) <> 0 Then r = Abs(Y2 - y1) / 2 Else Exit Sub
End If
If Abs(y1 - Y2) = Abs(X2 - x1) Then
 a = 1
Else
 If KeyDown = 17 Then
  a = 1
 Else
  If Abs(y1 - Y2) <> 0 And Abs(X2 - x1) <> 0 Then a = Abs(Abs(y1 - Y2) / Abs(X2 - x1)) Else Exit Sub
 End If
End If

xc = x1 + (X2 - x1) / 2
yc = y1 + (Y2 - y1) / 2

'Debug.Print KeyDown

If Fill = True Then

Dim H As Long, k As Long
H = picIcon(mCurrIcon).FillStyle
k = picIcon(mCurrIcon).FillColor

picIcon(mCurrIcon).FillStyle = 0
picLyr(mLayer).FillStyle = 0
picIcon(mCurrIcon).FillColor = mColor(mButton)
picLyr(mLayer).FillColor = mColor(mButton)
 picIcon(mCurrIcon).Circle (xc, yc), r, mColor(mButton), , , a
 picLyr(mLayer).Circle (xc, yc), r, mColor(mButton), , , a
picIcon(mCurrIcon).FillColor = k
picLyr(mLayer).FillColor = k
picIcon(mCurrIcon).FillStyle = H
picLyr(mLayer).FillStyle = H

Dim b As Integer
b = mButton
If b = 1 Then b = 2 Else b = 1
If OutLine = True Then
 picIcon(mCurrIcon).Circle (xc, yc), r, mColor(b), , , a
 picLyr(mLayer).Circle (xc, yc), r, mColor(b), , , a
End If


' Call Fill
Else
 picIcon(mCurrIcon).Circle (xc, yc), r, mColor(mButton), , , a
 picLyr(mLayer).Circle (xc, yc), r, mColor(mButton), , , a
End If
End Sub

Private Sub Tool_Line_Move(ByVal X As Integer, ByVal Y As Integer)
On Error Resume Next

Dim x3 As Integer, y3 As Integer
picEdit.Cls
X2 = Int(X / mZoom) * mZoom + mZoom - 1
Y2 = Int(Y / mZoom) * mZoom + mZoom - 1
picEdit.DrawMode = vbInvert
picEdit.Line (x1 + (mZoom / 2), y1 + (mZoom / 2))-(X2 - (mZoom / 2), Y2 - (mZoom / 2)), 0
picEdit.DrawMode = vbCopyPen
End Sub

Private Function Truncate(ByVal s As String) As String
If Len(s) > 8 Then Truncate = Left(s, 7) & "..." Else Truncate = s
End Function

Private Sub cbBox_Click(ByVal Index As Integer)
Call tbTools_ButtonClick(tbTools.Buttons(7))

shpToolSel(3).Visible = True
shpToolSel(3).Top = 20
shpToolSel(3).Left = 20

 SelTool = tCustom
 tCustomType = 1
 mCustomTool = cbBox.GetBrushData(Index)
End Sub

Private Sub cbPens_Click(ByVal Index As Integer)
Call tbTools_ButtonClick(tbTools.Buttons(3))

shpToolSel(1).Visible = True
shpToolSel(1).Top = 20
shpToolSel(1).Left = 0

 SelTool = tCustom
 tCustomType = 2
 mCustomTool = cbPens.GetBrushData(Index)
End Sub

Private Sub cp_Pick(ByVal Button As Integer, ByVal Clr As Long)
If Clr = -1 Then Exit Sub
  picClr(Button).BackColor = Clr
  mColor(Button) = Clr
 
  Call GradientClr(picClrSel(1), vbBlack, Clr)
  Call GradientClr(picClrSel(2), Clr, vbWhite)
  Call ShowRGBVal(Button)
End Sub

Private Sub Form_Click()
mOpacity = 0
End Sub

Private Sub Form_Load()
Call LoadPrefs
txtIcon(0).Caption = "untitled" & vbCrLf & "32x32"
  
Call GradientClr(picClrSel(1), vbBlack, vbWhite)
Call GradientClr(picClrSel(2), vbWhite, vbBlack)
  
cbPens.BrushType = btPen
cbPens.Path = App.Path & "\Brushes\"
cbBox.BrushType = btBox
cbBox.Path = App.Path & "\Brushes\"

ReDim mLPics(31, 9)
Dim i As Integer, b As Integer
For i = 0 To 31
 For b = 0 To 9
  Set mLPics(i, b) = LoadPicture()
 Next b
Next i

Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
Call DrawGrid
SelTool = tDraw
mColor(1) = 0
mColor(2) = vbWhite
picEdit.Refresh
Set picEdit.Picture = picEdit.Image
'draw colors
Dim X As Integer

cp.ClearColors

cp.AddColor vbBlack
cp.AddColor RGB(128, 128, 128)
cp.AddColor RGB(128, 0, 0)
cp.AddColor RGB(128, 128, 0)
cp.AddColor RGB(0, 128, 0)
cp.AddColor RGB(0, 128, 128) + 1
cp.AddColor RGB(0, 0, 128)
cp.AddColor RGB(128, 0, 128)
cp.AddColor RGB(128, 128, 64)
cp.AddColor RGB(0, 64, 64)
cp.AddColor RGB(0, 128, 255)
cp.AddColor RGB(0, 64, 128)
cp.AddColor RGB(64, 0, 255)
cp.AddColor RGB(128, 64, 0)


cp.AddColor vbWhite
cp.AddColor RGB(192, 192, 192)
cp.AddColor RGB(255, 0, 0)
cp.AddColor RGB(255, 255, 0)
cp.AddColor RGB(0, 255, 0)
cp.AddColor RGB(0, 255, 255)
cp.AddColor RGB(0, 0, 255)
cp.AddColor RGB(255, 0, 255)
cp.AddColor RGB(255, 255, 128)
cp.AddColor RGB(0, 255, 128)
cp.AddColor RGB(128, 255, 255)
cp.AddColor RGB(128, 128, 255)
cp.AddColor RGB(255, 0, 128)
cp.AddColor RGB(255, 128, 64)


cp.ShowColors

On Error Resume Next
Call MkDir(App.Path & "\Plugins\")
Call MkDir(App.Path & "\Filters\")

flbPlugin.Path = App.Path & "\Plugins\"
flbPlugin.Refresh

For X = 0 To flbPlugin.ListCount - 1
 Call NewPlugin(Left(flbPlugin.List(X), Len(flbPlugin.List(X)) - 4))
Next X
 
If X = 0 Then mnuPluginsRun(0).Caption = "(empty)": mnuPluginsRun(0).Enabled = False


flbPlugin.Path = App.Path & "\Filters\"
flbPlugin.Pattern = "*.ccf"
flbPlugin.Refresh

For X = 0 To flbPlugin.ListCount - 1
 Call NewFilter(Left(flbPlugin.List(X), Len(flbPlugin.List(X)) - 4))
Next X
 
If X = 0 Then mnuFiltersRun(0).Caption = "(empty)": mnuFiltersRun(0).Enabled = False


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call SavePrefs
Dim i As Integer, a As Integer
For i = 0 To picIcon.Count - 1
 If picIcon(i).Tag <> "" And FileChanged(i) = True Then
  a = MsgBox(ExtractFileName(picIcon(i).Tag) & " has changed." & vbCrLf & "Do you want to save it?", vbQuestion + vbYesNoCancel, "Save Changes")
  If a = vbYes Then 'save
   If InStr(picIcon(i).Tag, "ico") Then
    Open picIcon(i).Tag For Binary Access Write As #1
     Put #1, 1, GenerateIconForSave$(picIcon(i))
    Close #1
   Else
    Call mnuFileSaveAs_Click
   End If

  ElseIf a = vbCancel Then
   Cancel = -1
   Exit Sub
  End If
 End If
Next i
If FileExist(App.Path & "\temp.bmp") Then Kill App.Path & "\temp.bmp"
If FileExist(App.Path & "\temp.ico") Then Kill App.Path & "\temp.ico"
End
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then Exit Sub
If Me.ScaleWidth < 680 Then Me.Width = 680 * Screen.TwipsPerPixelX
If Me.ScaleHeight < 500 Then Me.Height = 500 * Screen.TwipsPerPixelY
On Error Resume Next
Dim W As Integer, H As Integer
W = Me.ScaleWidth
H = Me.ScaleHeight

picBack.Width = W - picBack.Left
picBack.Height = H - Frame1.Height + 4

Frame1.Top = H - Frame1.Height
Frame1.Width = picBack.Width

picIconsBack.Height = picBack.Height - picBack.Top - picIconsBack.Top
End Sub

Private Sub hsEdit_Change()
picEdit.Left = hsEdit.value + 8
End Sub

Private Sub lblSwitch_Click()
Dim l As Long
l = picClr(1).BackColor
picClr(1).BackColor = picClr(2).BackColor
picClr(2).BackColor = l

mColor(1) = picClr(1).BackColor
mColor(2) = picClr(2).BackColor
End Sub

Private Sub mnuEdit_Click()
Dim b As Boolean
b = picEdit.Visible
 mnuEditSelAll.Enabled = b
 mnuEditUndo.Enabled = b

 mnuEditCopy.Enabled = shpSel.Visible
 mnuEditCut.Enabled = shpSel.Visible
 mnuEditDelete.Enabled = shpSel.Visible
 mnuEditPaste.Enabled = Clipboard.GetFormat(2)
 
 mnuEditPaste.Enabled = b
End Sub

Private Sub mnuEditCopy_Click()
If shpSel.Visible = True Then
 Set picLyr(mLayer).Picture = picLyr(mLayer).Image
 ''Set picLyr(mLayer).Picture = picLyr(mLayer).Image:Set picLyr(mLayer + 5).Picture = picLyr(mLayer).Picture
 With picTemp
  .Width = Int(shpSel.Width / mZoom) '- 1
  .Height = Int(shpSel.Height / mZoom) ' - 1
  .PaintPicture picLyr(mLayer).Picture, 0, 0, picTemp.Width, picTemp.Height, Int(shpSel.Left / mZoom), Int(shpSel.Top / mZoom), Int(shpSel.Width / mZoom), Int(shpSel.Height / mZoom), vbSrcCopy
  Set .Picture = .Image
  .Refresh
  Clipboard.Clear
  Clipboard.SetData picTemp.Picture, 2
 End With
End If
End Sub

Private Sub mnuEditCut_Click()
If shpSel.Visible = True Then
 Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
 Set picLyr(mLayer).Picture = picLyr(mLayer).Image: Set picLyr(mLayer + 5).Picture = picLyr(mLayer).Picture
 With picTemp
  .Width = Int(shpSel.Width / mZoom) '- 1
  .Height = Int(shpSel.Height / mZoom) ' - 1
  .PaintPicture picLyr(mLayer).Picture, 0, 0, picTemp.Width, picTemp.Height, Int(shpSel.Left / mZoom), Int(shpSel.Top / mZoom), Int(shpSel.Width / mZoom), Int(shpSel.Height / mZoom), vbSrcCopy
  Set .Picture = .Image
  .Refresh
  Clipboard.Clear
  Clipboard.SetData picTemp.Picture, 2
 End With
 Call mnuEditDelete_Click
End If
End Sub

Private Sub mnuEditDelete_Click()
If shpSel.Visible = True Then
 FileChanged(mCurrIcon) = True
 Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
 Set picLyr(mLayer).Picture = picLyr(mLayer).Image: Set picLyr(mLayer + 5).Picture = picLyr(mLayer).Picture
 picLyr(mLayer).Line (Int(shpSel.Left / mZoom), Int(shpSel.Top / mZoom))-(Int(shpSel.Left / mZoom) + Int(shpSel.Width / mZoom) - 1, Int(shpSel.Top / mZoom) + Int(shpSel.Height / mZoom) - 1), picTrans.BackColor, BF
 Call SetAndDrawLayers
End If
End Sub

Private Sub mnuEditPaste_Click()
On Error GoTo 1
FileChanged(mCurrIcon) = True
Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
Set picLyr(mLayer).Picture = picLyr(mLayer).Image: Set picLyr(mLayer + 5).Picture = picLyr(mLayer).Picture
Set picTemp.Picture = Clipboard.GetData(2)

If shpSel.Visible = True And shpSel.Left <> -10000 Then
 picLyr(mLayer).PaintPicture picTemp.Picture, Int(shpSel.Left / mZoom), Int(shpSel.Top / mZoom), Int(shpSel.Width / mZoom), Int(shpSel.Height / mZoom), 0, 0, picTemp.Width, picTemp.Height, vbSrcCopy
Else
 Dim X As Integer, Y As Integer, c As Long
 For Y = 0 To picTemp.ScaleHeight - 1
  For X = 0 To picTemp.ScaleWidth - 1
   c = picTemp.Point(X, Y)
   If c <> picTrans.BackColor Then picLyr(mLayer).PSet (X, Y), c
  Next X
 Next Y
 'picIcon(mCurrIcon).PaintPicture picTemp.Picture, 0, 0, picTemp.Width, picTemp.Height, 0, 0, picTemp.Width, picTemp.Height, vbSrcCopy
End If
 Call SetAndDrawLayers
1
End Sub

Private Sub mnuEditSelAll_Click()
Dim b As MSComctlLib.Button
Set b = tbTools.Buttons(1)
Call tbTools_ButtonClick(b)
shpSel.Left = 1
shpSel.Top = 1
shpSel.Width = picIcon(mCurrIcon).Width * mZoom - 1
shpSel.Height = picIcon(mCurrIcon).Height * mZoom - 1
shpSel.Visible = True
End Sub

Private Sub mnuEditUndo_Click()
FileChanged(mCurrIcon) = True
Set picLyr(mLayer).Picture = picLyr(mLayer + 5).Image
 Call SetAndDrawLayers
End Sub

Private Sub mnuFile_Click()
Dim b As Boolean
b = picEdit.Visible

mnuFileClose.Enabled = b
mnuFileSave.Enabled = b
mnuFileSaveAs.Enabled = b
mnuFileSaveAllAs.Enabled = b
End Sub

Private Sub mnuFileClose_Click()
Dim a As Integer
 If picIcon(mCurrIcon).Tag <> "" And FileChanged(mCurrIcon) = True Then
  a = MsgBox(ExtractFileName(picIcon(mCurrIcon).Tag) & " has changed." & vbCrLf & "Do you want to save it?", vbQuestion + vbYesNoCancel, "Save Changes")
  If a = vbYes Then 'save
   If InStr(picIcon(mCurrIcon).Tag, "ico") Then
    Open picIcon(mCurrIcon).Tag For Binary Access Write As #1
     Put #1, 1, GenerateIconForSave$(picIcon(mCurrIcon))
    Close #1
   Else
    Call mnuFileSaveAs_Click
   End If
  ElseIf a = vbCancel Then
   Exit Sub
  End If
 End If

picEdit.Tag = ""
picEdit.Visible = False
picIcon(mCurrIcon).Tag = ""
picIcon(mCurrIcon).Visible = False
txtIcon(mCurrIcon).Caption = ""
txtIcon(mCurrIcon).Visible = False

Dim i As Integer, X As Integer, p As Integer
p = -1
For i = 0 To picIcon.Count - 1
 If picIcon(i).Tag <> "" Then
  picIcon(i).Top = X
  txtIcon(i).Top = X + picIcon(i).Height + 4
  X = X + picIcon(i).Height + txtIcon(i).Height + 8
 End If
Next i

picIconsMove.Height = X + 68
ReDim Preserve mIcons(mCurrIcon)
Set mIcons(mCurrIcon) = LoadPicture()

For i = 0 To 9
 Set picLyr(i).Picture = LoadPicture()
 Set mLPics(mCurrIcon, i) = LoadPicture()
Next i

End Sub

Private Sub mnuFileExit_Click()
Call Unload(Me)
End
End Sub

Private Sub mnuFileNew16_Click()
Call NewIcon(16)
Dim i As Integer
For i = 0 To 9
 picLyr(i).Width = 16
 picLyr(i).Height = 16
Next i
ReDim Preserve mIcons(mCurrIcon)
Set mIcons(mCurrIcon) = LoadPicture()

End Sub

Private Sub mnuFileNew32_Click()
Call NewIcon(32)
Dim i As Integer
For i = 0 To 9
 picLyr(i).Width = 32
 picLyr(i).Height = 32
Next i
ReDim Preserve mIcons(mCurrIcon)
Set mIcons(mCurrIcon) = LoadPicture()
End Sub

Private Sub mnuFileNew48_Click()
Call NewIcon(48)
Dim i As Integer
For i = 0 To 9
 picLyr(i).Width = 48
 picLyr(i).Height = 48
Next i
ReDim Preserve mIcons(mCurrIcon)
Set mIcons(mCurrIcon) = LoadPicture()
End Sub

Private Sub mnuFileOpen_Click()
On Error GoTo 1
cd.Filter = "Icon Files (*.ico, *.cur)|*.ico;*.cur|Bitmap Files (*.bmp)|*.bmp|JPEG Files (*.jpg)|*.jpg|GIF Files (*.gif)|*.gif|CozIcon (*.cci)|*.cci"
cd.ShowOpen
'Dim i As Integer
'i = GetIconSize(cd.FileName)

Select Case cd.FilterIndex
 Case 1
Dim d As ICONFILEDATA

d = LoadIcon(cd.Filename)

If d.ifdSuccess = False Then MsgBox "Icon could not be loaded.", vbCritical, "Error": Exit Sub

If d.ifdCount <> 1 Then
 If MsgBox("Multiple Icons were found in this file." & vbCrLf & "Would you like to see them all?", vbQuestion + vbYesNo, "Multiple Icons") = vbNo Then GoTo 2
 Temp_Icon = d
 Call Load(frmMIcons)
 frmMIcons.Show vbModal
 Dim c As ICONFILEDATA
 Temp_Icon = c
Else
2
 Call AddRecent(cd.Filename)
  'If i <> -1 Then
  Dim i As Integer, b As Integer
  b = mCurrIcon
 For i = 0 To picLyr.Count - 1
  Set picLyr(i).Picture = picLyr(i).Image
  picLyr(i).Width = picIcon(mCurrIcon).Width
  picLyr(i).Height = picIcon(mCurrIcon).Height
  Set mLPics(b, i) = picLyr(i).Picture
 Next i
  
  Call NewIcon(d.ifdIconData(0).idWidth)

 For i = 0 To picLyr.Count - 1
  Set picLyr(i).Picture = mLPics(mCurrIcon, i)
   picLyr(i).Width = picIcon(mCurrIcon).Width
  picLyr(i).Height = picIcon(mCurrIcon).Height
 Next i
  

  'Else Call NewIcon(32)
    
 'Set picUndo(mCurrIcon).Picture = d.ifdIcon(0) 'LoadPicture(cd.FileName)
 
 Set picIcon(mCurrIcon).Picture = d.ifdIcon(0) 'LoadPicture(cd.FileName)
 Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
 Set picLyr(0).Picture = picIcon(mCurrIcon).Picture
   
 picIcon(mCurrIcon).Refresh
 picIcon(mCurrIcon).Tag = cd.Filename
 txtIcon(mCurrIcon).Caption = Truncate(LCase(ExtractFileName(cd.Filename))) & vbCrLf & picIcon(mCurrIcon).Width & "x" & picIcon(mCurrIcon).Height
   
  Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
  Call DrawGrid
    
 picEdit.Refresh
 Set picEdit.Picture = picEdit.Image
 ReDim Preserve mIcons(mCurrIcon)
 Set mIcons(mCurrIcon) = d.ifdIcon(0)
End If
 Case 2 To 4
  Call NewIcon(32)
  Set picIcon(mCurrIcon).Picture = LoadPicture(cd.Filename)
 Case 5
  Call NewIcon(32)
  Dim l As Long, s As String
  l = FreeFile()
  Open cd.Filename For Binary Access Read As #l
   s = Input(LOF(l), #l)
  Close #l
  s = Mid(s, 4)
  Dim a() As String, v As Variant, j As Integer
  a() = Split(s, "&image&")
  For Each v In a()
   l = FreeFile()
   Call KillFile(App.Path & "\temp.bmp")
   Open App.Path & "\temp.bmp" For Binary Access Write As #l
    Put #l, , CStr(v)
   Close #l
   Set picLyr(j).Picture = LoadPicture(App.Path & "\temp.bmp")
   j = j + 1
  Next v
  Call SetAndDrawLayers
End Select
1
End Sub

Private Sub mnuFileRecentFiles_Click(Index As Integer)
Dim i As Integer
Dim s As String
s = mnuFileRecentFiles(Index).Caption
If FileExist(s) = False Then
 MsgBox s & vbCrLf & "Does not exist", vbInformation, "File Error"
 Exit Sub
End If

Dim d As ICONFILEDATA

d = LoadIcon(s)

If d.ifdSuccess = False Then MsgBox "Icon could not be loaded.", vbCritical, "Error": Exit Sub

If d.ifdCount <> 1 Then
 If MsgBox("Multiple Icons were found in this file." & vbCrLf & "Would you like to see them all?", vbQuestion + vbYesNo, "Multiple Icons") = vbNo Then GoTo 2
 cd.Filename = s
 Temp_Icon = d
 Call Load(frmMIcons)
 frmMIcons.Show vbModal
 Dim c As ICONFILEDATA
 Temp_Icon = c
Else
2
 Call AddRecent(s)
  'If i <> -1 Then
  Call NewIcon(d.ifdIconData(0).idWidth)
  'Else Call NewIcon(32)
    
 'Set picUndo(mCurrIcon).Picture = d.ifdIcon(0) 'LoadPicture(cd.FileName)
 Set picIcon(mCurrIcon).Picture = d.ifdIcon(0) 'LoadPicture(cd.FileName)
 Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
 For i = 0 To 9
  picLyr(i).Width = picIcon(mCurrIcon).Width
  picLyr(i).Height = picIcon(mCurrIcon).Height
 Next i
 Set picLyr(0).Picture = picIcon(mCurrIcon).Picture

 picIcon(mCurrIcon).Refresh
 picIcon(mCurrIcon).Tag = s
 txtIcon(mCurrIcon).Caption = Truncate(LCase(ExtractFileName(s))) & vbCrLf & picIcon(mCurrIcon).Width & "x" & picIcon(mCurrIcon).Height
   
  Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
  Call DrawGrid
    
 picEdit.Refresh
 Set picEdit.Picture = picEdit.Image
 ReDim Preserve mIcons(mCurrIcon)
 Set mIcons(mCurrIcon) = d.ifdIcon(0)
End If
End Sub

Private Sub mnuFileSave_Click()
If InStr(picIcon(mCurrIcon).Tag, "ico") Then
 If picIcon(mCurrIcon).Width <> 16 And picIcon(mCurrIcon).Height <> 16 And picIcon(mCurrIcon).Width <> 32 And picIcon(mCurrIcon).Height <> 32 And picIcon(mCurrIcon).Width <> 48 And picIcon(mCurrIcon).Height <> 48 Then MsgBox "CozIcon can only save 16x16, 32x32 or 48x48 icons, Sorry.", vbCritical, "Save Error": Exit Sub
 Open picIcon(mCurrIcon).Tag For Binary Access Write As #1
  Put #1, 1, GenerateIconForSave$(picIcon(mCurrIcon))
 Close #1
 FileChanged(mCurrIcon) = False
Else
 Call mnuFileSaveAs_Click
End If
End Sub

Private Sub mnuFileSaveAllAs_Click()
On Error GoTo 1
cd.Filter = "Icons (*.ico)|*.ico"
cd.ShowSave

Dim i As Integer, d As Integer
Dim temp As ICONFILEDATA, f As String
For i = 0 To picIcon.Count - 1
 If picIcon(i).Visible = True Then
  If picIcon(i).Width <> 32 And picIcon(i).Width <> 16 And picIcon(i).Width <> 48 Then
   MsgBox "CozIcon can only edit 48x48, 32x32 or 16x16 icons, Sorry", vbInformation, "Stop"
   GoTo 2
  End If
  ReDim Preserve temp.ifdIconData(temp.ifdCount)
  temp.ifdCount = temp.ifdCount + 1
  d = temp.ifdCount - 1
  temp.ifdIconData(d).idWidth = picIcon(i).Width
  temp.ifdIconData(d).idHeight = picIcon(i).Height
  temp.ifdIconData(d).idColorCount = 0
  temp.ifdIconData(d).idData = GenerateIconForSaveX(picIcon(i))
  temp.ifdIconData(d).idDataLength = Len(temp.ifdIconData(d).idData)
  'temp.ifdIconData(d).idDataOffset = 22 + (16 * d)

  f = f & Chr(temp.ifdIconData(d).idWidth) & Chr(temp.ifdIconData(d).idHeight) & Chr(temp.ifdIconData(d).idColorCount) & String(5, Chr(0)) & _
          Long2Chr(temp.ifdIconData(d).idDataLength) & Chr(0) & Chr(0) & "FFF" & (d + 1)
2
 End If
Next i

f = String(2, Chr(0)) & Chr(1) & Chr(0) & Chr(temp.ifdCount) & Chr(0) & f

For i = 0 To temp.ifdCount - 1
 f = Replace(f, "FFF" & (i + 1), Long2Chr(Len(f)) & String(4 - Len(Long2Chr(Len(f))), Chr(0)))
 f = f & temp.ifdIconData(i).idData
Next i

Call KillFile(cd.Filename)
Open cd.Filename For Binary Access Write As #1
 Put #1, , f
Close #1
1
End Sub

Private Sub mnuFileSaveAs_Click()
If picIcon(mCurrIcon).Width <> 32 And picIcon(mCurrIcon).Width <> 16 And picIcon(mCurrIcon).Width <> 48 Then
 MsgBox "CozIcon can only edit 48x48, 32x32 or 16x16 icons, Sorry", vbInformation, "Stop"
 Exit Sub
End If
On Error GoTo 1
cd.Filter = "True Color Icon (*.ico, *.cur)|*.ico;*.cur|Bitmap (*.bmp)|*.bmp|CozIcon (*.cci)|*.cci"
'cd.FilterIndex
cd.Filename = picIcon(mCurrIcon).Tag
cd.ShowSave
  Call AddRecent(cd.Filename)
  Dim l As Long
Select Case cd.FilterIndex
 Case 2
  Call SavePicture(picIcon(mCurrIcon), cd.Filename)
  Exit Sub
 Case 1
  l = FreeFile()
  Open cd.Filename For Binary Access Write As #l
   Put #l, 1, GenerateIconForSave$(picIcon(mCurrIcon))
  Close #l
 Case 3 'save layers of icon
  Dim i As Integer, s As String
  For i = 0 To 4
   Set picLyr(i).Picture = picLyr(i).Image
   Call SavePicture(picLyr(i), App.Path & "\temp.bmp")
   
   l = FreeFile()
   Open App.Path & "\temp.bmp" For Binary Access Read As #l
    s = s & Input(LOF(l), #l) & "&image&"
   Close #l
  Next i
   
   s = Left(s, Len(s) - 7)
   l = FreeFile()
   Open cd.Filename For Binary Access Write As #l
    Put #l, 1, CStr("CCI") & s$
   Close #l

  Exit Sub
'256 Color Icon (*.ico, *.cur)|*.ico;*.cur|16 Color Icon (*.ico, *.cur)|*.ico;*.cur|
End Select
picIcon(mCurrIcon).Tag = cd.Filename
txtIcon(mCurrIcon).Caption = ExtractFileName(cd.Filename) & vbCrLf & picIcon(mCurrIcon).Width & "x" & picIcon(mCurrIcon).Height
FileChanged(mCurrIcon) = False
1
End Sub

Private Sub mnuFilters_Click()
Dim i As Integer, b As Boolean
b = picEdit.Visible
For i = 0 To mnuFiltersRun.Count - 1
 mnuFiltersRun(i).Enabled = b
Next i
End Sub

Private Function GetFilterIndexByName(ByVal FName As String) As Integer
Dim i As Integer
For i = 0 To mnuFiltersRun.Count - 1
 If LCase(FName) = LCase(mnuFiltersRun(i).Caption) Then GetFilterIndexByName = i: Exit Function
Next i
GetFilterIndexByName = -1
End Function

Private Sub mnuFiltersRun_Click(Index As Integer)
'On Error Resume Next
Dim l As Long, X As Integer, Y As Integer
Dim s As String, arr() As String, a As Integer
l = FreeFile()

Open App.Path & "\filters\" & mnuFiltersRun(Index).Tag & ".ccf" For Input As #l
 s = Input(LOF(l), #l)
Close #l

arr() = Split(s, vbCrLf)
If arr(0) <> "filter" Then Exit Sub
For l = 1 To UBound(arr())
 If arr(l) = "#" & picLyr(mLayer).ScaleWidth Or arr(l) = "#*" Then a = l + 1: Exit For
Next l

Call DrawLayers(mCurrIcon, mLayer)
Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image

Set picLyr(mLayer).Picture = picLyr(mLayer).Image
Set picLyr(mLayer + 5).Picture = picLyr(mLayer).Picture

Dim c As Integer, d As Long
Dim sL As String, sR As String
Dim r As Double, g As Double, b As Double

Dim pX As Integer, pY As Integer
Dim sX As Integer, sY As Integer
Dim stX As Integer, stY As Integer
Dim UserInput As String, i As Integer

For l = a To UBound(arr())
 If Len(arr(l)) < 10 Then a = l: Exit For
 Select Case LCase(Left(arr(l), 10))
  Case "user_input"
   UserInput = Trim(InputBox(Mid(arr(l), 12), "User Input"))
   If UserInput = "" Then Exit Sub
  Case "pixelskipx"
   pX = CInt(Trim(Mid(arr(l), 12)))
  Case "pixelskipy"
   pY = CInt(Trim(Mid(arr(l), 12)))
  Case "pixelstrtx"
   sX = CInt(Trim(Mid(arr(l), 12)))
  Case "pixelstrty"
   sY = CInt(Trim(Mid(arr(l), 12)))
  Case "pixelstopx"
   stX = CInt(Trim(Mid(arr(l), 12)))
  Case "pixelstopy"
   stY = CInt(Trim(Mid(arr(l), 12)))
  Case "run_filter"
   i = GetFilterIndexByName(Trim(Mid(arr(l), 12)))
   If i <> -1 Then Call mnuFiltersRun_Click(i)
  Case Else
   a = l
   Exit For
 End Select
Next l
stY = IIf(stY = 0, picLyr(mLayer).ScaleHeight - 1, stY)
stX = IIf(stX = 0, picLyr(mLayer).ScaleHeight - 1, stX)

frmFilter.Show
frmFilter.Refresh
frmFilter.ProgressBar1.Max = stY

For Y = sY To stY
 For X = sX To stX
  For l = a To UBound(arr())
   If Left(arr(l), 1) = "#" Then Exit For
    c = InStr(arr(l), "=")
    'Debug.Print piclyr(mlayer).Point(X, Y), picTrans.BackColor, X, Y
    If c <> 0 And picLyr(mLayer).Point(X, Y) <> picTrans.BackColor Then
     sL = Trim(Left(arr(l), c - 1))
     sR = Trim(Mid(arr(l), c + 1))
     If UserInput <> "" And IsNumeric(UserInput) = True Then sR = Trim(Replace(sR, "ui", CDbl(UserInput), , , vbTextCompare))
     sR = Trim(Replace(sR, "x", X, , , vbTextCompare))
     sR = Trim(Replace(sR, "y", Y, , , vbTextCompare))
     sR = Trim(Replace(sR, "w", picLyr(mLayer).ScaleWidth, , , vbTextCompare))
     sR = Trim(Replace(sR, "h", picLyr(mLayer).ScaleHeight, , , vbTextCompare))
     'Debug.Print sR, Eval(sR)
      Select Case LCase(sL)
       Case "red"
        r = Abs(Eval(sR))
       Case "green"
        g = Abs(Eval(sR))
       Case "blue"
        b = Abs(Eval(sR))
      End Select
      If r > 255 Then r = 255
      If g > 255 Then g = 255
      If b > 255 Then b = 255
      If RGB(r, g, b) = picTrans.BackColor Then r = r + 1: g = g + 1: b = b + 1
            picLyr(mLayer).PSet (X, Y), RGB(r, g, b)
    End If
  Next l
   X = X + pX
 Next X
 frmFilter.ProgressBar1.value = Y
 Y = Y + pY
Next Y
frmFilter.Hide
 Call SetAndDrawLayers
FileChanged(mCurrIcon) = True
End Sub

Private Sub mnuHelp_Click()
Dim b As Boolean
b = picEdit.Visible

mnuHelpTest.Enabled = b
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub mnuHelpTest_Click()
Set picIcon(mCurrIcon).Picture = Me.Icon
Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
picEdit.Cls
Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
Call DrawGrid
picEdit.Refresh
Set picEdit.Picture = picEdit.Image
End Sub

Private Sub mnuImage_Click()
Dim b As Boolean
b = picEdit.Visible

mnuImageClear.Enabled = b
mnuImageFlipHo.Enabled = b
mnuImageFlipVert.Enabled = b
mnuImageRotate.Enabled = b
mnuImageOriginal.Enabled = b
mnuImageCopyTo.Enabled = b
End Sub

Private Sub mnuImageClear_Click()
FileChanged(mCurrIcon) = True
Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
Set picLyr(mLayer).Picture = picLyr(mLayer).Image: Set picLyr(mLayer + 5).Picture = picLyr(mLayer).Picture
Set picIcon(mCurrIcon).Picture = LoadPicture()
Dim i As Integer
For i = 0 To picLyr.Count - 1
 Set mLPics(mCurrIcon, i) = LoadPicture()
 Set picLyr(i).Picture = LoadPicture()
Next i
picEdit.Cls
Set picEdit.Picture = LoadPicture()
Call DrawGrid
picEdit.Refresh
Set picEdit.Picture = picEdit.Image
End Sub

Private Sub mnuImageCopyNew_Click()

Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
  Dim i As Integer, b As Integer
 i = mCurrIcon
 For b = 0 To picLyr.Count - 1
  Set picLyr(b).Picture = picLyr(b).Image
  picLyr(b).Width = picIcon(mCurrIcon).Width
  picLyr(b).Height = picIcon(mCurrIcon).Height
  Set mLPics(i, b) = picLyr(b).Picture
 Next b
  
 Call NewIcon(picIcon(mCurrIcon).ScaleWidth)

 For b = 0 To picLyr.Count - 1
  Set picLyr(b).Picture = mLPics(mCurrIcon, b)
   picLyr(b).Width = picIcon(mCurrIcon).Width
  picLyr(b).Height = picIcon(mCurrIcon).Height
 Next b

Call picIcon(mCurrIcon).PaintPicture(picIcon(i).Picture, 0, 0, picIcon(i).Width, picIcon(i).Height, 0, 0, picIcon(i).Width, picIcon(i).Height, vbSrcCopy)
Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
Set picLyr(mLayer).Picture = picIcon(mCurrIcon).Picture
txtIcon(mCurrIcon).Caption = Truncate("Copy of " & LCase(ExtractFileName(picIcon(i).Tag))) & vbCrLf & picIcon(mCurrIcon).Width & "x" & picIcon(mCurrIcon).Height
Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
Call DrawGrid
picEdit.Refresh
Set picEdit.Picture = picEdit.Image
End Sub

Private Sub mnuImageFlipHo_Click()
FileChanged(mCurrIcon) = True
Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
Set picLyr(mLayer).Picture = picLyr(mLayer).Image: Set picLyr(mLayer + 5).Picture = picLyr(mLayer).Picture
Call IconMirror(picIcon(mCurrIcon))
Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
picEdit.Cls
Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
Call DrawGrid
picEdit.Refresh
Set picEdit.Picture = picEdit.Image

End Sub

Private Sub mnuImageFlipVert_Click()
FileChanged(mCurrIcon) = True
Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
Set picLyr(mLayer).Picture = picLyr(mLayer).Image: Set picLyr(mLayer + 5).Picture = picLyr(mLayer).Picture
Call IconFlip(picIcon(mCurrIcon))
Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
picEdit.Cls
Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
Call DrawGrid
picEdit.Refresh
Set picEdit.Picture = picEdit.Image
End Sub

Private Sub mnuImageOriginal_Click()
On Error GoTo 1
Dim i As Integer
For i = 0 To 4
 Set picLyr(i).Picture = LoadPicture()
Next i
Set picLyr(0).Picture = mIcons(mCurrIcon)
 Call SetAndDrawLayers
FileChanged(mCurrIcon) = False
1
End Sub

Private Sub mnuImagePal_Click()
Call Load(frmPal)
Dim i As Integer, X As Integer, Y As Integer
If picIcon(mCurrIcon).Tag = "untitled" Then
2
Dim jpArr() As JPALETTE
Call GeneratePalette(picIcon(mCurrIcon), jpArr())
For i = 0 To UBound(jpArr())
 frmPal.picPal.Line (X, Y)-(X + 16, Y + 16), jpArr(i).jpColor, BF
 X = X + 16
 If X + 16 > frmPal.picPal.ScaleWidth Then X = 0: Y = Y + 16
Next i
Else
Dim dIcons As ICONFILEDATA
dIcons = LoadIcon(picIcon(mCurrIcon).Tag)
 If dIcons.ifdIconData(0).idColorCount2 <= 256 Then
For i = 0 To dIcons.ifdIconData(0).idColorCount2 - 1
 
 frmPal.picPal.Line (X, Y)-(X + 16, Y + 16), dIcons.ifdIconData(0).idPalette(i).jpColor, BF
 X = X + 16
 If X + 16 > frmPal.picPal.ScaleWidth Then X = 0: Y = Y + 16
Next i
 Else
  GoTo 2
 End If
End If
frmPal.Show vbModal
Call Unload(frmPal)
End Sub

Private Sub mnuImageRotate_Click()
Dim i As Integer, s As String
s = InputBox("Enter degree, 0 - 360", "Rotate Image")
If s = "" Or IsNumeric(s) = False Then Exit Sub
 If CInt(s) >= 0 And CInt(s) <= 360 Then
  Set picLyr(mLayer).Picture = picLyr(mLayer).Image
  Set picLyr(mLayer + 5).Picture = picLyr(mLayer).Picture
  Set picTemp.Picture = picLyr(mLayer).Picture
  Set picLyr(mLayer).Picture = LoadPicture()
  Call ImageRotate(picTemp, picLyr(mLayer), CInt(s), True)
  picLyr(mLayer).Refresh
 Call SetAndDrawLayers
  
 End If
End Sub

Private Sub mnuPlugins_Click()
Dim i As Integer, b As Boolean
b = picEdit.Visible
For i = 0 To mnuPluginsRun.Count - 1
 mnuPluginsRun(i).Enabled = b
Next i
End Sub

Private Sub mnuPluginsRun_Click(Index As Integer)
Call wc.Run(App.Path & "\plugins\" & mnuPluginsRun(Index).Tag & ".exe")
End Sub

Private Sub mnuView_Click()
Dim b As Boolean
b = picEdit.Visible

mnuViewZoom.Enabled = b
mnuViewGrid.Enabled = b
End Sub

Private Sub mnuViewBrushes_Click()
If mnuViewBrushes.Checked = True Then
 mnuViewBrushes.Checked = False
Else
 mnuViewBrushes.Checked = True
End If

picBrushes.Visible = mnuViewBrushes.Checked
Call picBack_Resize
End Sub

Private Sub mnuViewGrid_Click()
mDrawGrid = IIf(mDrawGrid = True, False, True)
mnuViewGrid.Checked = mDrawGrid
Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
picEdit.Cls
Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
Call DrawGrid
picEdit.Refresh
Set picEdit.Picture = picEdit.Image

End Sub

Private Sub mnuViewLayers_Click()
If mnuViewLayers.Checked = True Then
 mnuViewLayers.Checked = False
Else
 mnuViewLayers.Checked = True
End If

picLayers.Visible = mnuViewLayers.Checked
Call picBack_Resize
End Sub

Private Sub mnuViewZoomSize_Click(Index As Integer)
If picIcon(mCurrIcon).ScaleWidth > 32 And Index = 20 Then Call mnuViewZoomSize_Click(10): Exit Sub
Dim oz As Integer
oz = mZoom
mZoom = Index
mnuViewZoomSize(1).Checked = False
mnuViewZoomSize(5).Checked = False
mnuViewZoomSize(10).Checked = False
mnuViewZoomSize(20).Checked = False
mnuViewZoomSize(Index).Checked = True

 picEdit.Width = picIcon(mCurrIcon).Width * mZoom
 picEdit.Height = picIcon(mCurrIcon).Height * mZoom
 picIcon(mCurrIcon).Refresh
 Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
 picEdit.Cls
 Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
 Call DrawGrid
 picEdit.Refresh
 Set picEdit.Picture = picEdit.Image
 If shpSel.Visible = True Then
  shpSel.Left = Int(shpSel.Left / oz) * mZoom
  shpSel.Top = Int(shpSel.Top / oz) * mZoom
  shpSel.Width = Int(shpSel.Width / oz) * mZoom
  shpSel.Height = Int(shpSel.Height / oz) * mZoom
  OldSel(0) = shpSel.Left: OldSel(1) = shpSel.Top
  OldSel(2) = shpSel.Width: OldSel(3) = shpSel.Height
 End If
Call picBack_Resize
End Sub

Private Sub optEdit_Click()
MsgBox mCustomTool
optEdit.value = False
End Sub

Private Sub picBack_Click()
On Error GoTo 1
cd.ShowColor

picBack.BackColor = cd.Color
1
End Sub

Private Sub picBack_Resize()
If picLayers.Visible = True And picBrushes.Visible = False Then
 picLayers.Left = picBack.ScaleWidth - picLayers.Width
 picLayers.Height = picBack.ScaleHeight
 picLayers.Top = 0
ElseIf picLayers.Visible = False And picBrushes.Visible = True Then
 picBrushes.Left = picBack.ScaleWidth - picBrushes.Width
 picBrushes.Height = picBack.ScaleHeight
 picBrushes.Top = 0
Else
 picLayers.Left = picBack.ScaleWidth - picLayers.Width
 picLayers.Height = picBack.ScaleHeight / 2
 picLayers.Top = 0
 
 picBrushes.Left = picBack.ScaleWidth - picBrushes.Width
 picBrushes.Height = picBack.ScaleHeight / 2
 picBrushes.Top = picBrushes.Height
End If

hsEdit.Left = 0
hsEdit.Top = picBack.ScaleHeight - hsEdit.Height
hsEdit.Width = picBack.ScaleWidth - vsEdit.Width - IIf(picLayers.Visible = True Or picBrushes.Visible = True, picLayers.Width, 0)
vsEdit.Top = 0
vsEdit.Left = picBack.ScaleWidth - vsEdit.Width - IIf(picLayers.Visible = True Or picBrushes.Visible = True, picLayers.Width, 0)
vsEdit.Height = picBack.ScaleHeight - hsEdit.Height
optEdit.Width = vsEdit.Width
optEdit.Height = hsEdit.Height
optEdit.Top = hsEdit.Top
optEdit.Left = vsEdit.Left

If picEdit.Height > picBack.ScaleHeight - hsEdit.Height Then
 vsEdit.Max = picBack.ScaleHeight - picEdit.Height - hsEdit.Height - 16
 vsEdit.value = 0
Else
 vsEdit.Max = 0
 vsEdit.value = 0
End If
If picEdit.Width > picBack.ScaleWidth - vsEdit.Width - IIf(picLayers.Visible = True Or picBrushes.Visible = True, picLayers.Width, 0) Then
 hsEdit.Max = picBack.ScaleWidth - picEdit.Width - vsEdit.Width - 8 - IIf(picLayers.Visible = True Or picBrushes.Visible = True, picLayers.Width + 8, 0)
 hsEdit.value = 0
Else
 hsEdit.Max = 0
 hsEdit.value = 0
End If
End Sub

Private Sub picBrushes_Resize()
tabBrushes.Height = picBrushes.ScaleHeight
cbPens.Height = picBrushes.ScaleHeight - cbPens.Top - 6
cbBox.Height = cbPens.Height
End Sub

Private Sub picClr_Click(Index As Integer)
On Error GoTo 1
cd.Color = picClr(Index).BackColor
cd.ShowColor
Dim l As Long
If cd.Color = picTrans.BackColor Then l = cd.Color + 1 Else l = cd.Color

picClr(Index).BackColor = l
mColor(Index) = l

  Call GradientClr(picClrSel(1), vbBlack, l)
  Call GradientClr(picClrSel(2), l, vbWhite)

Call ShowRGBVal(Index)
1
End Sub

Private Sub picClrSel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
 Case 1
  If Button <> 0 Then picClr(Button).BackColor = picClrSel(1).Point(X, Y): mColor(Button) = picClrSel(1).Point(X, Y)
 Case 2
  If Button <> 0 Then picClr(Button).BackColor = picClrSel(2).Point(X, Y): mColor(Button) = picClrSel(2).Point(X, Y)
End Select

Call ShowRGBVal(Button)
End Sub

Private Sub picClrSel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 0 Then Exit Sub
If X < 0 Or Y < 0 Then Exit Sub
If X >= picClrSel(Index).ScaleWidth Or Y >= picClrSel(Index).ScaleHeight Then Exit Sub
Select Case Index
 Case 1
  If Button <> 0 Then picClr(Button).BackColor = picClrSel(1).Point(X, Y): mColor(Button) = picClrSel(1).Point(X, Y)
 Case 2
  If Button <> 0 Then picClr(Button).BackColor = picClrSel(2).Point(X, Y): mColor(Button) = picClrSel(2).Point(X, Y)
End Select

Call ShowRGBVal(Button)
End Sub

Private Sub picEdit_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 68 Then
 mColor(1) = 0
 mColor(2) = vbWhite
 picClr(1).BackColor = 0
 picClr(2).BackColor = vbWhite
 Call ShowRGBVal(1)
 Exit Sub
End If

 Dim X As Integer, Y As Integer, c As Long, b As Boolean

If Shift = 2 And KeyCode = 37 Then
 picTemp.Width = picIcon(mCurrIcon).Width - 1
 picTemp.Height = picIcon(mCurrIcon).Height
 Set picLyr(mLayer).Picture = picLyr(mLayer).Image
 picTemp.PaintPicture picLyr(mLayer).Picture, 0, 0, picTemp.Width, picTemp.Height, 1, 0, picTemp.Width, picTemp.Height, vbSrcCopy
 Set picLyr(mLayer).Picture = Nothing
 For Y = 0 To picTemp.ScaleHeight - 1
  For X = 0 To picTemp.ScaleWidth - 1
   c = picTemp.Point(X, Y)
   If c <> picTrans.BackColor Then picLyr(mLayer).PSet (X, Y), c
  Next X
 Next Y
 b = True
End If

If Shift = 2 And KeyCode = 38 Then
 picTemp.Width = picIcon(mCurrIcon).Width
 picTemp.Height = picIcon(mCurrIcon).Height - 1
 Set picLyr(mLayer).Picture = picLyr(mLayer).Image
 picTemp.PaintPicture picLyr(mLayer).Picture, 0, 0, picTemp.Width, picTemp.Height, 0, 1, picTemp.Width, picTemp.Height, vbSrcCopy
 Set picLyr(mLayer).Picture = Nothing
 For Y = 0 To picTemp.ScaleHeight - 1
  For X = 0 To picTemp.ScaleWidth - 1
   c = picTemp.Point(X, Y)
   If c <> picTrans.BackColor Then picLyr(mLayer).PSet (X, Y), c
  Next X
 Next Y
 b = True
End If

If Shift = 2 And KeyCode = 39 Then
 picTemp.Width = picIcon(mCurrIcon).Width - 1
 picTemp.Height = picIcon(mCurrIcon).Height
 Set picLyr(mLayer).Picture = picLyr(mLayer).Image
 picTemp.PaintPicture picLyr(mLayer).Picture, 0, 0, picTemp.Width, picTemp.Height, 0, 0, picTemp.Width, picTemp.Height, vbSrcCopy
 Set picLyr(mLayer).Picture = Nothing
 For Y = 0 To picTemp.ScaleHeight - 1
  For X = 0 To picTemp.ScaleWidth - 1
   c = picTemp.Point(X, Y)
   If c <> picTrans.BackColor Then picLyr(mLayer).PSet (X + 1, Y), c
  Next X
 Next Y
 b = True
End If

If Shift = 2 And KeyCode = 40 Then
 picTemp.Width = picIcon(mCurrIcon).Width
 picTemp.Height = picIcon(mCurrIcon).Height - 1
 Set picLyr(mLayer).Picture = picLyr(mLayer).Image
 picTemp.PaintPicture picLyr(mLayer).Picture, 0, 0, picTemp.Width, picTemp.Height, 0, 0, picTemp.Width, picTemp.Height, vbSrcCopy
 Set picLyr(mLayer).Picture = Nothing
 For Y = 0 To picTemp.ScaleHeight - 1
  For X = 0 To picTemp.ScaleWidth - 1
   c = picTemp.Point(X, Y)
   If c <> picTrans.BackColor Then picLyr(mLayer).PSet (X, Y + 1), c
  Next X
 Next Y
 b = True
End If

If b = True Then
 Call SetAndDrawLayers
 Exit Sub
End If

If mButton = 0 Then Exit Sub
KeyDown = Shift + KeyCode
Exit Sub
Select Case SelTool
 Case tBoxGrad, tBoxGradNS
  Call Tool_Box_Move(X2, Y2, True)
 Case tCircleGrad
  Call Tool_Circle_Move(X2, Y2)
 Case tBox To tBoxFill
  Call Tool_Box_Move(X2, Y2)
 Case tCircle To tCircleFill
  Call Tool_Circle_Move(X2, Y2)
End Select

End Sub

Private Sub picEdit_KeyUp(KeyCode As Integer, Shift As Integer)
KeyDown = 0
Exit Sub
Select Case SelTool
 Case tBoxGrad, tBoxGradNS
  Call Tool_Box_Move(X2, Y2, True)
 Case tCircleGrad
  Call Tool_Circle_Move(X2, Y2)
 Case tBox To tBoxFill
  Call Tool_Box_Move(X2, Y2)
 Case tCircle To tCircleFill
  Call Tool_Circle_Move(X2, Y2)
End Select

End Sub

Private Sub picEdit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If picIcon(mCurrIcon).Width <> 32 And picIcon(mCurrIcon).Width <> 16 And picIcon(mCurrIcon).Width <> 48 Then
 MsgBox "CozIcon can only edit 48x48, 32x32 or 16x16 icons, Sorry", vbInformation, "Stop"
 Exit Sub
End If
If Button <> 1 And Button <> 2 Then Exit Sub
Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
Set picLyr(mLayer).Picture = picLyr(mLayer).Image
Set picLyr(mLayer + 5).Picture = picLyr(mLayer).Picture
mButton = Button
x1 = Int(X / mZoom) * mZoom
y1 = Int(Y / mZoom) * mZoom
Select Case SelTool
 Case tSelect
  If x1 < shpSel.Left Or x1 > shpSel.Left + shpSel.Width Or _
     y1 < shpSel.Top Or y1 > shpSel.Top + shpSel.Height Then
      shpSel.Visible = False
      shpSel.Left = -10000
      shpSel.Top = -10000
      shpSel.Width = 10
      shpSel.Height = 20
      MovingSel = False
  End If
 Case tDraw
  x1 = Int(X / mZoom) * mZoom
  y1 = Int(Y / mZoom) * mZoom
  X2 = Int(X / mZoom) * mZoom + mZoom - 1
  Y2 = Int(Y / mZoom) * mZoom + mZoom - 1
  'mColor(mButton)
  'MsgBox mLayer & " - " & (mColor(Button) = picTrans.BackColor)
   picEdit.Line (x1, y1)-(X2, Y2), BlendColor(picIcon(mCurrIcon).Point(Int(x1 / mZoom), Int(y1 / mZoom)), mColor(mButton), mOpacity), BF
  Set picEdit.Picture = picEdit.Image
  picIcon(mCurrIcon).PSet (Int(x1 / mZoom), Int(Y / mZoom)), BlendColor(picIcon(mCurrIcon).Point(Int(x1 / mZoom), Int(y1 / mZoom)), mColor(mButton), mOpacity)
  picLyr(mLayer).PSet (Int(x1 / mZoom), Int(Y / mZoom)), BlendColor(picIcon(mCurrIcon).Point(Int(x1 / mZoom), Int(y1 / mZoom)), mColor(mButton), mOpacity)
 Case tEye
  mColor(Button) = picIcon(mCurrIcon).Point(Int(X / mZoom), Int(Y / mZoom))
  picClr(Button).BackColor = mColor(Button)
  Call ShowRGBVal(Button)
  Call GradientClr(picClrSel(1), vbBlack, mColor(Button))
  Call GradientClr(picClrSel(2), mColor(Button), vbWhite)
 Case tCustom
  If tCustomType = 2 Then
  X2 = Int(X / mZoom) * mZoom
  Y2 = Int(Y / mZoom) * mZoom
  Dim s As String, a As String, arr() As String, arrX() As String, v As Variant
  s = mCustomTool
  arrX() = Split(s, vbCrLf)
  For Each v In arrX()
  s = v
  s = Replace(s, "h", Int(Y2 / mZoom) - Int(y1 / mZoom))
  s = Replace(s, "w", Int(X2 / mZoom) - Int(x1 / mZoom))
  s = Replace(s, "x1", Int(x1 / mZoom))
  s = Replace(s, "x2", Int(X2 / mZoom))
  s = Replace(s, "y1", Int(y1 / mZoom))
  s = Replace(s, "y2", Int(Y2 / mZoom))
  If InStr(s, " ") <> 0 Then
  a = Left(s, InStr(s, " ") - 1)
  arr() = Split(Mid(s, InStr(s, " ") + 1), ",")
  ReDim Preserve arr(UBound(arr()) + 6)
  Select Case LCase(Trim(a))
   Case "line"
    picIcon(mCurrIcon).Line (Eval(arr(0)), Eval(arr(1)))-(Eval(arr(2)), Eval(arr(3))), IIf(arr(4) = "", mColor(Button), Eval(Replace(arr(4), "|", ">")))
    picLyr(mLayer).Line (Eval(arr(0)), Eval(arr(1)))-(Eval(arr(2)), Eval(arr(3))), IIf(arr(4) = "", mColor(Button), Eval(Replace(arr(4), "|", ">")))
    picIcon(mCurrIcon).PSet (Eval(arr(2)), Eval(arr(3))), IIf(arr(4) = "", mColor(Button), Eval(Replace(arr(4), "|", ">")))
    picLyr(mLayer).PSet (Eval(arr(2)), Eval(arr(3))), IIf(arr(4) = "", mColor(Button), Eval(Replace(arr(4), "|", ">")))
   Case "dot"
     picIcon(mCurrIcon).PSet (Eval(arr(0)), Eval(arr(1))), IIf(arr(2) = "", mColor(Button), Eval(Replace(arr(2), "|", ">")))
     picLyr(mLayer).PSet (Eval(arr(0)), Eval(arr(1))), IIf(arr(2) = "", mColor(Button), Eval(Replace(arr(2), "|", ">")))
   Case "box"
    picIcon(mCurrIcon).Line (Eval(arr(0)), Eval(arr(1)))-(Eval(arr(2)), Eval(arr(3))), IIf(arr(4) = "", mColor(Button), Eval(Replace(arr(4), "|", ">"))), B
    picLyr(mLayer).Line (Eval(arr(0)), Eval(arr(1)))-(Eval(arr(2)), Eval(arr(3))), IIf(arr(4) = "", mColor(Button), Eval(Replace(arr(4), "|", ">"))), B
   Case "filledbox"
    picIcon(mCurrIcon).Line (Eval(arr(0)), Eval(arr(1)))-(Eval(arr(2)), Eval(arr(3))), IIf(arr(4) = "", mColor(Button), Eval(Replace(arr(4), "|", ">"))), BF
    picLyr(mLayer).Line (Eval(arr(0)), Eval(arr(1)))-(Eval(arr(2)), Eval(arr(3))), IIf(arr(4) = "", mColor(Button), Eval(Replace(arr(4), "|", ">"))), BF
   Case "circle"
    picIcon(mCurrIcon).Circle (Eval(arr(0)), Eval(arr(1))), Eval(arr(2)), IIf(arr(3) = "", mColor(Button), Eval(Replace(arr(3), "|", ">")))
    picLyr(mLayer).Circle (Eval(arr(0)), Eval(arr(1))), Eval(arr(2)), IIf(arr(3) = "", mColor(Button), Eval(Replace(arr(3), "|", ">")))
  End Select
  End If
  Next v
  picIcon(mCurrIcon).Refresh
  Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
  picEdit.Cls
  Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
  Call DrawGrid
  picEdit.Refresh
  Set picEdit.Picture = picEdit.Image
  
  End If

End Select
End Sub

Private Sub picEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error Resume Next
Caption = ExtractFileName(picIcon(mCurrIcon).Tag) & IIf(FileChanged(mCurrIcon) = True, " *", "") & " - CozIcon [" & Int(X / mZoom) & ", " & Int(Y / mZoom) & "]"
If x1 = -1 Then Exit Sub
'picedit.Cls
If Button <> 1 And Button <> 2 Then Exit Sub

Select Case SelTool
 Case tSelect
  X2 = Int(X / mZoom) * mZoom + mZoom - 1
  Y2 = Int(Y / mZoom) * mZoom + mZoom - 1
    If MovingSel = True Then
     shpSel.Left = xz1 + (X2 - x1)
     shpSel.Top = yz1 + (Y2 - y1)
     Exit Sub
    End If
     
     xz1 = x1 - (mZoom - 1)
     yz1 = y1 - (mZoom - 1)
    shpSel.Left = x1
    shpSel.Top = y1
    If X2 > x1 Then shpSel.Width = Int(X2 - x1) + 2
    If shpSel.Width > picIcon(mCurrIcon).Width * mZoom Then shpSel.Width = picIcon(mCurrIcon).Width * mZoom
    If Y2 > y1 Then shpSel.Height = Int(Y2 - y1) + 2
    If shpSel.Height > picIcon(mCurrIcon).Height * mZoom Then shpSel.Height = picIcon(mCurrIcon).Height * mZoom
    OldSel(0) = shpSel.Left: OldSel(1) = shpSel.Top
    OldSel(2) = shpSel.Width: OldSel(3) = shpSel.Height
    shpSel.Visible = True
    MovingSel = False
 Case tDraw
  If Int(X / mZoom) * mZoom = x1 And Int(Y / mZoom) * mZoom = y1 Then Exit Sub
  x1 = Int(X / mZoom) * mZoom
  y1 = Int(Y / mZoom) * mZoom
  X2 = Int(X / mZoom) * mZoom
  Y2 = Int(Y / mZoom) * mZoom
  
  picEdit.Line (X2, Y2)-(X2 + mZoom, Y2 + mZoom), BlendColor(picIcon(mCurrIcon).Point(Int(x1 / mZoom), Int(y1 / mZoom)), mColor(mButton), mOpacity), BF
  Set picEdit.Picture = picEdit.Image
  picIcon(mCurrIcon).PSet (Int(X2 / mZoom), Int(Y2 / mZoom)), BlendColor(picIcon(mCurrIcon).Point(Int(x1 / mZoom), Int(y1 / mZoom)), mColor(mButton), mOpacity)
  picLyr(mLayer).PSet (Int(X2 / mZoom), Int(Y2 / mZoom)), BlendColor(picIcon(mCurrIcon).Point(Int(x1 / mZoom), Int(y1 / mZoom)), mColor(mButton), mOpacity)
 Case tEye
  If picIcon(mCurrIcon).Point(Int(X / mZoom), Int(Y / mZoom)) < 0 Then Exit Sub
  mColor(Button) = picIcon(mCurrIcon).Point(Int(X / mZoom), Int(Y / mZoom))
  picClr(Button).BackColor = mColor(Button)
  Call ShowRGBVal(Button)
  Call GradientClr(picClrSel(1), vbBlack, mColor(Button))
  Call GradientClr(picClrSel(2), mColor(Button), vbWhite)
 Case tLine, tLineX2, tLineX3
  Call Tool_Line_Move(X, Y)
 Case tBoxGrad
  drawing = True
  Call Tool_Box_Move(X, Y, 1)
 Case tBoxGradNS
  drawing = True
  Call Tool_Box_Move(X, Y, 2)
 Case tBoxGradNESW
  drawing = True
  Call Tool_Box_Move(X, Y, 4)
 Case tBoxGradNWSE
  drawing = True
  Call Tool_Box_Move(X, Y, 3)
 Case tCircleGrad
  drawing = True
  Call Tool_Circle_Move(X, Y)
 Case tBox To tBoxFill, tBoxFillX
  drawing = True
  Call Tool_Box_Move(X, Y)
 Case tCustom
  If tCustomType = 1 Then
   drawing = True
   Call Tool_Box_Move(X, Y)
  ElseIf tCustomType = 2 Then
  X2 = Int(X / mZoom) * mZoom
  Y2 = Int(Y / mZoom) * mZoom
  Dim s As String, a As String, arr() As String, arrX() As String, v As Variant
  s = mCustomTool
  arrX() = Split(s, vbCrLf)
  For Each v In arrX()
  s = v
  s = Replace(s, "h", Int(Y2 / mZoom) - Int(y1 / mZoom))
  s = Replace(s, "w", Int(X2 / mZoom) - Int(x1 / mZoom))
  s = Replace(s, "x1", Int(x1 / mZoom))
  s = Replace(s, "x2", Int(X2 / mZoom))
  s = Replace(s, "y1", Int(y1 / mZoom))
  s = Replace(s, "y2", Int(Y2 / mZoom))
  If InStr(s, " ") <> 0 Then
  a = Left(s, InStr(s, " ") - 1)
  arr() = Split(Mid(s, InStr(s, " ") + 1), ",")
  ReDim Preserve arr(UBound(arr()) + 6)
  Select Case LCase(Trim(a))
   Case "line"
    picIcon(mCurrIcon).Line (Eval(arr(0)), Eval(arr(1)))-(Eval(arr(2)), Eval(arr(3))), IIf(arr(4) = "", mColor(Button), Eval(Replace(arr(4), "|", ">")))
    picLyr(mLayer).Line (Eval(arr(0)), Eval(arr(1)))-(Eval(arr(2)), Eval(arr(3))), IIf(arr(4) = "", mColor(Button), Eval(Replace(arr(4), "|", ">")))
    picIcon(mCurrIcon).PSet (Eval(arr(2)), Eval(arr(3))), IIf(arr(4) = "", mColor(Button), Eval(Replace(arr(4), "|", ">")))
    picLyr(mLayer).PSet (Eval(arr(2)), Eval(arr(3))), IIf(arr(4) = "", mColor(Button), Eval(Replace(arr(4), "|", ">")))
   Case "dot"
     picIcon(mCurrIcon).PSet (Eval(arr(0)), Eval(arr(1))), IIf(arr(2) = "", mColor(Button), Eval(Replace(arr(2), "|", ">")))
     picLyr(mLayer).PSet (Eval(arr(0)), Eval(arr(1))), IIf(arr(2) = "", mColor(Button), Eval(Replace(arr(2), "|", ">")))
   Case "box"
    picIcon(mCurrIcon).Line (Eval(arr(0)), Eval(arr(1)))-(Eval(arr(2)), Eval(arr(3))), IIf(arr(4) = "", mColor(Button), Eval(Replace(arr(4), "|", ">"))), B
    picLyr(mLayer).Line (Eval(arr(0)), Eval(arr(1)))-(Eval(arr(2)), Eval(arr(3))), IIf(arr(4) = "", mColor(Button), Eval(Replace(arr(4), "|", ">"))), B
   Case "filledbox"
    picIcon(mCurrIcon).Line (Eval(arr(0)), Eval(arr(1)))-(Eval(arr(2)), Eval(arr(3))), IIf(arr(4) = "", mColor(Button), Eval(Replace(arr(4), "|", ">"))), BF
    picLyr(mLayer).Line (Eval(arr(0)), Eval(arr(1)))-(Eval(arr(2)), Eval(arr(3))), IIf(arr(4) = "", mColor(Button), Eval(Replace(arr(4), "|", ">"))), BF
   Case "circle"
    picIcon(mCurrIcon).Circle (Eval(arr(0)), Eval(arr(1))), Eval(arr(2)), IIf(arr(3) = "", mColor(Button), Eval(Replace(arr(3), "|", ">")))
    picLyr(mLayer).Circle (Eval(arr(0)), Eval(arr(1))), Eval(arr(2)), IIf(arr(3) = "", mColor(Button), Eval(Replace(arr(3), "|", ">")))
  End Select
  End If
  Next v
  picIcon(mCurrIcon).Refresh
  Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
  picEdit.Cls
  Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
  Call DrawGrid
  picEdit.Refresh
  Set picEdit.Picture = picEdit.Image
  
  End If
 Case tCircle To tCircleFill, tCircleFillX
  drawing = True
  Call Tool_Circle_Move(X, Y)
End Select

picEdit.Refresh
End Sub

Private Sub picEdit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case SelTool
 Case tSelect
  If MovingSel = False Then
   MovingSel = True
  Else
    xz1 = shpSel.Left - mZoom + 1
    yz1 = shpSel.Top - mZoom + 1
    If xz1 <= -9000 Or yz1 <= -9000 Then Exit Sub
    picTemp.Width = Int(shpSel.Width / mZoom)
    picTemp.Height = Int(shpSel.Height / mZoom)

    If OldSel(2) <= 0 Or OldSel(3) <= 0 Then Exit Sub
    picTemp.PaintPicture picIcon(mCurrIcon).Picture, 0, 0, picTemp.Width, picTemp.Height, Int(OldSel(0) / mZoom), Int(OldSel(1) / mZoom), Int(OldSel(2) / mZoom), Int(OldSel(3) / mZoom), vbSrcCopy
    Set picTemp.Picture = picTemp.Image
    If KeyDown <> 19 Then
     picIcon(mCurrIcon).Line (Int(OldSel(0) / mZoom), Int(OldSel(1) / mZoom))-(Int(OldSel(0) / mZoom) + Int(OldSel(2) / mZoom) - 1, Int(OldSel(1) / mZoom) + Int(OldSel(3) / mZoom) - 1), picTrans.BackColor, BF
     picLyr(mLayer).Line (Int(OldSel(0) / mZoom), Int(OldSel(1) / mZoom))-(Int(OldSel(0) / mZoom) + Int(OldSel(2) / mZoom) - 1, Int(OldSel(1) / mZoom) + Int(OldSel(3) / mZoom) - 1), picTrans.BackColor, BF
    End If
    picIcon(mCurrIcon).PaintPicture picTemp.Picture, Int(shpSel.Left / mZoom), Int(shpSel.Top / mZoom), Int(shpSel.Width / mZoom), Int(shpSel.Height / mZoom), 0, 0, picTemp.Width, picTemp.Height, vbSrcCopy
    picLyr(mLayer).PaintPicture picTemp.Picture, Int(shpSel.Left / mZoom), Int(shpSel.Top / mZoom), Int(shpSel.Width / mZoom), Int(shpSel.Height / mZoom), 0, 0, picTemp.Width, picTemp.Height, vbSrcCopy
    picEdit.Refresh
    
    OldSel(0) = shpSel.Left: OldSel(1) = shpSel.Top
    OldSel(2) = shpSel.Width: OldSel(3) = shpSel.Height
  End If
  'Exit Sub
 Case tLine, tLineX2 To tLineX3
  picIcon(mCurrIcon).Line (Int(x1 / mZoom), Int(y1 / mZoom))-(Int(X2 / mZoom), Int(Y2 / mZoom)), mColor(Button)
  picLyr(mLayer).Line (Int(x1 / mZoom), Int(y1 / mZoom))-(Int(X2 / mZoom), Int(Y2 / mZoom)), mColor(Button)
  picIcon(mCurrIcon).PSet (Int(X2 / mZoom), Int(Y2 / mZoom)), mColor(Button)
  picLyr(mLayer).PSet (Int(X2 / mZoom), Int(Y2 / mZoom)), mColor(Button)
  
  If SelTool >= tLineX2 Then
   If Abs(X2 - x1) > Abs(Y2 - y1) Then
    picIcon(mCurrIcon).Line (Int(x1 / mZoom), Int(y1 / mZoom) + 1)-(Int(X2 / mZoom), Int(Y2 / mZoom) + 1), mColor(Button)
    picLyr(mLayer).Line (Int(x1 / mZoom), Int(y1 / mZoom) + 1)-(Int(X2 / mZoom), Int(Y2 / mZoom) + 1), mColor(Button)
    picIcon(mCurrIcon).PSet (Int(X2 / mZoom), Int(Y2 / mZoom) + 1), mColor(Button)
    picLyr(mLayer).PSet (Int(X2 / mZoom), Int(Y2 / mZoom) + 1), mColor(Button)
   Else
    picIcon(mCurrIcon).Line (Int(x1 / mZoom) + 1, Int(y1 / mZoom))-(Int(X2 / mZoom) + 1, Int(Y2 / mZoom)), mColor(Button)
    picLyr(mLayer).Line (Int(x1 / mZoom) + 1, Int(y1 / mZoom))-(Int(X2 / mZoom) + 1, Int(Y2 / mZoom)), mColor(Button)
    picIcon(mCurrIcon).PSet (Int(X2 / mZoom) + 1, Int(Y2 / mZoom)), mColor(Button)
    picLyr(mLayer).PSet (Int(X2 / mZoom) + 1, Int(Y2 / mZoom)), mColor(Button)
   End If
  End If

  If SelTool >= tLineX3 Then
   If Abs(X2 - x1) > Abs(Y2 - y1) Then
    picIcon(mCurrIcon).Line (Int(x1 / mZoom), Int(y1 / mZoom) - 1)-(Int(X2 / mZoom), Int(Y2 / mZoom) - 1), mColor(Button)
    picLyr(mLayer).Line (Int(x1 / mZoom), Int(y1 / mZoom) - 1)-(Int(X2 / mZoom), Int(Y2 / mZoom) - 1), mColor(Button)
    picIcon(mCurrIcon).PSet (Int(X2 / mZoom), Int(Y2 / mZoom) - 1), mColor(Button)
    picLyr(mLayer).PSet (Int(X2 / mZoom), Int(Y2 / mZoom) - 1), mColor(Button)
   Else
    picIcon(mCurrIcon).Line (Int(x1 / mZoom) - 1, Int(y1 / mZoom))-(Int(X2 / mZoom) - 1, Int(Y2 / mZoom)), mColor(Button)
    picLyr(mLayer).Line (Int(x1 / mZoom) - 1, Int(y1 / mZoom))-(Int(X2 / mZoom) - 1, Int(Y2 / mZoom)), mColor(Button)
    picIcon(mCurrIcon).PSet (Int(X2 / mZoom) - 1, Int(Y2 / mZoom)), mColor(Button)
    picLyr(mLayer).PSet (Int(X2 / mZoom) - 1, Int(Y2 / mZoom)), mColor(Button)
   End If
  End If

 Case tFill
   Call FillRegion(X, Y)
   Call DrawLayers(mCurrIcon)
 Case tBox To tBoxFill
   If drawing = False Then Exit Sub
   Call Tool_Box_Up(X, Y, CBool(SelTool - 3))
 Case tBoxFillX
   If drawing = False Then Exit Sub
   Call Tool_Box_Up(X, Y, True, True)
 Case tBoxGrad, tBoxGradNS, tBoxGradNESW, tBoxGradNWSE
  If drawing = False Then Exit Sub
  If SelTool = tBoxGrad Then
   Call Tool_BoxGrad_Up(X, Y, gEW)
  ElseIf SelTool = tBoxGradNS Then
   Call Tool_BoxGrad_Up(X, Y, gNS)
  ElseIf SelTool = tBoxGradNWSE Then
   Call Tool_BoxGrad_Up(X, Y, gNWSE)
  ElseIf SelTool = tBoxGradNESW Then
   Call Tool_BoxGrad_Up(X, Y, gNESW)
  End If
 Case tCircleGrad
  If drawing = False Then Exit Sub
  Call Tool_CircleGrad_Up(X, Y)
 Case tCircle To tCircleFill
  If drawing = False Then Exit Sub
  Call Tool_Circle_Up(X, Y, CBool(SelTool - 5))
 Case tCircleFillX
  If drawing = False Then Exit Sub
  Call Tool_Circle_Up(X, Y, True, True)
 Case tCustom 'custom tool
  If tCustomType = 1 Then
  X2 = Int(X / mZoom) * mZoom
  Y2 = Int(Y / mZoom) * mZoom
  Dim s As String, a As String, arr() As String, arrX() As String, v As Variant
  s = mCustomTool
  arrX() = Split(s, vbCrLf)
  For Each v In arrX()
  s = v
  s = Replace(s, "w", Int(X2 / mZoom) - Int(x1 / mZoom))
  s = Replace(s, "h", Int(Y2 / mZoom) - Int(y1 / mZoom))
  s = Replace(s, "x1", Int(x1 / mZoom))
  s = Replace(s, "x2", Int(X2 / mZoom))
  s = Replace(s, "y1", Int(y1 / mZoom))
  s = Replace(s, "y2", Int(Y2 / mZoom))
  'MsgBox s
  If InStr(s, " ") <> 0 Then
  a = Left(s, InStr(s, " ") - 1)
  arr() = Split(Mid(s, InStr(s, " ") + 1), ",")
  ReDim Preserve arr(UBound(arr()) + 6)
  Select Case LCase(Trim(a))
   Case "line"
    picIcon(mCurrIcon).Line (Eval(arr(0)), Eval(arr(1)))-(Eval(arr(2)), Eval(arr(3))), IIf(arr(4) = "", mColor(Button), Eval(Replace(arr(4), "|", ">")))
    picLyr(mLayer).Line (Eval(arr(0)), Eval(arr(1)))-(Eval(arr(2)), Eval(arr(3))), IIf(arr(4) = "", mColor(Button), Eval(Replace(arr(4), "|", ">")))
    picIcon(mCurrIcon).PSet (Eval(arr(2)), Eval(arr(3))), IIf(arr(4) = "", mColor(Button), Eval(Replace(arr(4), "|", ">")))
    picLyr(mLayer).PSet (Eval(arr(2)), Eval(arr(3))), IIf(arr(4) = "", mColor(Button), Eval(Replace(arr(4), "|", ">")))
   Case "dot"
     picIcon(mCurrIcon).PSet (Eval(arr(0)), Eval(arr(1))), IIf(arr(2) = "", mColor(Button), Eval(Replace(arr(2), "|", ">")))
     picLyr(mLayer).PSet (Eval(arr(0)), Eval(arr(1))), IIf(arr(2) = "", mColor(Button), Eval(Replace(arr(2), "|", ">")))
   Case "box"
    picIcon(mCurrIcon).Line (Eval(arr(0)), Eval(arr(1)))-(Eval(arr(2)), Eval(arr(3))), IIf(arr(4) = "", mColor(Button), Eval(Replace(arr(4), "|", ">"))), B
    picLyr(mLayer).Line (Eval(arr(0)), Eval(arr(1)))-(Eval(arr(2)), Eval(arr(3))), IIf(arr(4) = "", mColor(Button), Eval(Replace(arr(4), "|", ">"))), B
   Case "filledbox"
    picIcon(mCurrIcon).Line (Eval(arr(0)), Eval(arr(1)))-(Eval(arr(2)), Eval(arr(3))), IIf(arr(4) = "", mColor(Button), Eval(Replace(arr(4), "|", ">"))), BF
    picLyr(mLayer).Line (Eval(arr(0)), Eval(arr(1)))-(Eval(arr(2)), Eval(arr(3))), IIf(arr(4) = "", mColor(Button), Eval(Replace(arr(4), "|", ">"))), BF
   Case "circle"
    picIcon(mCurrIcon).Circle (Eval(arr(0)), Eval(arr(1))), Eval(arr(2)), IIf(arr(3) = "", mColor(Button), Eval(Replace(arr(3), "|", ">")))
    picLyr(mLayer).Circle (Eval(arr(0)), Eval(arr(1))), Eval(arr(2)), IIf(arr(3) = "", mColor(Button), Eval(Replace(arr(3), "|", ">")))
  End Select
  End If
  Next v
  End If
  End Select
mButton = 0
drawing = False
 Call SetAndDrawLayers
 x1 = -1: y1 = -1
FileChanged(mCurrIcon) = True
End Sub

Private Sub picEdit_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 1
Dim ext As String, Filename As String
Filename = Data.Files(1)
ext = Right(Filename, 4)
If ext = ".ico" Then
Dim i As Integer
i = GetIconSize(Filename)

Call AddRecent(Filename)
If i <> -1 Then Call NewIcon(i) Else Call NewIcon(32)
Set picUndo(mCurrIcon).Picture = LoadPicture(Filename)
Set picIcon(mCurrIcon).Picture = LoadPicture(Filename)
Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
picIcon(mCurrIcon).Refresh
picIcon(mCurrIcon).Tag = Filename
txtIcon(mCurrIcon).Caption = ExtractFileName(Filename) & vbCrLf & picIcon(mCurrIcon).Width & "x" & picIcon(mCurrIcon).Height
Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
Call DrawGrid
picEdit.Refresh
Set picEdit.Picture = picEdit.Image
End If
1
End Sub

Private Sub picIcon_Click(Index As Integer)
 shpSel.Left = -10000
 shpSel.Top = -10000
 shpSel.Width = 10
 shpSel.Height = 20
 MovingSel = False


'On Error Resume Next
Dim i As Integer, b As Integer
 b = mCurrIcon
 For i = 0 To picLyr.Count - 1
 
  Set picLyr(i).Picture = picLyr(i).Image
  picLyr(i).Width = picIcon(mCurrIcon).Width
  picLyr(i).Height = picIcon(mCurrIcon).Height
  Set mLPics(b, i) = picLyr(i).Picture
 Next i
 
mCurrIcon = Index
 
 For i = 0 To picLyr.Count - 1
  Set picLyr(i).Picture = mLPics(mCurrIcon, i)
  picLyr(i).Width = picIcon(mCurrIcon).Width
  picLyr(i).Height = picIcon(mCurrIcon).Height
 Next i

mLayer = 0
shpLyrSel.Top = picLyr(0).Top - 5

picIcon(mCurrIcon).Refresh
Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
picEdit.Cls
picEdit.Width = picIcon(mCurrIcon).Width * mZoom
picEdit.Height = picIcon(mCurrIcon).Height * mZoom
Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
Call DrawGrid
picEdit.Refresh
Set picEdit.Picture = picEdit.Image
picEdit.Visible = True
End Sub

Private Sub picIconsBack_Resize()
vsIcons.Left = 0
vsIcons.Top = picIconsBack.Height - vsIcons.Height - 4
vsIcons.Width = picIconsBack.Width - 4
End Sub

Private Sub picIconsMove_Resize()
vsIcons.Max = picIconsMove.Height - picIconsBack.Height
End Sub

Private Sub picLayers_Resize()
vsLyr.Height = picLayers.ScaleHeight - tbLyr.Height - 4
If picLayers.Height < picLayersBack.Height + 10 Then vsLyr.Max = (picLayersBack.Height + 20 - picLayers.Height) Else vsLyr.Max = 0
End Sub

Private Sub picLyr_Click(Index As Integer)
mLayer = Index
shpLyrSel.Top = picLyr(Index).Top - 5
If picEdit.Visible = True Then picEdit.SetFocus
End Sub

Private Sub picToolSel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer, j As Integer
Dim l As Long
Dim arr() As String
     
shpToolSel(Index).Visible = True
shpToolSel(Index).Top = Int(Y / 20) * 20
shpToolSel(Index).Left = Int(X / 20) * 20
  
Select Case Index
 Case 1
 i = Int(Y / 20)
  Select Case i
   Case 0
    SelTool = tDraw
   Case 1
   On Error GoTo 1
    cd.Filter = "Custom Pens (*.ccp)|*.ccp"
    If mCustomTool = "" Then cd.Filename = App.Path
    cd.ShowOpen
    SelTool = tCustom
     l = FreeFile
    Open cd.Filename For Input As #l
     mCustomTool = Input(LOF(1), #1)
    Close #l
     arr() = Split(mCustomTool, vbCrLf)
     Select Case LCase(arr(0))
      Case "box"
       tCustomType = 1
      Case "draw"
       tCustomType = 2
      Case Else
       tCustomType = -1
     End Select
     mCustomTool = ""
     For i = 1 To UBound(arr())
      mCustomTool = mCustomTool & arr(i) & vbCrLf
     Next i
1
  End Select
 Case 3
 i = Int(Y / 20)
 j = Int(X / 20)
  Select Case i
   Case 0
    If j = 0 Then
     SelTool = tBox
    Else
     SelTool = tBoxFill
    End If
   Case 1
    If j = 0 Then
     SelTool = tBoxFillX
    Else
   On Error GoTo 2
    cd.Filter = "Custom Boxes (*.ccb)|*.ccb"
    If mCustomTool = "" Then cd.Filename = App.Path
    cd.ShowOpen
    SelTool = tCustom
     l = FreeFile
    Open cd.Filename For Input As #l
     mCustomTool = Input(LOF(1), #1)
    Close #l
     arr() = Split(mCustomTool, vbCrLf)
     Select Case LCase(arr(0))
      Case "box"
       tCustomType = 1
      Case "draw"
       tCustomType = 2
      Case Else
       tCustomType = -1
     End Select
     mCustomTool = ""
     For i = 1 To UBound(arr())
      mCustomTool = mCustomTool & arr(i) & vbCrLf
     Next i
2
    End If
   Case 2
    If j = 0 Then
     SelTool = tBoxGrad
    Else
     SelTool = tBoxGradNESW
    End If
   Case 3
    If j = 0 Then
     SelTool = tBoxGradNS
    Else
     SelTool = tBoxGradNWSE
    End If
  End Select
 Case 5
 i = Int(Y / 20)
  Select Case i
   Case 0
    SelTool = tCircle
   Case 1
    SelTool = tCircleFill
   Case 2
    SelTool = tCircleFillX
   Case 3
    SelTool = tCircleGrad
  End Select
 Case 10
 i = Int(Y / 20)
  Select Case i
   Case 0
    SelTool = tLine
   Case 1
    SelTool = tLineX2
   Case 2
    SelTool = tLineX3
  End Select
End Select
End Sub

Private Sub picToolSel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer, j As Integer
Dim l As Long
Dim arr() As String

 
Select Case Index
 Case 1
 i = Int(Y / 20)
  Select Case i
   Case 0
    picToolSel(Index).ToolTipText = "Pen Tool"
   Case 1
    picToolSel(Index).ToolTipText = "Custom Pen Tool"
  End Select
 Case 3
 i = Int(Y / 20)
 j = Int(X / 20)
  Select Case i
   Case 0
    If j = 0 Then
     picToolSel(Index).ToolTipText = "Box Tool"
    Else
     picToolSel(Index).ToolTipText = "Filled Box"
    End If
   Case 1
    If j = 0 Then
     picToolSel(Index).ToolTipText = "Filled Box w/ Border"
    Else
     picToolSel(Index).ToolTipText = "Custom Box Tool"
    End If
   Case 2
    If j = 0 Then
     picToolSel(Index).ToolTipText = "Gradient Filled Box EW"
    Else
     picToolSel(Index).ToolTipText = "Gradient Filled Box NESW"
    End If
   Case 3
    If j = 0 Then
     picToolSel(Index).ToolTipText = "Gradient Filled Box NS"
    Else
     picToolSel(Index).ToolTipText = "Gradient Filled Box NWSE"
    End If
  End Select
 Case 5
 i = Int(Y / 20)
  Select Case i
   Case 0
    picToolSel(Index).ToolTipText = "Circle Tool"
   Case 1
    picToolSel(Index).ToolTipText = "Filled Circle"
   Case 2
    picToolSel(Index).ToolTipText = "Filled Circle w/ Border"
   Case 3
    picToolSel(Index).ToolTipText = "Gradient Filled Circle"
  End Select
 Case 10
 i = Int(Y / 20)
  Select Case i
   Case 0
    picToolSel(Index).ToolTipText = "Line Tool"
   Case 1
    picToolSel(Index).ToolTipText = "Thick Line Tool"
   Case 2
    picToolSel(Index).ToolTipText = "Thicker Line Tool"
  End Select
End Select

End Sub

Private Sub picTrans_Click()
picClr(1).BackColor = picTrans.BackColor
mColor(1) = picClr(1).BackColor
Dim s As String
s = GetRGB(mColor(1)).Red & "," & GetRGB(mColor(1)).Green & "," & GetRGB(mColor(1)).Blue
Clipboard.Clear
Clipboard.SetText s
End Sub

Private Sub tabBrushes_Click()
cbBox.Visible = CBool(tabBrushes.SelectedItem.Index - 1)
'Select Case tabBrushes.SelectedItem.Index
' Case 1
'  cbBox.Visible = False
' Case 2
'  cbBox.Visible = True
'End Select
End Sub

Private Sub tbLyr_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo 2
Dim i As Integer, c As Integer, d As Integer, H As Integer, Y As Integer, b As Boolean
Select Case Button.Index
 Case 1 'delete
   If MsgBox("Are you sure you want to clear " & IIf(mLayer = 0, "the Background Layer", "Layer " & mLayer) & "?", vbQuestion + vbYesNo, "Delete Layer") = vbNo Then Exit Sub
    
    Set picLyr(mLayer).Picture = LoadPicture
    
    'picLyr(picLyr.Count - 1).Tag = ""
    b = True
 Case 2 'toggle visiblity
  If picLyr(mLayer).Tag = "invisible" Then
   picLyr(mLayer).Tag = "visible"
   lblLyr(mLayer).FontStrikethru = False
  Else
   picLyr(mLayer).Tag = "invisible"
   lblLyr(mLayer).FontStrikethru = True
  End If
  b = True
 Case 3 'move layer up
  If mLayer = 0 Then Exit Sub
  Dim p As Picture
   Set picLyr(mLayer - 1).Picture = picLyr(mLayer - 1).Image
   Set p = picLyr(mLayer - 1).Picture
   Set picLyr(mLayer).Picture = picLyr(mLayer).Image
   Set picLyr(mLayer - 1).Picture = picLyr(mLayer).Picture
   Set picLyr(mLayer).Picture = p
  Set p = Nothing
  mLayer = mLayer - 1
  b = True
 Case 4 'move layer down
  If mLayer >= 4 Then Exit Sub
  Dim e As Picture
   Set picLyr(mLayer + 1).Picture = picLyr(mLayer + 1).Image
   Set e = picLyr(mLayer + 1).Picture
   Set picLyr(mLayer).Picture = picLyr(mLayer).Image
   Set picLyr(mLayer + 1).Picture = picLyr(mLayer).Picture
   Set picLyr(mLayer).Picture = e
  Set e = Nothing
  mLayer = mLayer + 1
  b = True
 Case 5 'merge layer up
  If mLayer = 0 Then Exit Sub
  Call MergeLayers(mCurrIcon, mLayer - 1, mLayer)
  Call SetAndDrawLayers
 Case 6 'merge layer down
  If mLayer = 4 Then Exit Sub
  Call MergeLayers(mCurrIcon, mLayer, mLayer + 1, True)
  Call SetAndDrawLayers
 Case 7 'flatten image
  Call MergeLayers(mCurrIcon, 0, 4)
  Set picLyr(0).Picture = picLyr(0).Image
  Set picIcon(mCurrIcon).Picture = picLyr(0).Picture
End Select

If b = True Then
  shpLyrSel.Top = picLyr(mLayer).Top - 5
 Call SetAndDrawLayers
End If

2
End Sub

Private Sub tbTools_ButtonClick(ByVal Button As MSComctlLib.Button)
'Set picEdit.MouseIcon = ilIcons.ListImages(Button.Index).Picture
'picEdit.MousePointer = 99
picEdit.Cls
shpSel.Visible = False
shpSel.Left = -1000
shpSel.Top = -1000

If Button.Tag = tText Then
  Call Load(frmText)
  Set frmText.picTemp.Picture = picIcon(mCurrIcon).Picture
  Call frmText.Show(vbModal)
  Exit Sub
End If
SelTool = Button.Tag
Dim b As MSComctlLib.Button
For Each b In tbTools.Buttons
 b.value = tbrUnpressed
Next b
Button.value = tbrPressed
tbTools.Refresh
On Error Resume Next
Dim i As Integer
For i = 0 To 10
 picToolSel(i).Visible = False
Next i

shpToolSel(SelTool).Width = 20
shpToolSel(SelTool).Left = 0
shpToolSel(SelTool).Height = 20
shpToolSel(SelTool).Top = 0
shpToolSel(SelTool).Visible = True
picToolSel(SelTool).Visible = True
picToolSel(SelTool).Left = (picExtra.Width / 2) - (picToolSel(SelTool).Width / 2) - 4
End Sub

Private Sub vsEdit_Change()
picEdit.Top = vsEdit.value + 8
End Sub

Private Sub vsIcons_Change()
If picIconsMove.Height - picIconsBack.Height > 0 Then picIconsMove.Top = Val("-" & vsIcons.value) + 8 Else picIconsMove.Top = 8
End Sub

Private Sub vsLyr_Change()
picLayersBack.Top = -(vsLyr.value - 20)
End Sub

Private Sub wc_Got(ByVal Msg As String)
Dim X As Integer, Y As Integer
Dim i As Integer, b As Integer
Select Case Left(Msg, 1)
 Case "@" 'send icon
  Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
  Call SavePicture(picIcon(mCurrIcon).Picture, App.Path & "\temp.bmp")
  wc.mHwnd = Mid(Msg, 2)
  Call wc.Send("$" & App.Path & "\temp.bmp")
 Case "+" 'send all icons
  For i = 0 To picIcon.Count - 1
   If picIcon(i).Tag <> "" Then
    b = b + 1
    Set picIcon(i).Picture = picIcon(i).Image
    Call SavePicture(picIcon(i).Picture, App.Path & "\temp.bmp")

    Call wc.Send("-" & b & "-" & App.Path & "\temp.bmp")
   End If
  Next i
 Case ">" 'new 16x16 icon
 b = mCurrIcon
 For i = 0 To picLyr.Count - 1
  Set picLyr(i).Picture = picLyr(i).Image
  Set mLPics(b, i) = picLyr(i).Picture
 Next i
  
  Call NewIcon(16)

 For i = 0 To picLyr.Count - 1
  Set picLyr(i).Picture = mLPics(mCurrIcon, i)
  picLyr(i).Width = picIcon(mCurrIcon).Width
  picLyr(i).Height = picIcon(mCurrIcon).Height
 Next i

  FileChanged(mCurrIcon) = True
  Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
  Set picLyr(mLayer).Picture = picLyr(mLayer).Image: Set picLyr(mLayer + 5).Picture = picLyr(mLayer).Picture
  Set picIcon(mCurrIcon).Picture = LoadPicture(App.Path & "\temp.bmp")
  For Y = 0 To 15
   For X = 0 To 15

    If picIcon(mCurrIcon).Point(X, Y) = 8420352 Then picIcon(mCurrIcon).PSet (X, Y), picTrans.BackColor
    
   Next X
  Next Y
  picIcon(mCurrIcon).Refresh
  Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
  Set picLyr(mLayer).Picture = picIcon(mCurrIcon).Picture
    picEdit.Cls
    Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
    Call DrawGrid
    picEdit.Refresh
    Set picEdit.Picture = picEdit.Image
 Case "<" 'new 32x32 icon
 b = mCurrIcon
 For i = 0 To picLyr.Count - 1
  Set picLyr(i).Picture = picLyr(i).Image
  Set mLPics(b, i) = picLyr(i).Picture
 Next i
  
  Call NewIcon(32)

 For i = 0 To picLyr.Count - 1
  Set picLyr(i).Picture = mLPics(mCurrIcon, i)
  picLyr(i).Width = picIcon(mCurrIcon).Width
  picLyr(i).Height = picIcon(mCurrIcon).Height
 Next i

  FileChanged(mCurrIcon) = True
  Set picLyr(mLayer).Picture = picLyr(mLayer).Image: Set picLyr(mLayer + 5).Picture = picLyr(mLayer).Picture
  Set picIcon(mCurrIcon).Picture = LoadPicture(App.Path & "\temp.bmp")
  For Y = 0 To 31
   For X = 0 To 31

    If picIcon(mCurrIcon).Point(X, Y) = 8420352 Then picIcon(mCurrIcon).PSet (X, Y), picTrans.BackColor
    
   Next X
  Next Y
  picIcon(mCurrIcon).Refresh
  Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
  Set picLyr(mLayer).Picture = picIcon(mCurrIcon).Picture
    picEdit.Cls
    Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
    Call DrawGrid
    picEdit.Refresh
    Set picEdit.Picture = picEdit.Image
 Case "!" 'change current icon
  FileChanged(mCurrIcon) = True
  Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
  Set picLyr(mLayer).Picture = picLyr(mLayer).Image: Set picLyr(mLayer + 5).Picture = picLyr(mLayer).Picture
  Set picIcon(mCurrIcon).Picture = LoadPicture(App.Path & "\temp.bmp")
  For Y = 0 To 31
   For X = 0 To 31

    If picIcon(mCurrIcon).Point(X, Y) = 8420352 Then picIcon(mCurrIcon).PSet (X, Y), picTrans.BackColor
    
   Next X
  Next Y
  picIcon(mCurrIcon).Refresh
  Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
  Set picLyr(mLayer).Picture = picIcon(mCurrIcon).Picture
    picEdit.Cls
    Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
    Call DrawGrid
    picEdit.Refresh
    Set picEdit.Picture = picEdit.Image
    
End Select
End Sub

