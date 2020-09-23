VERSION 5.00
Begin VB.Form Form1 
   Caption         =   " [ jBistoGOOD@Hotmail.com ]"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   3840
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1395
      Left            =   300
      TabIndex        =   7
      Top             =   2700
      Width           =   3255
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   300
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   480
         Width           =   2655
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   735
      Left            =   1800
      TabIndex        =   6
      Top             =   1740
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   1740
      Width           =   1275
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   180
      TabIndex        =   4
      Top             =   1260
      Width           =   3255
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   375
      Left            =   1860
      TabIndex        =   3
      Top             =   660
      Width           =   1515
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   180
      Width           =   1515
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   660
      Width           =   1515
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   60
      Top             =   2460
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'######################################
'#                                    #
'# Windows XP theme style buttons etc #
'#    [ jBistoGOOD@Hotmail.com ]      #
'#                                    #
'######################################

'>>>>>>>>> THIS CODE SHOULD WORK 'OUT OF THE BOX', SO JUST COPY AND PASTE IN INTO UR APP...
'>>>>>>>>> THE ONLY THING YOU NEED TO ADD TO THE FORM IS A TIMER... SET IT TO AROUND 2000
'>>>>>>>>>> YOU MUST DISABLE THE TIMER!!!!!! OR YOUR APPILCATION WILL NOT WORK!!

Option Explicit ' EVERY VARIBLE MUST BE "DIM"-ED

Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long ' API FOR UPGRADING CONTROLS

Private Declare Function SetFileAttributes Lib "kernel32.dll" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Const FILE_ATTRIBUTE_HIDDEN = &H2 '                     API FOR SETTING THE MANIFEST AS HIDDEN



Private Sub Form_Initialize() 'BEFORE THE USER SEES FORM
Dim xptheme As Long
Dim manifestpth As String 'DIM THE VARIBLES ETC

On Error GoTo manifestdoesnotexisT 'IF NO MANIFEST THEME FILE HAS BEEN MADE YET

If Right(App.Path, 1) = "\" Then                                 '|
    manifestpth = App.Path & App.EXEName & ".exe.manifest"       '|
Else                                                             '|
    manifestpth = App.Path & "\" & App.EXEName & ".exe.manifest" '|  FIND OUT IF MANIFEST ALREADY EXISTS
End If                                                           '|

FileCopy manifestpth, "c:\checkexist.txt"
Kill "c:\checkexist.txt"
xptheme = InitCommonControls                        ' IF MANIFEST EXISTS, EXUCUTE CONTROL UPGRADE TO XP THEME STYLE
Exit Sub

manifestdoesnotexisT:
Call makeNEWmanifest   ' IF MANIFEST DOES NOT EXIST, AND ERROR OCURRS, GO AND MAKE A NEW ONE
End Sub

Sub makeNEWmanifest()

Dim NEWmanifestpth As String
Dim xptheme As Long             ' SET VARIBLES ETC...
Dim setAShidden As Long

On Error GoTo problemARGH ' ERROR HANDLING, GOTO PROBLEMARGH ON ERROR EVENT

If Right(App.Path, 1) = "\" Then                                        '|
    NEWmanifestpth = App.Path & App.EXEName & ".exe.manifest"           '|
Else                                                                    '| SET PATH OF MANIFEST THEME FILE
    NEWmanifestpth = App.Path & "\" & App.EXEName & ".exe.manifest"     '|
End If                                                                  '|

Open NEWmanifestpth For Output As #1  '     WRITE THE MANIFEST FILE BECAUSE IT DOES NOT YET EXIST.
Print #1, "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & " standalone=" & Chr(34) & "yes" & Chr(34) & "?><assembly xmlns=" & Chr(34) & "urn:schemas-microsoft-com:asm.v1" & Chr(34) & " manifestVersion=" & Chr(34) & "1.0" & Chr(34) & "><assemblyIdentity version=" & Chr(34) & "1.0.0.0" & Chr(34) & " processorArchitecture=" & Chr(34) & "X86" & Chr(34) & " name=" & Chr(34) & "HybridDesign.WindowsXP.Example" & Chr(34) & " type=" & Chr(34) & "win32" & Chr(34) & " /> <description>An example of windows XP theming.</description> <dependency> <dependentAssembly> <assemblyIdentity type=" & Chr(34) & "win32" & Chr(34) & " name=" & Chr(34) & "Microsoft.Windows.Common-Controls" & Chr(34) & " version=" & Chr(34) & "6.0.0.0" & Chr(34) & " processorArchitecture=" & Chr(34) & "X86" & Chr(34) & " publicKeyToken=" & Chr(34) & "6595b64144ccf1df" & Chr(34) & " language=" & Chr(34) & "*" & Chr(34) & " /> </dependentAssembly> </dependency> </assembly>" ' CONTENTS OF THE MANIFEST FILE...
Close #1 '                                  YOU NEED TO HAVE THIS FILE, OR THE THEME WILL NOT WORK!

xptheme = InitCommonControls                        ' IF MANIFEST EXISTS, EXUCUTE CONTROL UPGRADE TO XP THEME STYLE

setAShidden = SetFileAttributes(NEWmanifestpth, FILE_ATTRIBUTE_HIDDEN) ' HIDE THE MANIFEST THEME FILE

Timer1.Enabled = True ' START THE TIMER.... BECAUSE THE MANIFEST HAS JUST BEEN WRITTEN, YOUR PROGRAM NEEDS TO RESTART.. THIS DOES IT FOR YOU

Exit Sub ' SKIP ANYTHING AFTER THIS MARK IN CURRENT SUB

problemARGH: ' IF AN ERROR OCCURED DURING THE CREATION OF THE MANIFEST
MsgBox "Error creating Windows XP theme file. You may be running EXE file from a network drive with which you dont have write permissions. Themes will not be enabled.", vbExclamation, "Themeing Error!" ' TELLING USER THAT THEMES WILL NOT BE ENABLED
End Sub


Private Sub Timer1_Timer()
On Error GoTo error ' ERROR HANDLING
Dim myEXEpath As String ' DELCARE VARIBLES...

Timer1.Enabled = False

Unload Me 'CLOSE DOWN THE APPLICATION FOR RESTART

If Right(App.Path, 1) = "\" Then
    myEXEpath = App.Path & App.EXEName & ".exe"
Else                                                       '| GET THE PATH FOR YOUR APPLICATION
    myEXEpath = App.Path & "\" & App.EXEName & ".exe"
End If

Shell myEXEpath, vbNormalFocus          ' RESTART YOUR APPLICATION. THE THEME SHOULD NOW BE IN EFFECT! ENJOY!
Exit Sub

error:
MsgBox "Error exucuting the EXE file. This would be caused by you trying to compile the manifest file from inside Visual Basic. You can only see the theme when fully compiled, and ran as an .EXE file :)", vbExclamation, "Manifest Exucution Error!"

         '  tHe_cLeanER productions... [ jBistoGOOD@Hotmail.com ]  '
         
End Sub
