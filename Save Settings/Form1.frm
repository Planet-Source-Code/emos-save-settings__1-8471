VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Save Settings by EmoS"
   ClientHeight    =   1335
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   3165
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1335
   ScaleWidth      =   3165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_Exit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   975
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Resize the form then click on the Exit Button.. Then Load the exe and the height, width and its position will be saved"
      Height          =   795
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3195
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Exit_Click()
Rem: ____________the next lines save the values_____________
'(1) saves the settings
'(2) tells windows where to save it to
'(3) section where it save the values
'(4) sub-section where it saves the values
'(5) the values you want to save

Rem: NOTE: _
THE SECTIONS AND THE SUB-SECTIONS MUST BE THE SAME IN THE _
SAVESETTING AND GETSETTING OR YOU WONT GET THE VALUES

Rem:    (1)                 (2)                (3)              (4)              (5)
    SaveSetting App.EXEName, "Options", "My Height", Me.Height
    SaveSetting App.EXEName, "Options", "My Width", Me.Width
    SaveSetting App.EXEName, "Options", "My Top", Me.Top
    SaveSetting App.EXEName, "Options", "My Left", Me.Left
End
End Sub

Private Sub Form_Load()
Rem:_________The Next Line Sets Up The String__________________________
Dim Height As Long, Width As Long, Top As Long, Left As Long, Caption As Long


Rem:_________The Next Lines Place A Value to the String_______________
'(1) is the name of the string
'(2) gets the value from the Exe
'(3) tells windows to look in the exe for the values
'(4) is the section where the information gets saved
'(5) is the sub section where the information gets saved
'(6) is the default values incase they dont get saved when you las unloaded the EXE
Rem: NOTE: _
THE SECTION AND THE SUB-SECTION MUST BE THE SAME IN THE _
SAVESETTING AND GETSETTING OR YOU WONT GET THE VALUES

Rem:   (1)              (2)                (3)                 (4)             (5)            (6)
        Height = GetSetting(App.EXEName, "Options", "My Height", 1710)
        Width = GetSetting(App.EXEName, "Options", "My Width", 3285)
        Top = GetSetting(App.EXEName, "Options", "My Top", 500)
        Left = GetSetting(App.EXEName, "Options", "My Left", 500)
        
Rem: ___sets the values if the string to the form___________________________
            Me.Height = Height
            Me.Width = Width
            Me.Top = Top
            Me.Left = Left
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Rem: ____________the next lines save the values_____________
'(1) saves the settings
'(2) tells windows where to save it to
'(3) section where it save the values
'(4) sub-section where it saves the values
'(5) the values you want to save

Rem: NOTE: _
THE SECTIONS AND THE SUB-SECTIONS MUST BE THE SAME IN THE _
SAVESETTING AND GETSETTING OR YOU WONT GET THE VALUES

Rem:    (1)                 (2)                (3)              (4)              (5)
    SaveSetting App.EXEName, "Options", "My Height", Me.Height
    SaveSetting App.EXEName, "Options", "My Width", Me.Width
    SaveSetting App.EXEName, "Options", "My Top", Me.Top
    SaveSetting App.EXEName, "Options", "My Left", Me.Left
End Sub

Private Sub Form_Terminate()
Rem: ____________the next lines save the values_____________
'(1) saves the settings
'(2) tells windows where to save it to
'(3) section where it save the values
'(4) sub-section where it saves the values
'(5) the values you want to save

Rem: NOTE: _
THE SECTIONS AND THE SUB-SECTIONS MUST BE THE SAME IN THE _
SAVESETTING AND GETSETTING OR YOU WONT GET THE VALUES

Rem:    (1)                 (2)                (3)              (4)              (5)

    SaveSetting App.EXEName, "Options", "My Height", Me.Height
    SaveSetting App.EXEName, "Options", "My Width", Me.Width
    SaveSetting App.EXEName, "Options", "My Top", Me.Top
    SaveSetting App.EXEName, "Options", "My Left", Me.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
Rem: ____________the next lines save the values_____________
'(1) saves the settings
'(2) tells windows where to save it to
'(3) section where it save the values
'(4) sub-section where it saves the values
'(5) the values you want to save

Rem: NOTE: _
THE SECTIONS AND THE SUB-SECTIONS MUST BE THE SAME IN THE _
SAVESETTING AND GETSETTING OR YOU WONT GET THE VALUES

Rem:    (1)                 (2)                (3)              (4)              (5)
    SaveSetting App.EXEName, "Options", "My Height", Me.Height
    SaveSetting App.EXEName, "Options", "My Width", Me.Width
    SaveSetting App.EXEName, "Options", "My Top", Me.Top
    SaveSetting App.EXEName, "Options", "My Left", Me.Left
End Sub
