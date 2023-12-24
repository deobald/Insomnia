VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmAddStr 
   Caption         =   "Add a String to an Image"
   ClientHeight    =   4770
   ClientLeft      =   2520
   ClientTop       =   2070
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   6405
   Begin VB.CommandButton btnSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   1350
      TabIndex        =   6
      Top             =   75
      Width           =   765
   End
   Begin VB.PictureBox picEncrypt 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   4050
      ScaleHeight     =   1065
      ScaleWidth      =   1290
      TabIndex        =   5
      Top             =   975
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.CommandButton btnOpen 
      Caption         =   "Open"
      Height          =   315
      Left            =   600
      TabIndex        =   4
      ToolTipText     =   "Open"
      Top             =   75
      Width           =   765
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.CommandButton btnDecrypt 
      Caption         =   "Open and Decrypt"
      Height          =   390
      Left            =   450
      TabIndex        =   3
      Top             =   2625
      Width           =   1890
   End
   Begin VB.PictureBox picDecrypt 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   975
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   2
      Top             =   3150
      Width           =   1215
   End
   Begin VB.CommandButton btnEncrypt 
      Caption         =   "Encrypt"
      Height          =   390
      Left            =   2100
      TabIndex        =   1
      Top             =   1350
      Width           =   1215
   End
   Begin VB.PictureBox picView 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   750
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   0
      Top             =   1125
      Width           =   750
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   150
      X2              =   6225
      Y1              =   2325
      Y2              =   2325
   End
End
Attribute VB_Name = "frmAddStr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PicOrig As tArea

Private Sub Inits()
Dim Dummy

PicOrig.hDC = 0
PicOrig.Left = 0
PicOrig.Top = 0
PicOrig.Width = 32
PicOrig.Height = 32
PicOrig.hDC = CreateMemHdc(picView.hDC, 1600, 1600)

End Sub

Private Sub btnDecrypt_Click()
Dim Dummy

frmAddStr.MousePointer = 11 'Hourglass

EncryptImg.hDC = EncryptImg.hDC - 27
Dummy = BitBlt(picDecrypt.hDC, 0, 0, picView.Width, picView.Height, EncryptImg.hDC, 0, 0, SRCCOPY)

frmAddStr.MousePointer = 0

End Sub

Private Sub btnEncrypt_Click()
Dim Dummy

btnEncrypt.Enabled = False
frmAddStr.MousePointer = 11 'Hourglass

Dummy = BitBlt(picEncrypt.hDC, 0, 0, picView.Width, picView.Height, picView.hDC, 0, 0, SRCCOPY)
picEncrypt.hDC = picEncrypt.hDC + 27

frmAddStr.MousePointer = 0

End Sub


Private Sub btnOpen_Click()

Dim ErrHandler
CommonDialog.CancelError = True
On Error GoTo ErrHandler

CommonDialog.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
CommonDialog.Filter = "Windows Bitmap (*.BMP)|*.BMP|All Files (*.*)|*.*"
CommonDialog.FilterIndex = 1 'Set to .BMP as default
CommonDialog.DialogTitle = "Open Bitmap"
CommonDialog.ShowOpen 'Action = 1 'Open a "Open File" box
ImageFile = CommonDialog.filename 'MapFile equals the filename chosen

picView.Picture = LoadPicture(ImageFile)
PicOrig.Width = picView.Width
PicOrig.Height = picView.Height
picView.Width = 200
picView.Height = 200
Call LoadBmpToHdc(PicOrig.hDC, ImageFile)

picMask.Width = PicOrig.Width
picMask.Height = PicOrig.Height

btnMakeMask.Enabled = True

ErrHandler:
    Exit Sub

End Sub

Private Sub btnSave_Click()
Dim PalSet

Dim ErrHandler
CommonDialog.CancelError = True
On Error GoTo ErrHandler

CommonDialog.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
CommonDialog.Filter = "Windows Bitmap (*.BMP)|*.BMP|All Files (*.*)|*.*"
CommonDialog.FilterIndex = 1 'Set to .BMP as default
CommonDialog.DialogTitle = "Save Mask As..."
CommonDialog.ShowSave 'Action = 1 'Open a "Save" box
ImageFile = CommonDialog.filename 'File equals the filename chosen

Me.MousePointer = 11
picEncrypt.Picture = picEncrypt.Image

'[Set paths to game location on Hard Disk]
Dim Path
If (Right(App.Path, 1) <> "\") Then
    Path = App.Path & "\"
        Else
Path = App.Path
End If
'[End of Set Paths]

SavePicture picEncrypt.Picture, ImageFile
Me.MousePointer = 0

MsgBox "Mask is finished saving. It has been saved in Windows standard," + HR + " 16-Bit (64,000 color) format. It is advised that you convert it to" + HR + " your preferred format in an image-editing tool. Thank you.", vbOKOnly, "Mask Saved to Disk"

btnSave.Enabled = False

ErrHandler:
    Exit Sub

End Sub

Private Sub Form_Load()
    Call Inits
End Sub
