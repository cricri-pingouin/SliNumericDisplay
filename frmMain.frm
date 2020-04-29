VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LED display"
   ClientHeight    =   3144
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5412
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3144
   ScaleWidth      =   5412
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optColour 
      BackColor       =   &H00404040&
      Caption         =   "Blue"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   9
      Top             =   2040
      Width           =   615
   End
   Begin VB.OptionButton optColour 
      BackColor       =   &H00404040&
      Caption         =   "Red"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   7
      Top             =   2040
      Width           =   615
   End
   Begin VB.OptionButton optColour 
      BackColor       =   &H00404040&
      Caption         =   "Green"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   6
      Top             =   2040
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdMinus 
      Caption         =   "-"
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdPlus 
      Caption         =   "+"
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display"
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox txtNumber 
      Height          =   285
      Left            =   2760
      TabIndex        =   0
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label lblColour 
      BackColor       =   &H00404040&
      Caption         =   "Choose your display LEDs colour:"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label lblDigits 
      BackColor       =   &H00404040&
      Caption         =   "Use these buttons to set the number of digits to display:"
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   3975
   End
   Begin VB.Label lblNumber 
      BackColor       =   &H00404040&
      Caption         =   "Enter the 3 digits number to display:"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   2655
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'12 sprites HDCs: 4 for green (horizontal on/off, vertical on /off), 4 for red and same for blue
Dim SpritesHDC(12) As Long
'Number of digits allowed (user selected)
'Byte because it varies from 1 to 9
Dim Digits As Byte
'Array containing LEDs digits, coded as shown below
'Byte is enough as there are 7 digits only
'Minimum value is 80 for a "1", and maximum value is 127 for a "8"
Dim NumbersDigits(10) As Byte
'Width and height of a led
'Have to declare as Integer (and not byte) because of sprite functions reference type
Dim LedWidth, LedHeight As Integer

'The following numbering for the digits is used (powers of 2):
'
'    1
'  8   16
'    2
' 32   64
'    4
'
'See LEDs initialisation in the Load event if you don't understand

Private Sub Form_Load()
   'Initialise LEDs width and height, i.e. bitmap image size in pixels
   'These values are for horizontal LEDs
   'For vertical LEDs, these values will be inverted
   LedWidth = 40
   LedHeight = 10
   'Load sprites in memory
   'I HAD to use 40 and 10 instead of LedWidth and LedHeight
   'in the following HDCs and sprites declarations.
   'I've got no idea why, I'd appreciate any help on that :)
   SpritesHDC(0) = CreateMemHdc(Me.hdc, 40, 10)
   Call LoadBmpToHdc(SpritesHDC(0), "GHOff.bmp") 'Green, horizontal, off
   SpritesHDC(1) = CreateMemHdc(Me.hdc, 40, 10)
   Call LoadBmpToHdc(SpritesHDC(1), "GHOn.bmp")  'Green, horizontal, on
   SpritesHDC(2) = CreateMemHdc(Me.hdc, 10, 40)
   Call LoadBmpToHdc(SpritesHDC(2), "GVOff.bmp") 'Green, vertical, off
   SpritesHDC(3) = CreateMemHdc(Me.hdc, 10, 40)
   Call LoadBmpToHdc(SpritesHDC(3), "GVOn.bmp")  'Green, vertical, on
   SpritesHDC(4) = CreateMemHdc(Me.hdc, 40, 10)
   Call LoadBmpToHdc(SpritesHDC(4), "RHOff.bmp") 'Red, horizontal, off
   SpritesHDC(5) = CreateMemHdc(Me.hdc, 40, 10)
   Call LoadBmpToHdc(SpritesHDC(5), "RHOn.bmp")  'Red, horizontal, on
   SpritesHDC(6) = CreateMemHdc(Me.hdc, 10, 40)
   Call LoadBmpToHdc(SpritesHDC(6), "RVOff.bmp") 'Red, vertical, off
   SpritesHDC(7) = CreateMemHdc(Me.hdc, 10, 40)
   Call LoadBmpToHdc(SpritesHDC(7), "RVOn.bmp")  'Red, vertical, on
   SpritesHDC(8) = CreateMemHdc(Me.hdc, 40, 10)
   Call LoadBmpToHdc(SpritesHDC(8), "BHOff.bmp") 'Blue, horizontal, off
   SpritesHDC(9) = CreateMemHdc(Me.hdc, 40, 10)
   Call LoadBmpToHdc(SpritesHDC(9), "BHOn.bmp")  'Blue, horizontal, on
   SpritesHDC(10) = CreateMemHdc(Me.hdc, 10, 40)
   Call LoadBmpToHdc(SpritesHDC(10), "BVOff.bmp") 'Blue, vertical, off
   SpritesHDC(11) = CreateMemHdc(Me.hdc, 10, 40)
   Call LoadBmpToHdc(SpritesHDC(11), "BVOn.bmp")  'Blue, vertical, on
   'Determine LEDs to switch on for the different numbers 0 to 9
   NumbersDigits(0) = 125 '= 1 + 8 + 16 + 32 + 64 + 4
   NumbersDigits(1) = 80  '= 16 + 64
   NumbersDigits(2) = 55  '= 1 + 16 + 2 + 32 + 4
   NumbersDigits(3) = 87  '= 1 + 16 + 2 + 64 + 4
   NumbersDigits(4) = 90  '= 8 + 16 + 2 + 64
   NumbersDigits(5) = 79  '= 1 + 8 + 2 + 64 + 4
   NumbersDigits(6) = 111 '= 1 + 8 + 2 + 32 + 64 + 4 (substract 1 to remove top bar if you wish)
   NumbersDigits(7) = 81  '= 1 + 16 + 64
   NumbersDigits(8) = 127 '= 1 + 8 + 16 + 2 + 32 + 64 + 4
   NumbersDigits(9) = 95  '= 1 + 8 + 16 + 2 + 64 + 4 (substract 4 to remove bottom bar if you wish)
   'Set default digits number to 3. Change it to what you need.
   Digits = 3
   'Update label
   'Set by default on the label properties
   'Uncomment if you change default value above
   'lblNumber.Caption = "Enter the " + Str(Digits) + " number to display:"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'IMPORTANT! Release all created hDCs to prevent memory leaks!
   Call DestroyHdcs
   'And exit
   Unload Me
   Set frmMain = Nothing
End Sub

Private Sub cmdPlus_Click()
   'Allow maximum of 9 digits
   If Digits < 9 Then
      Digits = Digits + 1
   Else
      Exit Sub
   End If
   Call UpdateWithNewValue
End Sub

Private Sub cmdMinus_Click()
   'Allow minimum of 1 digits
   If Digits > 1 Then
      Digits = Digits - 1
   Else
      Exit Sub
   End If
   Call UpdateWithNewValue
End Sub

Private Sub UpdateWithNewValue()
   'Update label
   lblNumber.Caption = "Enter the " + Str(Digits) + " digits number to display:"
   'Redimension form (if 3 digits or more, else would lose buttons)
   If Digits > 2 Then Me.Width = 1000 + 1500 * Digits
End Sub

Private Sub cmdDisplay_Click()
   'Loops, byte is enough
   Dim i As Byte
   'Led colour: 0=green, 1=red, 2=blue
   'Thus, byte is enough
   Dim LedColour As Byte
   'Value for current digit to process, from 0 to 9 => byte is enough
   Dim DigitValue As Byte
   'Integer to be displayed. Use long because integer is not enough for 9 digits
   Dim NumberToDisplay As Long
   
   'Check if number contains "Digits" number of characters
   If Len(txtNumber.Text) <> Digits Then
      MsgBox "You MUST enter a" + Str(Digits) + " digits number.", vbExclamation, "Wrong size number!"
      txtNumber.Text = ""
      Exit Sub
   End If
   'Check if string contains numbers ONLY
   For i = 1 To Digits
      '0: Asc=48 ; 9: Asc=57
      If Asc(Mid(txtNumber.Text, i, 1)) < 48 Or Asc(Mid(txtNumber.Text, i, 1)) > 57 Then
         MsgBox "A number must consist in... numbers ONLY!", vbExclamation, "Invalid number!"
         txtNumber.Text = ""
         Exit Sub
      End If
   Next
   'String is a valid number: convert it to a long
   'BANG YOUR CHECKED NUMBER INTO NumberToDisplay HERE TO
   'ADAPT THIS CODE TO YOUR NEEDS. Piece of cake ain't it? :)
   NumberToDisplay = Int(txtNumber.Text)
   'Set LEDs colour according to option selected: 0=green, 1=red, 2=blue
   'You should use a "Case Select" if you want to use more sprites sets
   If optColour(0).Value Then
      LedColour = 0
   ElseIf optColour(1).Value Then
      LedColour = 1
   Else
      LedColour = 2
   End If
   'Clear form
   Me.Cls
   'Process digits one by one, starting from "higher figure"
   For i = 1 To Digits
      'Select only the biggest figure and store it
      DigitValue = Int(NumberToDisplay / 10 ^ (Digits - i))
      'Remove this figure from the number to display
      NumberToDisplay = NumberToDisplay - DigitValue * 10 ^ (Digits - i)
      'Check LEDs individually
      'Can't use a loop because the position is completely independent from the index
      'LED 1 (horizontal, top). Check if it has to be ON using "And 1) = 1"
      If (NumbersDigits(DigitValue) And 1) = 1 Then
         'Display horizontal on LED
         Call DisplayDigit(i * (LedWidth + 6 * LedHeight) - 40, 20, LedWidth, LedHeight, LedColour * 4 + 1)
      Else
         'Display horizontal off LED
         Call DisplayDigit(i * (LedWidth + 6 * LedHeight) - 40, 20, LedWidth, LedHeight, LedColour * 4)
      End If
      'LED 2 (horizontal, middle)
      If (NumbersDigits(DigitValue) And 2) = 2 Then
         'Display horizontal on LED
         Call DisplayDigit(i * (LedWidth + 6 * LedHeight) - 40, 20 + LedWidth + 0.25 * LedHeight, LedWidth, LedHeight, LedColour * 4 + 1)
      Else
         'Display horizontal off LED
         Call DisplayDigit(i * (LedWidth + 6 * LedHeight) - 40, 20 + LedWidth + 0.25 * LedHeight, LedWidth, LedHeight, LedColour * 4)
      End If
      'LED 3 (horizontal, bottom)
      If (NumbersDigits(DigitValue) And 4) = 4 Then
         'Display horizontal on LED
         Call DisplayDigit(i * (LedWidth + 6 * LedHeight) - 40, 20 + 2 * LedWidth + 0.5 * LedHeight, LedWidth, LedHeight, LedColour * 4 + 1)
      Else
         'Display horizontal off LED
         Call DisplayDigit(i * (LedWidth + 6 * LedHeight) - 40, 20 + 2 * LedWidth + 0.5 * LedHeight, LedWidth, LedHeight, LedColour * 4)
      End If
      'LED 4 (vertical, top left)
      If (NumbersDigits(DigitValue) And 8) = 8 Then
         'Display vertical on LED
         Call DisplayDigit(i * (LedWidth + 6 * LedHeight) - 40 - 1.5 * LedHeight, 20 + 0.5 * LedHeight, LedHeight, LedWidth, LedColour * 4 + 3)
      Else
         'Display vertical off LED
         Call DisplayDigit(i * (LedWidth + 6 * LedHeight) - 40 - 1.5 * LedHeight, 20 + 0.5 * LedHeight, LedHeight, LedWidth, LedColour * 4 + 2)
      End If
      'LED 5 (vertical, top right)
      If (NumbersDigits(DigitValue) And 16) = 16 Then
         'Display vertical on LED
         Call DisplayDigit(i * (LedWidth + 6 * LedHeight) - 40 + 0.5 * LedHeight + LedWidth, 20 + 0.5 * LedHeight, LedHeight, LedWidth, LedColour * 4 + 3)
      Else
         'Display vertical off LED
         Call DisplayDigit(i * (LedWidth + 6 * LedHeight) - 40 + 0.5 * LedHeight + LedWidth, 20 + 0.5 * LedHeight, LedHeight, LedWidth, LedColour * 4 + 2)
      End If
      'LED 6 (vertical, bottom left)
      If (NumbersDigits(DigitValue) And 32) = 32 Then
         'Display vertical on LED
         Call DisplayDigit(i * (LedWidth + 6 * LedHeight) - 40 - 1.5 * LedHeight, 20 + LedHeight + LedWidth, LedHeight, LedWidth, LedColour * 4 + 3)
      Else
         'Display vertical off LED
         Call DisplayDigit(i * (LedWidth + 6 * LedHeight) - 40 - 1.5 * LedHeight, 20 + LedHeight + LedWidth, LedHeight, LedWidth, LedColour * 4 + 2)
      End If
      'LED 7 (vertical, bottom right)
      If (NumbersDigits(DigitValue) And 64) = 64 Then
         'Display vertical on LED
         Call DisplayDigit(i * (LedWidth + 6 * LedHeight) - 40 + 0.5 * LedHeight + LedWidth, 20 + LedHeight + LedWidth, LedHeight, LedWidth, LedColour * 4 + 3)
      Else
         'Display vertical off LED
         Call DisplayDigit(i * (LedWidth + 6 * LedHeight) - 40 + 0.5 * LedHeight + LedWidth, 20 + LedHeight + LedWidth, LedHeight, LedWidth, LedColour * 4 + 2)
      End If
   Next
   'Refresh to show updated display: the lame and easy way :) + AutoRedraw! Blech!
   Me.Refresh
End Sub

Private Sub DisplayDigit(PosX, PosY, LedWidth, LedHeight, SpriteNumber As Byte)
   'Display sprite SpriteNumber at coordinates PosX,PosY on form (Me.hdc)
   'Also use LedWidth and LedHeight here, so it's 100% flexible
   'SpriteNumber specifies which LED to display in terms of colour and orientation
   BitBlt Me.hdc, PosX, PosY, LedWidth, LedHeight, SpritesHDC(SpriteNumber), 0, 0, vbSrcCopy
End Sub
