VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   256
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTest 
      Caption         =   "cmdTest"
      Height          =   465
      Left            =   3660
      TabIndex        =   0
      Top             =   3360
      Width           =   1125
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The following API calls are for:

'blitting
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'code timer
Private Declare Function GetTickCount Lib "kernel32" () As Long

'creating buffers / loading sprites
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

'loading sprites
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

'cleanup
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

'end of copy-paste here...

'our Buffer's DC
Public myBackBuffer As Long
Public myBufferBMP As Long

'The DC of our sprite/graphic
Public mySprite As Long

'coordinates of our sprite/graphic on the screen
Public SpriteX As Long
Public SpriteY As Long

'end of copy-paste here...


Public Function LoadGraphicDC(sFileName As String) As Long
'cheap error handling
On Error Resume Next

'temp variable to hold our DC address
Dim LoadGraphicDCTEMP As Long

'create the DC address compatible with
'the DC of the screen
LoadGraphicDCTEMP = CreateCompatibleDC(GetDC(0))

'load the graphic file into the DC...
SelectObject LoadGraphicDCTEMP, LoadPicture(sFileName)

'return the address of the file
LoadGraphicDC = LoadGraphicDCTEMP
End Function

Private Sub cmdTest_Click()
'=== THIS CODE GOES IN CMDTEXT_CLICK EVENT ===

'Timer variables...
Dim T1 As Long, T2 As Long

'create a compatable DC for the back buffer..
myBackBuffer = CreateCompatibleDC(GetDC(0))

'create a compatible bitmap surface for the DC
'that is the size of our form.. (320 X 256)
'NOTE - the bitmap will act as the actual graphics surface inside the DC
'because without a bitmap in the DC, the DC cannot hold graphical data..

myBufferBMP = CreateCompatibleBitmap(GetDC(0), 320, 256)

'final step of making the back buffer...
'load our created blank bitmap surface into our buffer
'(this will be used as our canvas to draw-on off screen)
SelectObject myBackBuffer, myBufferBMP

'before we can blit to the buffer, we should fill it with black
BitBlt myBackBuffer, 0, 0, 320, 256, 0, 0, 0, vbWhiteness

'load our sprite (using the function we made)
mySprite = LoadGraphicDC(App.Path & "\sprite1.bmp")
'MsgBox Dir$(App.Path & "\sprite1.bmp")
'ok now all the graphics are loaded so
'lets start our main loop..

'Disable cmdTest, because if the graphics are
'reloaded there will be memory leaks...
cmdTest.Enabled = False

'== START MAIN LOOP ==
'get current tickcount (this is used as a code timer)
T2 = GetTickCount
Do
DoEvents 'DoEvents makes sure that our mouse and keyboard dont freeze-up
T1 = GetTickCount

'if 15MS has gone by, execute our next frame
If (T1 - T2) >= 15 Then

'clear the place where the sprite used to be...
'(we do this by filling in the old sprites place
'with black... but in games you'll probably have
'a background tile that you would blit here)

BitBlt myBackBuffer, SpriteX - 1, SpriteY - 1, _
32, 32, 0, 0, 0, vbBlackness

'blit sprites to the back-buffer ***
'You could blit multiple sprites to the backbuffer,
'but in our example we only blit on...
BitBlt myBackBuffer, SpriteX, SpriteY, 32, 32, _
mySprite, 0, 0, vbSrcPaint

'now blit the backbuffer to the form...
BitBlt Me.hdc, 0, 0, 320, 256, myBackBuffer, _
0, 0, vbSrcCopy

'move our sprite down on a diagonal...
'Me.Caption = SpriteX & ", " & SpriteY
SpriteX = SpriteX + 1
SpriteY = SpriteY + 1

'update timer
T2 = GetTickCount
End If

'loop it until our sprite is off the screen...
Loop Until SpriteX = 320

'end of copy-paste here...
End Sub

Private Sub Form_Unload(Cancel As Integer)
'this clears up the memory we used to hold
'the graphics and the buffers we made

'Delete the bitmap surface that was in the backbuffer
DeleteObject myBufferBMP

'Delete the backbuffer HDC
DeleteDC myBackBuffer

'Delete the Sprite/Graphic HDC
DeleteDC mySprite
End
End Sub

