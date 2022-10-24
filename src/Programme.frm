VERSION 5.00
Begin VB.Form Programme 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Programme (naeem@email.com)"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   154
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   6360
      TabIndex        =   6
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton cmd_Calculate 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmd_Show 
      Caption         =   "Show"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
   End
   Begin VB.PictureBox picImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1095
      Left            =   6840
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   133
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txt_Message 
      Height          =   1335
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Programme.frx":0000
      Top             =   120
      Width           =   6255
   End
   Begin VB.CommandButton cmd_Hide 
      Caption         =   "Hide"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lbl_Display 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   6135
   End
End
Attribute VB_Name = "Programme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''' Programme  '''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
 Dim col_Target_Pixels As New Collection
 Dim T_Stream As TextStream, obj_F As New FileSystemObject
 Dim lng_Width As Long, lng_Height As Long
 Dim i As Integer
 Dim byte_Mask As Long, color_Mask As Integer
 Dim X As Integer, Y As Integer, pixel As Integer
 Dim strDummy As String
 Dim R As Byte, G As Byte, B As Byte
 Dim R_old As Byte, G_old As Byte, B_old As Byte
 Dim strD As String
 
 Const Max_Pixels = 8 '' 8 - 30   '' at 7 the data would be stored but noise would be embeded into the image
 Const Off_Set = 22 ' 0 - 22
 Const Dec_Min = 1 '&H1
  
  ' 256 -  2 = 254  , max_words=255 , no_of_pxls = 8
  ' 512 -  2 = 510  , max_words=511 , no_of_pxls = 8,9
  ' 1024 - 2 = 1022 , max_words=1023 , no_of_pxls = 10
  ' 2048 - 2 = 2046 , max_words= 2047 , no_of_pxls = 11
  ' 4096 - 2 = 4094 , max_words= 4097 , no_of_pxls = 12
  ' 8192 - 2 = 8190 , max_words= 8191 , no_of_pxls = 13
  ' 16384 - 2 = 16382 , max_words=  , no_of_pxls = 14
  ' 32768 - 2 = 32766 , max_words= 32767 , no_of_pxls = 15
  ' 65536 - 2 = 65534 , max_words= 65535 , no_of_pxls = 16
  ' 1073741824 - 2 =  1073741822, max_words = 1073741823 , no_of_pixels = 3
  
Private Function Encode(Value As Long, M_Pxls As Byte)
 Dim Enc_Max As Long
 
 Enc_Max = (2 ^ M_Pxls) - 2
 byte_Mask = 1
 For i = 1 To M_Pxls
     
   PickPosition col_Target_Pixels, lng_Width, lng_Height, X, Y, pixel
   Resolve_Color picImage.Point(X, Y), R, G, B
   R_old = R: G_old = G: B_old = B
     
  If Value And byte_Mask Then
    color_Mask = 1
   Else
    color_Mask = 0
  End If
    
  Select Case pixel
   Case 0
    R = (R And Enc_Max) Or color_Mask
   Case 1
    G = (G And Enc_Max) Or color_Mask
   Case 2
    B = (B And Enc_Max) Or color_Mask
  End Select

  'strDummy = x & " x " & y & " , " & pixel & " (" & R & "," & G & "," & B & ")"
  strDummy = R & "," & G & "," & B & "," & R_old & "," & G_old & "," & B_old
  T_Stream.WriteLine strDummy
  'Debug.Print strDummy
  
  picImage.PSet (X, Y), RGB(R, G, B) 'new colour
  byte_Mask = byte_Mask * 2
  
 Next
 'T_Stream.WriteLine "--------"

End Function

Private Sub PickPosition(ByVal Target_Pxls As Collection, ByVal int_Width As Integer, ByVal int_Height As Integer, ByRef Row As Integer, ByRef Col As Integer, ByRef Pxl As Integer)

Dim str_Temp As String, Run As Boolean
Run = True
On Error Resume Next

While Run
    Row = Int(Rnd * int_Width)
    Col = Int(Rnd * int_Height)
    Pxl = Int(Rnd * 3)
    
    str_Temp = Row & "," & Col & "," & Pxl
    Target_Pxls.Add str_Temp, str_Temp
    Run = Err.Number
    Err.Clear
Wend
End Sub


Private Sub Resolve_Color(ByVal Color As OLE_COLOR, ByRef R As Byte, ByRef G As Byte, ByRef B As Byte)
    R = Color And &HFF&
    G = (Color And &HFF00&) \ &H100&
    B = (Color And &HFF0000) \ &H10000
End Sub

Private Function Decode(M_Pxls As Byte) As Long

Dim Value As Long

byte_Mask = 1
For i = 1 To M_Pxls
 
 PickPosition col_Target_Pixels, lng_Width, lng_Height, X, Y, pixel
 Resolve_Color picImage.Point(X, Y), R, G, B
   
Select Case pixel
Case 0
    color_Mask = (R And Dec_Min)
Case 1
    color_Mask = (G And Dec_Min)
Case 2
    color_Mask = (B And Dec_Min)
End Select

If color_Mask Then
    Value = Value Or byte_Mask
End If

'strDummy = x & "," & y & "," & pixel & "," & R & "," & G & "," & B
'T_Stream.WriteLine strDummy
'Debug.Print strDummy
'Debug.Print Value

byte_Mask = byte_Mask * 2

Next i
'T_Stream.WriteLine "===================="
Decode = Value

End Function

Private Sub cmd_Show_Click()
 Dim L As Long, str_D As String, i As Long
 picImage.Picture = LoadPicture(App.Path & "\new.bmp")
 lng_Width = picImage.ScaleWidth
 lng_Height = picImage.ScaleHeight
 'Initialize_Files ("dec_data.txt")
 L = Decode(Max_Pixels + Off_Set)

 For i = 1 To L
  str_D = str_D & Chr(Decode(Max_Pixels))
 Next
 txt_Message.Text = str_D
 cmd_Show.Visible = False
End Sub

Private Sub cmd_Hide_Click()
 Dim lng_Msg As Long, i   As Long
 picImage.Picture = LoadPicture(App.Path & "\old.jpg")
 lng_Width = picImage.ScaleWidth
 lng_Height = picImage.ScaleHeight
 
 lng_Msg = Len(txt_Message.Text)
 Initialize_Files ("enc_data.txt")
 
 Encode lng_Msg, Max_Pixels + Off_Set
 For i = 1 To lng_Msg
  Encode Asc(Mid(txt_Message.Text, i, 1)), Max_Pixels
 Next
  
 picImage.Picture = picImage.Image
 SavePicture picImage.Picture, App.Path & "\new.bmp"
 End
End Sub

Private Sub Initialize_Files(strFile)
 Set T_Stream = obj_F.OpenTextFile(App.Path & "\" & strFile, ForWriting, True)
 T_Stream.Write vbNullString
 T_Stream.Close
 Set T_Stream = obj_F.OpenTextFile(App.Path & "\" & strFile, ForAppending, True)
End Sub

Private Sub cmd_Calculate_Click()
 Dim A As Variant, strDum As String
 Dim R As Integer, G As Integer, B As Integer
 Dim R1 As Integer, G1 As Integer, B1 As Integer
 Dim R2 As Integer, G2 As Integer, B2 As Integer
 Dim Sum As Long, Change As Long, UnChange As Long
    
   
 Set T_Stream = obj_F.OpenTextFile(App.Path & "\enc_data.txt", ForReading, True)
 While Not T_Stream.AtEndOfStream
  A = Split(T_Stream.ReadLine, ",")
  R1 = A(0): G1 = A(1): B1 = A(2)
  R2 = A(3): G2 = A(4): B2 = A(5)
  R = R1 - R2: G = G1 - G2: B = B1 - B2
  strDum = "(" & R1 & "," & G1 & "," & B1 & ")"
  strDum = strDum & " - " & "(" & R2 & "," & G2 & "," & B2 & ")"
  strDum = strDum & " - " & "(" & R & "," & G & "," & B & ")"
  Sum = R + G + B
  strDum = strDum & " Sum = " & Sum
  If Sum = 0 Then
    UnChange = UnChange + 1
  Else
   Change = Change + 1
  End If
  'List1.AddItem strDum
 Wend
 
  Sum = Change + UnChange
  'List1.AddItem "=================================="
  List1.AddItem "Change= " & Change & " ( " & 100 * (Change / Sum) & " % ) "
  List1.AddItem "UnChange= " & UnChange & " ( " & 100 * (UnChange / Sum) & " % ) "
  
 T_Stream.Close
 
End Sub

Private Sub txt_Message_Change()
 lbl_Display.Caption = Len(txt_Message.Text)
End Sub
