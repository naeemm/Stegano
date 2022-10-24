VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Performance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Performance (naeem@email.com)"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   393
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   468
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdl_File 
      Left            =   7320
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   0
      TabIndex        =   3
      Top             =   1920
      Width           =   6975
   End
   Begin VB.PictureBox picImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1095
      Left            =   6840
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   133
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txt_Message 
      Height          =   1335
      Left            =   0
      MaxLength       =   65535
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Performance.frx":0000
      Top             =   120
      Width           =   6975
   End
   Begin VB.Label lbl_Characters 
      Height          =   255
      Left            =   5640
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lbl_Display 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   1560
      Width           =   4695
   End
   Begin VB.Menu mnu_Load_Image 
      Caption         =   "Load Image"
   End
   Begin VB.Menu mnu_Hide 
      Caption         =   "Hide"
   End
   Begin VB.Menu mnu_Show 
      Caption         =   "Show"
   End
   Begin VB.Menu mnu_Calculate 
      Caption         =   "Calculate"
   End
   Begin VB.Menu mnu_Test 
      Caption         =   "Test"
   End
End
Attribute VB_Name = "Performance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '''''''''''''''''''''''' Performacne  '''''''''''''''''''
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
 Dim strD As String, lng_Performance As Long
 Dim lng_Msg As Long
 
 ''''''''''''' performance declarations.....
 Dim Ch_Pixels As Long, UnCh_Pixels As Long
 Dim col_Dummy As New Collection
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ''Formula  for Storage of Characters in BMP
 ''Truncate(answer) = ((width * height * 3 ) - (Target_Pixels+Off_Set)) / Target_Pixels
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 Const Max_Pixels = 8 '' 8 - 30   '' at 7 the data would be stored but noise would be embeded into the image
 Const Off_Set = 30 - Max_Pixels   ' 0 - 22
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
 
 On Error Resume Next
 
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
   
  If (R > R_old) Or (G > G_old) Or (B > B_old) Or (R < R_old) Or (G < G_old) Or (B < B_old) Then
     Ch_Pixels = Ch_Pixels + 1
     str_dd = X & "," & Y
     col_Dummy.Add str_dd, str_dd
  Else
     UnCh_Pixels = UnCh_Pixels + 1
  End If
       
  'strDummy = R & "," & G & "," & B & "," & R_old & "," & G_old & "," & B_old
  'T_Stream.WriteLine strDummy
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
    'lng_Performance = lng_Performance + 1
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

Private Sub mnu_Show_Click()
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
 mnu_Show.Visible = False
End Sub

Private Sub mnu_Hide_Click()
 Dim i As Long, Msg As String, str_Perf As String
    
    str_Perf = lng_Width & "=X," & lng_Height & "=Y,"
    str_Perf = str_Perf & Max_Pixels & "=Pxl/Ch,"
    str_Perf = str_Perf & lng_Msg & "=Max_Ch,"
 Msg = Trim(Left(txt_Message.Text, lng_Msg))
 lng_Msg = Len(Msg)
    str_Perf = str_Perf & lng_Msg & "=Enc_Ch,"
 Initialize_Files ("Performance.txt")
    
 Encode lng_Msg, Max_Pixels + Off_Set
      
 For i = 1 To lng_Msg
  Encode Asc(Mid(txt_Message.Text, i, 1)), Max_Pixels
 Next
  
 picImage.Picture = picImage.Image
 SavePicture picImage.Picture, App.Path & "\new.bmp"
 
 str_Perf = str_Perf & UnCh_Pixels & "=UnCh_Pxl,"
 str_Perf = str_Perf & Ch_Pixels & "=Ch_Pxl,"
 str_Perf = str_Perf & col_Dummy.Count & "=Uniq_Pxl"
 T_Stream.WriteLine str_Perf
   
 End
End Sub

Private Sub Initialize_Files(strFile)
 'Set T_Stream = obj_F.OpenTextFile(App.Path & "\" & strFile, ForWriting, True)
 'T_Stream.Write vbNullString
 'T_Stream.Close
 Set T_Stream = obj_F.OpenTextFile(App.Path & "\" & strFile, ForAppending, True)
End Sub

Private Sub mnu_Calculate_Click()
 Dim A As Variant, strDum As String
 Dim X As Long, Y As Long, PpC As Long
 Dim MxC As Long, EC As Long
 Dim UChgPxl As Long, ChgPxl As Long, UniqPxl As Long
 Dim Sum As Long
 
 Set T_Stream = obj_F.OpenTextFile(App.Path & "\performance.txt", ForReading, True)
 
 While Not T_Stream.AtEndOfStream
  A = Split(T_Stream.ReadLine, ",")
  X = Val(A(0)):         Y = Val(A(1)):       PpC = Val(A(2))
  MxC = Val(A(3)):       EC = Val(A(4))
  UChgPxl = Val(A(5)):   ChgPxl = Val(A(6)):  UniqPxl = Val(A(7))
  
  
  
  
  strDum = X & "x" & Y & " PpC=" & PpC
  strDum = strDum & " MxC=" & MxC & " EC=" & EC
  strDum = strDum & " UChgP=" & UChgPxl & " ChgP=" & ChgPxl & " UniqP=" & UniqPxl
      
  'Sum = CLng(A(0)) + CLng(A(1))
  List1.AddItem "=================================="
  List1.AddItem strDum
 Wend
 T_Stream.Close
 
End Sub


Private Function Validate()
  Dim F As Long, S As Long
  F = (lng_Width / Max_Pixels) * lng_Height * 3
  S = Off_Set / Max_Pixels
  lng_Msg = F - S - 2
  lbl_Display.Caption = "Max Characters = " & lng_Msg
End Function

Private Sub mnu_Load_Image_Click()
 Dim i   As Long, str_File As String
 cdl_File.ShowOpen
 str_File = cdl_File.FileName
 picImage.Picture = LoadPicture(str_File)
 lng_Width = picImage.ScaleWidth
 lng_Height = picImage.ScaleHeight
 Validate
End Sub

Private Sub mnu_Test_Click()
 Dim i As Long
 Initialize_Files ("Numbers.txt")
  For i = 1 To 65535
   T_Stream.WriteLine i
  Next
  
End Sub

Private Sub txt_Message_Change()
 lbl_Characters.Caption = Len(txt_Message.Text)
End Sub
