Attribute VB_Name = "Inpainting"
Option Explicit
'Real Image InPainting
'VB Version 2009 by Scythe
'Thanks goes to Qiushuang Zhang
'Who made the Original as C++ Source
Private Type Bitmap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Type RGBQUAD
    rgbBlue As Byte
    rgbgreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private PicInfo As Bitmap
Private PicAr1() As RGBQUAD
Public StopIt As Boolean 'Stop the inpainting

Private Type gradient
    grad_x As Double
    grad_y As Double
End Type 'the structure that record the gradient

Private Type norm
    norm_x As Double
    norm_y As Double
End Type ' the structure that record the norm

    Const Source = 0
Private Winsize As Long
Private m_width As Long ' image width
Private m_height As Long ' image height
Private m_color() As RGBQUAD
Private m_r() As Double
Private m_g() As Double
Private m_b() As Double
Private m_top As Integer ' the rectangle of inpaint area
Private m_bottom As Integer
Private m_left As Integer
Private m_right As Integer
Private m_mark() As Integer ' mark it as source(0) or to-be-inpainted area(-1) or bondary(-2).
Private m_confid() As Double ' record the confidence for every pixel
Private m_pri() As Double ' record the priority for pixels. only boudary pixels will be used
Private m_gray() As Double ' the gray image
Private m_source() As Boolean ' whether this pixel can be used as an example texture center
Private PatchL As Long
Private PatchR As Long
Private PatchT As Long
Private PatchB As Long


Public Function DoInPaint(InPicture As PictureBox, ResultPicture As PictureBox, MaskRed As Byte, MaskGreen As Byte, MaskBlue As Byte, Optional Preview As Boolean = False, Optional BlockSize As Long = 4, Optional BorderSize As Long = 10000) As Long

Dim BufSize As Long
Dim X As Long
Dim Y As Long
Dim j As Long
Dim i As Long
Dim Count As Long
'Get the Picture
    Pic2Array InPicture, PicAr1
'Fill our Used Variables
    m_width = UBound(PicAr1, 1) + 1
    m_height = UBound(PicAr1, 2) + 1
    Winsize = BlockSize
    BufSize = m_width
    BufSize = BufSize * m_height - 1
    ReDim m_mark(BufSize)
    ReDim m_confid(BufSize)
    ReDim m_pri(BufSize)
    ReDim m_gray(BufSize)
    ReDim m_source(BufSize)
    ReDim m_color(BufSize)
    ReDim m_r(BufSize)
    ReDim m_g(BufSize)
    ReDim m_b(BufSize)
    ReDim m_confid(BufSize)
Dim max_pri As Double
Dim pri_x As Long
Dim pri_y As Long
Dim patch_x As Long
Dim patch_y As Long
Dim Jidx As Long

    m_top = m_height ' initialize the rectangle area

    m_bottom = 0
    m_left = m_width
    m_right = 0
'Now fill some of the Variables
    For Y = 0 To m_height - 1
        j = Y * m_width
        For X = 0 To m_width - 1
            i = j + X
            m_color(i) = PicAr1(X, Y)
            m_r(i) = PicAr1(X, Y).rgbRed
            m_g(i) = PicAr1(X, Y).rgbgreen
            m_b(i) = PicAr1(X, Y).rgbBlue
        Next X
    Next Y
    Convert2Gray  ' convert it to gray image
    DrawBoundary MaskRed, MaskGreen, MaskBlue ' first time draw boundary
'Set Boundary for PatchTexture
    PatchL = IIf(m_left - BorderSize < 0, 0, m_left - BorderSize)
    PatchR = IIf(m_right + BorderSize > m_width - 1, m_width - 1, m_right + BorderSize)
    PatchT = IIf(m_top - BorderSize < 0, 0, m_top - BorderSize)
    PatchB = IIf(m_bottom + BorderSize > m_height - 1, m_height - 1, m_bottom + BorderSize)
    draw_source ' find the patches that can be used as sample texture
    For j = m_top To m_bottom
        Y = j * m_width
        For i = m_left To m_right  'if it is boundary, calculate the priority
            If m_mark(Y + i) = -2 Then
                m_pri(Y + i) = priority(i, j)
            End If
        Next i
    Next j
'Now the real function
    Do While TargetExist()
        max_pri = 0
        Count = Count + 1
        For j = m_top To m_bottom
            Jidx = j * m_width
            For i = m_left To m_right
                If m_mark(Jidx + i) = -2 And m_pri(Jidx + i) > max_pri Then ' find the boundary pixel with highest priority
                    pri_x = i
                    pri_y = j
                    max_pri = m_pri(Jidx + i)
                End If
            Next i
        Next j
        DoEvents
        If StopIt Then Exit Function
        PatchTexture pri_x, pri_y, patch_x, patch_y ' find the most similar source patch
        DoEvents
        If StopIt Then Exit Function
        update pri_x, pri_y, patch_x, patch_y, ComputeConfidence(pri_x, pri_y) ' inpaint this area and update confidence
        DoEvents
        If StopIt Then Exit Function
        UpdateBoundary pri_x, pri_y, MaskRed, MaskGreen, MaskBlue ' update boundary near the changed area
        DoEvents
        If StopIt Then Exit Function
        UpdatePri pri_x, pri_y  ' update priority near the changed area
        DoEvents
        If StopIt Then Exit Function
        If Preview Then
            Array2Pic ResultPicture, PicAr1
            ResultPicture.Refresh
            DoEvents
        End If
    Loop
    DoInPaint = Count
    Array2Pic ResultPicture, PicAr1
    ResultPicture.Picture = ResultPicture.Image
    ResultPicture.Refresh

End Function
Private Sub DrawBoundary(MaskRed As Byte, MaskGreen As Byte, MaskBlue As Byte)

Dim X As Long
Dim Y As Long
Dim j As Long
Dim i As Long
Dim Found As Boolean
    On Error Resume Next

    For Y = 0 To m_height - 1
        For X = 0 To m_width - 1
            If PicAr1(X, Y).rgbRed = MaskRed And PicAr1(X, Y).rgbgreen = MaskGreen And PicAr1(X, Y).rgbBlue = MaskBlue Then ' if the pixel is specified as boundary
                m_mark(Y * m_width + X) = -1
                m_confid(Y * m_width + X) = 0
                Else
                m_mark(Y * m_width + X) = Source
                m_confid(Y * m_width + X) = 1
            End If
        Next X
    Next Y
    For j = 0 To m_height - 1
        For i = 0 To m_width - 1
            If m_mark(j * m_width + i) = -1 Then
                If i < m_left Then ' resize the rectangle to the range of target area
                    m_left = i
                End If
                If i > m_right Then
                    m_right = i
                End If
                If j > m_bottom Then
                    m_bottom = j
                End If
                If j < m_top Then
                    m_top = j
                End If
'if one of the four neighbours is source pixel, then this should be a boundary
                If j = m_height - 1 Or j = 0 Or i = 0 Or i = m_width - 1 Then Found = True
                If m_mark(j * m_width + i + 1) = Source Then Found = True
                If m_mark(j * m_width + i - 1) = Source Then Found = True
                If m_mark((j + 1) * m_width + i) = Source Then Found = True
                If m_mark((j - 1) * m_width + i) = Source Then Found = True
                If Found Then
                    Found = False
                    m_mark(j * m_width + i) = -2
                End If
            End If
        Next i
    Next j

End Sub
Private Function ComputeConfidence(ByVal i As Long, ByVal j As Long) As Double

Dim confidence As Double
Dim X As Long
Dim Y As Long

    For Y = (IIf(((j - Winsize) > (0)), (j - Winsize), (0))) To (IIf(((j + Winsize) < (m_height - 1)), (j + Winsize), (m_height - 1)))
        For X = (IIf(((i - Winsize) > (0)), (i - Winsize), (0))) To (IIf(((i + Winsize) < (m_width - 1)), (i + Winsize), (m_width - 1)))
            confidence = confidence + m_confid(Y * m_width + X)
        Next X
    Next Y
    confidence = confidence / ((Winsize * 2 + 1) * (Winsize * 2 + 1))
    ComputeConfidence = confidence

End Function
Private Function priority(ByVal i As Long, ByVal j As Long) As Double

Dim confidence As Double
Dim data As Double
    confidence = ComputeConfidence(i, j) ' confidence term

    data = ComputeData(i, j) ' data term
    priority = confidence * data

End Function
Private Function ComputeData(ByVal i As Long, ByVal j As Long) As Double

Dim grad As gradient
Dim temp As gradient
Dim grad_T As gradient
    grad.grad_x = 0

    grad.grad_y = 0
Dim result As Double
Dim magnitude As Double
Dim max As Double
Dim X As Long
Dim Y As Long
Dim nn As norm
Dim Found As Boolean
    On Error Resume Next

    For Y = (IIf(((j - Winsize) > (0)), (j - Winsize), (0))) To (IIf(((j + Winsize) < (m_height - 1)), (j + Winsize), (m_height - 1)))
        For X = (IIf(((i - Winsize) > (0)), (i - Winsize), (0))) To (IIf(((i + Winsize) < (m_width - 1)), (i + Winsize), (m_width - 1)))
' find the greatest gradient in this patch, this will be the gradient of this pixel
            If m_mark(Y * m_width + X) >= 0 Then ' source pixel
'since I use four neighbors to calculate the gradient, make sure this four neighbors do not touch target region(big jump in gradient)
                Found = False
                If m_mark(Y * m_width + X + 1) < 0 Then Found = True
                If m_mark(Y * m_width + X - 1) < 0 Then Found = True
                If m_mark((Y + 1) * m_width + X) < 0 Then Found = True
                If m_mark((Y - 1) * m_width + X) < 0 Then Found = True
                If Found = False Then
                    temp = GetGradient(X, Y)
                    magnitude = temp.grad_x * temp.grad_x + temp.grad_y * temp.grad_y
                    If magnitude > max Then
                        grad.grad_x = temp.grad_x
                        grad.grad_y = temp.grad_y
                        max = magnitude
                    End If
                End If
            End If
        Next X
    Next Y
    grad_T.grad_x = grad.grad_y ' perpendicular to the gradient: (x,y)->(y, -x)
    grad_T.grad_y = -grad.grad_x
    nn = GetNorm(i, j)
    result = nn.norm_x * grad_T.grad_x + nn.norm_y * grad_T.grad_y ' dot product
    result = result / 255 '"alpha" in the paper: normalization factor
    result = Abs(result)
    ComputeData = result

End Function
'Get a gray Picture
Private Sub Convert2Gray()

Dim r As Double
Dim g As Double
Dim b As Double
Dim X As Long
Dim Y As Long

    For Y = 0 To m_height - 1
        For X = 0 To m_width - 1
            r = PicAr1(X, Y).rgbRed
            g = PicAr1(X, Y).rgbgreen
            b = PicAr1(X, Y).rgbBlue
            m_gray(Y * m_width + X) = CDbl((r * 3735 + g * 19267 + b * 9765) / 32767)
        Next X
    Next Y

End Sub
'Calculate the Gradient
Private Function GetGradient(ByVal i As Long, ByVal j As Long) As gradient

Dim result As gradient
    result.grad_x = (m_gray(j * m_width + i + 1) - m_gray(j * m_width + i - 1)) / 2#

    result.grad_y = (m_gray((j + 1) * m_width + i) - m_gray((j - 1) * m_width + i)) / 2#
    If i = 0 Then
        result.grad_x = m_gray(j * m_width + i + 1) - m_gray(j * m_width + i)
    End If
    If i = m_width - 1 Then
        result.grad_x = m_gray(j * m_width + i) - m_gray(j * m_width + i - 1)
    End If
    If j = 0 Then
        result.grad_y = m_gray((j + 1) * m_width + i) - m_gray(j * m_width + i)
    End If
    If j = m_height - 1 Then
        result.grad_y = m_gray(j * m_width + i) - m_gray((j - 1) * m_width + i)
    End If
    GetGradient = result

End Function
'Find the Normals
Private Function GetNorm(ByVal i As Long, ByVal j As Long) As norm

Dim result As norm
Dim num As Long
Dim neighbor_x(8) As Long
Dim neighbor_y(8) As Long
Dim record(8) As Long
Dim Count As Long
Dim X As Long
Dim Y As Long
Dim n_x As Long
Dim n_y As Long
Dim temp As Long
Dim square As Double

    For Y = (IIf(((j - 1) > (0)), (j - 1), (0))) To (IIf(((j + 1) < (m_height - 1)), (j + 1), (m_height - 1)))
        For X = (IIf(((i - 1) > (0)), (i - 1), (0))) To (IIf(((i + 1) < (m_width - 1)), (i + 1), (m_width - 1)))
            Count = Count + 1
            If X <> i Or Y <> j Then
                If m_mark(Y * m_width + X) = -2 Then
                    num = num + 1
                    neighbor_x(num) = X
                    neighbor_y(num) = Y
                    record(num) = Count
                End If
            End If
        Next X
    Next Y
    If num = 0 Or num = 1 Then ' if it doesn't have two neighbors, give it a random number to proceed
        result.norm_x = 0.6
        result.norm_y = 0.8
        GetNorm = result
        Exit Function
    End If
' draw a line between the two neighbors of the boundary pixel, then the norm is the perpendicular to the line
    n_x = neighbor_x(2) - neighbor_x(1)
    n_y = neighbor_y(2) - neighbor_y(1)
    temp = n_x
    n_x = n_y
    n_y = temp
    square = CDbl(n_x * n_x + n_y * n_y) ^ 0.5
    If n_x = 0 Then
        result.norm_x = 0
        Else
        result.norm_x = n_x / square
    End If
    If n_y = 0 Then
        result.norm_y = 0
        Else
        result.norm_y = n_y / square
    End If
    GetNorm = result

End Function
Private Function draw_source() As Boolean

' draw a window around the pixel, if all of the points within the window are source pixels, then this patch can be used as a source patch
Dim X As Long
Dim Y As Long
Dim i As Long
Dim j As Long
Dim flag As Boolean

    For j = 0 To m_height - 1
        For i = 0 To m_width - 1
            flag = True
            If i < Winsize Or j < Winsize Or i >= m_width - Winsize Or j >= m_height - Winsize Then 'cannot form a complete window
                m_source(j * m_width + i) = False
                Else
                For Y = j - Winsize To j + Winsize
                    For X = i - Winsize To i + Winsize
                        If m_mark(Y * m_width + X) <> Source Then
                            m_source(j * m_width + i) = False
                            flag = False
                            Exit For
                        End If
                    Next X
                    If flag = False Then
                        Exit For
                    End If
                Next Y
                If flag <> False Then
                    m_source(j * m_width + i) = True
                End If
            End If
        Next i
    Next j
    draw_source = True

End Function
Private Function PatchTexture(ByVal X As Long, ByVal Y As Long, ByRef patch_x As Long, ByRef patch_y As Long) As Boolean

Dim temp_r As Double
Dim temp_g As Double
Dim temp_b As Double
Dim i As Long
Dim j As Long
Dim iter_y As Long
Dim iter_x As Long
Dim min As Double
Dim sum As Double
Dim source_x As Long
Dim source_y As Long
Dim target_x As Long
Dim target_y As Long
Dim Tidx As Long
Dim Sidx As Long
Dim Jidx As Long
    
    min = 99999999

    For j = PatchT To PatchB
        Jidx = j * m_width
        For i = PatchL To PatchR
            If m_source(Jidx + i) Then
                sum = 0
                If StopIt Then Exit Function
                For iter_y = -Winsize To Winsize
                    target_y = Y + iter_y
                    If target_y > 0 And target_y < m_height Then
                        target_y = target_y * m_width
                        For iter_x = -Winsize To Winsize
                            target_x = X + iter_x
                            If target_x > 0 And target_x < m_width Then
                                Tidx = target_y + target_x
                                If m_mark(Tidx) >= 0 Then
                                    source_x = i + iter_x
                                    source_y = j + iter_y
                                    Sidx = source_y * m_width + source_x
                                    temp_r = m_r(Tidx) - m_r(Sidx)
                                    temp_g = m_g(Tidx) - m_g(Sidx)
                                    temp_b = m_b(Tidx) - m_b(Sidx)
                                    sum = sum + temp_r * temp_r + temp_g * temp_g + temp_b * temp_b
                                End If
                            End If
                        Next iter_x
                    End If
                Next iter_y
                If sum < min Then
                    min = sum
                    patch_x = i
                    patch_y = j
                End If
            End If
        Next i
    Next j
    PatchTexture = True

End Function
Private Function update(ByVal target_x As Long, ByVal target_y As Long, ByVal source_x As Long, ByVal source_y As Long, ByVal confid As Double) As Boolean

Dim color As Integer
Dim r As Double
Dim g As Double
Dim b As Double
Dim x0 As Long
Dim y0 As Long
Dim x1 As Long
Dim y1 As Long
Dim iter_y As Long
Dim iter_x As Long
Dim X0idx As Long
Dim X1idx As Long

    x0 = (-1) * Winsize

    On Error Resume Next
    For iter_y = (-1) * Winsize To Winsize
        For iter_x = (-1) * Winsize To Winsize
            x0 = source_x + iter_x
            y0 = source_y + iter_y
            x1 = target_x + iter_x
            y1 = target_y + iter_y
            X1idx = y1 * m_width + x1
            If m_mark(X1idx) < 0 Then
                X0idx = y0 * m_width + x0
                PicAr1(x1, y1) = m_color(X0idx) ' inpaint the color
                m_color(X1idx) = m_color(X0idx)
                m_r(X1idx) = m_r(X0idx)
                m_g(X1idx) = m_g(X0idx)
                m_b(X1idx) = m_b(X0idx)
                m_gray(X1idx) = CDbl((m_r(X0idx) * 3735 + m_g(X0idx) * 19267 + m_b(X0idx) * 9765) / 32767) ' update gray image
                m_confid(X1idx) = confid ' update the confidence
            End If
        Next iter_x
    Next iter_y
    update = True

End Function
'Is there still someting to do
Private Function TargetExist() As Boolean

Dim i As Long
Dim j As Long

    For j = m_top To m_bottom
        For i = m_left To m_right
            If m_mark(j * m_width + i) < 0 Then
                TargetExist = True
                Exit Function
            End If
        Next i
    Next j
    TargetExist = False

End Function
Private Sub UpdateBoundary(ByVal i As Long, ByVal j As Long, MaskRed As Byte, MaskGreen As Byte, MaskBlue As Byte)

Dim X As Long
Dim Y As Long
Dim Yidx As Long
Dim Found As Boolean
    On Error Resume Next

    For Y = (IIf(((j - Winsize - 2) > (0)), (j - Winsize - 2), (0))) To (IIf(((j + Winsize + 2) < (m_height - 1)), (j + Winsize + 2), (m_height - 1)))
        Yidx = Y * m_width
        For X = (IIf(((i - Winsize - 2) > (0)), (i - Winsize - 2), (0))) To (IIf(((i + Winsize + 2) < (m_width - 1)), (i + Winsize + 2), (m_width - 1)))
            If PicAr1(X, Y).rgbRed = MaskRed And PicAr1(X, Y).rgbgreen = MaskGreen And PicAr1(X, Y).rgbBlue = MaskBlue Then ' if the pixel is specified as boundary
                m_mark(Yidx + X) = -1
                Else
                m_mark(Yidx + X) = Source
            End If
        Next X
    Next Y
    For Y = (IIf(((j - Winsize - 2) > (0)), (j - Winsize - 2), (0))) To (IIf(((j + Winsize + 2) < (m_height - 1)), (j + Winsize + 2), (m_height - 1)))
        Yidx = Y * m_width
        For X = (IIf(((i - Winsize - 2) > (0)), (i - Winsize - 2), (0))) To (IIf(((i + Winsize + 2) < (m_width - 1)), (i + Winsize + 2), (m_width - 1)))
            If m_mark(Yidx + X) = -1 Then
                If Y = m_height - 1 Or Y = 0 Or X = 0 Or X = m_width - 1 Then Found = True
                If m_mark(Yidx + X - 1) = Source Then Found = True
                If m_mark(Yidx + X + 1) = Source Then Found = True
                If m_mark((Y - 1) * m_width + X) = Source Then Found = True
                If m_mark((Y + 1) * m_width + X) = Source Then Found = True
                If Found Then
                    Found = False
                    m_mark(Yidx + X) = -2
                End If
            End If
        Next X
    Next Y

End Sub
Private Sub UpdatePri(ByVal i As Long, ByVal j As Long)

Dim X As Long
Dim Y As Long
Dim Yidx As Long

    For Y = (IIf(((j - Winsize - 3) > (0)), (j - Winsize - 3), (0))) To (IIf(((j + Winsize + 3) < (m_height - 1)), (j + Winsize + 3), (m_height - 1)))
        Yidx = Y * m_width
        For X = (IIf(((i - Winsize - 3) > (0)), (i - Winsize - 3), (0))) To (IIf(((i + Winsize + 3) < (m_width - 1)), (i + Winsize + 3), (m_width - 1)))
            If m_mark(Yidx + X) = -2 Then
                m_pri(Yidx + X) = priority(X, Y)
            End If
        Next X
    Next Y

End Sub
Private Sub Pic2Array(Picbox As PictureBox, ByRef PicArray() As RGBQUAD)

    GetObject Picbox.Image, Len(PicInfo), PicInfo
    ReDim PicArray(0 To PicInfo.bmWidth - 1, 0 To PicInfo.bmHeight - 1) As RGBQUAD
    GetBitmapBits Picbox.Image, PicInfo.bmWidth * PicInfo.bmHeight * 4, PicArray(0, 0)

End Sub
Private Sub Array2Pic(Picbox As PictureBox, ByRef PicArray() As RGBQUAD)

    GetObject Picbox.Image, Len(PicInfo), PicInfo
    SetBitmapBits Picbox.Image, PicInfo.bmWidth * PicInfo.bmHeight * 4, PicArray(0, 0)

End Sub
