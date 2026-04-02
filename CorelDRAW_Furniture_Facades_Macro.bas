Option Explicit

Public mPlaced As Long
Public mSkipped As Long
Public mSheets As Long
Public mUsedArea As Double

Public Sub RunModernShapeBuilder()
    On Error GoTo EH
    UserForm1.Show vbModal
    Exit Sub
EH:
    MsgBox "Shape Builder: " & Err.Description, vbCritical
End Sub

Public Sub RunArtFacadePanel()
    RunModernShapeBuilder
End Sub

Public Sub DrawRect(ByVal lay As Layer, ByVal x As Double, ByVal yTop As Double, ByVal w As Double, ByVal h As Double)
    On Error Resume Next
    Dim r As Shape
    Set r = lay.CreateRectangle(x, yTop, x + w, yTop - h)
    If Not r Is Nothing Then
        r.Outline.Width = 0.3
        r.Outline.Color.RGBAssign 0, 0, 0
        r.Fill.UniformColor.RGBAssign 240, 240, 250
    End If
    On Error GoTo 0
End Sub

Public Sub RunCutPRO(ByVal sheetW As Double, ByVal sheetH As Double, ByVal gap As Double, _
                     ByVal fc As Long, fw() As Double, fh() As Double, fq() As Long)

    Dim allW() As Double
    Dim allH() As Double
    Dim pending() As Boolean
    Dim total As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim tmpW As Double
    Dim tmpH As Double
    Dim remaining As Long
    Dim doc As Document
    Dim lay As Layer
    Dim curX As Double
    Dim curY As Double
    Dim rowH As Double
    Dim placedOnSheet As Long
    Dim w As Double
    Dim h As Double
    Dim rotated As Boolean
    Dim tmp As Double
    Dim tryAgain As Boolean

    total = 0
    For i = 0 To fc - 1
        total = total + fq(i)
    Next i
    If total = 0 Then Exit Sub

    ReDim allW(0 To total - 1)
    ReDim allH(0 To total - 1)

    k = 0
    For i = 0 To fc - 1
        For j = 1 To fq(i)
            allW(k) = fw(i)
            allH(k) = fh(i)
            k = k + 1
        Next j
    Next i

    For i = 0 To total - 2
        For j = i + 1 To total - 1
            If allW(j) * allH(j) > allW(i) * allH(i) Then
                tmpW = allW(i): allW(i) = allW(j): allW(j) = tmpW
                tmpH = allH(i): allH(i) = allH(j): allH(j) = tmpH
            End If
        Next j
    Next i

    ReDim pending(0 To total - 1)
    For i = 0 To total - 1
        pending(i) = True
    Next i

    remaining = total

    If ActiveDocument Is Nothing Then
        Set doc = Documents.Add
    Else
        Set doc = ActiveDocument
    End If
    On Error Resume Next
    doc.Unit = cdrMillimeter
    On Error GoTo 0

    Set lay = doc.ActivePage.ActiveLayer

    Dim sheetOffX As Double
    Dim sr As Shape

    mPlaced = 0
    mSkipped = 0
    mSheets = 0
    mUsedArea = 0

    Do While remaining > 0

        sheetOffX = mSheets * (sheetW + 40)

        On Error Resume Next
        Set sr = lay.CreateRectangle(sheetOffX, sheetH, sheetOffX + sheetW, 0)
        If Not sr Is Nothing Then
            sr.Outline.Width = 0.5
            sr.Outline.Color.RGBAssign 0, 0, 180
            sr.Fill.UniformColor.RGBAssign 248, 248, 255
        End If
        On Error GoTo 0

        curX = gap
        curY = sheetH - gap
        rowH = 0
        placedOnSheet = 0

        For i = 0 To total - 1
            If Not pending(i) Then GoTo SkipThis

            w = allW(i)
            h = allH(i)
            rotated = False
            tryAgain = True

            Do While tryAgain
                tryAgain = False

                If curX + w + gap > sheetW Then
                    curX = gap
                    curY = curY - rowH - gap
                    rowH = 0
                End If

                If curY - h < 0 Then
                    If Not rotated Then
                        tmp = w: w = h: h = tmp
                        rotated = True
                        tryAgain = True
                    End If
                End If
            Loop

            If curY - h < 0 Then GoTo SkipThis

            DrawRect lay, sheetOffX + curX, curY, w, h

            curX = curX + w + gap
            If h > rowH Then rowH = h

            pending(i) = False
            remaining = remaining - 1
            mPlaced = mPlaced + 1
            mUsedArea = mUsedArea + w * h
            placedOnSheet = placedOnSheet + 1

SkipThis:
        Next i

        mSheets = mSheets + 1

        If placedOnSheet = 0 Then
            mSkipped = remaining
            Exit Do
        End If

        If mSheets > 50 Then
            mSkipped = remaining
            Exit Do
        End If

    Loop

End Sub
