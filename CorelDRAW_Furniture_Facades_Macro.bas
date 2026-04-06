Option Explicit

Public mPlaced As Long
Public mSkipped As Long
Public mSheets As Long
Public mUsedArea As Double
Public mCostTotal As Double
' Стиль отрисовки деталей (ставит BuildFacade из формы; DistributeOnSheet сбрасывает к умолчанию)
Public mOutlineStyle As Long  ' 0 без обводки, 1 тонкая, 2 средняя
Public mOrnMode As Long       ' 0 нет, 1 углы (метки), 2 классика (внутр. рамка)
' Геометрия «рама / филёнка» (мм; подписи «см» на форме — ввод как число в мм)
Public mInsetX As Double
Public mInsetY As Double
Public mArcTop As Double
Public mArcSide As Double
Public mOuterQty As Long
Public mOuterMm As Double
Public mInnerQty As Long
Public mInnerMm As Double
' ЧПУ: детали на CUT, подписи на TEXT, контур листа на SHEET (экспорт DXF)
Public mCncLayers As Boolean
' Guillotine: если True — справа остаток на всю высоту зоны (лучше для широких деталей у края)
Public mGuillotineTallRightStrip As Boolean
' Тип фасада из формы: 0=Прямой, 1=Классика 1, 2=Классика 2, 3=Неоклассика 1 (задаётся перед RunPlacement)
Public mFormFaceKind As Long
' Неоклассика 1: длины сегментов вдоль верхнего ребра (мм), 2…8 шт., задаётся с формы
Public mNeoSegN As Long
Public mNeoSeg() As Double

' Эвристика RunCutOptimizer / NestingPieceScore (подстройка утилизации листа)
Private Const NEST_SCORE_SNUG As Double = 0.12
Private Const NEST_COL_COEF As Double = 0.19
Private Const NEST_COL_TALL_EXTRA As Double = 0.11
Private Const NEST_FLOOR_FRAC As Double = 0.125
Private Const NEST_FLOOR_MIN_MM2 As Double = 3500#

' Неоклассика 1: горизонтали вдоль верхнего реза (мм, сумма 84), разрез для направляющих/фрезы
Private Const NEO_L1 As Double = 57#
Private Const NEO_L2 As Double = 7#
Private Const NEO_L3 As Double = 8#
Private Const NEO_L4 As Double = 12#
Private Const NEO_PATTERN As Double = 84#
' Мин. остаток внутри вложенного прямоугольника неоклассики (мм); иначе кольцо не рисуем
Private Const NEO_INSET_WALL_MIN_MM As Double = 0.25

Private mPartCounter As Long

'--- comdlg32: объявления только в начале модуля (до любого Sub/Function) ---
Private Type tagOPENFILENAME
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As String
End Type

#If VBA7 Then
Private Declare PtrSafe Function apiGetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (OFN As tagOPENFILENAME) As Long
Private Declare PtrSafe Function apiGetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (OFN As tagOPENFILENAME) As Long
#End If

Private Const OFN_PATHMUSTEXIST As Long = &H800
Private Const OFN_FILEMUSTEXIST As Long = &H1000
Private Const OFN_EXPLORER As Long = &H80000
Private Const OFN_OVERWRITEPROMPT As Long = &H2
Private Const OFN_HIDEREADONLY As Long = &H4

' Версия линейки: Function (не Const) — в Corel после AddFromString Const иногда «теряется», даёт Variable not defined
' Optional — чтобы не дублировать запись в «Запустить макрос» (только Фасады — публичный вход).
Public Function MEBEL_MACRO_VERSION(Optional ByVal MacroListHide As Byte = 0) As String
    MEBEL_MACRO_VERSION = "0.1.4"
End Function

Private Sub RunModernShapeBuilder()
    On Error GoTo EH
    UserForm1.Show vbModal
    Exit Sub
EH:
    MsgBox "Shape Builder: " & Err.Description & " (" & CStr(Err.Number) & ")", vbCritical, "Shape Builder"
End Sub

' Единственный пункт в списке «Сервис — Макросы — Запустить» для GlobalMacros (остальные входы — Private).
Public Sub Фасады()
    RunModernShapeBuilder
End Sub

Private Sub ShapeBuilderAbout()
    MsgBox "Shape Builder Pro" & vbCrLf & _
           "Версия линейки: " & MEBEL_MACRO_VERSION() & vbCrLf & vbCrLf & _
           "Если в заголовке формы старая версия (например v2.2.0) — заливка не применилась." & vbCrLf & _
           "Запустите deploy_direct.ps1 из папки проекта (или ЗАПУСК_ДЕПЛОЯ.bat), затем в VBA: Сохранить GlobalMacros." & vbCrLf & vbCrLf & _
           "Типы фасадов — в CSV (кнопки «+ сохранить тип», «Типы…»; удаление — в окне «Типы»)." & vbCrLf & _
           "Неоклассика: превью — сечение шпунта (мм), поле сегментов на «Основное»." & vbCrLf & vbCrLf & _
           "Запуск: только макрос «Фасады» (Module1.Фасады).", _
           vbInformation, "Shape Builder"
End Sub

Public Sub SetDrawStyleFromForm(ByVal borderTxt As String, ByVal ornTxt As String)
    Dim b As String, o As String
    b = LCase$(Trim$(borderTxt))
    o = LCase$(Trim$(ornTxt))
    mOutlineStyle = 1
    If InStr(1, b, "без", vbTextCompare) > 0 And InStr(1, b, "обв", vbTextCompare) > 0 Then mOutlineStyle = 0
    If InStr(1, b, "тонк", vbTextCompare) > 0 Then mOutlineStyle = 1
    If InStr(1, b, "сред", vbTextCompare) > 0 Then mOutlineStyle = 2
    mOrnMode = 0
    If InStr(1, o, "угол", vbTextCompare) > 0 Or InStr(1, o, "угл", vbTextCompare) > 0 Then mOrnMode = 1
    If InStr(1, o, "класс", vbTextCompare) > 0 Then mOrnMode = 2
End Sub

Private Sub ResetDrawStyleDefaults()
    mOutlineStyle = 1
    mOrnMode = 0
    mInsetX = 0#
    mInsetY = 0#
    mArcTop = 0#
    mArcSide = 0#
    mOuterQty = 0
    mOuterMm = 0#
    mInnerQty = 0
    mInnerMm = 0#
End Sub

Public Sub SetPanelGeometry( _
    ByVal ix As Double, ByVal iy As Double, _
    ByVal arcTop As Double, ByVal arcSide As Double, _
    ByVal oQty As Long, ByVal oMm As Double, _
    ByVal iQty As Long, ByVal iMm As Double)
    mInsetX = ix
    mInsetY = iy
    mArcTop = arcTop
    mArcSide = arcSide
    mOuterQty = oQty
    mOuterMm = oMm
    mInnerQty = iQty
    mInnerMm = iMm
    If mOuterQty < 0 Then mOuterQty = 0
    If mOuterQty > 24 Then mOuterQty = 24
    If mInnerQty < 0 Then mInnerQty = 0
    If mInnerQty > 24 Then mInnerQty = 24
End Sub

Private Sub ApplyPartFill(ByVal r As Shape)
    On Error Resume Next
    If mCncLayers Then
        r.Fill.ApplyNoFill
    Else
        r.Fill.UniformColor.RGBAssign 240, 240, 250
    End If
End Sub

Private Sub ApplyShapeOutline(ByVal r As Shape)
    On Error Resume Next
    Select Case mOutlineStyle
        Case 0
            r.Outline.SetNoOutline
        Case 2
            r.Outline.Width = 0.6
            r.Outline.Color.RGBAssign 0, 0, 0
        Case Else
            r.Outline.Width = 0.28
            r.Outline.Color.RGBAssign 0, 0, 0
    End Select
End Sub

Private Sub StyleOrnDot(ByVal dot As Shape)
    On Error Resume Next
    dot.Fill.UniformColor.RGBAssign 200, 110, 35
    dot.Outline.SetNoOutline
End Sub

Private Sub StyleGuideShape(ByVal sh As Shape)
    On Error Resume Next
    sh.Outline.Width = 0.2
    sh.Outline.Color.RGBAssign 120, 85, 40
    sh.Fill.ApplyNoFill
End Sub

Private Function NeoXu(ByVal x0 As Double, ByVal u As Double, ByVal w As Double, ByVal mirrored As Boolean) As Double
    If mirrored Then NeoXu = x0 + w - u Else NeoXu = x0 + u
End Function

Private Sub DrawNeoGuideVert(ByVal lay As Layer, ByVal xv As Double, ByVal yTop As Double, ByVal yBottom As Double)
    Dim ln As Shape
    On Error Resume Next
    Set ln = lay.CreateLine(xv, yTop, xv, yBottom)
    If Not ln Is Nothing Then
        ln.Outline.Width = 0.18
        ln.Outline.Color.RGBAssign 0, 128, 72
        ln.Fill.ApplyNoFill
    End If
End Sub

Private Sub DrawNeoProfileSeg(ByVal lay As Layer, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double)
    Dim ln As Shape
    On Error Resume Next
    Set ln = lay.CreateLine(x1, y1, x2, y2)
    If Not ln Is Nothing Then
        ln.Outline.Width = 0.35
        ln.Outline.Color.RGBAssign 210, 72, 38
        ln.Fill.ApplyNoFill
    End If
End Sub

Public Sub AssignNeoProfileSegments(ByRef valsMm() As Double, ByVal n As Long)
    Dim i As Long
    mNeoSegN = 0
    Erase mNeoSeg
    If n < 2 Or n > 8 Then Exit Sub
    For i = 1 To n
        If valsMm(i) <= 0# Then Exit Sub
    Next i
    mNeoSegN = n
    ReDim mNeoSeg(1 To n)
    For i = 1 To n
        mNeoSeg(i) = valsMm(i)
    Next i
End Sub

' Прямоугольник + вертикали по стыкам сегментов; при ровно 4 сегментах — фрезерный «разрез» как раньше
Private Sub DrawNeoclassic1Facade(ByVal lay As Layer, ByVal x As Double, ByVal yTop As Double, ByVal w As Double, ByVal h As Double, ByVal mirrored As Boolean)
    On Error Resume Next
    Dim locN As Long, locS() As Double
    Dim i As Long, j As Long
    Dim pat As Double, sc As Double
    Dim ss() As Double, pos() As Double
    Dim d As Double
    Dim r As Shape, g As Shape
    Dim ya As Double, yMid As Double, yBot As Double
    Dim k As Long, cum As Double
    Dim s1 As Double, s2 As Double, s3 As Double, s4 As Double
    Dim u0 As Double, u1 As Double, u2 As Double, u3 As Double, u4 As Double

    locN = mNeoSegN
    If locN < 2 Then
        locN = 4
        ReDim locS(1 To 4)
        locS(1) = NEO_L1: locS(2) = NEO_L2: locS(3) = NEO_L3: locS(4) = NEO_L4
    Else
        ReDim locS(1 To locN)
        For i = 1 To locN
            locS(i) = mNeoSeg(i)
        Next i
    End If

    pat = 0#
    For i = 1 To locN
        pat = pat + locS(i)
    Next i
    If pat < 0.01 Then Exit Sub

    sc = 1#
    If w < pat - 0.01 Then sc = w / pat
    ReDim ss(1 To locN)
    ReDim pos(0 To locN)
    pos(0) = 0#
    For i = 1 To locN
        ss(i) = locS(i) * sc
        pos(i) = pos(i - 1) + ss(i)
    Next i

    d = h * 0.06
    If d > 4.5 Then d = 4.5
    If d < 1.2 Then d = 1.2
    ya = yTop
    yMid = yTop - d * 0.38
    yBot = yTop - d

    Set r = lay.CreateRectangle(x, yTop, x + w, yTop - h)
    If Not r Is Nothing Then
        ApplyPartFill r
        ApplyShapeOutline r
        AddOrnamentMarks lay, x, yTop, w, h
        AddPanelGeometry lay, x, yTop, w, h
    End If

    ' Вложенные контуры: k = 1…locN (все сегменты, в т.ч. последний «12»). Вызывается только при mFormFaceKind = 3 (Неоклассика 1).
    For k = 1 To locN
        cum = pos(k)
        If (w - 2# * cum) < NEO_INSET_WALL_MIN_MM Or (h - 2# * cum) < NEO_INSET_WALL_MIN_MM Then Exit For
        Set g = lay.CreateRectangle(x + cum, ya - cum, x + w - cum, ya - h + cum)
        If Not g Is Nothing Then
            StyleGuideShape g
            g.Outline.Color.RGBAssign 55, 85, 140
            g.Outline.Width = 0.22
        End If
    Next k

    ' Вертикали по стыкам сегментов — на всю высоту детали (раньше только ~28 мм от верха, на фасаде почти незаметно)
    For j = 1 To locN - 1
        DrawNeoGuideVert lay, NeoXu(x, pos(j), w, mirrored), ya, ya - h
    Next j

    If locN = 4 Then
        s1 = ss(1): s2 = ss(2): s3 = ss(3): s4 = ss(4)
        u0 = 0#: u1 = pos(1): u2 = pos(2): u3 = pos(3): u4 = pos(4)
        DrawNeoProfileSeg lay, NeoXu(x, u0, w, mirrored), ya, NeoXu(x, u1, w, mirrored), ya
        DrawNeoProfileSeg lay, NeoXu(x, u1, w, mirrored), ya, NeoXu(x, u1 + s2 * 0.5, w, mirrored), yMid
        DrawNeoProfileSeg lay, NeoXu(x, u1 + s2 * 0.5, w, mirrored), yMid, NeoXu(x, u2, w, mirrored), yTop - d * 0.12
        DrawNeoProfileSeg lay, NeoXu(x, u2, w, mirrored), yTop - d * 0.12, NeoXu(x, u3, w, mirrored), yBot
        DrawNeoProfileSeg lay, NeoXu(x, u3, w, mirrored), yBot, NeoXu(x, u4, w, mirrored), ya
        DrawNeoProfileSeg lay, NeoXu(x, u4, w, mirrored), ya, NeoXu(x, w, w, mirrored), ya
    Else
        DrawNeoProfileSeg lay, NeoXu(x, pos(0), w, mirrored), ya, NeoXu(x, pos(locN), w, mirrored), ya
        If pos(locN) + 0.02 < w Then
            DrawNeoProfileSeg lay, NeoXu(x, pos(locN), w, mirrored), ya, NeoXu(x, w, w, mirrored), ya
        End If
    End If
End Sub

' SetRoundness: 0–100% от радиуса = половина короткой стороны (док. Corel)
Private Sub TryApplyArcToRect(ByVal r As Shape, ByVal w As Double, ByVal h As Double)
    On Error Resume Next
    Dim a As Double
    If mArcTop > 0.05 And mArcSide > 0.05 Then
        a = (mArcTop + mArcSide) / 2#
    ElseIf mArcTop > 0.05 Then
        a = mArcTop
    ElseIf mArcSide > 0.05 Then
        a = mArcSide
    Else
        Exit Sub
    End If
    Dim halfL As Double
    halfL = w * 0.5
    If h * 0.5 < halfL Then halfL = h * 0.5
    If halfL < 0.2 Then Exit Sub
    If a > halfL - 0.05 Then a = halfL - 0.05
    Dim pct As Long
    pct = CLng((a / halfL) * 100#)
    If pct < 1 Then pct = 1
    If pct > 100 Then pct = 100
    If r.Type = cdrRectangleShape Then
        r.Rectangle.SetRoundness pct
    End If
End Sub

Private Sub AddPanelGeometry(ByVal lay As Layer, ByVal x As Double, ByVal yTop As Double, ByVal w As Double, ByVal h As Double)
    On Error Resume Next
    Dim ix As Double, iy As Double
    ix = mInsetX: iy = mInsetY
    If ix < 0# Then ix = 0#
    If iy < 0# Then iy = 0#
    Dim k As Long, j As Long
    Dim off As Double, inn As Double
    Dim g As Shape

    If mOuterQty > 0 And mOuterMm > 0.05 Then
        For k = 1 To mOuterQty
            off = k * mOuterMm
            If w <= off * 2.05 Or h <= off * 2.05 Then Exit For
            Set g = lay.CreateRectangle(x + off, yTop - off, x + w - off, yTop - h + off)
            StyleGuideShape g
        Next k
    End If

    If ix > 0.05 Or iy > 0.05 Then
        If w > ix * 2.05 And h > iy * 2.05 Then
            Set g = lay.CreateRectangle(x + ix, yTop - iy, x + w - ix, yTop - h + iy)
            StyleGuideShape g
            g.Outline.Color.RGBAssign 55, 85, 140

            If mInnerQty > 0 And mInnerMm > 0.05 Then
                For j = 1 To mInnerQty
                    inn = j * mInnerMm
                    If w <= (ix + inn) * 2.05 Or h <= (iy + inn) * 2.05 Then Exit For
                    Set g = lay.CreateRectangle(x + ix + inn, yTop - iy - inn, x + w - ix - inn, yTop - h + iy + inn)
                    StyleGuideShape g
                Next j
            End If
        End If
    End If
End Sub

Private Sub AddOrnamentMarks(ByVal lay As Layer, ByVal x As Double, ByVal yTop As Double, ByVal w As Double, ByVal h As Double)
    On Error Resume Next
    If mOrnMode = 0 Then Exit Sub
    Dim yb As Double
    yb = yTop - h
    If mOrnMode = 1 Then
        Dim s As Double
        s = 1.2
        Dim dot As Shape
        Set dot = lay.CreateRectangle(x, yTop, x + s, yTop - s): StyleOrnDot dot
        Set dot = lay.CreateRectangle(x + w - s, yTop, x + w, yTop - s): StyleOrnDot dot
        Set dot = lay.CreateRectangle(x, yb + s, x + s, yb): StyleOrnDot dot
        Set dot = lay.CreateRectangle(x + w - s, yb + s, x + w, yb): StyleOrnDot dot
    ElseIf mOrnMode = 2 Then
        If mInsetX > 0.05 Or mInsetY > 0.05 Then Exit Sub
        Dim inset As Double
        inset = 5#
        If w <= inset * 2.2 Or h <= inset * 2.2 Then Exit Sub
        Dim inn As Shape
        Set inn = lay.CreateRectangle(x + inset, yTop - inset, x + w - inset, yTop - h + inset)
        inn.Outline.Width = 0.22
        inn.Outline.Color.RGBAssign 95, 70, 35
        inn.Fill.ApplyNoFill
    End If
End Sub

Public Sub DrawRect(ByVal lay As Layer, ByVal x As Double, ByVal yTop As Double, ByVal w As Double, ByVal h As Double)
    On Error Resume Next
    If mFormFaceKind = 3 Then
        DrawNeoclassic1Facade lay, x, yTop, w, h, False
        On Error GoTo 0
        Exit Sub
    End If
    Dim r As Shape
    Set r = lay.CreateRectangle(x, yTop, x + w, yTop - h)
    If Not r Is Nothing Then
        TryApplyArcToRect r, w, h
        ApplyPartFill r
        ApplyShapeOutline r
        AddOrnamentMarks lay, x, yTop, w, h
        AddPanelGeometry lay, x, yTop, w, h
    End If
    On Error GoTo 0
End Sub

Private Function GetOrCreateLayer(ByVal pg As Page, ByVal nm As String) As Layer
    Dim L As Layer
    On Error Resume Next
    For Each L In pg.Layers
        If StrComp(L.Name, nm, vbTextCompare) = 0 Then
            Set GetOrCreateLayer = L
            Exit Function
        End If
    Next L
    Set GetOrCreateLayer = pg.CreateLayer(nm)
End Function

Private Sub DrawCutLabel(ByVal lay As Layer, ByVal xLeft As Double, ByVal yTop As Double, ByVal w As Double, ByVal h As Double, Optional ByVal sheetIdx As Long = -1)
    On Error Resume Next
    Dim t As Shape
    Dim s As String
    Dim sz As Long
    If mCncLayers Then
        mPartCounter = mPartCounter + 1
        If sheetIdx > 0 Then
            s = "S" & CStr(sheetIdx) & "  #" & CStr(mPartCounter) & vbCrLf & CStr(CLng(w + 0.5)) & " x " & CStr(CLng(h + 0.5))
        Else
            s = "#" & CStr(mPartCounter) & vbCrLf & CStr(CLng(w + 0.5)) & " x " & CStr(CLng(h + 0.5))
        End If
        sz = 8
    Else
        s = CStr(CLng(w + 0.5)) & " x " & CStr(CLng(h + 0.5))
        sz = 7
    End If
    Set t = lay.CreateArtisticText(xLeft + 2, yTop - h + 3, s)
    If Not t Is Nothing Then
        On Error Resume Next
        t.Text.Story.Size = sz
        t.Fill.UniformColor.RGBAssign 25, 25, 110
        t.Outline.SetNoOutline
    End If
End Sub

Public Sub RunCutPRO(ByVal sheetW As Double, ByVal sheetH As Double, ByVal gap As Double, _
                     ByVal fc As Long, facW() As Double, facH() As Double, fq() As Long)

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
            allW(k) = facW(i)
            allH(k) = facH(i)
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

    Dim laySheet As Layer
    Dim layCut As Layer
    Dim layText As Layer
    Dim sheetNo As Long

    If mCncLayers Then
        Set laySheet = GetOrCreateLayer(doc.ActivePage, "SHEET")
        Set layCut = GetOrCreateLayer(doc.ActivePage, "CUT")
        Set layText = GetOrCreateLayer(doc.ActivePage, "TEXT")
        mPartCounter = 0
    Else
        Set laySheet = lay
        Set layCut = lay
        Set layText = lay
    End If

    Dim sheetOffX As Double
    Dim sr As Shape

    mPlaced = 0
    mSkipped = 0
    mSheets = 0
    mUsedArea = 0
    mCostTotal = 0

    Do While remaining > 0

        sheetNo = mSheets + 1
        sheetOffX = mSheets * (sheetW + 40)

        On Error Resume Next
        Set sr = laySheet.CreateRectangle(sheetOffX, sheetH, sheetOffX + sheetW, 0)
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

            DrawRect layCut, sheetOffX + curX, curY, w, h
            DrawCutLabel layText, sheetOffX + curX, curY, w, h, sheetNo

            curX = curX + w + gap
            If h > rowH Then rowH = h

            pending(i) = False
            remaining = remaining - 1
            mPlaced = mPlaced + 1
            mUsedArea = mUsedArea + w * h
            placedOnSheet = placedOnSheet + 1
            If mPlaced Mod 80 = 0 Then DoEvents

SkipThis:
        Next i

        mSheets = mSheets + 1
        DoEvents

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

' --- Cut Optimizer (guillotine + free-rect) ---

Public Sub DrawRectMirror(ByVal lay As Layer, ByVal x As Double, ByVal yTop As Double, ByVal w As Double, ByVal h As Double, ByVal mirrored As Boolean)
    On Error Resume Next
    If mFormFaceKind = 3 Then
        DrawNeoclassic1Facade lay, x, yTop, w, h, mirrored
        On Error GoTo 0
        Exit Sub
    End If
    Dim r As Shape
    Set r = lay.CreateRectangle(x, yTop, x + w, yTop - h)
    If Not r Is Nothing Then
        TryApplyArcToRect r, w, h
        ApplyPartFill r
        ApplyShapeOutline r
        If mirrored Then
            r.Flip cdrFlipHorizontal
        End If
        AddOrnamentMarks lay, x, yTop, w, h
        AddPanelGeometry lay, x, yTop, w, h
    End If
    On Error GoTo 0
End Sub

Private Sub ParseStock(ByVal txt As String, ByVal defH As Double, ByVal defW As Double, _
                       ByRef qH() As Double, ByRef qW() As Double, ByRef qN As Long)
    Dim lines() As String
    Dim i As Long
    Dim s As String
    Dim p1 As Long
    Dim p2 As Long
    Dim rest As String
    Dim hh As Double
    Dim ww As Double
    Dim cnt As Long
    Dim k As Long
    Dim j As Long
    qN = 0
    txt = Replace(Replace(Trim$(txt), vbCrLf, "|"), vbLf, "|")
    If Len(txt) = 0 Then Exit Sub
    lines = Split(txt, "|")
    For i = LBound(lines) To UBound(lines)
        s = Trim$(lines(i))
        If Len(s) > 0 Then
            p1 = InStr(1, s, "x", vbTextCompare)
            If p1 = 0 Then p1 = InStr(1, s, "X", vbTextCompare)
            If p1 > 0 Then
                hh = Val(Trim$(Left$(s, p1 - 1)))
                rest = Mid$(s, p1 + 1)
                p2 = InStr(1, rest, "*")
                If p2 > 0 Then
                    ww = Val(Trim$(Left$(rest, p2 - 1)))
                    cnt = CLng(Val(Trim$(Mid$(rest, p2 + 1))))
                Else
                    ww = Val(Trim$(rest))
                    cnt = 1
                End If
                If hh > 0 And ww > 0 Then
                    If cnt < 1 Then cnt = 1
                    If cnt > 5000 Then cnt = 5000
                    qN = qN + cnt
                End If
            End If
        End If
    Next i
    If qN = 0 Then Exit Sub
    ReDim qH(0 To qN - 1)
    ReDim qW(0 To qN - 1)
    j = 0
    For i = LBound(lines) To UBound(lines)
        s = Trim$(lines(i))
        If Len(s) > 0 Then
            p1 = InStr(1, s, "x", vbTextCompare)
            If p1 = 0 Then p1 = InStr(1, s, "X", vbTextCompare)
            If p1 > 0 Then
                hh = Val(Trim$(Left$(s, p1 - 1)))
                rest = Mid$(s, p1 + 1)
                p2 = InStr(1, rest, "*")
                If p2 > 0 Then
                    ww = Val(Trim$(Left$(rest, p2 - 1)))
                    cnt = CLng(Val(Trim$(Mid$(rest, p2 + 1))))
                Else
                    ww = Val(Trim$(rest))
                    cnt = 1
                End If
                If hh > 0 And ww > 0 Then
                    If cnt < 1 Then cnt = 1
                    If cnt > 5000 Then cnt = 5000
                    For k = 1 To cnt
                        qH(j) = hh
                        qW(j) = ww
                        j = j + 1
                    Next k
                End If
            End If
        End If
    Next i
End Sub

Private Sub RemoveFreeRect(ByRef frL() As Double, ByRef frB() As Double, ByRef frW() As Double, ByRef frH() As Double, ByRef frN As Long, ByVal idx As Long)
    Dim k As Long
    If idx < 0 Or idx >= frN Then Exit Sub
    For k = idx To frN - 2
        frL(k) = frL(k + 1)
        frB(k) = frB(k + 1)
        frW(k) = frW(k + 1)
        frH(k) = frH(k + 1)
    Next k
    frN = frN - 1
End Sub

' Расширение массивов свободных прямоугольников (раньше жёсткий лимит 400 молча отбрасывал зоны → сбои/зависания).
Private Sub EnsureFreeRectCap(ByRef frL() As Double, ByRef frB() As Double, ByRef frW() As Double, ByRef frH() As Double, ByVal minUB As Long)
    Const MAX_FR_CAP As Long = 8000
    Dim ub As Long
    On Error Resume Next
    ub = UBound(frL)
    On Error GoTo 0
    If minUB <= ub Then Exit Sub
    If minUB > MAX_FR_CAP Then minUB = MAX_FR_CAP
    Dim nu As Long
    nu = ub + 250
    If nu < minUB Then nu = minUB
    If nu > MAX_FR_CAP Then nu = MAX_FR_CAP
    ReDim Preserve frL(0 To nu)
    ReDim Preserve frB(0 To nu)
    ReDim Preserve frW(0 To nu)
    ReDim Preserve frH(0 To nu)
End Sub

Private Sub AddFreeRect(ByRef frL() As Double, ByRef frB() As Double, ByRef frW() As Double, ByRef frH() As Double, ByRef frN As Long, ByVal L As Double, ByVal B As Double, ByVal W As Double, ByVal H As Double)
    Dim ub As Long
    If W <= 0.001 Or H <= 0.001 Then Exit Sub
    EnsureFreeRectCap frL, frB, frW, frH, frN
    On Error Resume Next
    ub = UBound(frL)
    On Error GoTo 0
    If frN > ub Then Exit Sub
    frL(frN) = L
    frB(frN) = B
    frW(frN) = W
    frH(frN) = H
    frN = frN + 1
End Sub

Private Function DblMin2(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then DblMin2 = a Else DblMin2 = b
End Function

Private Function DblMax2(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then DblMax2 = a Else DblMax2 = b
End Function

' Оценка размещения: локальный отход + «раскройщик» — узкие полосы, вытянутые детали не в широкую «полку».
' guillotineTallForced: если в форме включена жёсткая «полоса справа» — сильнее тянем детали в узкую колонку (с поворотом).
Private Function NestingPieceScore( _
    ByVal freeW As Double, ByVal freeH As Double, _
    ByVal w As Double, ByVal h As Double, _
    ByVal L As Double, ByVal B As Double, _
    ByVal pw0 As Double, ByVal ph0 As Double, _
    Optional ByVal guillotineTallForced As Boolean = False) As Double
    Dim fa As Double, waste As Double
    Dim snug As Double
    Dim mn0 As Double, mx0 As Double, elong As Double
    Dim slabPen As Double
    Dim colFit As Double
    Dim colBonus As Double
    fa = freeW * freeH
    waste = fa - w * h
    If freeW <= freeH Then
        If freeW > 0.001 Then snug = w / freeW Else snug = 0#
    Else
        If freeH > 0.001 Then snug = h / freeH Else snug = 0#
    End If
    If snug > 1# Then snug = 1#
    waste = waste - w * h * NEST_SCORE_SNUG * snug
    ' Вертикальная полоса (узкий freeW): не оставлять пустой, если деталь влезает с поворотом
    If freeH > freeW * 1.65 And freeW > 0.001 Then
        colFit = w / freeW
        If colFit > 1# Then colFit = 1#
        colBonus = w * h * NEST_COL_COEF * colFit * colFit
        If guillotineTallForced Then colBonus = colBonus + w * h * NEST_COL_TALL_EXTRA * colFit
        waste = waste - colBonus
    End If
    mn0 = DblMin2(pw0, ph0)
    mx0 = DblMax2(pw0, ph0)
    If mn0 > 0.001 Then elong = mx0 / mn0 Else elong = 1#
    slabPen = 0#
    If elong > 3.2 Then
        If freeW > freeH * 2.1 Then
            If w > freeW * 0.42 And h < freeH * 0.92 Then
                slabPen = DblMin2(fa * 0.22, freeW * DblMax2(0#, freeH - h) * 0.55)
            End If
        End If
    End If
    NestingPieceScore = waste + slabPen + 0.000000000001 * (L + B)
End Function

' Выбор разреза guillotine: «высокая полоса справа» vs классика — по max(min(S1,S2)) остатков.
Private Function GuillotinePickTallRightStrip( _
    ByVal freeW As Double, ByVal freeH As Double, _
    ByVal pw As Double, ByVal ph As Double, ByVal g As Double) As Boolean
    Dim c1 As Double, c2 As Double, t1 As Double, t2 As Double
    Dim rw As Double, rh As Double
    Dim mC As Double, mT As Double
    rw = freeW - pw - g
    rh = freeH - ph - g
    If rw > 0.001 Then
        c1 = rw * ph
        t1 = rw * freeH
    Else
        c1 = 0#
        t1 = 0#
    End If
    If rh > 0.001 Then
        c2 = freeW * rh
        t2 = pw * rh
    Else
        c2 = 0#
        t2 = 0#
    End If
    mC = DblMin2(c1, c2)
    mT = DblMin2(t1, t2)
    If mT > mC + 0.0000000001 Then
        GuillotinePickTallRightStrip = True
    ElseIf mC > mT + 0.0000000001 Then
        GuillotinePickTallRightStrip = False
    Else
        ' При равенстве остатков — классика, чтобы реже появлялась пустая высокая полоса справа
        GuillotinePickTallRightStrip = False
    End If
End Function

Public Sub RunCutOptimizer(ByVal sheetW As Double, ByVal sheetH As Double, ByVal gap As Double, _
                           ByVal fc As Long, facW() As Double, facH() As Double, fq() As Long, _
                           ByVal allowRotate As Boolean, ByVal allowMirror As Boolean, ByVal grainLock As Boolean, _
                           ByVal sheetStockText As String)

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
    Dim pw As Double
    Dim ph As Double
    Dim useRot As Boolean
    Dim sheetOffX As Double
    Dim sr As Shape
    Dim innerW As Double
    Dim innerH As Double
    Dim frL() As Double
    Dim frB() As Double
    Dim frW() As Double
    Dim frH() As Double
    Dim frN As Long
    Dim pi As Long
    Dim fr As Long
    Dim bestPi As Long
    Dim bestFr As Long
    Dim bestW As Double
    Dim bestH As Double
    Dim bestScore As Double
    Dim sc As Double
    Dim fa As Double
    Dim waste As Double
    Dim mirDraw As Boolean
    Dim placedThis As Long
    Dim didPlace As Boolean
    Dim L As Double
    Dim B As Double
    Dim freeW As Double
    Dim freeH As Double
    Dim qH() As Double
    Dim qW() As Double
    Dim qN As Long
    Dim stockIdx As Long
    Dim curSH As Double
    Dim curSW As Double
    Dim sheetBaseX As Double
    Dim laySheet As Layer
    Dim layCut As Layer
    Dim layText As Layer
    Dim curSheet As Long
    Dim maxA As Double
    Dim fa2 As Double
    Dim faFloor As Double
    Dim tryTier As Long
    Dim pig As Long

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
            allW(k) = facW(i)
            allH(k) = facH(i)
            k = k + 1
        Next j
    Next i

    For i = 1 To total - 1
        tmpW = allW(i)
        tmpH = allH(i)
        k = i - 1
        Do While k >= 0
            If allW(k) * allH(k) >= tmpW * tmpH Then Exit Do
            allW(k + 1) = allW(k)
            allH(k + 1) = allH(k)
            k = k - 1
        Loop
        allW(k + 1) = tmpW
        allH(k + 1) = tmpH
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

    If mCncLayers Then
        Set laySheet = GetOrCreateLayer(doc.ActivePage, "SHEET")
        Set layCut = GetOrCreateLayer(doc.ActivePage, "CUT")
        Set layText = GetOrCreateLayer(doc.ActivePage, "TEXT")
        mPartCounter = 0
    Else
        Set laySheet = lay
        Set layCut = lay
        Set layText = lay
    End If

    mPlaced = 0
    mSkipped = 0
    mSheets = 0
    mUsedArea = 0
    mCostTotal = 0

    ParseStock sheetStockText, sheetH, sheetW, qH, qW, qN

    stockIdx = 0
    sheetBaseX = 0

    ReDim frL(0 To 255)
    ReDim frB(0 To 255)
    ReDim frW(0 To 255)
    ReDim frH(0 To 255)

    useRot = allowRotate And Not grainLock

    Do While remaining > 0

        If qN > 0 And stockIdx < qN Then
            curSH = qH(stockIdx)
            curSW = qW(stockIdx)
            stockIdx = stockIdx + 1
        Else
            curSH = sheetH
            curSW = sheetW
        End If

        innerW = curSW - 2 * gap
        innerH = curSH - 2 * gap
        If innerW <= 0 Or innerH <= 0 Then
            mSkipped = remaining
            Exit Do
        End If

        sheetOffX = sheetBaseX
        curSheet = mSheets + 1

        On Error Resume Next
        Set sr = laySheet.CreateRectangle(sheetOffX, curSH, sheetOffX + curSW, 0)
        If Not sr Is Nothing Then
            sr.Outline.Width = 0.5
            sr.Outline.Color.RGBAssign 0, 0, 180
            sr.Fill.UniformColor.RGBAssign 248, 248, 255
        End If
        On Error GoTo 0

        frN = 1
        frL(0) = 0
        frB(0) = 0
        frW(0) = innerW
        frH(0) = innerH

        placedThis = 0

        Do
            didPlace = False
            bestPi = -1
            bestFr = -1
            bestScore = 1E+300
            bestW = 0
            bestH = 0

            maxA = 0#
            For pig = 0 To total - 1
                If pending(pig) Then
                    fa2 = allW(pig) * allH(pig)
                    If fa2 > maxA Then maxA = fa2
                End If
            Next pig
            faFloor = maxA * NEST_FLOOR_FRAC
            If faFloor < NEST_FLOOR_MIN_MM2 Then faFloor = 0#

            For tryTier = 0 To 1
                If tryTier = 1 Then faFloor = 0#
                bestScore = 1E+300
                bestPi = -1
                bestFr = -1
                bestW = 0
                bestH = 0

                For pi = 0 To total - 1
                    If pending(pi) Then
                        fa2 = allW(pi) * allH(pi)
                        If Not (faFloor > 0.5 And fa2 + 0.001 < faFloor) Then
                            pw = allW(pi)
                            ph = allH(pi)
                            For fr = 0 To frN - 1
                                L = frL(fr)
                                B = frB(fr)
                                freeW = frW(fr)
                                freeH = frH(fr)
                                If pw <= freeW And ph <= freeH Then
                                    sc = NestingPieceScore(freeW, freeH, pw, ph, L, B, pw, ph, mGuillotineTallRightStrip)
                                    If sc < bestScore Then
                                        bestScore = sc
                                        bestPi = pi
                                        bestFr = fr
                                        bestW = pw
                                        bestH = ph
                                    End If
                                End If
                                If useRot Then
                                    If Abs(pw - ph) > 0.001 Then
                                        If ph <= freeW And pw <= freeH Then
                                            sc = NestingPieceScore(freeW, freeH, ph, pw, L, B, pw, ph, mGuillotineTallRightStrip)
                                            If sc < bestScore Then
                                                bestScore = sc
                                                bestPi = pi
                                                bestFr = fr
                                                bestW = ph
                                                bestH = pw
                                            End If
                                        End If
                                    End If
                                End If
                            Next fr
                        End If
                    End If
                Next pi

                If bestPi >= 0 Then Exit For
            Next tryTier

            If bestPi < 0 Then Exit Do

            L = frL(bestFr)
            B = frB(bestFr)
            freeW = frW(bestFr)
            freeH = frH(bestFr)

            RemoveFreeRect frL, frB, frW, frH, frN, bestFr

            Dim useTall As Boolean
            If mGuillotineTallRightStrip Then
                useTall = True
            Else
                useTall = GuillotinePickTallRightStrip(freeW, freeH, bestW, bestH, gap)
            End If
            If useTall Then
                AddFreeRect frL, frB, frW, frH, frN, L + bestW + gap, B, freeW - bestW - gap, freeH
                AddFreeRect frL, frB, frW, frH, frN, L, B + bestH + gap, bestW, freeH - bestH - gap
            Else
                AddFreeRect frL, frB, frW, frH, frN, L + bestW + gap, B, freeW - bestW - gap, bestH
                AddFreeRect frL, frB, frW, frH, frN, L, B + bestH + gap, freeW, freeH - bestH - gap
            End If

            mirDraw = False
            If allowMirror Then mirDraw = ((mPlaced Mod 2) = 1)
            DrawRectMirror layCut, sheetOffX + gap + L, gap + B + bestH, bestW, bestH, mirDraw
            DrawCutLabel layText, sheetOffX + gap + L, gap + B + bestH, bestW, bestH, curSheet

            pending(bestPi) = False
            remaining = remaining - 1
            mPlaced = mPlaced + 1
            mUsedArea = mUsedArea + allW(bestPi) * allH(bestPi)
            placedThis = placedThis + 1
            didPlace = True
            If mPlaced Mod 80 = 0 Then DoEvents

        Loop While didPlace

        mSheets = mSheets + 1
        sheetBaseX = sheetBaseX + curSW + 40

        If placedThis = 0 Then
            mSkipped = remaining
            Exit Do
        End If

        If mSheets > 50 Then
            mSkipped = remaining
            Exit Do
        End If

        DoEvents

    Loop

End Sub

' nestMode: 0 = полка nesting (RunCutOptimizer), 1 = Guillotine, 2 = строчно RunCutPRO (больше отхода)
Public Sub RunPlacement(ByVal nestMode As Long, _
    ByVal sheetW As Double, ByVal sheetH As Double, ByVal gap As Double, _
    ByVal fc As Long, facW() As Double, facH() As Double, fq() As Long, _
    ByVal allowRot As Boolean, ByVal allowMir As Boolean, ByVal grainL As Boolean, _
    ByVal sheetStockText As String)
    If nestMode = 2 Then
        RunCutPRO sheetW, sheetH, gap, fc, facW, facH, fq
    Else
        RunCutOptimizer sheetW, sheetH, gap, fc, facW, facH, fq, allowRot, allowMir, grainL, sheetStockText
    End If
End Sub

' Optional — вызывается из формы; в списке «Запустить макрос» не показывать как отдельный пункт.
Public Sub ExportCncDxf(Optional ByVal MacroListHide As Byte = 0)
    Dim p As String
    Dim doc As Document
    On Error GoTo EH
    Set doc = ActiveDocument
    If doc Is Nothing Then
        MsgBox "Нет открытого документа.", vbExclamation
        Exit Sub
    End If
    p = ShowDxfSaveDialog(Environ$("USERPROFILE") & "\Desktop\", "layout.dxf")
    p = Trim$(p)
    If Len(p) = 0 Then Exit Sub
    If LCase$(Right$(p, 4)) <> ".dxf" Then p = p & ".dxf"
    On Error Resume Next
    doc.Export p, cdrDXF, cdrCurrentPage
    If Err.Number <> 0 Then
        MsgBox "Экспорт DXF: " & Err.Description & vbCrLf & "Проверьте единицы документа (мм) и наличие фильтра DXF в CorelDRAW.", vbExclamation
        Err.Clear
        Exit Sub
    End If
    On Error GoTo EH
    MsgBox "DXF сохранён (текущая страница, мм):" & vbCrLf & p, vbInformation
    Exit Sub
EH:
    MsgBox "Экспорт DXF: " & Err.Description, vbExclamation
End Sub

Public Function ShowDxfSaveDialog(ByVal defDir As String, ByVal defName As String) As String
    On Error GoTo EH
    Dim OFN As tagOPENFILENAME
    Dim strFile As String
    Dim strTitle As String
    Dim flt As String
    Dim ok As Long

    flt = "DXF (*.dxf)" & Chr$(0) & "*.dxf" & Chr$(0) & "All (*.*)" & Chr$(0) & "*.*" & Chr$(0) & Chr$(0)
    strFile = Left$(defName & String$(256, 0), 256)
    strTitle = String$(256, 0)

    With OFN
        .lStructSize = Len(OFN)
        .hwndOwner = 0
        .hInstance = 0
        .lpstrFilter = flt
        .nFilterIndex = 1
        .lpstrFile = strFile
        .nMaxFile = 260
        .lpstrFileTitle = strTitle
        .nMaxFileTitle = 256
        .lpstrInitialDir = defDir
        .lpstrTitle = "Сохранить DXF для ЧПУ"
        .Flags = OFN_EXPLORER Or OFN_PATHMUSTEXIST Or OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY
        .lpstrDefExt = "dxf"
        ok = apiGetSaveFileName(OFN)
        If ok <> 0 Then
            ShowDxfSaveDialog = TrimNull(.lpstrFile)
            Exit Function
        End If
    End With
EH:
    ShowDxfSaveDialog = ShowDxfSaveDialog_InputBox(defDir, defName)
End Function

Private Function ShowDxfSaveDialog_InputBox(ByVal defDir As String, ByVal defName As String) As String
    Dim s As String
    s = Trim$(InputBox("Полный путь к файлу DXF:", "Сохранить DXF", defDir & defName))
    ShowDxfSaveDialog_InputBox = s
End Function

'--- UserForm2: расклад «полка» (классика). Доп. параметры — для будущей отрисовки рамок/филёнок ---
Public Sub DistributeOnSheet( _
    ByVal sheetW As Double, ByVal sheetH As Double, ByVal gap As Double, _
    ByVal fc As Long, facW() As Double, facH() As Double, fq() As Long, _
    ByVal formIndex As Long, ByVal formName As String, _
    ByVal borderVal As String, ByVal ornVal As String, _
    ByVal insetX As Double, ByVal insetY As Double, _
    ByVal arcTop As Double, ByVal arcSide As Double, _
    ByVal outerQty As Long, ByVal outerMm As Double, _
    ByVal innerQty As Long, ByVal innerMm As Double, _
    ByVal noHandle As Boolean, ByVal fv As Long, _
    Optional ByVal useCncLayers As Boolean = False)

    Dim saveCnc As Boolean
    saveCnc = mCncLayers
    mCncLayers = useCncLayers

    SetDrawStyleFromForm borderVal, ornVal
    SetPanelGeometry insetX, insetY, arcTop, arcSide, outerQty, outerMm, innerQty, innerMm
    RunPlacement 0, sheetW, sheetH, gap, fc, facW, facH, fq, False, False, False, ""

    mCncLayers = saveCnc
End Sub

'--- Файловый диалог для UserForm2 (импорт/экспорт CSV) ---
Private Function TrimNull(ByVal s As String) As String
    Dim p As Long
    p = InStr(1, s, Chr$(0), vbBinaryCompare)
    If p > 0 Then TrimNull = Left$(s, p - 1) Else TrimNull = s
End Function

Public Function ShowFileDialogAPI(ByVal isSave As Boolean, ByVal defDir As String, ByVal defName As String) As String
    On Error GoTo EH
    Dim OFN As tagOPENFILENAME
    Dim strFile As String
    Dim strTitle As String
    Dim flt As String
    Dim ok As Long

    flt = "CSV (*.csv)" & Chr$(0) & "*.csv" & Chr$(0) & "All (*.*)" & Chr$(0) & "*.*" & Chr$(0) & Chr$(0)
    strFile = Left$(defName & String$(256, 0), 256)
    strTitle = String$(256, 0)

    With OFN
        .lStructSize = Len(OFN)
        .hwndOwner = 0
        .hInstance = 0
        .lpstrFilter = flt
        .nFilterIndex = 1
        .lpstrFile = strFile
        .nMaxFile = 260
        .lpstrFileTitle = strTitle
        .nMaxFileTitle = 256
        .lpstrInitialDir = defDir
        If isSave Then
            .lpstrTitle = "Сохранить CSV"
            .Flags = OFN_EXPLORER Or OFN_PATHMUSTEXIST Or OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY
            .lpstrDefExt = "csv"
            ok = apiGetSaveFileName(OFN)
        Else
            .lpstrTitle = "Открыть CSV"
            .Flags = OFN_EXPLORER Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
            ok = apiGetOpenFileName(OFN)
        End If
        If ok <> 0 Then
            ShowFileDialogAPI = TrimNull(.lpstrFile)
            Exit Function
        End If
    End With
EH:
    ShowFileDialogAPI = ShowFileDialogAPI_InputBox(isSave, defDir, defName)
End Function

Private Function ShowFileDialogAPI_InputBox(ByVal isSave As Boolean, ByVal defDir As String, ByVal defName As String) As String
    Dim ttl As String
    If isSave Then ttl = "Сохранить CSV" Else ttl = "Открыть CSV"
    Dim s As String
    s = Trim$(InputBox("Полный путь к файлу (" & defName & "):", ttl, defDir & defName))
    ShowFileDialogAPI_InputBox = s
End Function

' База пользовательских типов фасадов (параметры на вкладке «Цена»)
Public Function FacadeTypesCsvPath(Optional ByVal MacroListHide As Byte = 0) As String
    FacadeTypesCsvPath = Environ("APPDATA") & "\MebelShapeBuilder\facade_types.csv"
End Function

Private Sub EnsureFacadeTypesAppFolder()
    On Error Resume Next
    MkDir Environ("APPDATA") & "\MebelShapeBuilder"
End Sub

Public Sub DeleteFacadeTypeByName(ByVal nm As String)
    On Error Resume Next
    Dim p As String, fh As Integer
    Dim line As String, parts() As String
    Dim out As String, hit As Boolean
    p = FacadeTypesCsvPath()
    If Dir(p) = "" Then Exit Sub
    fh = FreeFile
    Open p For Input As #fh
    Line Input #fh, line
    out = line & vbCrLf
    Do While Not EOF(fh)
        Line Input #fh, line
        If Len(Trim$(line)) > 0 Then
            parts = Split(line, ";")
            If UBound(parts) >= 0 Then
                If StrComp(Trim$(parts(0)), nm, vbTextCompare) = 0 Then
                    hit = True
                Else
                    out = out & line & vbCrLf
                End If
            End If
        End If
    Loop
    Close #fh
    If Not hit Then Exit Sub
    fh = FreeFile
    Open p For Output As #fh
    Print #fh, out;
    Close #fh
End Sub

Public Function FacadeTypesCsvHeaderLine(Optional ByVal MacroListHide As Byte = 0) As String
    FacadeTypesCsvHeaderLine = "Name;Kind;InsetX;InsetY;ArcTop;ArcSide;OuterQty;OuterMm;InnerQty;InnerMm;NeoSegIdx;NeoSegMm;BorderIdx;OrnIdx"
End Function

Public Sub UpsertFacadeTypeRow(ByVal csvLine As String)
    On Error Resume Next
    Dim parts() As String
    parts = Split(csvLine, ";")
    If UBound(parts) < 1 Then Exit Sub
    Dim nm As String
    nm = Trim$(parts(0))
    If Len(nm) = 0 Then Exit Sub
    EnsureFacadeTypesAppFolder
    DeleteFacadeTypeByName nm
    Dim p As String, fh As Integer
    p = FacadeTypesCsvPath()
    If Dir(p) = "" Then
        fh = FreeFile
        Open p For Output As #fh
        Print #fh, FacadeTypesCsvHeaderLine()
        Print #fh, csvLine
        Close #fh
    Else
        fh = FreeFile
        Open p For Append As #fh
        Print #fh, csvLine
        Close #fh
    End If
End Sub
