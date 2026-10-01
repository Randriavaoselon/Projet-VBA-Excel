Attribute VB_Name = "modGraphiques"
Option Explicit

' ============================================================
'  Graphiques : colonnes (Top 15), camembert (Top 8), courbe mensuelle
'  Les colonnes utilisees (categorie/valeur/date) sont detectees
'  automatiquement a chaque chargement de table, et modifiables
'  via des listes deroulantes sur le Dashboard.
' ============================================================

Private Const NB_TOP_BARRES As Long = 15
Private Const NB_TOP_CAMEMBERT As Long = 8

Public Sub SupprimerGraphiques()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(FEUILLE_DASH)
    Do While ws.ChartObjects.Count > 0
        ws.ChartObjects(1).Delete
    Loop
End Sub

' Devine des colonnes categorie/valeur/date/agregation raisonnables si les
' cellules du Dashboard sont vides ou ne correspondent plus a la table chargee.
' L'utilisateur garde la main : une valeur qu'il a choisie et qui existe
' toujours dans la table n'est jamais ecrasee.
Public Sub AssurerParametresGraphiques(ByVal wc As Worksheet)
    Dim ws As Worksheet, i As Long, n As Long
    Dim premT As String, premN As String, premAutreN As String, premD As String
    Dim nom As String, cls As String, valeurAuto As String

    Set ws = ThisWorkbook.Worksheets(FEUILLE_DASH)
    n = 0
    Do While wc.Cells(n + 2, 1).Value <> ""
        n = n + 1
    Loop

    For i = 1 To n
        nom = wc.Cells(i + 1, 1).Value
        cls = wc.Cells(i + 1, 2).Value
        If cls = "T" And premT = "" Then premT = nom
        If cls = "N" Then
            If premN = "" Then
                premN = nom
            ElseIf premAutreN = "" Then
                premAutreN = nom
            End If
        End If
        If cls = "D" And premD = "" Then premD = nom
    Next i

    ' On evite en priorite la 1re colonne numerique (souvent un identifiant)
    If premAutreN <> "" Then valeurAuto = premAutreN Else valeurAuto = premN

    CorrigerCelluleParam ws.Range(CELL_CAT), premT, wc
    CorrigerCelluleParam ws.Range(CELL_VAL), valeurAuto, wc
    CorrigerCelluleParam ws.Range(CELL_DATECOL), premD, wc

    If Trim$(ws.Range(CELL_AGG).Value) = "" Then
        ws.Range(CELL_AGG).Value = IIf(valeurAuto = "", "NOMBRE", "SOMME")
    End If
    If Trim$(ws.Range(CELL_DOUBLONS).Value) = "" Then
        ws.Range(CELL_DOUBLONS).Value = "Surligner"
    End If
End Sub

Private Sub CorrigerCelluleParam(ByVal cel As Range, ByVal valeurAuto As String, ByVal wc As Worksheet)
    Dim actuel As String, r As Variant
    actuel = Trim$(CStr(cel.Value))
    If actuel <> "" Then
        r = Application.Match(actuel, wc.Range("A2:A2000"), 0)
        If Not IsError(r) Then Exit Sub ' la colonne choisie existe encore -> on la garde
    End If
    cel.Value = valeurAuto
End Sub

' Retourne un message si les graphiques n'ont pas pu etre construits ("" sinon)
Public Function MettreAJourGraphiques(ByVal lo As ListObject, ByRef estDbl() As Boolean) As String
    Dim ws As Worksheet, wc As Worksheet
    Dim d As Variant, tmp() As Variant, v As Variant
    Dim iCat As Long, iVal As Long, iDate As Long
    Dim nomCat As String, nomVal As String, nomDate As String, agg As String, nomSerie As String
    Dim dSom As Object, dNb As Object, mSom As Object, mNb As Object
    Dim r As Long, k As String, x As Double, ok As Boolean
    Dim nbCat As Long, nbPie As Long, nbMois As Long, idx As Long
    Dim nbLignesValides As Long, totalSomme As Double, totalCountValeur As Long

    Set ws = ThisWorkbook.Worksheets(FEUILLE_DASH)
    Set wc = ThisWorkbook.Worksheets(FEUILLE_CALC)
    SupprimerGraphiques

    nomCat = Trim$(ws.Range(CELL_CAT).Value)
    nomVal = Trim$(ws.Range(CELL_VAL).Value)
    nomDate = Trim$(ws.Range(CELL_DATECOL).Value)
    iCat = IndexColonne(lo, nomCat)
    iVal = IndexColonne(lo, nomVal)
    iDate = IndexColonne(lo, nomDate)

    If iCat = 0 And iDate = 0 Then
        MettreAJourGraphiques = "Graphiques : choisissez une colonne Categorie et/ou Date valide dans les parametres du Dashboard"
        Exit Function
    End If

    agg = UCase$(Trim$(ws.Range(CELL_AGG).Value))
    If agg <> "MOYENNE" And agg <> "NOMBRE" Then agg = "SOMME"
    If iVal = 0 Then agg = "NOMBRE"
    nomSerie = LibelleAgg(agg, nomVal)

    ' Lecture des donnees du tableau
    d = lo.DataBodyRange.Value
    If Not IsArray(d) Then
        ReDim tmp(1 To 1, 1 To 1)
        tmp(1, 1) = d
        d = tmp
    End If

    Set dSom = NouveauDico(): Set dNb = NouveauDico()
    Set mSom = NouveauDico(): Set mNb = NouveauDico()

    For r = 1 To UBound(d, 1)
        If estDbl(r) Then GoTo Suivant
        nbLignesValides = nbLignesValides + 1

        ok = True
        x = 0
        If iVal > 0 Then
            v = d(r, iVal)
            If IsEmpty(v) Then
                ok = False
            ElseIf IsNumeric(v) Then
                x = CDbl(v)
                totalSomme = totalSomme + x
                totalCountValeur = totalCountValeur + 1
            Else
                ok = False
            End If
        End If

        If ok Or agg = "NOMBRE" Then
            If iCat > 0 Then
                k = Txt(d(r, iCat))
                If k = "" Then k = "(vide)"
                Accumuler dSom, dNb, k, x
            End If
            If iDate > 0 Then
                v = d(r, iDate)
                If IsDate(v) Then Accumuler mSom, mNb, Format(CDate(v), "yyyy-mm"), x
            End If
        End If
Suivant:
    Next r

    ' Resume chiffre (KPI), mis a jour a chaque chargement/recherche
    ws.Range(CELL_KPI_LIGNES).Value = nbLignesValides
    If iVal > 0 And totalCountValeur > 0 Then
        ws.Range(CELL_KPI_SOMME).Value = Format(totalSomme, "#,##0.##")
        ws.Range(CELL_KPI_MOYENNE).Value = Format(totalSomme / totalCountValeur, "#,##0.##")
    Else
        ws.Range(CELL_KPI_SOMME).Value = "-"
        ws.Range(CELL_KPI_MOYENNE).Value = "-"
    End If
    If iCat > 0 Then
        ws.Range(CELL_KPI_CATEGORIES).Value = dSom.Count
    Else
        ws.Range(CELL_KPI_CATEGORIES).Value = "-"
    End If
    ws.Cells(LIGNE_KPI_LABELS, 3).Value = "Somme" & IIf(nomVal <> "", " de " & nomVal, "")
    ws.Cells(LIGNE_KPI_LABELS, 5).Value = "Moyenne" & IIf(nomVal <> "", " de " & nomVal, "")
    ws.Cells(LIGNE_KPI_LABELS, 7).Value = "Categories" & IIf(nomCat <> "", " (" & nomCat & ")", " uniques")

    ' Ecriture des donnees agregees
    If iCat > 0 Then
        EcrireTop wc, 4, dSom, dNb, agg, nomSerie, NB_TOP_BARRES, False, nbCat
        If agg <> "MOYENNE" Then EcrireTop wc, 7, dSom, dNb, agg, nomSerie, NB_TOP_CAMEMBERT, True, nbPie
    End If
    If iDate > 0 Then EcrireMois wc, 10, mSom, mNb, agg, nomSerie, nbMois

    ' Creation des graphiques
    idx = 0
    If nbCat > 0 Then
        AjouterGraphique ws, xlColumnClustered, idx, "chCategories", _
            nomSerie & " par " & nomCat & " (Top " & NB_TOP_BARRES & ")", _
            wc.Range(wc.Cells(2, 4), wc.Cells(nbCat + 1, 4)), _
            wc.Range(wc.Cells(2, 5), wc.Cells(nbCat + 1, 5)), nomSerie
        idx = idx + 1
    End If
    If nbPie > 0 Then
        AjouterGraphique ws, xlPie, idx, "chRepartition", _
            "Repartition par " & nomCat & " (Top " & NB_TOP_CAMEMBERT & ")", _
            wc.Range(wc.Cells(2, 7), wc.Cells(nbPie + 1, 7)), _
            wc.Range(wc.Cells(2, 8), wc.Cells(nbPie + 1, 8)), nomSerie
        idx = idx + 1
    End If
    If nbMois > 0 Then
        AjouterGraphique ws, xlLineMarkers, idx, "chEvolution", _
            "Evolution mensuelle - " & nomSerie, _
            wc.Range(wc.Cells(2, 10), wc.Cells(nbMois + 1, 10)), _
            wc.Range(wc.Cells(2, 11), wc.Cells(nbMois + 1, 11)), nomSerie
    End If

    ' Segment cliquable (filtre visuel) sur la colonne categorie, si raisonnable
    GererSegmentCategorie lo, nomCat, IIf(iCat > 0, dSom.Count, 0)
End Function

' ---------- Utilitaires ----------

Private Function IndexColonne(ByVal lo As ListObject, ByVal nom As String) As Long
    Dim r As Variant
    If nom = "" Then Exit Function
    r = Application.Match(nom, lo.HeaderRowRange, 0)
    If Not IsError(r) Then IndexColonne = CLng(r)
End Function

Private Function NouveauDico() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set NouveauDico = d
End Function

Private Sub Accumuler(ByVal dSom As Object, ByVal dNb As Object, ByVal k As String, ByVal x As Double)
    If dSom.Exists(k) Then
        dSom(k) = dSom(k) + x
        dNb(k) = dNb(k) + 1
    Else
        dSom.Add k, x
        dNb.Add k, 1
    End If
End Sub

Private Function ValeurAgg(ByVal dSom As Object, ByVal dNb As Object, ByVal k As String, ByVal agg As String) As Double
    Select Case agg
    Case "MOYENNE"
        If dNb(k) > 0 Then ValeurAgg = dSom(k) / dNb(k)
    Case "NOMBRE"
        ValeurAgg = dNb(k)
    Case Else
        ValeurAgg = dSom(k)
    End Select
End Function

Private Function LibelleAgg(ByVal agg As String, ByVal nomVal As String) As String
    Select Case agg
    Case "SOMME"
        LibelleAgg = "Somme de " & nomVal
    Case "MOYENNE"
        LibelleAgg = "Moyenne de " & nomVal
    Case Else
        LibelleAgg = "Nombre de lignes"
    End Select
End Function

Private Sub PreparerZone(ByVal wc As Worksheet, ByVal colDeb As Long)
    wc.Range(wc.Cells(1, colDeb), wc.Cells(5000, colDeb + 1)).ClearContents
    wc.Range(wc.Cells(1, colDeb), wc.Cells(5000, colDeb)).NumberFormat = "@"
End Sub

' Ecrit les N plus grandes categories (+ ligne "Autres" optionnelle)
Private Sub EcrireTop(ByVal wc As Worksheet, ByVal colDeb As Long, ByVal dSom As Object, ByVal dNb As Object, _
                      ByVal agg As String, ByVal nomSerie As String, ByVal n As Long, _
                      ByVal avecAutres As Boolean, ByRef nbLignes As Long)
    Dim m As Long, i As Long, j As Long, best As Long, lim As Long
    Dim cles() As String, vals() As Double, kk As Variant
    Dim tk As String, tv As Double, autres As Double

    nbLignes = 0
    PreparerZone wc, colDeb
    m = dSom.Count
    If m = 0 Then Exit Sub

    ReDim cles(0 To m - 1)
    ReDim vals(0 To m - 1)
    i = 0
    For Each kk In dSom.Keys
        cles(i) = CStr(kk)
        vals(i) = ValeurAgg(dSom, dNb, cles(i), agg)
        i = i + 1
    Next kk

    lim = n
    If m < lim Then lim = m

    ' Tri partiel decroissant (les 'lim' premiers)
    For i = 0 To lim - 1
        best = i
        For j = i + 1 To m - 1
            If vals(j) > vals(best) Then best = j
        Next j
        If best <> i Then
            tk = cles(i): cles(i) = cles(best): cles(best) = tk
            tv = vals(i): vals(i) = vals(best): vals(best) = tv
        End If
    Next i

    wc.Cells(1, colDeb).Value = "Categorie"
    wc.Cells(1, colDeb + 1).Value = nomSerie
    For i = 0 To lim - 1
        wc.Cells(i + 2, colDeb).Value = cles(i)
        wc.Cells(i + 2, colDeb + 1).Value = vals(i)
    Next i
    nbLignes = lim

    If avecAutres And m > lim And agg <> "MOYENNE" Then
        For j = lim To m - 1
            autres = autres + vals(j)
        Next j
        wc.Cells(lim + 2, colDeb).Value = "Autres"
        wc.Cells(lim + 2, colDeb + 1).Value = autres
        nbLignes = lim + 1
    End If
End Sub

' Ecrit la serie mensuelle triee par ordre chronologique
Private Sub EcrireMois(ByVal wc As Worksheet, ByVal colDeb As Long, ByVal mSom As Object, ByVal mNb As Object, _
                       ByVal agg As String, ByVal nomSerie As String, ByRef nbLignes As Long)
    Dim cles() As String, i As Long, j As Long, t As String, kk As Variant, m As Long

    nbLignes = 0
    PreparerZone wc, colDeb
    m = mSom.Count
    If m = 0 Then Exit Sub

    ReDim cles(0 To m - 1)
    i = 0
    For Each kk In mSom.Keys
        cles(i) = CStr(kk)
        i = i + 1
    Next kk

    ' Tri par insertion (format aaaa-mm : l'ordre alphabetique = ordre chronologique)
    For i = 1 To m - 1
        t = cles(i)
        j = i - 1
        Do While j >= 0
            If cles(j) <= t Then Exit Do
            cles(j + 1) = cles(j)
            j = j - 1
        Loop
        cles(j + 1) = t
    Next i

    wc.Cells(1, colDeb).Value = "Mois"
    wc.Cells(1, colDeb + 1).Value = nomSerie
    For i = 0 To m - 1
        wc.Cells(i + 2, colDeb).Value = cles(i)
        wc.Cells(i + 2, colDeb + 1).Value = ValeurAgg(mSom, mNb, cles(i), agg)
    Next i
    nbLignes = m
End Sub

Private Sub AjouterGraphique(ByVal ws As Worksheet, ByVal typeGraph As Long, ByVal idx As Long, _
                             ByVal nom As String, ByVal titre As String, _
                             ByVal rCat As Range, ByVal rVal As Range, ByVal nomSerie As String)
    Dim co As ChartObject

    Set co = ws.ChartObjects.Add(ws.Columns(1).Left + 5 + idx * 385, ws.Rows(LIGNE_GRAPHIQUES).Top, 375, 270)
    co.Name = nom
    co.Placement = xlFreeFloating

    With co.Chart
        .ChartType = typeGraph
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        With .SeriesCollection.NewSeries
            .Name = nomSerie
            .Values = rVal
            .XValues = rCat
        End With

        .HasTitle = True
        .ChartTitle.Text = titre
        .ChartTitle.Font.Size = 12

        If typeGraph = xlPie Then
            .HasLegend = True
            .Legend.Position = xlLegendPositionRight
            .SeriesCollection(1).HasDataLabels = True
            With .SeriesCollection(1).DataLabels
                .ShowPercentage = True
                .ShowValue = False
                .ShowCategoryName = False
            End With
        Else
            .HasLegend = False
            If typeGraph = xlColumnClustered Then
                .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(31, 78, 121)
            Else
                .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(31, 78, 121)
            End If
        End If
    End With
End Sub

' ---------- Segment cliquable (filtre visuel de type "dashboard") ----------

' Recree le segment attache a la colonne categorie a chaque chargement.
' Si la colonne a trop de valeurs distinctes (au-dela de MAX_CATEGORIES_SEGMENT),
' ou si aucune categorie n'est choisie, le segment est simplement retire :
' un segment avec des centaines de boutons serait illisible.
Private Sub GererSegmentCategorie(ByVal lo As ListObject, ByVal nomCat As String, ByVal nbCategories As Long)
    Dim sc As SlicerCache, sl As Slicer, ws As Worksheet

    SupprimerSegmentsExistants

    If nomCat = "" Or nbCategories = 0 Or nbCategories > MAX_CATEGORIES_SEGMENT Then Exit Sub

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(FEUILLE_DASH)
    Set sc = ThisWorkbook.SlicerCaches.Add2(lo, nomCat)
    If Not sc Is Nothing Then
        Set sl = sc.Slicers.Add(ws, , sc.Name, nomCat, _
            ws.Cells(LIGNE_KPI_TITRE, 17).Top, ws.Cells(LIGNE_KPI_TITRE, 17).Left, 170, 210)
    End If
    On Error GoTo 0
End Sub

Private Sub SupprimerSegmentsExistants()
    Dim i As Long
    On Error Resume Next
    For i = ThisWorkbook.SlicerCaches.Count To 1 Step -1
        If InStr(1, ThisWorkbook.SlicerCaches(i).SourceName, NOM_TABLEAU, vbTextCompare) > 0 Then
            ThisWorkbook.SlicerCaches(i).Delete
        End If
    Next i
    On Error GoTo 0
End Sub
