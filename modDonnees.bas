Attribute VB_Name = "modDonnees"
Option Explicit

' ============================================================
'  Chargement des donnees, gestion des doublons, recherche avancee
' ============================================================

' ---------- Macros liees aux boutons ----------

Public Sub Charger_Table()
    If Not ApplicationPrete() Then Exit Sub
    If TableSelectionnee() = "" Then
        MsgBox "Cliquez d'abord sur 'Connecter', puis choisissez une table dans la liste.", vbInformation
        Exit Sub
    End If
    ViderCriteres
    ChargerDonnees ""
End Sub

Public Sub Rechercher()
    Dim clause As String

    If Not ApplicationPrete() Then Exit Sub
    If ThisWorkbook.Worksheets(FEUILLE_CALC).Range("A2").Value = "" Then
        MsgBox "Cliquez d'abord sur 'Connecter' puis 'Charger la table' pour recuperer la liste des colonnes.", vbInformation
        Exit Sub
    End If

    On Error GoTo Erreur
    clause = ConstruireWhere()
    ChargerDonnees clause
    Exit Sub

Erreur:
    MsgBox Err.Description, vbExclamation, "Recherche"
End Sub

Public Sub Effacer_Criteres()
    If Not ApplicationPrete() Then Exit Sub
    ViderCriteres
    ThisWorkbook.Worksheets(FEUILLE_DASH).Range("B" & LIGNE_RESULTAT).Value = _
        "Criteres effaces. Cliquez sur 'Charger la table' pour recharger toutes les donnees."
End Sub

Private Sub ViderCriteres()
    With ThisWorkbook.Worksheets(FEUILLE_DASH)
        .Range("B" & LIGNE_RECH_CRIT1 & ":D" & (LIGNE_RECH_CRIT1 + 2)).ClearContents
        .Range("B" & LIGNE_RECH_PERIODE).ClearContents
        .Range("D" & LIGNE_RECH_PERIODE).ClearContents
        .Range("A" & (LIGNE_RECH_CRIT1 + 1) & ":A" & (LIGNE_RECH_CRIT1 + 2)).Value = "ET"
        .Range(CELL_RECHERCHE_RAPIDE).ClearContents
    End With
    RechercheRapide_Filtrer   ' retire le filtre instantane eventuellement actif
End Sub

' ---------- Chargement ----------

Private Sub ChargerDonnees(ByVal clauseWhere As String)
    Dim cn As Object, rs As Object
    Dim ws As Worksheet, wc As Worksheet, lo As ListObject
    Dim sql As String, msg As String, msgGraph As String, modeDbl As String
    Dim brut As Variant, sortie() As Variant, estDbl() As Boolean
    Dim noms() As String
    Dim nbCols As Long, nbBrut As Long, nbSortie As Long, nbDbl As Long
    Dim limite As Long, i As Long

    On Error GoTo Erreur

    If TableSelectionnee() = "" Then
        MsgBox "Choisissez une table dans la liste du Dashboard (bouton 'Connecter' d'abord).", vbExclamation
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets(FEUILLE_DASH)
    Set wc = ThisWorkbook.Worksheets(FEUILLE_CALC)
    Application.ScreenUpdating = False
    Application.Cursor = xlWait
    Application.StatusBar = "Connexion a la base de donnees..."

    Set cn = OuvrirConnexion()
    ChargerSchema cn
    AssurerParametresGraphiques wc

    sql = "SELECT * FROM " & Q(TableSelectionnee())
    If clauseWhere <> "" Then sql = sql & " WHERE " & clauseWhere

    limite = LIMITE_LIGNES_DEFAUT
    Set rs = CreateObject("ADODB.Recordset")
    If limite > 0 Then rs.MaxRecords = limite

    Application.StatusBar = "Lecture des donnees..."
    rs.Open sql, cn, ADO_STATIC, ADO_LECTURE_SEULE

    nbCols = rs.Fields.Count
    ReDim noms(0 To nbCols - 1)
    For i = 0 To nbCols - 1
        noms(i) = rs.Fields(i).Name
    Next i
    If Not rs.EOF Then brut = rs.GetRows
    rs.Close
    cn.Close

    ViderTableau ws

    If IsEmpty(brut) Then
        SupprimerGraphiques
        ws.Range("B" & LIGNE_RESULTAT).Value = "Aucun resultat pour cette requete."
        GoTo Fin
    End If

    nbBrut = UBound(brut, 2) + 1
    modeDbl = ModeDoublons()

    Application.StatusBar = "Traitement des doublons..."
    PreparerSortie brut, noms, sortie, estDbl, nbSortie, nbDbl, modeDbl

    ' En-tetes puis donnees
    Application.StatusBar = "Ecriture dans la feuille..."
    ws.Range(ws.Cells(LIGNE_TABLEAU, 1), ws.Cells(LIGNE_TABLEAU, nbCols)).NumberFormat = "@"
    For i = 0 To nbCols - 1
        ws.Cells(LIGNE_TABLEAU, i + 1).Value = noms(i)
    Next i
    ws.Range(ws.Cells(LIGNE_TABLEAU + 1, 1), ws.Cells(LIGNE_TABLEAU + nbSortie, nbCols)).Value = sortie

    Set lo = ws.ListObjects.Add(xlSrcRange, _
        ws.Range(ws.Cells(LIGNE_TABLEAU, 1), ws.Cells(LIGNE_TABLEAU + nbSortie, nbCols)), , xlYes)
    lo.Name = NOM_TABLEAU
    lo.TableStyle = "TableStyleMedium2"

    ' Surlignage des doublons (mode Surligner)
    If nbDbl > 0 Then
        For i = 1 To nbSortie
            If estDbl(i) Then lo.DataBodyRange.Rows(i).Interior.Color = RGB(255, 199, 206)
        Next i
    End If

    AjusterLargeurs ws, nbCols, nbSortie
    AppliquerMiseEnForme lo, wc
    AjouterColonneRecherche lo, sortie, nbCols, nbSortie
    RechercheRapide_Filtrer   ' reapplique le filtre instantane si l'utilisateur avait deja tape quelque chose

    ' Graphiques et resume chiffre (KPI)
    Application.StatusBar = "Mise a jour des graphiques..."
    msgGraph = MettreAJourGraphiques(lo, estDbl)

    ' Message de resultat
    msg = Format(nbSortie, "#,##0") & " ligne(s) affichee(s)"
    If nbDbl > 0 Then
        If modeDbl = "SUPPRIMER" Then
            msg = msg & " | " & nbDbl & " doublon(s) supprime(s)"
        Else
            msg = msg & " | " & nbDbl & " doublon(s) surligne(s) en rose (exclus des graphiques)"
        End If
    End If
    If limite > 0 And nbBrut >= limite Then
        msg = msg & " | LIMITE ATTEINTE (" & limite & ") : d'autres lignes existent en base"
    End If
    If clauseWhere <> "" Then msg = msg & " | recherche appliquee"
    If msgGraph <> "" Then msg = msg & " | " & msgGraph
    ws.Range("B" & LIGNE_RESULTAT).Value = msg & " | " & Format(Now, "dd/mm/yyyy hh:nn:ss")

Fin:
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    Application.StatusBar = False
    Exit Sub

Erreur:
    msg = "Erreur : " & Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
    End If
    If Not cn Is Nothing Then
        If cn.State <> 0 Then cn.Close
    End If
    On Error GoTo 0
    MsgBox msg, vbCritical, "Chargement des donnees"
    Resume Fin
End Sub

Private Sub ViderTableau(ByVal ws As Worksheet)
    Dim derniere As Long
    Do While ws.ListObjects.Count > 0
        ws.ListObjects(1).Delete
    Loop
    derniere = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If derniere >= LIGNE_TABLEAU Then ws.Rows(LIGNE_TABLEAU & ":" & derniere).Clear
End Sub

Private Sub AjusterLargeurs(ByVal ws As Worksheet, ByVal nbCols As Long, ByVal nbLignes As Long)
    Dim i As Long
    ws.Range(ws.Cells(LIGNE_TABLEAU, 1), ws.Cells(LIGNE_TABLEAU + nbLignes, nbCols)).Columns.AutoFit
    For i = 1 To nbCols
        If ws.Columns(i).ColumnWidth > 45 Then ws.Columns(i).ColumnWidth = 45
        If i <= 4 Then
            If ws.Columns(i).ColumnWidth < 22 Then ws.Columns(i).ColumnWidth = 22
        ElseIf ws.Columns(i).ColumnWidth < 12 Then
            ws.Columns(i).ColumnWidth = 12
        End If
    Next i
End Sub

' ---------- Doublons ----------
' Produit le tableau 'sortie' (Variant 2D) et le repere des doublons.
' modeDbl : SUPPRIMER (retire les doublons), SURLIGNER (les garde, marques), CONSERVER (ne fait rien)
' Un doublon = une ligne identique a une ligne precedente sur TOUTES les colonnes
' (plus simple et fiable qu'une configuration de colonnes cles a saisir).

Private Sub PreparerSortie(ByRef brut As Variant, ByRef noms() As String, ByRef sortie() As Variant, _
                           ByRef estDbl() As Boolean, ByRef nbSortie As Long, ByRef nbDbl As Long, _
                           ByVal modeDbl As String)
    Dim nbCols As Long, nbBrut As Long, r As Long, c As Long, o As Long
    Dim dic As Object, cle As String
    Dim dblBrut() As Boolean, v As Variant

    nbCols = UBound(brut, 1) + 1
    nbBrut = UBound(brut, 2) + 1
    nbDbl = 0

    ReDim dblBrut(0 To nbBrut - 1)
    If modeDbl <> "CONSERVER" Then
        Set dic = CreateObject("Scripting.Dictionary")
        dic.CompareMode = vbTextCompare
        For r = 0 To nbBrut - 1
            cle = ""
            For c = 0 To nbCols - 1
                cle = cle & Txt(brut(c, r)) & Chr$(1)
            Next c
            If dic.Exists(cle) Then
                dblBrut(r) = True
                nbDbl = nbDbl + 1
            Else
                dic.Add cle, 1
            End If
        Next r
    End If

    If modeDbl = "SUPPRIMER" Then nbSortie = nbBrut - nbDbl Else nbSortie = nbBrut

    ReDim sortie(1 To nbSortie, 1 To nbCols)
    ReDim estDbl(1 To nbSortie)

    o = 0
    For r = 0 To nbBrut - 1
        If Not (modeDbl = "SUPPRIMER" And dblBrut(r)) Then
            o = o + 1
            estDbl(o) = dblBrut(r)
            For c = 0 To nbCols - 1
                v = brut(c, r)
                If IsNull(v) Then
                    v = Empty
                ElseIf IsArray(v) Then
                    v = "(binaire)"
                ElseIf VarType(v) = vbDecimal Then
                    v = CDbl(v)
                ElseIf VarType(v) = vbString Then
                    If Left$(v, 1) = "=" Then v = "'" & v      ' evite l'execution de formules
                End If
                sortie(o, c + 1) = v
            Next c
        End If
    Next r
End Sub

' ---------- Recherche avancee ----------

Private Function ConstruireWhere() As String
    Dim ws As Worksheet, i As Long
    Dim col As String, op As String, valeur As String, lien As String
    Dim cond As String, clause As String
    Dim d1 As Variant, d2 As Variant, colDate As String, colReelle As String, cls As String

    Set ws = ThisWorkbook.Worksheets(FEUILLE_DASH)

    ' Criteres 1 a 3
    For i = LIGNE_RECH_CRIT1 To LIGNE_RECH_CRIT1 + 2
        col = Trim$(CStr(ws.Cells(i, 2).Value))
        op = Trim$(CStr(ws.Cells(i, 3).Value))
        valeur = Trim$(CStr(ws.Cells(i, 4).Value))
        If col <> "" And op <> "" Then
            If valeur <> "" Or op = "est vide" Or op = "n'est pas vide" Then
                cond = ConditionSQL(col, op, valeur)
                If clause = "" Then
                    clause = cond
                Else
                    If UCase$(Trim$(CStr(ws.Cells(i, 1).Value))) = "OU" Then lien = "OR" Else lien = "AND"
                    clause = "(" & clause & ") " & lien & " (" & cond & ")"
                End If
            End If
        End If
    Next i

    ' Periode (colonne date choisie pour les graphiques)
    d1 = ws.Range("B" & LIGNE_RECH_PERIODE).Value
    d2 = ws.Range("D" & LIGNE_RECH_PERIODE).Value
    If Trim$(CStr(d1)) <> "" Or Trim$(CStr(d2)) <> "" Then
        colDate = Trim$(ws.Range(CELL_DATECOL).Value)
        If colDate = "" Then Err.Raise vbObjectError + 300, , "Choisissez une 'Date' dans les parametres des graphiques pour filtrer par periode."
        colReelle = ResoudreColonne(colDate, cls)
        If colReelle = "" Then Err.Raise vbObjectError + 301, , "Colonne date introuvable : " & colDate
        If cls <> "D" Then Err.Raise vbObjectError + 302, , "La colonne " & colReelle & " n'est pas de type date."

        If Trim$(CStr(d1)) <> "" Then
            If Not IsDate(d1) Then Err.Raise vbObjectError + 303, , "Date de debut invalide."
            cond = Q(colReelle) & " >= " & LitteralDate(CDate(Int(CDate(d1))))
            If clause = "" Then clause = cond Else clause = "(" & clause & ") AND " & cond
        End If
        If Trim$(CStr(d2)) <> "" Then
            If Not IsDate(d2) Then Err.Raise vbObjectError + 304, , "Date de fin invalide."
            cond = Q(colReelle) & " < " & LitteralDate(CDate(Int(CDate(d2))) + 1)
            If clause = "" Then clause = cond Else clause = "(" & clause & ") AND " & cond
        End If
    End If

    ConstruireWhere = clause
End Function

Private Function ConditionSQL(ByVal colSaisie As String, ByVal op As String, ByVal valeur As String) As String
    Dim col As String, cls As String, id As String, v As String, sym As String
    Dim d As Date, dLit As String, dSuiv As String

    col = ResoudreColonne(colSaisie, cls)
    If col = "" Then Err.Raise vbObjectError + 400, , "Colonne inconnue : " & colSaisie
    id = Q(col)

    If op = "est vide" Then
        If cls = "T" Then
            ConditionSQL = "(" & id & " IS NULL OR " & id & " = '')"
        Else
            ConditionSQL = id & " IS NULL"
        End If
        Exit Function
    End If
    If op = "n'est pas vide" Then
        If cls = "T" Then
            ConditionSQL = "(" & id & " IS NOT NULL AND " & id & " <> '')"
        Else
            ConditionSQL = id & " IS NOT NULL"
        End If
        Exit Function
    End If

    sym = Symbole(op)

    Select Case cls
    Case "N"
        If sym = "" Then Err.Raise vbObjectError + 401, , "L'operateur '" & op & "' ne s'applique pas a la colonne numerique " & col & "."
        v = Replace(valeur, ",", ".")
        If Not EstNombre(v) Then Err.Raise vbObjectError + 402, , "Valeur numerique attendue pour la colonne " & col & "."
        ConditionSQL = id & " " & sym & " " & v

    Case "D"
        If sym = "" Then Err.Raise vbObjectError + 403, , "L'operateur '" & op & "' ne s'applique pas a la colonne date " & col & "."
        If Not IsDate(valeur) Then Err.Raise vbObjectError + 404, , "Date attendue (jj/mm/aaaa) pour la colonne " & col & "."
        d = CDate(Int(CDate(valeur)))
        dLit = LitteralDate(d)
        dSuiv = LitteralDate(d + 1)
        Select Case sym
        Case "="
            ConditionSQL = "(" & id & " >= " & dLit & " AND " & id & " < " & dSuiv & ")"
        Case "<>"
            ConditionSQL = "(" & id & " < " & dLit & " OR " & id & " >= " & dSuiv & ")"
        Case ">"
            ConditionSQL = id & " >= " & dSuiv
        Case ">="
            ConditionSQL = id & " >= " & dLit
        Case "<"
            ConditionSQL = id & " < " & dLit
        Case "<="
            ConditionSQL = id & " < " & dSuiv
        End Select

    Case Else   ' texte
        v = EchapperTexte(valeur)
        Dim joker As String
        joker = JokerMulti()   ' "%" pour MySQL/SQLServer/PostgreSQL, "*" pour Access
        If sym <> "" Then
            ConditionSQL = id & " " & sym & " " & Lit(v)
        Else
            Select Case op
            Case "contient"
                ConditionSQL = id & " LIKE " & Lit(joker & v & joker)
            Case "ne contient pas"
                ConditionSQL = id & " NOT LIKE " & Lit(joker & v & joker)
            Case "commence par"
                ConditionSQL = id & " LIKE " & Lit(v & joker)
            Case "se termine par"
                ConditionSQL = id & " LIKE " & Lit(joker & v)
            Case Else
                Err.Raise vbObjectError + 405, , "Operateur inconnu : " & op
            End Select
        End If
    End Select
End Function

Private Function Symbole(ByVal op As String) As String
    Select Case op
    Case "égal à":                   Symbole = "="
    Case "différent de":             Symbole = "<>"
    Case "supérieur à":              Symbole = ">"
    Case "supérieur ou égal à":      Symbole = ">="
    Case "inférieur à":              Symbole = "<"
    Case "inférieur ou égal à":      Symbole = "<="
    End Select
End Function

' Retrouve le nom exact et la classe (N/D/T) d'une colonne. Retourne "" si inconnue.
Private Function ResoudreColonne(ByVal saisie As String, ByRef classe As String) As String
    Dim wc As Worksheet, r As Variant
    Set wc = ThisWorkbook.Worksheets(FEUILLE_CALC)
    r = Application.Match(saisie, wc.Range("A2:A2000"), 0)
    If IsError(r) Then Exit Function
    ResoudreColonne = CStr(wc.Cells(CLng(r) + 1, 1).Value)
    classe = CStr(wc.Cells(CLng(r) + 1, 2).Value)
End Function

Private Function EstNombre(ByVal s As String) As Boolean
    Dim i As Long, c As String, pts As Long, chiffres As Long
    s = Trim$(s)
    If s = "" Then Exit Function
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        Select Case c
        Case "0" To "9"
            chiffres = chiffres + 1
        Case "."
            pts = pts + 1
        Case "-", "+"
            If i <> 1 Then Exit Function
        Case Else
            Exit Function
        End Select
    Next i
    EstNombre = (chiffres > 0 And pts <= 1)
End Function

' Echappe une valeur texte pour l'inserer dans un litteral SQL
Private Function EchapperTexte(ByVal s As String) As String
    If TypeSGBD() = "MYSQL" Then s = Replace(s, "\", "\\")
    EchapperTexte = Replace(s, "'", "''")
End Function

Private Function Lit(ByVal texteEchappe As String) As String
    If TypeSGBD() = "SQLSERVER" Then
        Lit = "N'" & texteEchappe & "'"
    Else
        Lit = "'" & texteEchappe & "'"
    End If
End Function

' ---------- Mise en forme automatique et recherche instantanee ----------

' Applique un format de nombre/date adapte a chaque colonne (selon sa classe N/D/T),
' l'alignement correspondant, et une ligne de totaux (nombre de lignes + sommes).
Private Sub AppliquerMiseEnForme(ByVal lo As ListObject, ByVal wc As Worksheet)
    Dim i As Long, nom As String, cls As String

    For i = 1 To lo.ListColumns.Count
        nom = lo.ListColumns(i).Name
        cls = ClasseColonne(wc, nom)
        Select Case cls
        Case "N"
            lo.ListColumns(i).DataBodyRange.NumberFormat = "#,##0.##"
            lo.ListColumns(i).DataBodyRange.HorizontalAlignment = xlRight
        Case "D"
            lo.ListColumns(i).DataBodyRange.NumberFormat = "dd/mm/yyyy"
            lo.ListColumns(i).DataBodyRange.HorizontalAlignment = xlCenter
        Case Else
            lo.ListColumns(i).DataBodyRange.HorizontalAlignment = xlLeft
        End Select
    Next i

    On Error Resume Next
    lo.ShowTotals = True
    For i = 1 To lo.ListColumns.Count
        nom = lo.ListColumns(i).Name
        cls = ClasseColonne(wc, nom)
        If i = 1 Then
            lo.ListColumns(i).TotalsCalculation = xlTotalsCalculationCount
        ElseIf cls = "N" Then
            lo.ListColumns(i).TotalsCalculation = xlTotalsCalculationSum
        Else
            lo.ListColumns(i).TotalsCalculation = xlTotalsCalculationNone
        End If
    Next i
    On Error GoTo 0
End Sub

Private Function ClasseColonne(ByVal wc As Worksheet, ByVal nom As String) As String
    Dim r As Variant
    r = Application.Match(nom, wc.Range("A2:A2000"), 0)
    If IsError(r) Then Exit Function
    ClasseColonne = CStr(wc.Cells(CLng(r) + 1, 2).Value)
End Function

' Ajoute une colonne technique masquee concatenant toutes les valeurs de chaque
' ligne (en majuscules) : c'est sur cette colonne que porte le filtre instantane
' de la case "Recherche rapide", ce qui evite une nouvelle requete a la base.
Private Sub AjouterColonneRecherche(ByVal lo As ListObject, ByRef sortie As Variant, _
                                    ByVal nbCols As Long, ByVal nbSortie As Long)
    Dim colRech As ListColumn, r As Long, c As Long, texteLigne As String
    Dim sortie2D() As Variant

    If nbSortie = 0 Then Exit Sub

    ReDim sortie2D(1 To nbSortie, 1 To 1)
    For r = 1 To nbSortie
        texteLigne = ""
        For c = 1 To nbCols
            texteLigne = texteLigne & "|" & UCase$(Txt(sortie(r, c)))
        Next c
        sortie2D(r, 1) = texteLigne
    Next r

    Set colRech = lo.ListColumns.Add
    colRech.Name = COL_RECHERCHE_CACHEE
    colRech.DataBodyRange.NumberFormat = "@"
    colRech.DataBodyRange.Value = sortie2D
    colRech.Range.EntireColumn.Hidden = True
End Sub

' Filtre instantanement le tableau deja charge selon le texte tape dans la case
' de recherche rapide, sans requeter la base (appelee automatiquement a chaque
' frappe via le code de la feuille Dashboard, et reappelee apres un rechargement).
Public Sub RechercheRapide_Filtrer()
    Dim ws As Worksheet, lo As ListObject, texte As String, idx As Long

    If Not ExisteFeuille(FEUILLE_DASH) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(FEUILLE_DASH)
    If ws.ListObjects.Count = 0 Then Exit Sub
    Set lo = ws.ListObjects(1)

    idx = 0
    On Error Resume Next
    idx = lo.ListColumns(COL_RECHERCHE_CACHEE).Index
    On Error GoTo 0
    If idx = 0 Then Exit Sub

    texte = TexteRechercheRapide()
    Application.ScreenUpdating = False
    On Error Resume Next
    If texte = "" Then
        lo.Range.AutoFilter Field:=idx
    Else
        lo.Range.AutoFilter Field:=idx, Criteria1:="*" & UCase$(texte) & "*"
    End If
    On Error GoTo 0
    Application.ScreenUpdating = True
End Sub
