Attribute VB_Name = "modSetup"
Option Explicit

' ============================================================
'  Initialisation : cree les feuilles Config, Dashboard et _Calc
'  (a executer une seule fois). Pour tout reinitialiser,
'  supprimez la feuille concernee puis relancez la macro.
' ============================================================

Public Sub InitialiserApplication()
    Dim wsC As Worksheet, wsX As Worksheet, wsD As Worksheet
    Dim cree As Boolean, injecte As Boolean

    Application.ScreenUpdating = False

    Set wsC = ObtenirFeuille(FEUILLE_CONFIG, cree)
    If cree Then ConstruireConfig wsC

    Set wsX = ObtenirFeuille(FEUILLE_CALC, cree)
    If cree Then ConstruireCalc wsX

    Set wsD = ObtenirFeuille(FEUILLE_DASH, cree)
    If cree Then ConstruireDashboard wsD

    wsD.Move Before:=ThisWorkbook.Worksheets(1)
    wsC.Move After:=wsD
    wsX.Visible = xlSheetHidden

    injecte = InjecterCodeFeuilleDashboard(wsD)

    wsC.Activate
    Application.ScreenUpdating = True

    Dim msg As String
    msg = "Application initialisee." & vbCrLf & vbCrLf & _
          "1) Dans la feuille Config, collez le lien de votre base et son mot de passe." & vbCrLf & _
          "2) Allez sur Dashboard, cliquez sur 'Connecter', puis choisissez une table."
    If injecte Then
        msg = msg & vbCrLf & vbCrLf & "Le chargement automatique et la recherche instantanee sont deja actifs."
    Else
        msg = msg & vbCrLf & vbCrLf & _
              "Pour activer le chargement automatique et la recherche instantanee," & vbCrLf & _
              "collez le code fourni dans le module de la feuille Dashboard" & vbCrLf & _
              "(voir Dashboard_CodeDeLaFeuille.txt), OU cochez 'Acces approuve au modele" & vbCrLf & _
              "d'objet du projet VBA' (Fichier > Options > Centre de gestion de la" & vbCrLf & _
              "confidentialite > Parametres des macros) puis relancez cette macro."
    End If
    MsgBox msg, vbInformation, "Initialisation"
End Sub

Private Function ObtenirFeuille(ByVal nom As String, ByRef cree As Boolean) As Worksheet
    Dim ws As Worksheet
    If ExisteFeuille(nom) Then
        Set ws = ThisWorkbook.Worksheets(nom)
        cree = False
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nom
        cree = True
    End If
    Set ObtenirFeuille = ws
End Function

' Tente d'inserer directement le code evenementiel dans le module de la feuille
' Dashboard (chargement automatique + recherche instantanee), via le modele
' d'objet VBA. Necessite que l'utilisateur ait coche, une seule fois :
' Fichier > Options > Centre de gestion de la confidentialite >
' Parametres des macros > "Acces approuve au modele d'objet de projet VBA".
' Si ce n'est pas le cas, echoue proprement (Function retourne False) et
' l'utilisateur suit alors les instructions manuelles fournies a part.
Private Function InjecterCodeFeuilleDashboard(ByVal wsD As Worksheet) As Boolean
    Dim comp As Object, mod_ As Object, i As Long, debut As Long, fin As Long
    Dim code As String

    On Error GoTo Echec
    Set comp = ThisWorkbook.VBProject.VBComponents(wsD.CodeName)
    Set mod_ = comp.CodeModule

    ' Retire une eventuelle version deja presente pour eviter les doublons
    debut = 0
    For i = 1 To mod_.CountOfLines
        If InStr(mod_.Lines(i, 1), "Worksheet_Change") > 0 Then
            debut = i
            Exit For
        End If
    Next i
    If debut > 0 Then
        fin = debut
        Do While fin < mod_.CountOfLines And InStr(mod_.Lines(fin, 1), "End Sub") = 0
            fin = fin + 1
        Loop
        mod_.DeleteLines debut, fin - debut + 1
    End If

    code = "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf & _
           "    On Error GoTo Fin" & vbCrLf & _
           "    If Not Intersect(Target, Me.Range(CELL_TABLE)) Is Nothing Then" & vbCrLf & _
           "        Application.EnableEvents = False" & vbCrLf & _
           "        modDonnees.Charger_Table" & vbCrLf & _
           "    ElseIf Not Intersect(Target, Me.Range(CELL_RECHERCHE_RAPIDE)) Is Nothing Then" & vbCrLf & _
           "        Application.EnableEvents = False" & vbCrLf & _
           "        modDonnees.RechercheRapide_Filtrer" & vbCrLf & _
           "    End If" & vbCrLf & _
           "Fin:" & vbCrLf & _
           "    Application.EnableEvents = True" & vbCrLf & _
           "End Sub"

    mod_.AddFromString code
    InjecterCodeFeuilleDashboard = True
    Exit Function

Echec:
    InjecterCodeFeuilleDashboard = False
End Function

' ---------- Feuille Config : uniquement le lien et le mot de passe ----------

Private Sub ConstruireConfig(ByVal ws As Worksheet)
    ws.Cells.Font.Name = "Calibri"
    With ws.Range("A1")
        .Value = "Connexion a la base de donnees"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(31, 78, 121)
    End With
    With ws.Range("A2")
        .Value = "Renseignez ces deux informations, puis allez dans l'onglet Dashboard."
        .Font.Italic = True
        .Font.Color = RGB(100, 100, 100)
    End With

    ws.Range("A3:C3").Value = Array("Parametre", "Valeur", "Exemple / Aide")
    With ws.Range("A3:C3")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
    End With

    ws.Cells(CFG_LIEN, 1).Value = "Lien de la base de donnees"
    ws.Cells(CFG_LIEN, 3).Value = "mysql://utilisateur@serveur:3306/nombase (ou postgresql://... ou sqlserver://...)" & _
        " -- pour Access, indiquez simplement le chemin du fichier, ex. C:\Dossier\Base.accdb"
    ws.Cells(CFG_MDP, 1).Value = "Mot de passe"
    ws.Cells(CFG_MDP, 3).Value = "Laissez vide pour qu'il soit demande a chaque connexion (plus sur). Pour Access, laissez vide si le fichier n'a pas de mot de passe."

    ws.Range("A4:A5").Font.Bold = True
    ws.Range("B4:B5").NumberFormat = "@"
    ws.Range("C4:C5").Font.Italic = True
    ws.Range("C4:C5").Font.Color = RGB(100, 100, 100)
    With ws.Range("B4:B5")
        .Interior.Color = RGB(255, 250, 220)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
    End With

    With ws.Range("A" & CFG_PILOTE)
        .Value = "Pilote ODBC (optionnel)"
        .Font.Bold = True
        .Font.Italic = True
        .Font.Color = RGB(140, 140, 140)
    End With
    ws.Cells(CFG_PILOTE, 3).Value = "A remplir SEULEMENT si 'Connecter' echoue malgre un lien et un mot de passe corrects." & _
        " Copiez le nom exact affiche dans Windows > Sources de donnees ODBC > onglet Pilotes."
    ws.Cells(CFG_PILOTE, 3).Font.Italic = True
    ws.Cells(CFG_PILOTE, 3).Font.Color = RGB(100, 100, 100)
    ws.Cells(CFG_PILOTE, 2).NumberFormat = "@"
    With ws.Cells(CFG_PILOTE, 2)
        .Interior.Color = RGB(240, 240, 240)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
    End With

    ws.Columns("A").ColumnWidth = 28
    ws.Columns("B").ColumnWidth = 40
    ws.Columns("C").ColumnWidth = 95
End Sub

' ---------- Feuille _Calc (masquee) ----------

Private Sub ConstruireCalc(ByVal ws As Worksheet)
    ws.Range("A:B").NumberFormat = "@"
    ws.Range("A1").Value = "Colonne"
    ws.Range("B1").Value = "Classe (N=numerique, D=date, T=texte)"
    ws.Range("D1").Value = "Donnees graphique colonnes"
    ws.Range("G1").Value = "Donnees graphique camembert"
    ws.Range("J1").Value = "Donnees courbe mensuelle"
    ws.Range("M1").Value = "Table"
    ThisWorkbook.Names.Add Name:="ListeColonnes", RefersTo:="='" & FEUILLE_CALC & "'!$A$2:$A$2"
    ThisWorkbook.Names.Add Name:="ListeTables", RefersTo:="='" & FEUILLE_CALC & "'!$M$2:$M$2"
End Sub

' ---------- Feuille Dashboard ----------
' Mise en page en "panneaux" cote a cote plutot qu'empiles verticalement :
' Connexion (gauche) + Recherche rapide (droite) partagent la meme bande.

Private Sub ConstruireDashboard(ByVal ws As Worksheet)
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    ws.Columns("A:D").ColumnWidth = 20
    ws.Columns("E:P").ColumnWidth = 13

    With ws.Range("A1")
        .Value = "Tableau de bord - Base de donnees"
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(31, 78, 121)
    End With

    ' --- Panneau CONNEXION (gauche) ---
    TitreSection ws.Range("A" & LIGNE_CONNEXION_TITRE & ":D" & LIGNE_CONNEXION_TITRE), "CONNEXION"
    ws.Cells(LIGNE_CONNEXION, 1).Value = "Table"
    ws.Cells(LIGNE_CONNEXION, 1).Font.Bold = True
    ws.Cells(LIGNE_CONNEXION, 3).Value = "Connecte a"
    ws.Cells(LIGNE_CONNEXION, 3).Font.Bold = True
    ws.Range(CELL_BASE_INFO).Value = "(non connecte)"
    ws.Range(CELL_BASE_INFO).Font.Italic = True
    With ws.Range(CELL_TABLE)
        .Interior.Color = RGB(255, 250, 220)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
    End With
    EncadrerPanneau ws.Range("A" & LIGNE_CONNEXION_TITRE & ":D" & LIGNE_CONNEXION)

    ' --- Panneau RECHERCHE RAPIDE (droite, meme bande que Connexion) ---
    TitreSection ws.Range("F" & LIGNE_CONNEXION_TITRE & ":J" & LIGNE_CONNEXION_TITRE), "RECHERCHE RAPIDE (filtre instantane)"
    ws.Cells(LIGNE_CONNEXION, 6).Value = "Taper ici :"
    ws.Cells(LIGNE_CONNEXION, 6).Font.Bold = True
    With ws.Range(CELL_RECHERCHE_RAPIDE & ":J" & LIGNE_CONNEXION)
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
    End With
    EncadrerPanneau ws.Range("F" & LIGNE_CONNEXION_TITRE & ":J" & LIGNE_CONNEXION)

    ' --- Boutons (bande "barre d'outils") ---
    With ws.Range("A" & LIGNE_BOUTONS & ":N" & LIGNE_BOUTONS)
        .Interior.Color = RGB(245, 247, 250)
    End With
    ws.Rows(LIGNE_BOUTONS).RowHeight = 34
    AjouterBouton ws, "Connecter", "modDB.Connecter", ws.Range("A" & LIGNE_BOUTONS).Left + 6, 120, LIGNE_BOUTONS
    AjouterBouton ws, "Charger / Actualiser", "modDonnees.Charger_Table", ws.Range("A" & LIGNE_BOUTONS).Left + 132, 155, LIGNE_BOUTONS
    AjouterBouton ws, "Rechercher", "modDonnees.Rechercher", ws.Range("A" & LIGNE_BOUTONS).Left + 297, 105, LIGNE_BOUTONS
    AjouterBouton ws, "Effacer criteres", "modDonnees.Effacer_Criteres", ws.Range("A" & LIGNE_BOUTONS).Left + 412, 125, LIGNE_BOUTONS

    ' --- Panneau RECHERCHE AVANCEE ---
    TitreSection ws.Range("A" & LIGNE_RECH_TITRE & ":D" & LIGNE_RECH_TITRE), "RECHERCHE AVANCEE (interroge la base)"
    ws.Range("A" & LIGNE_RECH_HEADER & ":D" & LIGNE_RECH_HEADER).Value = Array("Lien", "Colonne", "Operateur", "Valeur")
    With ws.Range("A" & LIGNE_RECH_HEADER & ":D" & LIGNE_RECH_HEADER)
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With
    ws.Cells(LIGNE_RECH_CRIT1, 1).Value = "Ou"
    ws.Cells(LIGNE_RECH_CRIT1 + 1, 1).Value = "ET"
    ws.Cells(LIGNE_RECH_CRIT1 + 2, 1).Value = "ET"
    ws.Cells(LIGNE_RECH_PERIODE, 1).Value = "Date debut"
    ws.Cells(LIGNE_RECH_PERIODE, 3).Value = "Date fin"
    ws.Cells(LIGNE_RECH_PERIODE, 2).NumberFormat = "dd/mm/yyyy"
    ws.Cells(LIGNE_RECH_PERIODE, 4).NumberFormat = "dd/mm/yyyy"

    ws.Range("A" & LIGNE_RECH_CRIT1 & ":A" & LIGNE_RECH_PERIODE).Font.Bold = True
    ws.Cells(LIGNE_RECH_PERIODE, 3).Font.Bold = True
    ws.Range("B" & LIGNE_RECH_CRIT1 & ":D" & (LIGNE_RECH_CRIT1 + 2)).Interior.Color = RGB(255, 250, 220)
    ws.Range("A" & (LIGNE_RECH_CRIT1 + 1) & ":A" & (LIGNE_RECH_CRIT1 + 2)).Interior.Color = RGB(255, 250, 220)
    ws.Cells(LIGNE_RECH_PERIODE, 2).Interior.Color = RGB(255, 250, 220)
    ws.Cells(LIGNE_RECH_PERIODE, 4).Interior.Color = RGB(255, 250, 220)

    AjouterListe ws.Range("A" & (LIGNE_RECH_CRIT1 + 1) & ":A" & (LIGNE_RECH_CRIT1 + 2)), "ET|OU"
    AjouterListePlage ws.Range("B" & LIGNE_RECH_CRIT1 & ":B" & (LIGNE_RECH_CRIT1 + 2)), "=ListeColonnes"
    AjouterListe ws.Range("C" & LIGNE_RECH_CRIT1 & ":C" & (LIGNE_RECH_CRIT1 + 2)), _
        "contient|ne contient pas|commence par|se termine par|égal à|différent de|supérieur à|supérieur ou égal à|inférieur à|inférieur ou égal à|est vide|n'est pas vide"

    ws.Cells(LIGNE_RESULTAT, 1).Value = "Resultat :"
    ws.Cells(LIGNE_RESULTAT, 1).Font.Bold = True
    ws.Cells(LIGNE_RESULTAT, 2).Font.Color = RGB(31, 78, 121)
    ws.Cells(LIGNE_RESULTAT, 2).Value = "Cliquez sur 'Connecter' pour commencer."
    EncadrerPanneau ws.Range("A" & LIGNE_RECH_TITRE & ":D" & LIGNE_RESULTAT)

    ' --- Panneau RESUME (KPI) ---
    TitreSection ws.Range("A" & LIGNE_KPI_TITRE & ":H" & LIGNE_KPI_TITRE), "RESUME"
    EcrireCarteKPI ws, "A", "Lignes affichees", CELL_KPI_LIGNES
    EcrireCarteKPI ws, "C", "Somme", CELL_KPI_SOMME
    EcrireCarteKPI ws, "E", "Moyenne", CELL_KPI_MOYENNE
    EcrireCarteKPI ws, "G", "Categories uniques", CELL_KPI_CATEGORIES
    EncadrerPanneau ws.Range("A" & LIGNE_KPI_TITRE & ":H" & LIGNE_KPI)

    ' --- Panneau GRAPHIQUES : parametres detectes automatiquement, modifiables ---
    TitreSection ws.Range("A" & LIGNE_GRAPH_TITRE & ":P" & LIGNE_GRAPH_TITRE), "GRAPHIQUES"
    ws.Cells(LIGNE_GRAPH_PARAMS, 6).Value = "Categorie :"
    ws.Cells(LIGNE_GRAPH_PARAMS, 8).Value = "Valeur :"
    ws.Cells(LIGNE_GRAPH_PARAMS, 10).Value = "Date :"
    ws.Cells(LIGNE_GRAPH_PARAMS, 12).Value = "Agregation :"
    ws.Cells(LIGNE_GRAPH_PARAMS, 14).Value = "Doublons :"
    With ws.Range("F" & LIGNE_GRAPH_PARAMS & ":N" & LIGNE_GRAPH_PARAMS)
        .Font.Italic = True
        .Font.Size = 9
    End With

    Dim cellesParam As Range
    Set cellesParam = Union(ws.Range(CELL_CAT), ws.Range(CELL_VAL), ws.Range(CELL_DATECOL), _
                             ws.Range(CELL_AGG), ws.Range(CELL_DOUBLONS))
    cellesParam.Interior.Color = RGB(255, 250, 220)
    cellesParam.Borders.LineStyle = xlContinuous
    cellesParam.Borders.Color = RGB(200, 200, 200)

    AjouterListePlage ws.Range(CELL_CAT), "=ListeColonnes"
    AjouterListePlage ws.Range(CELL_VAL), "=ListeColonnes"
    AjouterListePlage ws.Range(CELL_DATECOL), "=ListeColonnes"
    AjouterListe ws.Range(CELL_AGG), "SOMME|MOYENNE|NOMBRE"
    AjouterListe ws.Range(CELL_DOUBLONS), "Surligner|Supprimer|Conserver"
    EncadrerPanneau ws.Range("A" & LIGNE_GRAPH_TITRE & ":P" & LIGNE_GRAPH_PARAMS)

    TitreSection ws.Range("A" & LIGNE_DONNEES_TITRE & ":P" & LIGNE_DONNEES_TITRE), "DONNEES"

    AjouterListePlage ws.Range(CELL_TABLE), "=ListeTables"
End Sub

' Encadre un bloc pour lui donner un effet "carte" (au lieu d'un simple empilement)
Private Sub EncadrerPanneau(ByVal rng As Range)
    rng.BorderAround LineStyle:=xlContinuous, Weight:=xlThin, Color:=RGB(210, 210, 210)
End Sub

' Petite "carte" KPI : libelle au-dessus, grande valeur en dessous, encadree
Private Sub EcrireCarteKPI(ByVal ws As Worksheet, ByVal colonne As String, ByVal libelle As String, ByVal celluleValeur As String)
    With ws.Range(colonne & LIGNE_KPI_LABELS)
        .Value = libelle
        .Font.Bold = True
        .Font.Size = 9
        .Font.Color = RGB(100, 100, 100)
    End With
    With ws.Range(celluleValeur)
        .Value = "-"
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = RGB(31, 78, 121)
        .Interior.Color = RGB(240, 245, 250)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
        .HorizontalAlignment = xlCenter
    End With
End Sub

Private Sub TitreSection(ByVal rng As Range, ByVal texte As String)
    rng.Cells(1, 1).Value = texte
    With rng
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
    End With
End Sub

Private Sub AjouterBouton(ByVal ws As Worksheet, ByVal texte As String, ByVal macro As String, _
                          ByVal gauche As Double, ByVal largeur As Double, ByVal ligne As Long)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, gauche, ws.Cells(ligne, 1).Top + 4, largeur, 26)
    shp.Name = "btn_" & Replace(macro, ".", "_")
    shp.Fill.ForeColor.RGB = RGB(31, 78, 121)
    shp.Line.Visible = msoFalse
    shp.Shadow.Visible = msoFalse
    With shp.TextFrame2.TextRange
        .Text = texte
        .Font.Size = 10
        .Font.Bold = msoTrue
        .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
    shp.OnAction = macro
End Sub

' Liste deroulante a partir d'elements separes par "|"
' (Excel attend le separateur de liste de la langue : ";" en francais)
Private Sub AjouterListe(ByVal rng As Range, ByVal elements As String)
    Dim sep As String
    sep = Application.International(xlListSeparator)
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=Replace(elements, "|", sep)
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

Private Sub AjouterListePlage(ByVal rng As Range, ByVal formule As String)
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=formule
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub
