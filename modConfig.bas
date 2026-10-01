Attribute VB_Name = "modConfig"
Option Explicit

' ============================================================
'  Constantes et utilitaires communs
'  L'utilisateur ne saisit que le lien de connexion et le mot de
'  passe. Tout le reste (type de SGBD, pilote ODBC, liste des
'  tables, colonnes des graphiques, mise en forme) est detecte
'  ou genere automatiquement.
' ============================================================

Public Const FEUILLE_CONFIG As String = "Config"
Public Const FEUILLE_DASH As String = "Dashboard"
Public Const FEUILLE_CALC As String = "_Calc"

Public Const NOM_TABLEAU As String = "tblDonnees"
Public Const COL_RECHERCHE_CACHEE As String = "_RechercheTexte"   ' colonne technique, masquee

' Constantes ADO (liaison tardive : aucune reference a ajouter)
Public Const ADO_STATIC As Long = 3
Public Const ADO_LECTURE_SEULE As Long = 1
Public Const ADO_CLIENT As Long = 3

' Limite de lignes lues, fixee dans le code (plus simple pour l'utilisateur)
Public Const LIMITE_LIGNES_DEFAUT As Long = 50000

' Au-dela de ce nombre de categories distinctes, on n'affiche pas de segment
' (boutons de filtre) pour la colonne categorie : ce serait illisible.
Public Const MAX_CATEGORIES_SEGMENT As Long = 15

' ---------- Feuille Config ----------
Public Const CFG_LIEN As Long = 4
Public Const CFG_MDP As Long = 5
Public Const CFG_PILOTE As Long = 7   ' optionnel : nom exact du pilote, a ne remplir qu'en cas d'echec de connexion

' ---------- Mise en page de la feuille Dashboard ----------
' Bande CONNEXION (a gauche, colonnes A-D) et RECHERCHE RAPIDE (a droite,
' colonnes F-J) partagent les memes lignes pour eviter d'empiler les blocs.
Public Const LIGNE_CONNEXION_TITRE As Long = 3
Public Const LIGNE_CONNEXION As Long = 4
Public Const LIGNE_BOUTONS As Long = 6
Public Const LIGNE_RECH_TITRE As Long = 8
Public Const LIGNE_RECH_HEADER As Long = 9
Public Const LIGNE_RECH_CRIT1 As Long = 10        ' 3 lignes de criteres : 10, 11, 12
Public Const LIGNE_RECH_PERIODE As Long = 13
Public Const LIGNE_RESULTAT As Long = 14
Public Const LIGNE_KPI_TITRE As Long = 16
Public Const LIGNE_KPI_LABELS As Long = 17
Public Const LIGNE_KPI As Long = 18
Public Const LIGNE_GRAPH_TITRE As Long = 20
Public Const LIGNE_GRAPH_PARAMS As Long = 21
Public Const LIGNE_GRAPHIQUES As Long = 22        ' ligne de placement des graphiques
Public Const LIGNE_DONNEES_TITRE As Long = 44
Public Const LIGNE_TABLEAU As Long = 46           ' ligne d'en-tete du tableau de donnees

' Cellules de parametres sur le Dashboard (remplies automatiquement, modifiables)
Public Const CELL_TABLE As String = "B4"
Public Const CELL_BASE_INFO As String = "D4"
Public Const CELL_RECHERCHE_RAPIDE As String = "F4"
Public Const CELL_CAT As String = "G21"
Public Const CELL_VAL As String = "I21"
Public Const CELL_DATECOL As String = "K21"
Public Const CELL_AGG As String = "M21"
Public Const CELL_DOUBLONS As String = "O21"

' Cellules du resume chiffre (KPI), mises a jour a chaque chargement/recherche
Public Const CELL_KPI_LIGNES As String = "B18"
Public Const CELL_KPI_SOMME As String = "D18"
Public Const CELL_KPI_MOYENNE As String = "F18"
Public Const CELL_KPI_CATEGORIES As String = "H18"

' Lit un parametre de la feuille Config (Lien, Mot de passe, Pilote)
Public Function Cfg(ByVal ligne As Long) As String
    Cfg = Trim$(CStr(ThisWorkbook.Worksheets(FEUILLE_CONFIG).Cells(ligne, 2).Value))
End Function

' Table actuellement choisie sur le Dashboard
Public Function TableSelectionnee() As String
    TableSelectionnee = Trim$(CStr(ThisWorkbook.Worksheets(FEUILLE_DASH).Range(CELL_TABLE).Value))
End Function

' Mode de gestion des doublons choisi sur le Dashboard (defaut : SURLIGNER)
Public Function ModeDoublons() As String
    Dim v As String
    v = UCase$(Trim$(CStr(ThisWorkbook.Worksheets(FEUILLE_DASH).Range(CELL_DOUBLONS).Value)))
    If v = "" Then v = "SURLIGNER"
    ModeDoublons = v
End Function

' Texte actuellement tape dans la case de recherche rapide (filtre instantane)
Public Function TexteRechercheRapide() As String
    TexteRechercheRapide = Trim$(CStr(ThisWorkbook.Worksheets(FEUILLE_DASH).Range(CELL_RECHERCHE_RAPIDE).Value))
End Function

' Convertit une valeur quelconque en texte (Null, Empty et tableaux -> "")
Public Function Txt(ByVal v As Variant) As String
    If IsArray(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    Txt = CStr(v)
End Function

Public Function ExisteFeuille(ByVal nom As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nom)
    On Error GoTo 0
    ExisteFeuille = Not ws Is Nothing
End Function

Public Function ApplicationPrete() As Boolean
    ApplicationPrete = ExisteFeuille(FEUILLE_DASH) And ExisteFeuille(FEUILLE_CONFIG) And ExisteFeuille(FEUILLE_CALC)
    If Not ApplicationPrete Then
        MsgBox "Executez d'abord la macro InitialiserApplication (module modSetup).", vbExclamation
    End If
End Function
