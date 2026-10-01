Attribute VB_Name = "modDB"
Option Explicit

' ============================================================
'  Connexion a la base de donnees (ADODB)
'  L'utilisateur fournit uniquement :
'    - un lien   : mysql://utilisateur@serveur:port/base
'                  postgresql://utilisateur@serveur:port/base
'                  sqlserver://utilisateur@serveur:port/base
'                  ou un chemin de fichier Access (.accdb / .mdb)
'    - un mot de passe (facultatif pour Access)
'  Le type de SGBD, le port par defaut et le pilote a utiliser
'  sont detectes automatiquement : le code interroge le
'  gestionnaire de pilotes ODBC de Windows (API SQLGetInstalledDrivers)
'  pour connaitre les pilotes reellement installes sur le poste et
'  choisit celui qui correspond, avec une liste de secours si la
'  detection ne renvoie rien.
' ============================================================

#If VBA7 Then
    Private Declare PtrSafe Function SQLGetInstalledDrivers Lib "odbccp32.dll" _
        (ByVal lpszBuf As String, ByVal cbBufMax As Integer, ByRef pcbBufOut As Integer) As Long
#Else
    Private Declare Function SQLGetInstalledDrivers Lib "odbccp32.dll" _
        (ByVal lpszBuf As String, ByVal cbBufMax As Integer, ByRef pcbBufOut As Integer) As Long
#End If

' Decompose un lien du type  sgbd://utilisateur@serveur:port/base
' Retourne False si le format n'est pas reconnu (lien avance ou fichier Access).
' Accepte aussi bien mysql://utilisateur@serveur:port/base que
' mysql://utilisateur:motdepasse@serveur:port/base (le mot de passe inclus
' dans le lien n'est utilise que si la case "Mot de passe" de Config est vide).
Public Function AnalyserLien(ByVal lien As String, ByRef typ As String, ByRef utilisateur As String, _
                              ByRef motDePasseLien As String, ByRef serveur As String, ByRef port As String, _
                              ByRef base As String) As Boolean
    Dim reste As String, arobase As Long, deuxPoints As Long, barre As Long, segUtilisateur As String, dpU As Long

    typ = "" : utilisateur = "" : motDePasseLien = "" : serveur = "" : port = "" : base = ""

    If LCase$(Left$(lien, 8)) = "mysql://" Then
        typ = "MYSQL" : port = "3306" : reste = Mid$(lien, 9)
    ElseIf LCase$(Left$(lien, 13)) = "postgresql://" Then
        typ = "POSTGRESQL" : port = "5432" : reste = Mid$(lien, 14)
    ElseIf LCase$(Left$(lien, 11)) = "postgres://" Then
        typ = "POSTGRESQL" : port = "5432" : reste = Mid$(lien, 12)
    ElseIf LCase$(Left$(lien, 12)) = "sqlserver://" Then
        typ = "SQLSERVER" : port = "1433" : reste = Mid$(lien, 13)
    ElseIf LCase$(Left$(lien, 8)) = "mssql://" Then
        typ = "SQLSERVER" : port = "1433" : reste = Mid$(lien, 9)
    Else
        Exit Function
    End If

    arobase = InStr(reste, "@")
    If arobase = 0 Then Exit Function
    segUtilisateur = Left$(reste, arobase - 1)
    dpU = InStr(segUtilisateur, ":")
    If dpU > 0 Then
        utilisateur = Left$(segUtilisateur, dpU - 1)
        motDePasseLien = Mid$(segUtilisateur, dpU + 1)
    Else
        utilisateur = segUtilisateur
    End If
    reste = Mid$(reste, arobase + 1)

    barre = InStr(reste, "/")
    If barre = 0 Then Exit Function
    base = Mid$(reste, barre + 1)
    reste = Left$(reste, barre - 1)

    deuxPoints = InStr(reste, ":")
    If deuxPoints > 0 Then
        serveur = Left$(reste, deuxPoints - 1)
        port = Mid$(reste, deuxPoints + 1)
    Else
        serveur = reste
    End If

    AnalyserLien = (typ <> "" And utilisateur <> "" And serveur <> "" And base <> "")
End Function

' Un lien qui pointe vers un fichier Access (local ou reseau), avec ou sans prefixe access://
Private Function EstFichierAccess(ByVal lien As String) As Boolean
    Dim l As String
    l = LCase$(Trim$(lien))
    EstFichierAccess = (Right$(l, 6) = ".accdb") Or (Right$(l, 4) = ".mdb")
End Function

Private Function CheminAccess(ByVal lien As String) As String
    Dim c As String
    c = Trim$(lien)
    If LCase$(Left$(c, 9)) = "access://" Then c = Mid$(c, 10)
    If LCase$(Left$(c, 7)) = "file://" Then c = Mid$(c, 8)
    CheminAccess = c
End Function

' Type de SGBD detecte a partir du lien : "MYSQL", "SQLSERVER", "POSTGRESQL", "ACCESS" ou "PERSONNALISE"
Public Function TypeSGBD() As String
    Dim typ As String, u As String, mp As String, s As String, p As String, b As String, lien As String
    lien = Cfg(CFG_LIEN)
    If EstFichierAccess(lien) Then
        TypeSGBD = "ACCESS"
    ElseIf AnalyserLien(lien, typ, u, mp, s, p, b) Then
        TypeSGBD = typ
    Else
        TypeSGBD = "PERSONNALISE"
    End If
End Function

' Chaine de connexion (sans le DRIVER=, ajoute ensuite selon le pilote qui fonctionne)
' Ne s'utilise que pour MYSQL / SQLSERVER / POSTGRESQL / PERSONNALISE (pas ACCESS).
Private Function ChaineBase() As String
    Dim lien As String, mdp As String, typ As String, utilisateur As String, motDePasseLien As String
    Dim serveur As String, port As String, base As String
    Dim c As String

    lien = Cfg(CFG_LIEN)
    mdp = Cfg(CFG_MDP)
    If lien = "" Then Err.Raise vbObjectError + 90, , "Renseignez le lien de connexion dans la feuille Config."

    If AnalyserLien(lien, typ, utilisateur, motDePasseLien, serveur, port, base) Then
        If mdp = "" Then mdp = motDePasseLien   ' mot de passe fourni directement dans le lien (utilisateur:motdepasse@...)
        If mdp = "" Then mdp = InputBox("Mot de passe de la base de donnees :", "Connexion")
        ChaineBase = "SERVER=" & serveur & ";PORT=" & port & ";DATABASE=" & base & _
            ";UID=" & utilisateur & ";PWD=" & mdp & ";OPTION=3;"
    ElseIf InStr(lien, "=") > 0 Then
        ' Mode avance : l'utilisateur a colle une chaine de connexion complete
        If mdp = "" Then mdp = InputBox("Mot de passe de la base de donnees :", "Connexion")
        c = lien
        If Right$(Trim$(c), 1) <> ";" Then c = c & ";"
        If InStr(1, c, "PWD=", vbTextCompare) = 0 And InStr(1, c, "Password=", vbTextCompare) = 0 Then
            c = c & "PWD=" & mdp & ";"
        End If
        ChaineBase = c
    Else
        Err.Raise vbObjectError + 92, , "Lien non reconnu." & vbCrLf & _
            "Format attendu : mysql://utilisateur@serveur:port/nombase" & vbCrLf & _
            "(remplacez mysql par postgresql ou sqlserver selon votre base, ou indiquez un fichier .accdb/.mdb pour Access)."
    End If
End Function

' ---------- Detection automatique des pilotes ODBC installes ----------

' Interroge le gestionnaire de pilotes ODBC de Windows pour connaitre
' les pilotes reellement installes sur ce poste (fiable, sans dependre
' de WScript.Shell ni d'une politique de securite qui pourrait le bloquer).
Private Function PilotesInstalles() As Collection
    Dim col As New Collection
    Dim tampon As String, tailleTampon As Integer, tailleUtile As Integer, ok As Long
    Dim morceaux() As String, i As Long, nom As String

    On Error GoTo Fin
    tailleTampon = 8000
    tampon = String$(tailleTampon, vbNullChar)
    ok = SQLGetInstalledDrivers(tampon, tailleTampon, tailleUtile)
    On Error GoTo 0

    If ok <> 0 And tailleUtile > 0 Then
        tampon = Left$(tampon, tailleUtile)
        morceaux = Split(tampon, vbNullChar)
        For i = LBound(morceaux) To UBound(morceaux)
            nom = Trim$(morceaux(i))
            If nom <> "" Then
                On Error Resume Next
                col.Add nom, nom
                On Error GoTo 0
            End If
        Next i
    End If

    Set PilotesInstalles = col
    Exit Function

Fin:
    Set PilotesInstalles = col
End Function

' Pilotes a essayer pour un type de SGBD : d'abord ceux reellement installes
' (detectes via le registre), puis une liste de secours au cas ou la detection
' echouerait (regles de securite bloquant WScript.Shell, par exemple).
Private Function PilotesCandidats(ByVal typ As String) As Variant
    Dim piloteImpose As String
    piloteImpose = Trim$(Cfg(CFG_PILOTE))
    If piloteImpose <> "" Then
        ' L'utilisateur a indique lui-meme le pilote exact (feuille Config) : on l'utilise directement.
        PilotesCandidats = Array(piloteImpose)
        Exit Function
    End If

    Dim motCle As String, secours As Variant, installes As Collection, resultat As New Collection
    Dim i As Long, nom As String

    Select Case typ
    Case "MYSQL"
        motCle = "mysql"
        secours = Array("MySQL ODBC 9.0 Unicode Driver", "MySQL ODBC 8.0 Unicode Driver", _
            "MySQL ODBC 8.0 ANSI Driver", "MySQL ODBC 5.3 Unicode Driver", "MySQL ODBC 5.3 ANSI Driver", _
            "MySQL ODBC 5.1 Driver")
    Case "SQLSERVER"
        motCle = "sql server"
        secours = Array("ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server", _
            "ODBC Driver 13 for SQL Server", "ODBC Driver 11 for SQL Server", _
            "SQL Server Native Client 11.0", "SQL Server")
    Case "POSTGRESQL"
        motCle = "postgres"
        secours = Array("PostgreSQL Unicode(x64)", "PostgreSQL Unicode", _
            "PostgreSQL ANSI(x64)", "PostgreSQL ANSI")
    Case Else
        PilotesCandidats = Array()
        Exit Function
    End Select

    Set installes = PilotesInstalles()
    For i = 1 To installes.Count
        nom = installes(i)
        If InStr(1, nom, motCle, vbTextCompare) > 0 Then
            On Error Resume Next
            resultat.Add nom, nom
            On Error GoTo 0
        End If
    Next i
    For i = LBound(secours) To UBound(secours)
        On Error Resume Next
        resultat.Add secours(i), secours(i)
        On Error GoTo 0
    Next i

    Dim tableau() As String
    ReDim tableau(0 To resultat.Count - 1)
    For i = 1 To resultat.Count
        tableau(i - 1) = resultat(i)
    Next i
    PilotesCandidats = tableau
End Function

' Message d'aide affiche quand aucun pilote n'a fonctionne
Private Function AideInstallationPilote(ByVal typ As String) As String
    Dim msg As String
    Select Case typ
    Case "MYSQL"
        msg = "Aucun pilote ODBC MySQL utilisable n'a ete trouve sur ce PC." & vbCrLf & _
              "Telechargez et installez 'MySQL Connector/ODBC' (gratuit) :" & vbCrLf & _
              "https://dev.mysql.com/downloads/connector/odbc/"
    Case "SQLSERVER"
        msg = "Aucun pilote ODBC SQL Server utilisable n'a ete trouve sur ce PC." & vbCrLf & _
              "Telechargez 'Microsoft ODBC Driver for SQL Server' :" & vbCrLf & _
              "https://learn.microsoft.com/sql/connect/odbc/download-odbc-driver-for-sql-server"
    Case "POSTGRESQL"
        msg = "Aucun pilote ODBC PostgreSQL utilisable n'a ete trouve sur ce PC." & vbCrLf & _
              "Telechargez 'psqlODBC' :" & vbCrLf & _
              "https://www.postgresql.org/ftp/odbc/versions/"
    Case Else
        msg = "Installez le pilote correspondant a votre base de donnees."
    End Select
    AideInstallationPilote = msg & vbCrLf & vbCrLf & _
        "Astuce : la version du pilote (32 ou 64 bits) doit correspondre a celle d'Office" & vbCrLf & _
        "(dans Excel : Fichier > Compte > A propos d'Excel, en haut de la fenetre)."
End Function

' ---------- Guillemets d'identifiants selon le SGBD ----------

Public Function Q(ByVal ident As String) As String
    Dim parts() As String, i As Long, typ As String
    typ = TypeSGBD()
    parts = Split(ident, ".")
    For i = LBound(parts) To UBound(parts)
        Select Case typ
        Case "MYSQL"
            parts(i) = "`" & Replace(parts(i), "`", "``") & "`"
        Case "SQLSERVER", "ACCESS"
            parts(i) = "[" & Replace(parts(i), "]", "]]") & "]"
        Case "POSTGRESQL"
            parts(i) = """" & Replace(parts(i), """", """""") & """"
        End Select
    Next i
    Q = Join(parts, ".")
End Function

' Litteral de date valable pour le SGBD detecte (utilise pour les filtres de date)
Public Function LitteralDate(ByVal d As Date) As String
    If TypeSGBD() = "ACCESS" Then
        LitteralDate = "#" & Format(d, "yyyy-mm-dd") & "#"
    Else
        LitteralDate = "'" & Format(d, "yyyymmdd") & "'"
    End If
End Function

' Caractere joker a utiliser avec LIKE selon le SGBD (Access/Jet utilise * et ? en mode natif)
Public Function JokerMulti() As String
    If TypeSGBD() = "ACCESS" Then JokerMulti = "*" Else JokerMulti = "%"
End Function

' ---------- Ouverture de connexion ----------

' Ouvre et retourne une connexion ADODB (l'appelant doit la fermer).
Public Function OuvrirConnexion() As Object
    Dim cn As Object, typ As String, base As String, pilotes As Variant, i As Long
    Dim dernierMsg As String, essais As String

    typ = TypeSGBD()

    If typ = "ACCESS" Then
        Set OuvrirConnexion = OuvrirConnexionAccess()
        Exit Function
    End If

    base = ChaineBase()

    If typ = "PERSONNALISE" Then
        Set cn = CreateObject("ADODB.Connection")
        cn.ConnectionTimeout = 15
        cn.CommandTimeout = 60
        cn.CursorLocation = ADO_CLIENT
        cn.Open base
        Set OuvrirConnexion = cn
        Exit Function
    End If

    pilotes = PilotesCandidats(typ)
    For i = LBound(pilotes) To UBound(pilotes)
        Set cn = CreateObject("ADODB.Connection")
        cn.ConnectionTimeout = 15
        cn.CommandTimeout = 60
        cn.CursorLocation = ADO_CLIENT
        On Error Resume Next
        Err.Clear
        cn.Open "DRIVER={" & pilotes(i) & "};" & base
        If Err.Number = 0 Then
            On Error GoTo 0
            Set OuvrirConnexion = cn
            Exit Function
        End If
        dernierMsg = Err.Description
        essais = essais & "  - " & pilotes(i) & vbCrLf
        On Error GoTo 0
    Next i

    Err.Raise vbObjectError + 91, , "Impossible de se connecter." & vbCrLf & _
        "Pilotes ODBC essayes :" & vbCrLf & essais & vbCrLf & _
        "Derniere erreur : " & dernierMsg & vbCrLf & vbCrLf & _
        AideInstallationPilote(typ)
End Function

' Connexion a un fichier Access (.accdb / .mdb) via le moteur ACE ou, a defaut, Jet
Private Function OuvrirConnexionAccess() As Object
    Dim chemin As String, mdp As String, cn As Object
    Dim fournisseurs As Variant, i As Long, dernierMsg As String, essais As String, cs As String

    chemin = CheminAccess(Cfg(CFG_LIEN))
    mdp = Cfg(CFG_MDP)

    fournisseurs = Array("Microsoft.ACE.OLEDB.16.0", "Microsoft.ACE.OLEDB.12.0", "Microsoft.Jet.OLEDB.4.0")
    For i = LBound(fournisseurs) To UBound(fournisseurs)
        Set cn = CreateObject("ADODB.Connection")
        cn.ConnectionTimeout = 15
        cn.CommandTimeout = 60
        cn.CursorLocation = ADO_CLIENT
        cs = "Provider=" & fournisseurs(i) & ";Data Source=" & chemin & ";Persist Security Info=False;"
        If mdp <> "" Then cs = cs & "Jet OLEDB:Database Password=" & mdp & ";"
        On Error Resume Next
        Err.Clear
        cn.Open cs
        If Err.Number = 0 Then
            On Error GoTo 0
            Set OuvrirConnexionAccess = cn
            Exit Function
        End If
        dernierMsg = Err.Description
        essais = essais & "  - " & fournisseurs(i) & vbCrLf
        On Error GoTo 0
    Next i

    Err.Raise vbObjectError + 93, , "Impossible d'ouvrir le fichier Access :" & vbCrLf & chemin & vbCrLf & vbCrLf & _
        "Fournisseurs essayes :" & vbCrLf & essais & vbCrLf & "Derniere erreur : " & dernierMsg & vbCrLf & vbCrLf & _
        "Verifiez que le chemin est accessible depuis ce PC et qu'Access (ou le" & vbCrLf & _
        "'Microsoft Access Database Engine', telechargeable gratuitement chez Microsoft) est installe."
End Function

' ---------- Structure de la table et liste des tables ----------

' Lit la structure de la table selectionnee (noms + classes de types) dans la feuille _Calc
Public Sub ChargerSchema(ByVal cn As Object)
    Dim rs As Object, ws As Worksheet, i As Long, n As Long

    Set ws = ThisWorkbook.Worksheets(FEUILLE_CALC)
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM " & Q(TableSelectionnee()) & " WHERE 1=0", cn, ADO_STATIC, ADO_LECTURE_SEULE

    n = rs.Fields.Count
    ws.Range("A1:B5000").ClearContents
    ws.Range("A1:B5000").NumberFormat = "@"
    ws.Range("A1").Value = "Colonne"
    ws.Range("B1").Value = "Classe (N=numerique, D=date, T=texte)"
    For i = 0 To n - 1
        ws.Cells(i + 2, 1).Value = rs.Fields(i).Name
        ws.Cells(i + 2, 2).Value = ClasseType(CLng(rs.Fields(i).Type))
    Next i
    rs.Close

    ThisWorkbook.Names("ListeColonnes").RefersTo = "='" & FEUILLE_CALC & "'!$A$2:$A$" & (n + 1)
End Sub

Private Function ClasseType(ByVal t As Long) As String
    Select Case t
    Case 2, 3, 4, 5, 6, 14, 16, 17, 18, 19, 20, 21, 131
        ClasseType = "N"
    Case 7, 64, 133, 135
        ClasseType = "D"
    Case Else
        ClasseType = "T"
    End Select
End Function

' Liste les tables de la base connectee (essaie OpenSchema, sinon requete de secours)
Public Function ListerTables(ByVal cn As Object) As Collection
    Dim rs As Object, col As New Collection, nom As String

    On Error Resume Next
    Set rs = cn.OpenSchema(20) ' adSchemaTables
    On Error GoTo 0

    If rs Is Nothing Then Set rs = RequeteListeTablesSecours(cn)
    If rs Is Nothing Then
        Set ListerTables = col
        Exit Function
    End If

    Do While Not rs.EOF
        On Error Resume Next
        If UCase$(Txt(rs.Fields("TABLE_TYPE").Value)) = "TABLE" Then
            nom = Txt(rs.Fields("TABLE_NAME").Value)
            If nom <> "" Then col.Add nom, nom
        End If
        On Error GoTo 0
        rs.MoveNext
    Loop
    rs.Close
    Set ListerTables = col
End Function

Private Function RequeteListeTablesSecours(ByVal cn As Object) As Object
    Dim rs As Object, sql As String
    Select Case TypeSGBD()
    Case "MYSQL"
        sql = "SELECT TABLE_NAME, 'TABLE' AS TABLE_TYPE FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = DATABASE()"
    Case "POSTGRESQL"
        sql = "SELECT TABLE_NAME, 'TABLE' AS TABLE_TYPE FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = 'public'"
    Case "SQLSERVER"
        sql = "SELECT TABLE_NAME, 'TABLE' AS TABLE_TYPE FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'"
    Case Else
        Exit Function
    End Select

    On Error Resume Next
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, cn, ADO_STATIC, ADO_LECTURE_SEULE
    If Err.Number <> 0 Then Set rs = Nothing
    On Error GoTo 0
    Set RequeteListeTablesSecours = rs
End Function

' Bouton "Connecter" : teste la connexion et remplit la liste des tables a choisir
Public Sub Connecter()
    Dim cn As Object, tables As Collection, ws As Worksheet, wc As Worksheet
    Dim i As Long, typ As String, u As String, mp As String, s As String, p As String, b As String

    If Not ApplicationPrete() Then Exit Sub
    On Error GoTo Erreur

    Application.Cursor = xlWait
    Application.StatusBar = "Connexion..."
    Set cn = OuvrirConnexion()
    Set tables = ListerTables(cn)
    cn.Close
    Application.Cursor = xlDefault
    Application.StatusBar = False

    If tables.Count = 0 Then
        MsgBox "Connexion reussie, mais aucune table trouvee." & vbCrLf & _
               "Verifiez les droits de l'utilisateur sur cette base.", vbExclamation, "Connexion"
        Exit Sub
    End If

    Set wc = ThisWorkbook.Worksheets(FEUILLE_CALC)
    wc.Range("M1:M2000").ClearContents
    wc.Range("M1").Value = "Table"
    For i = 1 To tables.Count
        wc.Cells(i + 1, 13).Value = tables(i)
    Next i
    ThisWorkbook.Names("ListeTables").RefersTo = "='" & FEUILLE_CALC & "'!$M$2:$M$" & (tables.Count + 1)

    Set ws = ThisWorkbook.Worksheets(FEUILLE_DASH)
    typ = TypeSGBD()
    If typ = "ACCESS" Then
        ws.Range(CELL_BASE_INFO).Value = CheminAccess(Cfg(CFG_LIEN))
    Else
        AnalyserLien Cfg(CFG_LIEN), typ, u, mp, s, p, b
        ws.Range(CELL_BASE_INFO).Value = IIf(b <> "", b, "(connexion personnalisee)")
    End If
    If ws.Range(CELL_TABLE).Value = "" Then ws.Range(CELL_TABLE).Value = tables(1)

    MsgBox tables.Count & " table(s) trouvee(s)." & vbCrLf & _
           "Choisissez une table dans la liste, puis cliquez sur 'Charger la table'.", vbInformation, "Connexion"
    Exit Sub

Erreur:
    Application.Cursor = xlDefault
    Application.StatusBar = False
    MsgBox "Echec de la connexion :" & vbCrLf & Err.Description, vbCritical, "Connexion"
End Sub
