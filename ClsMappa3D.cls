VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMappa3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'----------------------------------------------------------------------------------
'Dichiaro tutti gli oggetti e le variabili private della classe,ovvero tutti quei
'componenti che sono accedibili solamente dall'interno della classe stessa
'----------------------------------------------------------------------------------
'Dichiaro l'oggetto principale di questa classe, infatti i muri della mappa
'3D verranno costruiti proprio all'interno dello stesso.
Private Struttura As New Mesh8
'Dichiro un'altro oggetto in cui crere tutti qui muri pavimenti e soffitti con
'propriet� SphereMapping.
'Lo SphereMapping � una particolare funzione che permett� di simulare una superficie lucida
'con un falso effetto di riflessione
Private StrutturaMap As New Mesh8
'Quest'altro oggetto invece mi servir� per contenere tutti qui muri,pavimenti e soffitti
'con propriet� Trasparenza.
Private Trasparenza As New Mesh8
'L'oggeto FabbricaMateriali, mi servir� per creare diversi tipi di materiali da impostare ai muri pavimenti e
'soffitti, impostando alcuni parametri come la luce ambientale, la luce diffuse, ecc.
Private FabbricaMateriali As New MaterialFactory8

'----------------------------------------------------------------------------------
'Ora dichiaro i metodi pubblici necessari alla creazione della Mappa 3D da
'visualizzare in modalita Anteprima 3D
'----------------------------------------------------------------------------------
Public Sub Crea_Mappa_3D()
    Dim ColoreAppoggio As TColor
    'Dichiaro la variabile Vertice, la quale terr� conto del numero di vertici presenti all'interno dell'oggetto
    'Trasparenza.
    'Questa variabile � MOLTO importante, in quanto mi torner� utile nell'assegnazione dei colori di trasparenza
    'a tutti qui muri,pavimenti o soffitti che presentano il valore di propriet� Trasparenza.
    'In poche parole, questa variabile mi servir� per identificare singolarmente ogni vertice presente
    'all'interno dell'ogetto Trasparenza al fine di assegnare ad ognuno di esso il rispettivo colore
    Dim Vertice As Long
    'Dichiaro un indice che mi servir� nel caricamento all'interno del programma e l'assegnazione delle Texture e dei materiali dei muri,
    'pavimenti e soffitti
    Dim S As Single
    'Dichiaro la variabile di appoggio X1 che mi servir� per contenere il valore di Riga(I).X1 o SoP(J).CR.X1
    Dim X1 As Single
    'Dichiaro la variabile di appoggio X2 che mi servir� per contenere il valore di Riga(I).X2 o SoP(J).CR.X2
    Dim X2 As Single
    'Dichiaro la variabile di appoggio X3 che mi servir� per contenere solo il valore di SoP(J).X3
    Dim X3 As Single
    'Dichiaro la variabile di appoggio X4 che mi servir� per contenere solo il valore di SoP(J).X4
    Dim X4 As Single
    'Dichiaro la variabile di appoggio Z1 che mi servir� per contenere il valore di Riga(I).Y1 o SoP(J).CR.Y1
    Dim Z1 As Single
    'Dichiaro la variabile di appoggio Z2 che mi servir� per contenere il valore di Riga(I).Y2 o SoP(J).CR.Y2
    Dim Z2 As Single
    'Dichiaro la variabile di appoggio Z3 che mi servir� per contenere solo il valore di SoP(J).Y3
    Dim Z3 As Single
    'Dichiaro la variabile di appoggio Z4 che mi servir� per contenere solo il valore di SoP(J).Y4
    Dim Z4 As Single
    'Dichiaro la variabile di appoggio Altezza che mi servir� per contenere il valore di Riga(I).Altezza o SoP(J).CR.Altezza
    Dim Altezza As Single
    'Dichiaro la variabile di appoggio Altitudine che mi servir� per contenere il valore di Riga(I).Altitudine o SoP(J).CR.Altitudine
    Dim Altitudine As Single
    'Dichiaro la variabile di appoggio Mattonelle_in_Larghezza che mi servir� per contenere il valore di Riga(I).NMattonelleAltezza o SoP(J).CR.NMattonelleLarghezza
    Dim Mattonelle_in_Larghezza As Single
    'Dichiaro la variabile di appoggio Mattonelle_in_Altezza che mi servir� per contenere il valore di Riga(I).NMattonelleAltezza o SoP(J).CR.NMattonelleAltezza
    Dim Mattonelle_in_Altezza As Single
    'Dichiaro la variabile di appoggio Propriet� che mi servir� per contenere il valore di Riga(I).Propriet� o SoP(J).CR.Propriet�
    Dim Propriet� As String
    'Dichiaro la variabile di appoggio Texture che mi servir� per contenere il valore di Riga(I).Texture o SoP(J).CR.Texture
    Dim Texture As String
    'Inizializzo la variabile Vertice con un valore iniziale uguale a 0.
    'Questo sar� il primo vertice dell'oggetto Trasparenza a cui assegnare il rispettivo colore selezionato e il grado di solidit�
    Vertice = 0
    'Creo all'interno della scena un gruppo di oggetti in cui verranno costruiti tutti i muri,pavimenti o soffitti che avranno un
    'valore di propriet� Normale
    Set Struttura = Scena.CreateMeshBuilder("Struttura")
    'Creo all'interno della scena un gruppo di oggetti in cui verranno costruiti tutti i muri,pavimenti o soffitti che avranno un
    'valore di propriet� SphereMapping (false riflessione)
    Set StrutturaMap = Scena.CreateMeshBuilder("StruttureMap")
    'Creo all'interno della scena un gruppo di oggetti in cui verranno costruiti tutti i muri,pavimenti o soffitti che avranno un
    'valore di propriet� Trasparenza
    Set Trasparenza = Scena.CreateMeshBuilder("Trasparenza")
    'Dichiaro tutti i vertici dell'oggetto Trasparenza di tipo colorati.
    'Questo � NECESSARIO per impostare ad ognuno di questi colori differenti
    Trasparenza.SetVertexType TV_COLOR_VERTEX
    'Assegno all'oggetto Trasparenza la propriet� BlendingMode uguale a TV_BLEND_ALPHA.
    'Anche questa istruzione � NECESSARIA per far si che l'oggetto Trasparenza si appunto
    'realmente trasparente
    Trasparenza.SetBlendingMode TV_BLEND_ALPHA
    'Quest'altra istruzione invece, anch'essa NECESSARIA, fa s� che tutti i muri,pavimenti o soffitti creati
    'su quest'oggetto presentino un effetto di falsa riflessione
    StrutturaMap.EnableSphereMapping True
    'Avvio un ciclo che scandir� tutte le linee create all'interno della mappa attuale
    'ATTENZIONE: le variabili temporanee assumeranno il valore della rispettiva coordinata,per� apportando
    '            qualche modifica:
    '1) Tutte le coordinate verranno divise per la variabile Molt al fine di ripristinare la loro corretta
    '   dimensione,la quale era stata ventualmente modificata dalle operazioni di Zoom
    '2) Verranno tutte moltiplicate per il valore contenuto nella variabile VScale.
    '   Questa � una delle operazioni pi� significative,in quando mi permetter� di ingrandire la lunghezza dei muri
    '   tanto quanto � stato richiesto dall'utente mediante l'apposito controllo presente nel Form_Opzioni
    '3) Infine a questo risultato verr� completamente troncata la parte decimale,in modo da avere un valore
    '   completamente approssimato ad intero
    For I = 1 To Max
        'Assegno alla variabile temporanea X1 il rispettivo valore della coordinata X1 del muro analizzato
        X1 = Fix(Riga(I).X1 / Molt) * VScale
        'Assegno alla variabile temporanea X2 il rispettivo valore della coordinata X2 del muro analizzato
        X2 = Fix(Riga(I).X2 / Molt) * VScale
        'Assegno alla variabile temporanea Z1 il rispettivo valore della coordinata Y1 del muro analizzato
        Z1 = Fix(-Riga(I).Y1 / Molt) * VScale
        'Assegno alla variabile temporanea Z2 il rispettivo valore della coordinata Y2 del muro analizzato
        Z2 = Fix(-Riga(I).Y2 / Molt) * VScale
        'Assegno alla variabile temporanea Altezza,l'altezza del muro analizzato
        'Questo valore come tutti i successivi,verranno salvati all'interno delle rispettive variabili temporanee
        '"puri",ovvero senza apportarne alcuna modifica
        Altezza = Riga(I).Altezza
        'Assegno alla variabile temporanea Altitudine,la rispettiva altitudine dal terreno del muro analizzato
        Altitudine = Riga(I).Altitudine
        'Assegno alla variabile temporanea Mattonelle_in_Larghezza,il numero di mattonelle disposte in larghezza
        'su tutta la superficie del muro analizzato
        Mattonelle_in_Larghezza = Riga(I).NMAttonelleLarghezza
        'Assegno alla variabile temporanea Mattonelle_in_Altezza,il numero di mattonelle disposte in altezza
        'su tutta la superficie del muro analizzato
        Mattonelle_in_Altezza = Riga(I).NMattonelleALtezza
        'Assegno alla variabile temporanea propriet�,appunto la propriet� che il muro analizzato adotter�
        'A questo proposito si ricorda che quelle disponibili sono tre e sono:
        '1) Normale: non credo ci sia bisogno di spiegazioni
        '2) Trasparenza: al muro con questa propriet� sar� possibile "vederci attraverso" appunto perch�
        '   a differenza degli altri due tipi,questo non sar� solido,ma trasparente.
        '3) Sphere Mapping: questa propriet� permette di simulare sul muro che ne � in possesso,un suggestivo effetto
        '   di falsa riflessione
        Propriet� = RTrim(Riga(I).Propriet�)
        'Assegno alla variabile temporanea Texture,la rispettiva Texture(Immagine applicata) del muro analizzato
        Texture = RTrim(Riga(I).Texture)
        'Finita l'assegnazione dei valori alle rispettive variabili temporanee, carico all'interno del programma
        'la Texture assegnata al muro corrente,in modo da poterla applicare come mattonella appunto al muro corrente.
        'Maggiore saranno il numero di mattonelle,maggiore sar� appunto il numero di volte che questa immagine comparir�
        'sulla superficie del muro
        Scena.LoadTexture Texture, , , "Texture" + Str(S)
        'Ora creo il materiale assegnato al muro corrente
        FabbricaMateriali.CreateMaterial "Materiale" + Str(S)
        '---------------------------------
        With ColoreAppoggio
            ColoreAppoggio = Trasferisci_Colore(Riga(I).Materiale.Ambiente)
            FabbricaMateriali.SetAmbient GetMat("Materiale" + Str(S)), .R, .G, .B, .A
            ColoreAppoggio = Trasferisci_Colore(Riga(I).Materiale.Diffusa)
            FabbricaMateriali.SetDiffuse GetMat("Materiale" + Str(S)), .R, .G, .B, .A
            ColoreAppoggio = Trasferisci_Colore(Riga(I).Materiale.Emissiva)
            FabbricaMateriali.SetEmissive GetMat("Materiale" + Str(S)), .R, .G, .B, .A
            ColoreAppoggio = Trasferisci_Colore(Riga(I).Materiale.Speculare)
            FabbricaMateriali.SetSpecular GetMat("Materiale" + Str(S)), .R, .G, .B, .A
        End With
        '---------------------------------
        FabbricaMateriali.SetPower GetMat("Materiale" + Str(S)), Riga(I).Materiale.Potenza
        'Se il muro analizzato possiede una propriet� di tipo Normale,allora...
        If Propriet� = "Normale" Then
            'lo aggiunger� all'oggetto addetto a ricreare in 3D i muri con questa propriet�
            Struttura.AddWall GetTex("Texture" + Str(S)), X1, Z1, X2, Z2, Altezza, Altitudine, Mattonelle_in_Larghezza, Mattonelle_in_Altezza
            'Assegno al muro analizzato il suo stesso materiale precedentemente impostato
            Struttura.SetMaterial GetMat("Materiale" + Str(S)), S + 2
        'Altrimenti se il muro analizzato possiede una propriet� di tipo SphereMapping,allora...
        ElseIf Propriet� = "SphereMapping" Then
            'Anche questa volta lo aggiunger� all'oggetto addetto a ricreare in 3D i muri con questa
            'propriet�
            StrutturaMap.AddWall GetTex("Texture" + Str(S)), X1, Z1, X2, Z2, Altezza, Altitudine, Mattonelle_in_Larghezza, Mattonelle_in_Altezza
            'Assegno anche in questo caso al muro analizzato,il suo corrispondente materiale
            'precedentemente settato
            StrutturaMap.SetMaterial GetMat("Materiale" + Str(S)), S + 2
        'Nel caso in cui invece il muro analizzato avesse una propriet� del tipo Trasparenza,allora...
        ElseIf Propriet� = "Trasparenza" Then
            'Lo aggiunger� all'oggetto addetto a contenere appunto i muri,i pavimenti e i soffitti di questo genere
            Trasparenza.AddWall GetTex("Texture" + Str(S)), X1, Z1, X2, Z2, Altezza, Altitudine, Mattonelle_in_Larghezza, Mattonelle_in_Altezza
            '--------------------------------------------------------------------------------------------------
            ' Da qui iniziano le operazioni pi� importanti al fine di ricreare un ottimo effetto di trasparenza
            '--------------------------------------------------------------------------------------------------
            With ColoreAppoggio
                'Inizio con il settare il colore di trasparenza del vertice del muro posto a Sud - Ovest
                ColoreAppoggio = Trasferisci_Colore(Riga(I).ColVertici(2))
                Trasparenza.SetColor RGBA(.R, .G, .B, .A), Vertice
                'Poi faccio lo stesso anche con il vertice posto a Nord - Ovest...
                ColoreAppoggio = Trasferisci_Colore(Riga(I).ColVertici(0))
                Trasparenza.SetColor RGBA(.R, .G, .B, .A), Vertice + 1
                '...quello a Sud - Est...
                ColoreAppoggio = Trasferisci_Colore(Riga(I).ColVertici(3))
                Trasparenza.SetColor RGBA(.R, .G, .B, .A), Vertice + 2
                '...e per finire setto il colore di trasparenza dell'ultimo vertice del muro analizzato,cio� quello posto
                'a Nord - Est
                ColoreAppoggio = Trasferisci_Colore(Riga(I).ColVertici(1))
                Trasparenza.SetColor RGBA(.R, .G, .B, .A), Vertice + 3
            End With
            '--------------------------------------------------------------------------------------------------
            ' Qui finisce l'impostazione dei colori di trasparenza di un eventuale muro avente questa propriet�
            '--------------------------------------------------------------------------------------------------
            'Assegno al muro analizzato il suo corrispondente materiale precedentemente impostato
            Trasparenza.SetMaterial GetMat("Materiale" + Str(S)), S + 2
            'Ora incremente la variabile Vertice di quattro.Vi spiego perch�:
            'Essendo il muro composto da quattro vertici differenti,ed avendogli appena assegnato i rispettivi
            'colori di trasparenza,bisogner� settare in seguito i colori di un'altro muro,pavimento o soffitto,
            'quindi bisogner� appunto saltare a quattro a quattro i vertici
            Vertice = Vertice + 4
        End If
        'Ora incremento la variabile S.
        'Questo mi permetter� di assegnare con assoluta precisione la giusta Texture e materiale al muro corrispondente
        S = S + 1
    'Passo ad analizzare il muro successivo
    Next
    'Finita la conversione in 3D di tutti i muri presenti all'interno della mappa attuale,avvio un ciclo al fine di
    'convertire questa volta,tutti i pavimenti e soffitti in 3D.
    'Anche questa volta tutti i valori verranno salvati all'interno delle rispettive variabili temporanee,con le stesse
    'modifiche apportate anche per i muri (VEDI SOPRA)
    For J = 1 To Max2
            'Assegno alla variabile temporanea X1 il rispettivo valore della coordinata X1 del pavimento / soffitto analizzato
            X1 = Fix(SoP(J).CR.X1 / Molt) * VScale
            'Assegno alla variabile temporanea X2 il rispettivo valore della coordinata X2 del pavimento / soffitto analizzato
            X2 = Fix(SoP(J).CR.X2 / Molt) * VScale
            'Assegno alla variabile temporanea X3 il rispettivo valore della coordinata X3 del pavimento / soffitto analizzato
            X3 = Fix(SoP(J).X3 / Molt) * VScale
            'Assegno alla variabile temporanea X4 il rispettivo valore della coordinata X4 del pavimento / soffitto analizzato
            X4 = Fix(SoP(J).X4 / Molt) * VScale
            'Assegno alla variabile temporanea Z1 il rispettivo valore della coordinata Y1 del pavimento / soffitto analizzato
            Z1 = Fix(-SoP(J).CR.Y1 / Molt) * VScale
            'Assegno alla variabile temporanea Z2 il rispettivo valore della coordinata Y2 del pavimento / soffitto analizzato
            Z2 = Fix(-SoP(J).CR.Y2 / Molt) * VScale
            'Assegno alla variabile temporanea Z3 il rispettivo valore della coordinata Y3 del pavimento / soffitto analizzato
            Z3 = Fix(-SoP(J).Y3 / Molt) * VScale
            'Assegno alla variabile temporanea Z4 il rispettivo valore della coordinata Y4 del pavimento / soffitto analizzato
            Z4 = Fix(-SoP(J).Y4 / Molt) * VScale
            'Assegno alla variabile temporanea Altitudine,la rispettiva altitudine da cui il pavimento o soffitto,dista da terra
            Altitudine = SoP(J).CR.Altitudine
            'Assegno alla variabile temporanea Mattonelle_in_Larghezza,il numero di mattonelle disposte in larghezza
            'su tutta la superficie del pavimento / soffitto analizzato
            Mattonelle_in_Larghezza = SoP(J).CR.NMAttonelleLarghezza
            'Assegno alla variabile temporanea Mattonelle_in_Altezza,il numero di mattonelle disposte in altezza
            'su tutta la superficie del pavimento / soffitto analizzato
            Mattonelle_in_Altezza = SoP(J).CR.NMattonelleALtezza
            'Assegno alla variabile temporanea propriet�,appunto la propriet� che il pavimento / soffitto analizzato adotter�
            'Le propriet� disponibili sono le stesse messe a disposizione per i muri (VEDI SOPRA)
            Propriet� = RTrim(SoP(J).CR.Propriet�)
            'Assegno alla variabile temporanea Texture,la rispettiva Texture(Immagine applicata) del pavimento / soffitto analizzato
            Texture = RTrim(SoP(J).CR.Texture)
            'Finita l'assegnazione dei valori alle rispettive variabili temporanee, carico all'interno del programma
            'la Texture assegnata al pavimento / soffitto corrente,in modo da poterla applicare come mattonella appunto al pavimento / soffitto corrente.
            Scena.LoadTexture Texture, , , "Texture" + Str(S)
            'Ora creo il materiale assegnato al pavimento / soffitto corrente
            FabbricaMateriali.CreateMaterial "Materiale" + Str(S)
            '---------------------------------
            With ColoreAppoggio
                ColoreAppoggio = Trasferisci_Colore(SoP(J).CR.Materiale.Ambiente)
                FabbricaMateriali.SetAmbient GetMat("Materiale" + Str(S)), .R, .G, .B, .A
                ColoreAppoggio = Trasferisci_Colore(SoP(J).CR.Materiale.Diffusa)
                FabbricaMateriali.SetDiffuse GetMat("Materiale" + Str(S)), .R, .G, .B, .A
                ColoreAppoggio = Trasferisci_Colore(SoP(J).CR.Materiale.Emissiva)
                FabbricaMateriali.SetEmissive GetMat("Materiale" + Str(S)), .R, .G, .B, .A
                ColoreAppoggio = Trasferisci_Colore(SoP(J).CR.Materiale.Speculare)
                FabbricaMateriali.SetSpecular GetMat("Materiale" + Str(S)), .R, .G, .B, .A
            End With
            '---------------------------------
            'Se il pavimento / soffitto possiede una propriet� di trasparenza,allora
            If Propriet� = "Trasparenza" Then
                'Lo aggiunger� all'oggetto addetto a contenere appunto i muri,i pavimenti e i soffitti di questo genere
                Trasparenza.AddFaceFromPoint GetTex("Texture" + Str(S)), X1, Altitudine, Z1, X2, Altitudine, Z2, X3, Altitudine, Z3, X4, Altitudine, Z4, Mattonelle_in_Larghezza, Mattonelle_in_Altezza
                '----------------------------------------------------------------------------------------------------
                ' Da qui come gi� fatto per i muri,cominciano le operazione di assegnazione dei colori di trasparenza
                ' del pavimento / soffitto analizzato
                '----------------------------------------------------------------------------------------------------
                With ColoreAppoggio
                    'Inizio con il settare il colore di trasparenza del vertice del muro posto a Sud - Ovest
                    ColoreAppoggio = Trasferisci_Colore(SoP(J).CR.ColVertici(2))
                    Trasparenza.SetColor RGBA(.R, .G, .B, .A), Vertice
                    'Poi faccio lo stesso anche con il vertice posto a Nord - Ovest...
                    ColoreAppoggio = Trasferisci_Colore(SoP(J).CR.ColVertici(0))
                    Trasparenza.SetColor RGBA(.R, .G, .B, .A), Vertice + 1
                    '...quello a Sud - Est...
                    ColoreAppoggio = Trasferisci_Colore(SoP(J).CR.ColVertici(3))
                    Trasparenza.SetColor RGBA(.R, .G, .B, .A), Vertice + 2
                    '...e per finire setto il colore di trasparenza dell'ultimo vertice del muro analizzato,cio� quello posto
                    'a Nord - Est
                    ColoreAppoggio = Trasferisci_Colore(SoP(J).CR.ColVertici(1))
                    Trasparenza.SetColor RGBA(.R, .G, .B, .A), Vertice + 3
                End With
                'Assegno al muro analizzato il suo corrispondente materiale precedentemente impostato
                Trasparenza.SetMaterial GetMat("Materiale" + Str(S)), S + 2
                'Anche qui incremento la variabile Vertice di 4 (VEDI SOPRA)
                Vertice = Vertice + 4
            'Se il pavimento / soffitto analizzato possiede una propriet� di tipo Normale,allora...
            ElseIf Propriet� = "Normale" Then
                'Lo aggiunger� all'oggetto addetto a ricreare i muri,pavimenti e soffitti aventi questo tipo di propriet�
                Struttura.AddFaceFromPoint GetTex("Texture" + Str(S)), X1, Altitudine, Z1, X2, Altitudine, Z2, X3, Altitudine, Z3, X4, Altitudine, Z4, Mattonelle_in_Larghezza, Mattonelle_in_Altezza
                'Assegno anche in questo caso al pavimento / soffitto analizzato,il suo corrispondente materiale
                'precedentemente settato
                Struttura.SetMaterial GetMat("Materiale" + Str(S)), S + 2
            'Se invece il pavimento / soffitto analizzato possiede una propriet� di tipo SphereMapping,allora...
            ElseIf Propriet� = "SphereMapping" Then
                'Lo aggiungo all'oggetto addetto a ricreare muri,pavimenti e soffitti aventi questa propriet�
                StrutturaMap.AddFaceFromPoint GetTex("Texture" + Str(S)), X1, Altitudine, Z1, X2, Altitudine, Z2, X3, Altitudine, Z3, X4, Altitudine, Z4, Mattonelle_in_Larghezza, Mattonelle_in_Altezza
                'Assegno anche in questo caso al pavimento / soffitto analizzato,il suo corrispondente materiale
                'precedentemente settato
                StrutturaMap.SetMaterial GetMat("Materiale" + Str(S)), S + 2
            End If
            'Incremento la variabile S (VEDI SOPRA)
            S = S + 1
    'Passo ad analizzare il pavimento / soffitto successivo
    Next
End Sub

Private Function Trasferisci_Colore(ColoreOriginale As TColor) As TColor
    'Questa funzione ha il compito di duplicare i valori delle rispettive quantit� di Rosso,Verde,Blu e Alpha
    'contenuti appunto nel colore passato alla funzione stessa,e mantenere il suo duplicato in se stessa,
    'in modo da poterlo passare alla funzione pubblica addetta alla trasformazione del formato RGB in RGBA.
    'Questo far� in modo che i dati non vengano modificati direttamente all'interno del record,ma verranno
    'salvati su variabili di appoggio
    With ColoreOriginale
        'Copio la quantit� di Rosso
        Trasferisci_Colore.R = .R
        'Copio la quantit� di Verde
        Trasferisci_Colore.G = .G
        'Copio la quantit� di Blu
        Trasferisci_Colore.B = .B
        'Copio il valore dell'Alpha
        Trasferisci_Colore.A = .A
    End With
    'Finito l'operazione di copiatura,sono pronto a trasformare RGB in RGBA tramite la chiamata
    'della funzione addetta
    RGB_To_RGBA Trasferisci_Colore
End Function

Public Sub Distruggi_Mappa_3D()
    '---------------------------------------------------------------------------------------------------------------
    ' Resetto il contenuto di tutti quegli oggetti che mi erano serviti per poter ricreare muri,pavimenti e soffitti
    ' in 3D
    '---------------------------------------------------------------------------------------------------------------
    'Resetto l'oggetto contenente tutti quie muri,pavimenti e sofffitti avente una propriet� di tipo Normale
    Struttura.ResetMesh
    'Resetto l'oggetto contenente tutti quie muri,pavimenti e sofffitti avente una propriet� di tipo SphereMapping
    StrutturaMap.ResetMesh
    'Resetto l'oggetto contenente tutti quie muri,pavimenti e sofffitti avente una propriet� di tipo Trasparente
    Trasparenza.ResetMesh
    'Distruggo l'oggetto contenente tutti i muri creati
    '-----------------------------------------------------------
    ' Infine proseguo con la distruzione definitiva degli stessi
    '-----------------------------------------------------------
    'Distruggo l'oggetto Struttura
    Set Struttura = Nothing
    'Distruggo l'oggetto StruttraMap
    Set StrutturaMap = Nothing
    'Distruggo l'oggetto Trasparenza
    Set Trasparenza = Nothing
    'Distruggo l'oggetto FabbricaMateriali,il cui compito � quello di ricreare i materiali,precedentemente
    'settati da applicare a tutti i muri,pavimenti e soffitti analizzati
    Set FabbricaMateriali = Nothing
End Sub
