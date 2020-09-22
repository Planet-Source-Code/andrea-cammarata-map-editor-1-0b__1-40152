Attribute VB_Name = "Module1"
'Inizializzo l 'oggetto principale dal quale diperderanno tutte le operazioni di
'rendering e non solo
Public TV8 As New TrueVision8
'Creo una nuova scena
Public Scena As New Scene8
'Dichiaro di tipo pubblico l'oggetto Suoni di tipo Classe_sonora,la quale mi servirà
'a riprodurre dei passi quando ci si sposterà all'interno della mappa
Public Suoni As New ClsClasse_sonora
'Definisco: un nuovo schermo su cui potrò tracciare le righe che serviranno per costruire
'i muri della nuova mappa...
Public Schermo As New Screen8
'Dichiaro un oggetto che mi servirà a selezionare ogni singolo oggetto caricato all'interno della
'mappa semplicemente cliccandoci sopra in modalità Anteprima 3D
Public Collisione As CollisionResult8
'Definisco il tipo TColor il quale mi servirà per impostare tutti i parametri di materiale
'nelle rispettive voci,e il colore delle linee da disegnare sull'editor
Type TColor
    R As Single
    G As Single
    B As Single
    A As Single
End Type
'Definisco il tipo TMateriale, il quale mi servirà per impostare le qualità del materiale
'del muro,pavimento o soffitto selezionato
Type TMateriale
    Ambiente As TColor
    Diffusa As TColor
    Emissiva As TColor
    Potenza As Single
    Speculare As TColor
End Type
'Creo il tipo Coordinate_Riga che mi servirà per impostare le coordinate della nuova
'riga che verrà creata
Type Coordinate_Riga
    X1 As Single
    Y1 As Single
    X2 As Single
    Y2 As Single
    SpigoloI As Boolean
    SpigoloF As Boolean
    Altezza As Single
    Altitudine As Single
    Texture As String * 500
    Nome As String * 20
    Proprietà As String * 20
    Materiale As TMateriale
    NMattonelleALtezza As Single
    NMAttonelleLarghezza As Single
    ColVertici(0 To 3) As TColor
End Type
'Creo il tipo Coordinate_Sop il quale mi servirà per impostare tutte le coordinate
'e i rispettivi parametri del nuovo soffitto o pavimento che verrà creato
Type Coordinate_Sop
    'Ho aggiunto al tipo Coordinate_Sop il record """Ereditato""" Coordinate_Sop.
    'Questo mi ha evitato di riscrivere nuovamente all'interno di questo tipo tutti
    'i parametri che alla fine sono gli stessi che servono alla costruzione dei muri.
    'PS: l'unico parametro che non mi servirà sarà quello di Altezza
    CR As Coordinate_Riga
    Tipo As String * 10
    X3 As Single
    Y3 As Single
    X4 As Single
    Y4 As Single
End Type
'Dichiaro il Tipo Type_Ogg il quale mi servirà per contenere i valori di riferimento di tutti
'gli oggetti caricati all'interno della mappa 3d,al fine di poter effettuare le dovute
'operazioni di salvataggio
Type Type_Ogg
    Appartenenza As String * 50
    Percorso As String * 300
    Texture As String * 300
    Key As String * 50
    X As Single
    Y As Single
    Z As Single
    ScaleX As Single
    ScaleY As Single
    ScaleZ As Single
    RotationX As Single
    RotationY As Single
    RotationZ As Single
End Type
'Creo un array contenente 10000 elementi del tipo Coordinate_Riga
Public Riga(0 To 10000) As Coordinate_Riga
'Creo un array contenente 10000 elementi del tipo Coordinate_SoP
Public SoP(0 To 10000) As Coordinate_Sop
'Dichiaro una varibile uguale al tipo definito Coordinate_Riga.Questa mi servirà per modificare i
'record del file che andrò a salvare
Public Muro As Coordinate_Riga
'Dichiaro una varibile uguale al tipo definito Coordinate_Sop.Anche questa mi servirà per modificare i
'record del file che andrò a salvare
Public SofPav As Coordinate_Sop
'Dichiaro una variabile del tipo definito TColor,la quale mi servirà per impostare il colore di alcune linee
'da disegnare all'interno dell'editor
Public Colore As TColor
'Dichiaro una variabile di tipo definito Type_Ogg,al fine di poter salvare all'interno di un file tutti i
'riferimenti necessari agli oggetti caricati all'interno della mappa 3d
Public Ogg As Type_Ogg
'--------------------------------------------------------------------------------------------------------
' Ora dichiaro 11 variabili del tipo definito TColor,ognuna delle quali mi servirà per contenere il colore
' del rispettivo componente all'interno dell'editor
'--------------------------------------------------------------------------------------------------------
'La variabile CM mi servirà per contenere il colore di tutte quelle linee che rappresentano i muri
'all'interno dell'editor
Public CM As TColor
'La variabile CMS mi servirà invece per contenere il colore della linea che rappresenta il muro
'selezionato dall'oggetto ElencoMuri all'inerno del Form_opzioni
Public CMS As TColor
'La variabile CSM invece mi servirà per contenere il colore di tutti quei quadratini che si formeranno
'in seguito all'intersezione di due o più muri,in pratica di tutti quelli che formeranno uno spigolo
Public CSM As TColor
'Quest'altra variabile,la CS,mi servirà per contenere il colore di tutte quelle linee che rappresentano
'i soffitti all'interno della superficie dell'editor
Public CS As TColor
'Mentre la CP conterrà il colore di tutte quelle linee che rappresentano i pavimenti all'interno
'dell'editor
Public CP As TColor
'La variabile CSOPS conterrà il colore dei pavimenti o soffitti che verranno selezionati al fine di apportare
'modifiche dall'oggetto Elenco_SoP presente nel Form_Opzioni
Public CSOPS As TColor
'La variabile CAM conterrà il colore dei quadratini che si formeranno in corrispondenza a tutti quei muri
'che presentano una stessa cordinato del mouse dell'editor
Public CAM As TColor
'La variabile CASOP invece è identica alla CAM,solo che questa conterrà il colore dell'allineamento
'tra i pavimenti / soffitti e il mouse dell'editor
Public CASOP As TColor
'La variabile CLG conterrà il colore delle linee guida che si formeranno perpendicolarmente alle coordinate del mouse
Public CLG As TColor
'La variabile CSFM conterrà il colore di sfondo del menù posto in alto all'editor
Public CSFM As TColor
'Mentre la variabile C1M conterrà il 1° colore utilizzato dal menù...
Public C1M As TColor
'...E la variabile C2M conterrà a sua volta il secondo colore utilizzato dallo stesso Menù
Public C2M As TColor
'--------------------------------------------------------------------------------------------------------
'Quest'altra variabile serve invece funge da contatore dei muri presenti
Public Max As Integer
'Questa invece serve per contare il numero di pavimenti + il numero dei soffitti presenti
Public Max2 As Integer
'Quest'altra per contare il numero massimo di pavimenti presenti
Public Max3 As Integer
'E quest'altra per contare il numero massimo di soffitti presenti
Public Max4 As Integer
'Definisco un indice per identificare ogni singolo muro
Public I As Integer
'..un indice per i soffitto e i pavimenti...
Public J As Integer
'Dichiaro una variabile che permetterà di far capire al programma quale oggetto
'dovrà essere inserito
Public Scelta_Oggetto As String
'Dichiaro un indice che mi servirà per identificare ogni singolo elemento
'dell'oggetto ComboBox ElencoMuri
Public IndiceLista As Long
'Dichiaro un indice che mi servirà per identificare ogni singolo elemento
'dell'oggetto ComboBox ElencoSoP
Public IndiceLista2 As Long
'Dichiaro una variabile che conterrà il valore dello scale selezionato dal form opzioni
Public VScale As Integer
'Dichiaro una variabile che mi servirà per contenere le coordinate della telecamera
Public PosizioneTelecamera As D3DVECTOR
'Dichiaro una variabile che mi servirà per capire quale file è stato salvato.Questa
'mi tornerà utile se dovessi decidere di salvare nuovamente il file ma non con nome
Public FileSalvato As String
'Dichiaro una variabile che servirà per capire in quale file è stata convertita la mappa
'corrente
Public FileConvertito As String
'Dichiaro una variabile che mi servirà nella funzione di zoom.
'Ora Spiego:
'Quando veniva premuto il tasto Zoom+ o Zoom-,le coordinate delle righe presenti all'interno della
'mappa attuale venivano modificate,e quindi quando avveniva un'operazione di salvataggio o di conversione
'venivano riscontrati dei problemi,in quanto, le coordinate delle righe,non rispettavano effettivamente
'quelle definite.
'Dichiarando questa variabile,invece si tiene conto di quanto sono state modificate le coordinate delle righe,
'in modo che,al momento del salvataggio o della connversione,queste vengano divise per
'il valore contenuto in Molt affinchè vengano ripristinate alle loro dimensioni originali
Public Molt As Double
'Dichiaro una variabile che mi servirà per contenere i cambiamenti della griglia effettuati tramite le
'operazioni di Zoom.
'Infatti se nel form opzioni è selezionata la voce Ridimensiona griglia con Zoom,questa,ogni qualvolta
'che verrà effettuato uno zoom in ingrandirà i "quadrati della griglia",in caso contrario
'li ridurrà
Public VCambiamentiGriglia  As Single
'Dichiaro una variabile che mi servirà per capire in che modalità avviare il Form_Assegnazione_Multipla
'Le modalità sono due:
'1 - Per i muri
'2 - Per i pavimenti e soffitti
Public ModalitàAssegnazioneMultipla As String
'Dichiaro come sopra, un'altra variabile che mi servirà per capire in che modalità è stato avviato il form di
'assegnazione del materiale.
'Anche qui le modalità sono le stesse di quelle per l'assegnazione multipla
Public ModalitàGestioneMateriale As String
'Dichiaro una variabile che mi servirà per capire quale immagine si dovrà caricare come fondale dell'editor
Public ImmagineSfondo As String
'Ora creo un vettore di classi Oggetti,in modo che ognuno di questi possa essere indipendente,ognuno
'con le sue coordinate spaziali,dimensione e angoli di rotazione
Public Oggetto(0 To 10000) As New ClsOggetti
'Dichiaro un indice che mi servirà per tenere conto del numero di oggetti inseriti all'interno della
'mappa 3D
Public IOg As Integer
'QUest'altra variabile mi servirà invece per contenere il suo indice all'interno dell'array Oggetto
Public IndiceOggettoSelezionato As Integer
'Dichiaro un oggetto che mi permetterà di caricare all'interno della Scena di Rendering tutte quelle Texture
'che andranno applicate sulla superficie degli oggetti caricati all'interno della mappa 3D
Public FabbricaTexture As New TextureFactory8
'Dichiaro una variabile che servirà al programma per capire qual'è appunto la lingua corrente del programma.
'Le lingue disponibili sono due:
'- Italiano
'- Inglese
Public LinguaS As String

Public Sub Reimposta_materiale(Materiale As TMateriale)
    'Questa funzione ha l'utilità di reimpostare un determinato materiale passato alla funzione stessa con il colore nullo
    'che in questo caso è il bianco
    'Risetto i parametri del materiale passato alla funzione stessa alla voce Ambiente con il colore bianco
    With Materiale.Ambiente
        .R = 255
        .G = 255
        .B = 255
        .A = 1
    End With
    'Risetto i parametri del materiale passato alla funzione stessa alla voce Diffusa con il colore bianco
    With Materiale.Diffusa
        .R = 255
        .G = 255
        .B = 255
        .A = 1
    End With
    'Risetto i parametri del materiale passato alla funzione stessa alla voce Emissiva con il colore bianco
    With Materiale.Emissiva
        .R = 255
        .G = 255
        .B = 255
        .A = 1
    End With
    'Risetto i parametri del materiale passato alla funzione stessa alla voce Speculare con il colore bianco
    With Materiale.Speculare
        .R = 255
        .G = 255
        .B = 255
        .A = 1
    End With
    'Risetto la potenza del materiale passato alla funzione stessa alla voce Potenza con un valore iniziale
    'pari a 500
    Materiale.Potenza = 500
End Sub

Public Sub Preleva_RGB(Componente As TColor)
    Set SceltaColori = Form_Materiali.ControlloSceltaColori
    'Questa particolare funzione mi permetterà di estrarre i rispettivi valori RGB dal colore LONG
    'appena selezionato, e assegnarli, in base al valore passato, al vertice del muro,pavimento o soffitto
    'selezionato
    With Componente
        'Estraggo la quantità di Blu
        .B = ((SceltaColori.Color And 16711680) / 65536)
        'Estraggo la quantità di Verde
        .G = (((SceltaColori.Color And 65280) / 256) Mod 256)
        'Estraggo la quantità di Rosso
        .R = (SceltaColori.Color Mod 256)
    End With
End Sub

Public Sub RGB_To_RGBA(Componente As TColor)
    'Questa particolare è utilissima funzione da me ideata e realizzata mi permetterà di convertire dei colori
    'espressi secondo il loro originale formato RGB in formato RGBA.
    'Questo mi serve perchè l'editor lavora in un componente semi 3D, cioè un particoloare oggetto
    'chiamato Screen8 aovvero uno schermo,e a differenza della comunissima grafica di Visual BAsic
    'i colori devono essere espressi in valori RGBA [R(RED),G(GREEN),B(BLUE),A(ALPHA)].
    'Quest'ultimo parametro può essere considerato come l'indice di trasparenza o di luminosità,
    'infatti tanto più alto sarà questo valore,tanto meno trasparente sarà l'oggetto che lo possederà,
    'oppure meno solido.
    'Altra differenza di questo formato rispetto all'RGB,è che questo contiene una tripletta di valori
    'che partono da 0 e arrivano a 255 e sono di tipo Long,mentre il tipo RGBA,contiene una quadripetta
    'di valori che partono da 0 e arrivano fino a 1,passando da 0.1,0.2,0.3 ecc.e sono quindi di tipo
    'single.
    'In pratica questa funzione si propone di suddividere il valore massimo dei valori RGB (255) in
    '10 parti dato che l'RGBA parte da 0 e giunge sino al valore 1.
    'A questo punto se un valore dell'RGB sarà inferiore o uguale alla prima partizione del valore 255,
    'ovvero 25.5 questo assumerà il valore della prima partizione dell'RGBA ovvero 0.1,se invece
    'sarà maggiore della prima partizione dell'RGB ma minore o uguale alla sua secondo,questo assumerà
    'ovviamente il valore della seconda partizione dell'RGBA ovvero 0.2 e così via.
    With Componente
        'Avvio la trasformazione del valore espresso in RGB del Rosso in RGBA
        If .R <= 25.5 Then .R = 0.1
        If .R > 25.5 And .R <= 51 Then .R = 0.2
        If .R > 51 And .R <= 76.5 Then .R = 0.3
        If .R > 76.5 And .R <= 102 Then .R = 0.4
        If .R > 102 And .R <= 127.5 Then .R = 0.5
        If .R > 127.5 And .R <= 153 Then .R = 0.6
        If .R > 153 And .R <= 178.5 Then .R = 0.7
        If .R > 178.5 And .R <= 204 Then .R = 0.8
        If .R > 204 And .R <= 229.5 Then .R = 0.9
        If .R > 229.5 And .R <= 255 Then .R = 1
        'Ora trasformo anche la quantità di Verde
        If .G <= 25.5 Then .G = 0.1
        If .G > 25.5 And .G <= 51 Then .G = 0.2
        If .G > 51 And .G <= 76.5 Then .G = 0.3
        If .G > 76.5 And .G <= 102 Then .G = 0.4
        If .G > 102 And .G <= 127.5 Then .G = 0.5
        If .G > 127.5 And .G <= 153 Then .G = 0.6
        If .G > 153 And .G <= 178.5 Then .G = 0.7
        If .G > 178.5 And .G <= 204 Then .G = 0.8
        If .G > 204 And .G <= 229.5 Then .G = 0.9
        If .G > 229.5 And .G <= 255 Then .G = 1
        'Infine trasformo anche la quantità di Blu
        If .B <= 25.5 Then .B = 0.1
        If .B > 25.5 And .B <= 51 Then .B = 0.2
        If .B > 51 And .B <= 76.5 Then .B = 0.3
        If .B > 76.5 And .B <= 102 Then .B = 0.4
        If .B > 102 And .B <= 127.5 Then .B = 0.5
        If .B > 127.5 And .B <= 153 Then .B = 0.6
        If .B > 153 And .B <= 178.5 Then .B = 0.7
        If .B > 178.5 And .B <= 204 Then .B = 0.8
        If .B > 204 And .B <= 229.5 Then .B = 0.9
        If .B > 229.5 And .B <= 255 Then .B = 1
    End With
End Sub
