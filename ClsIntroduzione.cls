VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsIntroduzione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Dichiaro un oggetto che mi servirà per carica appunto l'anello 3D che mi sono creto al fine
'di migliorare la presentazione del programma
Private Anello3D As New Mesh8
'Ora dichiaro lo stesso tipo di oggetto di quello sopra indicato,con la sola differenza che
'questo mi servirà per costruire il rettangolo che ruoterà all'interno dell'anello
Private OggettoTesto As New Mesh8
'Questi tre ogetti invece mi serviranno per dare all'oggetto sopra indicato un aspetto personalizzato
'permettendomi di scriverne sulla sua superficie il mio nome e cognome ed altro ancora
Private Testo(0 To 2) As New RenderSurface8
'Ora dichiaro uno scehrmo da unitizzare per creare la texture(l'immagine) contenente
'il mio nome e altro da applicare sui due tre oggetti appena dichiarati
Private SchermoIntroduzione As New Screen8
'Dichiaro una variabile che terrà conto della rotazione dell'anello 3D,questa verrà sempre incrementata
'di 0,01 al fine di non farlo mai fermare
Private Rotazione As Single
'Quest'altra variabbile,come Rotazione,tiene conto della rotazione del rettangolo costruito
'all'interno dell'anello 3D
Private RotazioneTesto As Single
'L'ultimo oggetto,dichiarato pubblico perchè mi dovrò richiamarlo dall'interno del Form1
'mi serve per creare un suggestivo effetto neve all'interno di questa animazione
Public Terreno As New Landscape8

Public Sub Start()
    'Per prima cosa richiamo la funzione che mi permetterà di caricare tutte le Texture(immagini) necessarie
    'alla creazione dell'animazione
    Carica_Textures
    'Ora setto l'oggetto Anello3D in modo che possa essere caricato al suo interno l'anello 3D che mi ero precedentemente
    'creato.
    'ATTENZIONE: Questa istruzione è fondamentale!
    Set Anello3D = Scena.CreateMeshBuilder
    'Ora settato l'oggetto Anello3D sono pronto a caricare al suo interno il "famoso" anello 3D
    Anello3D.LoadXFile "Anello3d.x"
    'Imposto la posizione dell'anello in coordinate X = 0,Y = 0 e Z = 10
    Anello3D.SetPosition 0, 0, 10
    'Ora ingrandisco l'anello 3D:
    '- 20 volte più grande in larghezza
    '- 20 volte più grande in altezza
    '- 10 volte più grande in spessore
    Anello3D.ScaleMesh 20, 20, 10
    'Attivo una particolare funzione 3D che mi permetterà di simulare una falsa riflessione della
    'Texture (immagine) che verrò applicata subito dopo sulla superficie dell'anello
    Anello3D.EnableSphereMapping True
    'Ora carico la texture che era stata caricata all'interno del programma dalla funzione Carica_Texture
    'e la applico sulla superficie dell'anello 3D
    Anello3D.SetTexture GetTex("TextureAnello")
    'Ora come già fatto per l'anello 3D setto l'oggetto OggettoTesto in modo da poter creare il rettangolo che si vedrà
    'ruotare all'interno dell'anello 3D,indicando il mio nome,cognome e altro
    Set OggettoTesto = Scena.CreateMeshBuilder
    'Ora creo la prima Texture del rettangolo con l'aiuto dallo schermo precedentemente dichiarato rappresentante il mio nome e cognome
    Set Testo(0) = Scena.CreateRenderSurface(300, 50, False)
    'Da qui ha inizio la creazione della prima Texture
    Testo(0).StartRender
        'Creo all'interno dello schermo precedentemente dichiarato un "box",cioè un piccolo rettangolo riempito da quattro colori,
        'uno per ogni vertice dello stesso su cui scrivere il mio nome e cognome
        SchermoIntroduzione.DrawFilledColorBox 0, 0, 300, 80, RGBA(0, 0, 1, 0.6), RGBA(0, 1, 1, 0.6), RGBA(0.5, 0.3, 0, 0.6), RGBA(0, 0.8, 0.8, 0.6)
        'Creato ilrettangolo multi colore sono pronto per scrivere su di esso e quindi anche sulla prima texture
        'il mio nome e cognome
        SchermoIntroduzione.DrawText "Andrea Cammarata", 1, 1, RGBA(1, 1, 1, 1), "Carattere_Animazione_1"
    'Qui termina la creazione della prima Texture
    Testo(0).EndRender
    'Ora come già fatto per la prima,inizio a costruire con l'aiuto dello stesso schermo la seconda Texture
    Set Testo(1) = Scena.CreateRenderSurface(300, 50, False)
    'Da qui ha invece inizio la crezione della seconda Texture
    Testo(1).StartRender
        'Come già fatto prima,creo sempre lo stesso rettangolo multicolore su cui disegnare la scritta "Programma scritto da:"
        SchermoIntroduzione.DrawFilledColorBox 0, 0, 300, 80, RGBA(0, 0, 1, 0.6), RGBA(0, 1, 1, 0.6), RGBA(0.5, 0.3, 0, 0.6), RGBA(0, 0.8, 0.8, 0.6)
        'Ora posso disegnare la scritta
        SchermoIntroduzione.DrawText "Program written by:", 10, 5, RGBA(1, 1, 1, 1), "Carattere_Animazione_2"
    'Da qui finisce anche la creazione della seconda Texture
    Testo(1).EndRender
    'Ormai rimane solamente più una Texture da creare,quindi incomincio a crearla,sempre con l'aiuto dello
    'schermo precedentemente creato
    Set Testo(2) = Scena.CreateRenderSurface(30, 50, False)
    'Inizio la crezione della terza e ultima Texture
    Testo(2).StartRender
        'Questa volta disegno solamente il solito rettangolo multiicolore,ma non verrà scritto niente su di esso
        SchermoIntroduzione.DrawFilledColorBox 0, 0, 30, 80, RGBA(0, 0, 1, 0.6), RGBA(0, 1, 1, 0.6), RGBA(0.5, 0.3, 0, 0.6), RGBA(0, 0.8, 0.8, 0.6)
    'Ecco anche finita la creazione dell'ultima Texture
    Testo(2).EndRender
    'Ora che mi sono creato tutte le Texture da applicare al rettangolo rotante,non mi resta che applicarle sulle relative
    'facce...
    With OggettoTesto
        'Creo la prima faccia su cui applicherò la Texture indicante il mio nome e cognome
        .AddWall Testo(0).GetTexture, -10, 0, 10, 0, 5
        'Creo la seconda faccia,parallela e posta alle spalle della prima su cui avevo creato la
        'scritta "Programma scritto da:"
        .AddWall Testo(1).GetTexture, 10, -2, -10, -2, 5
        '-----------------------------------------------------------------------------------
        ' Ora le due facce principali sono state create,non mi resta che creare le rimanenti
        ' quattro,ovvero le due laterali quella superiore e quella inferiore.
        '-----------------------------------------------------------------------------------
        'Creo la faccia laterale sinistra su cui applicherò l'unica Texture senza scritta
        .AddWall Testo(2).GetTexture, -10, 0, -10, -2, 5
        'Creo la faccia laterale destra e anche qui come la precedente e le prossime due
        'applicherò su di essa la Texture senza scritta
        .AddWall Testo(2).GetTexture, 10, 0, 10, -2, 5
        'Creo la faccia inferiore,applicando su di essa sempre la stessa Texture
        .AddFloor Testo(2).GetTexture, -10, 0, 10, -2, 0
        'Creo la faccia inferiore e come le precedenti tre,applicherò sempre la solita Texture senza scritta
        .AddFloor Testo(2).GetTexture, -10, 0, 10, -2, 5
        .SetPosition 0, 0, 20
    End With
    'Mediante questa istruzione, viene simulato un effetto di caduta neve al fine di migliore l'animazione.
    'La neve viene creata mediante la Texture caricata all'interno del programma mediante la funzione
    'Carica_Texture
    Terreno.InitRain 100, GetTex("Neve"), , , , 1
End Sub

Public Sub Ruota_anelli()
    '--------------------------------------------------------------------------------------------
    ' Questa funzione dichiarata pubblica in quanto mi servirà richiamarla dall'interno del Form1
    ' Mi permette di far effettuare una rotazione dell'anello 3D e delrettangolo indicante il mio
    ' nome,cognome e altro
    '--------------------------------------------------------------------------------------------
    'Incremento la variabile Rotazione di 0.01 in modo che l'annello 3D non stia mai fermo ma simuli
    'continuamente un effetto di rotazione
    Rotazione = Rotazione + 0.01 + TV8.TimeElapsed / 20
    'Ora posso effettuare una rotazione dell'anello 3D sull'asse delle Y grazie alla variabile rotazione
    Anello3D.SetRotation 0, Rotazione, 0
    'Incremento,come già fatto per la variabile Rotazione,anche RotazioneTesto,anche qui per non fare
    'mai fermare il rettangolo
    RotazioneTesto = RotazioneTesto + 0.01 + TV8.TimeElapsed / 50
    'Ora effettuo la rotazione del rettangolo sull'asse X e sull'asse Z.
    'Quest'ultima rotazione avverrà nel senso opposto rispetto a quella effettuata sull'asse X
    OggettoTesto.SetRotation RotazioneTesto, 0, -RotazioneTesto
End Sub

Public Sub Distruggi_Animazione()
    'Resetto l'oggetto Anello 3D,in modo che l'annello 3F precedentemente caricato,non sia più visibile
    Anello3D.ResetMesh
    'Resetto anche l'oggetto OggettoTesto, per eliminare tutte le faccie che si erano create
    OggettoTesto.ResetMesh
    '-----------------------------------------------------------------------------------------------------
    ' Infine distruggo tutti gli oggetti che erano stati istanziati al fine di costruire questa animazione
    '-----------------------------------------------------------------------------------------------------
    'Distruggo l'oggetto Anello 3D
    Set Anello3D = Nothing
    'Distruggo l'oggetto OggettoTesto
    Set OggettoTesto = Nothing
    'Distruggo l'oggetto Testo(0)
    Set Testo(0) = Nothing
    'Distruggo l'oggetto Testo(1)
    Set Testo(1) = Nothing
    'Distruggo l'oggetto Testo(2)
    Set Testo(2) = Nothing
    'Distruggo l'oggetto Terreno
    Set Terreno = Nothing
End Sub

Private Sub Carica_Textures()
    '------------------------------------------------------------------------------------------------------
    ' Qusta funzione mi serve al fine di caricare all'interno del programma tutte quelle Texture (immagini)
    ' e caratteri,necessari alla creazione dell'animazione
    '------------------------------------------------------------------------------------------------------
    'Carico la Texture da applicare sulla superficie dell'anello 3D
    Scena.LoadTexture "Images/lblue005.jpg", , , "TextureAnello"
    'Carico la Texture necessaria a simulare l'effetto neve
    Scena.LoadTexture "Images/Neve.bmp", , , "Neve"
    'Creo il primo carattere con cui disegnerò la scritta contenente il mio nome e cognome
    SchermoIntroduzione.CreateUserFont "Carattere_Animazione_1", "Comic sans MS", 20, True, False, False
    'Creo il secondo carattere'un pò più piccolo del primo per disegnare la scritta "Programma scritto da:"
    SchermoIntroduzione.CreateUserFont "Carattere_Animazione_2", "Comic sans MS", 16, True, False, False
End Sub



