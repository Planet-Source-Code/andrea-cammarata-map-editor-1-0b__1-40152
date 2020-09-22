VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsOggetti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Da qui in poi comincia la creazione della classe oggetti.
Option Explicit
'Inizioinnanzitutto con il dichiarare l'oggetto che mi permetterà di contenere all'interno
'della mappa 3D l'oggetto che desidererò inserire
Public Scheletro As New Mesh8
'Dichiaro una variabile che mi servirà per identificare inequivocabilmente ogni singolo
'oggetto caricato all'interno della mappa 3D,infatti,ognuno di esso avrà una chiave
'differente da qualsiasi altro
Public Key As String
'La variabile Percorso invece servirà al programma per sapere la fonte di origine
'(CD,Hard Disk,Floppy,etc.) dal quale andare a prelevare l'oggetto da caricare
Public Percorso As String
'La variabile Texture come la variabile Percorso serve al programma per "ricordare"
'dove andare a prelevare l'immagine che verrà applicata sull'oggetto come Texture
Public Texture As String
'Dichiaro una variabile di tipo alfanumerico che permetterà all'utente di aggiungere una breve
'descrizione di Max 200 caratteri.
'Questa descrizione verrà visualizzata in modalità Anteprima 3D ogni qual volta che l'utente
'cliccherà su uno di esso con il bottone sinistra del mouse
Public Decrizione As String
'------------------------------------------------------------------------------------------------
' Ora invece,finito di dichiarare tutte quelle variabili necessarie al caricamento dell'oggetto
' all'interno della mappa,passerò a istanziare tutte quelle addette a posizionare,rimpicciolire
' o ingrandire e ruotare lo stesso secondo volontà dell'utente
'------------------------------------------------------------------------------------------------
'Dichiaro le tre variabili spaziali (X,Y,Z),addette a posizionare l'oggetto desiderato in un punto ben
'preciso all'interno della mappa
Private X As Single
Private Y As Single
Private Z As Single
'Le tre variabili che seguono invece,mi serviranno invece,una volta caricato l'oggetto,
'di poterlo ingrandire o rimpicciolire a mio piacimento.
'Ovviamente,nell'ordine in cui sono state dichiarate le tre variabili servono rispettivamente per
'- Scalare l'oggetto nella sua larghezza
'- Scalare l'oggetto nella sua altezza
'- Scalare l'oggetto nel suo spessore
Private ScaleX As Single
Private ScaleY As Single
Private ScaleZ As Single
'Le ultime tre variabili invece mi servianno per porre all'oggetto appena caricato,la sua giusta
'angolazione di posizione all'interno della mappa.
'Ovviamente anche qui ognuna della tre variabili servirà per ruotare l'oggetto su ognuna delle
'rispettive assi
Private RotationX As Single
Private RotationY As Single
Private RotationZ As Single

Public Sub Carica_Oggetto(Nome_File As String, Chiave As String)
    'Salvo all'interno della variabile Percorso,il valore passato al parametro della funzione chiave.
    'Qusto mi servirà quando si dovranno attivare gli oggetti in modalità Anteprima 3D,in modo che il programma
    'sappia dove andare a cercare l'oggetto da caricare
    Percorso = Nome_File
    'Salvo all'interno della variabile Key il valore passato al parametro Chiave.
    Key = Chiave
    'imposto una grandezza iniziale dell'oggetto appena caricato
    ScaleX = 20
    ScaleY = 20
    ScaleZ = 10
End Sub

Public Sub Carica_Texture(Nome_File As String)
    'Salvo il parametro passato alla funzione stessa,all'interno della variabile Texture
    Texture = Nome_File
End Sub

Public Sub Attiva_Oggetto()
    'Dichiaro una variabile che mi servirà per capire quanto è lunga la stringa percorso
    Dim LunghezzaPercorso As Integer
    'Dichiaro una variabile che mi servirà per controllare l'ultima lettera della stringa
    'percorso.Questa operazione mi serve perchè possono essere inseriti due tipi di oggetti in
    'due formati diversi all'interno della mappa e per ognuno di essi occorre una funzione
    'differente per poterli caricare
    Dim Carattere As String
    'Prelevo la lunghezza della stringa Percorso e la salvo all'interno della variabile
    'LunghezzaPercorso
    LunghezzaPercorso = Len(Percorso)
    'Prelevo l'ultimo carattere della stringa Percorso e lo salvo all'interno della
    'stringa Carattere
    Carattere = Mid(Percorso, LunghezzaPercorso, 1)
    'Se la vaiabile Carattere è uguale a x,quindi si è scelto di caricare un DirectX File,allora...
    If Carattere = "x" Then
        'Carico l'oggetto scelto all'interno della mappa 3D grazie all'ausilio della variabile Percorso,che,ricordo,
        'contiene la path dello stesso
        Scheletro.LoadXFile Percorso
    'Altrimenti...
    Else
        'Carico l'oggetto scelto all'interno della mappa 3D grazie all'ausilio della variabile Percorso,che,ricordo,
        'contiene la path dello stesso
        Scheletro.Load3DsMesh Percorso
    End If
    If Texture <> "" Then
        'Carico la Texture selezionata all'interno della scena,affinchè questa possa essere applicata
        'sull'oggetto
        FabbricaTexture.LoadTexture Texture, "TextureOggetti" + Str(IOg - 1)
        'Applico la Texture selezionata sulla superficie dell'oggetto
        Scheletro.SetTexture GetTex("TextureOggetti" + Str(IOg - 1))
    End If
    'Posiziono l'oggetto nella sua giusta collocazione spaziale selezionata dall'utente
    Scheletro.SetPosition X * VScale, Y, -Z * VScale
    'Imposto i gradi di rotazione su ciascun asse dell'oggetto selezionato
    Scheletro.SetRotation RotationX, RotationY, RotationZ
    'Imposto la dimensione dell'oggetto su ciascun asse
    Scheletro.ScaleMesh ScaleX, ScaleY, ScaleZ
End Sub

Public Function Ricava_Posizione() As D3DVECTOR
    With Ricava_Posizione
        'Restituisco il valore della variabile X la quale contiene la coordinata spaziale sull'asse appunto X
        'dell'oggetto all'interno della mappa 3D
        .X = X
        'Restituisco il valore della variabile Y la quale contiene la coordinata spaziale sull'asse appunto Y
        'dell'oggetto all'interno della mappa 3D
        .Y = Y
        'Restituisco il valore della variabile Z la quale contiene la coordinata spaziale sull'asse appunto Z
        'dell'oggetto all'interno della mappa 3D
        .Z = Z
    End With
End Function

Public Function Ricava_Dimensione() As D3DVECTOR
    With Ricava_Dimensione
        'Restituisco il valore della variabile ScaleX,la quale contiene appunto la larghezza dell'oggetto
        .X = ScaleX
        'Restituisco il valore della variabile ScaleY,la quale contiene appunto l'altezza dell'oggetto
        .Y = ScaleY
        'Restituisco il valore della variabile ScaleZ,la quale contiene appunto lo spessore dell'oggetto
        .Z = ScaleZ
    End With
End Function

Public Function Ricava_Rotazione() As D3DVECTOR
    With Ricava_Rotazione
        'Restituisco il valore della variabile RotationX,la quale contiene appunto i gradi di rotazione
        'dell'oggetto sull'asse appunto delle X
        .X = RotationX
        'Restituisco il valore della variabile RotationY,la quale contiene appunto i gradi di rotazione
        'dell'oggetto sull'asse appunto delle Y
        .Y = RotationY
        'Restituisco il valore della variabile RotationZ,la quale contiene appunto i gradi di rotazione
        'dell'oggetto sull'asse appunto delle Z
        .Z = RotationZ
    End With
End Function

Public Sub Setta_Dimensione(Asse As String, NuovoValore As Single)
    'Esamino i vari valori che può aver assunto il paramtro Asse passato alla funzione stessa
    Select Case Asse
    'Se il parametro ha assunto il valore X...
    Case Is = "X"
        'Setto la dimensione in larghezza dell'oggetto con il suo nuovo valore passato alla funzione
        'stessa tramite il parametro appunto NuovoValore
        ScaleX = NuovoValore
    'Se il parametro ha assunto il valore Y...
    Case Is = "Y"
        'Setto la dimensione in altezza dell'oggetto con il suo nuovo valore passato alla funzione
        'stessa tramite il parametro appunto NuovoValore
        ScaleY = NuovoValore
    'Se il parametro ha assunto il valore Z...
    Case Is = "Z"
        'Setto la dimensione in spessore dell'oggetto con il suo nuovo valore passato alla funzione
        'stessa tramite il parametro appunto NuovoValore
        ScaleZ = NuovoValore
    End Select
    'Infine imposto la nuova dimensione dell'oggetto con le grandezze aggiornate
    Scheletro.ScaleMesh ScaleX, ScaleY, ScaleZ
End Sub

Public Sub Setta_Posizione(Asse As String, NuovoValore As Single)
    'Esamino i vari valori che può aver assunto il paramtro Asse passato alla funzione stessa
    Select Case Asse
    'Se il parametro asse ha assunto il valore X allora...
    Case Is = "X"
    'Setto la coordinata spaziale X dell'oggetto con il suo nuovo valore passato
    'alla funzione stessa tramite il parametro appunto NuovoValore
    X = NuovoValore
    'Se il parametro asse ha assunto il valore Y allora...
    Case Is = "Y"
        'Setto la coordinata spaziale Y dell'oggetto con il suo nuovo valore passato
        'alla funzione stessa tramite il parametro appunto NuovoValore
        Y = NuovoValore
    'Se il parametro asse ha assunto il valore Z allora...
    Case Is = "Z"
        'Setto la coordinata spaziale Z dell'oggetto con il suo nuovo valore passato
        'alla funzione stessa tramite il parametro appunto NuovoValore
        Z = NuovoValore
    End Select
    'Infine imposto la nuova posizione dell'oggetto con le coordinate aggiornate
    Scheletro.SetPosition X * VScale, Y, -Z * VScale
End Sub

Public Sub Setta_Rotazione(Asse As String, NuovaAngolazione As Single)
    'Esamino i vari valori che può aver assunto il paramtro Asse passato alla funzione stessa
    Select Case Asse
    'Se il parametro asse ha assunto il valore X allora...
    Case Is = "X"
        'Setto la rotazione dell'oggetto sull'asse delle X con il nuovo angolo
        'passato alla funzione stessa tramite il parametro NuovaAngolazione
        RotationX = NuovaAngolazione
    'Se il parametro asse ha assunto il valore Y allora...
    Case Is = "Y"
        'Setto la rotazione dell'oggetto sull'asse delle Y con il nuovo angolo
        'passato alla funzione stessa tramite il parametro NuovaAngolazione
        RotationY = NuovaAngolazione
    'Se il parametro asse ha assunto il valore Z allora...
    Case Is = "Z"
        'Setto la rotazione dell'oggetto sull'asse delle Z con il nuovo angolo
        'passato alla funzione stessa tramite il parametro NuovaAngolazione
        RotationZ = NuovaAngolazione
    End Select
    'Ora pongo un limite a tutte e tre le variabile di rotazione,in modo che nessuna di questa
    'supera i 360°. Non avrebbe senso ruotare l'oggetto di 360°,in quanto questo farebbe un giro
    'su se stesso e poi tornerebbe alla sua rotazione di origine,come non avrebbe senso
    'ruotare l'oggetto di 340°,basterebbe schivere 10°
    If RotationX >= 360 Then RotationX = RotationX - 360
    If RotationY >= 360 Then RotationY = RotationY - 360
    If RotationZ >= 360 Then RotationZ = RotationZ - 360
    'Infine imposto la nuova rotazione dell'oggetto con i gradi aggiornati
    Scheletro.SetRotation RotationX, RotationY, RotationZ
End Sub

Public Sub Distruggi_Oggetto()
    'iniziamo con azzerare le variabili alfanumeriche contenenti la path dell'oggetto
    'e della Texture
    Percorso = ""
    Texture = ""
    'Ora tocca riportare a 0 le variabili contenenti le variabili spaziali dell'oggetto...
    X = 0
    Y = 0
    Z = 0
    '...azzerare le sue dimensioni...
    ScaleX = 0
    ScaleY = 0
    ScaleZ = 0
    '...riportare a 0 gli angoli di rotazione...
    RotationX = 0
    RotationY = 0
    RotationZ = 0
    '...ed infine distruggere l'oggetto
    Set Scheletro = Nothing
End Sub

