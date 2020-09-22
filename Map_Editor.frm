VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Map_Editor 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Map Editor 1.0"
   ClientHeight    =   8055
   ClientLeft      =   600
   ClientTop       =   1770
   ClientWidth     =   8310
   Icon            =   "Map_Editor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   8310
   Begin MSComDlg.CommonDialog Operazioni 
      Left            =   7680
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame Frame1 
      Caption         =   "Editor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin VB.PictureBox Editor 
         BorderStyle     =   0  'None
         Height          =   7575
         Left            =   120
         ScaleHeight     =   505
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   537
         TabIndex        =   1
         Top             =   240
         Width           =   8055
      End
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu Nuovo 
         Caption         =   "&Nuovo"
      End
      Begin VB.Menu Carica_mappa 
         Caption         =   "Carica &Mappa"
         Visible         =   0   'False
      End
      Begin VB.Menu Barra1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu Salva 
         Caption         =   "&Salva"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu Salva_con_nome 
         Caption         =   "Salva con &nome"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu Barra 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu Converti_mappa_in_3D 
         Caption         =   "&Converti mappa in &3D"
         Enabled         =   0   'False
         Shortcut        =   ^C
         Visible         =   0   'False
      End
      Begin VB.Menu Linea4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu Stampa 
         Caption         =   "S&tampa"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu Linea3 
         Caption         =   "-"
      End
      Begin VB.Menu Esci 
         Caption         =   "&Esci"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu Visualizza 
      Caption         =   "&Visualizza"
      Begin VB.Menu Opzioni 
         Caption         =   "&Opzioni"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu m3D 
      Caption         =   "&3D"
      Begin VB.Menu AvviaAnteprima 
         Caption         =   "A&vvia Anteprima"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu StoppaAnteprima 
         Caption         =   "S&toppa Anteprima"
         Enabled         =   0   'False
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu Lingua 
      Caption         =   "&Lingua"
      Begin VB.Menu LItaliano 
         Caption         =   "&Italiano"
      End
      Begin VB.Menu LInglese 
         Caption         =   "In&glese"
      End
   End
   Begin VB.Menu Info 
      Caption         =   "&?"
      Visible         =   0   'False
      Begin VB.Menu Registra 
         Caption         =   "&Registra"
      End
      Begin VB.Menu InfoMapEditor 
         Caption         =   "&Informazioni su Map Editor 1.0"
      End
   End
End
Attribute VB_Name = "Map_Editor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dichiaro un oggetto di tipo Introduzione,ovvero la classe che mi sono costruito per
'creare una semplice animazione all'avvio del programma
Dim Animazione As New ClsIntroduzione
'Dichiaro un'oggetto di tipo ClsMappa3D,ovvero la classe che mi sono costruito al fine di
'visualizzare un'anteprima 3D della mappa corrente
Dim Mappa3D As New clsMappa3D
'Dichiaro un'oggetto di tipo Tele,ovvero la classe che mi permetterà di muovermi all'interno
'della mappa appena creata
Dim Telecamera As New ClsTele
'...e l'oggetto comandi che mi permetterà di ricevere input direttamente da mouse e tastiera
Dim Comandi As New InputEngine8
'Dichiaro un oggetto che mi servirà per effettuare alcuni effetti grafici come una sfumatura
'in entrata e in uscita
Dim Effetti As New GraphicEffect8
'La variabile Continua,serve da riconoscimento al motore 3D per verificare se siamo all'interno
'del ciclo principale,in cui vengono svolte tutte le funzioni del programma.
'Se questa assumerà il valore di false allora si uscirà dal ciclo e quindi dal programma
Dim Continua As Boolean
'Dichiaro una variabile che riferirà al programma quando viene scelto di costruire
'una nuova mappa e quindi quando dovrà terminare l'animazione
Dim Continua_Animazione As Boolean
'Dichiaro un'altra variabile simile alle due precedenti,con la sola differenza che questa
'riferirà al programma se si dovrà continuare a ciclare nel ciclo che permetterà di
'visualizzare la mappa corrente in 3D
Dim Continua_Anteprima_3d As Boolean
'Definisco due variabili di appoggio che manterranno il valore di larghezza e altezza della
'finestra 3D e più precisamente della picturebox "Editor".Queste mi torneranno utili quando
'dovrò disegnare il menù su schermo
Dim Larghezza As Single
Dim Altezza As Single
'Ora inizializzo le variabili che mi restituiranno le coordinate del mouse...
Dim MouseX As Long
Dim MouseY As Long
'...e quelle che riferiranno al motore 3D se ho premuto o il bottone sinistro del mouse,
'o quello destro,rispettivamente B1 e B2
Dim B1 As Integer
Dim B2 As Integer
'Ora mi servono altre due variabile che conterranno le coordinate del mouse però questa volta
'le definisco Single.Tutto ciò perchè molte funzioni predefinite richiedono i valori
'espressi in single
Dim SmouseX As Single
Dim SmouseY As Single
'Le variabili che seguono sono riferite alle righe e mi servono per salvare le coordinate
'temporanee della riga che stò analizzando
Dim TmpX1 As Single
Dim TmpX2 As Single
Dim TmpY1 As Single
Dim TmpY2 As Single
'...un indice di appoggio per i muri...
Dim AppI As Integer
'...e un indice di appoggio per i pavimenti e soffitti
Dim AppJ As Integer
'Definisco una variabile che indicherà lo stato di costruzione della nuova Riga
Dim Stato_riga As Integer
'Definisco una variabile che indicherà lo stato di costruzione del nuovo soffitto o pavimento
Dim Stato_sop As Integer
'Questa variabile invece mi serve per capire se è stata scelta la voce di menù
'Salva o Salva con nome
Dim SCN As Boolean
'Questa variabile invece mi serve per capire se è stata scelta la voce di menù Converti Mappa in 3D o
'aggiorna mappa 3D
Dim CM3D As Boolean
'Dichiaro una variabile booleana che mi servirà per capire se il menù di aiuto in modalità Anteprima 3D
'è aperto oppure chiuso
Dim MenùAiuto As Boolean
'Definisco una variabile che conterrà le coordinate del cursore di windows
Dim tmpWindowsMousePosition As POINTAPI
'Definisco la funzione che mi servirà per rilevare la posizione del cursore di windows
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Sub AvviaAnteprima_Click()
    'Le istruzioni che seguono servono ad avvisare l'utente che la scalatura selezionata non è una delle migliori
    'al fine di una corretta visualizzazione della mappa 3D in modalità Anteprima 3D,tuttavia egli potrà scegliere
    'se continuare con l'operazione di Conversione oppure annullare e quindi selezionare una scala migliore
    If VScale <= 9 Then
        Dim RispostaContinuaAnteprima As VbMsgBoxResult
        If LinguaS = "Italiano" Then RispostaContinuaAnteprima = MsgBox("Per una buona visualizzazione della mappa 3d in modalità Anteprima 3D e consigliabile selezionare una scalatura maggiore o uguale a 10!" + Chr(13) + " Vuoi comunque avviare la modalità Anteprima 3D in scala 1 :" + Str(VScale) + "? ", vbYesNo, "Avvio modalità Anteprima 3D")
        If LinguaS = "Inglese" Then RispostaContinuaAnteprima = MsgBox("For a better view of 3d map in 3D preview modality you must choose a scale value major or equal then 10!" + Chr(13) + "Do you want aniway start 3D preview modality with 1 :" + Str(VScale) + " scale? ", vbYesNo, "3D preview modality start")
        If RispostaContinuaAnteprima = vbNo Then GoTo Annulla
    End If
    'Questa condizione fa in modo che se è già stato visualizzato il messaggio dell'errore dello scale, non verrà visualizzato
    'l'altra di conferma di entrata in modalità Anteprima 3D
    If RispostaContinuaAnteprima <> vbYes Then
        'Dichiaro una variabile che mi servirà per capire se cerrà deciso di proseguire con
        'l'avvio della modalità anteprima 3D
        Dim RispostaModalitàAnteprima3D As VbMsgBoxResult
        'Assegno alla variabile sopra dichiarata il valore del tasto premuto in corrispondenza
        'del messaggio visualizzato
        If LinguaS = "Italiano" Then RispostaModalitàAnteprima3D = MsgBox("Attenzione! Si stà passando alla modalità Anteprima 3D!" + Chr(13) + " vuoi proseguire?", vbYesNo, "Avvio della modalità Anteprima 3D")
        If LinguaS = "Inglese" Then RispostaModalitàAnteprima3D = MsgBox("Attention! The 3D preview modality will be started!" + Chr(13) + "Do you want continue?", vbYesNo, "3D preview modality start")
    End If
    'Richiamo dall'oggetto Mappa3D il metodo che mi permetterà di costruire una vera
    'e propria mappa3d in base alle righe costruite sullo schermo
    Mappa3D.Crea_Mappa_3D
    'Se si ha il consenso da parte dell'utente allora si procederà con le operazioni
    'di inizializzazione
    If RispostaModalitàAnteprima3D = vbYes Or RispostaContinuaAnteprima = vbYes Then
        'Riproduco un file sonoro riservato all'avvio della modalità Anteprima 3D
        Suoni.Esegui_suono_AvviaAnteprima
        'Impongo l avariabile continua con un valore false,in modo che si uscirà temporaneamete dall'editor
        'e si avvierà il ciclo che mi permetterà di gironzolare tranquillamente all'interno
        'della mappa appena creata
        Continua = False
        'Attivo la voce di menà StoppaAnteprima,in modo da permettere all'utente di tornare
        'quando vuole a visualizzare l'editor vero e proprio
        StoppaAnteprima.Enabled = True
        'Inizializzo la variabile Continua_Anteprima_3D = True,in modo che si continuerà a ciclare
        'finchè questa non assumerà il valore opposto = False
        Continua_Anteprima_3d = True
        'Disabilitò la voce di menù Avvia Anteprima in modo che non sia possibile selezionarla
        'quando ci si trova in modalità Anteprima 3D
        AvviaAnteprima.Enabled = False
        'Avvio il ciclo di anteprima della mappa 3D
        CicloAnteprima3D
        'Quando si selezionerà la voce di menù StoppaAnteprima, si uscirà dal ciclo di anteprima 3D
        'per ritornare all'editor vero e proprio.
        'Inoltre reinizializzo la variabile continua con un valore booleano = true
        Continua = True
        'Si torna al ciclo dell'editor vero e proprio
        Ciclo
    End If
'Label Annulla
Annulla:
End Sub

Sub CicloAnteprima3D()
    '---------------------------------------------------------------------------------------------
    ' Le suguenti variabili mi servono per ricavare tutti gli imput forniti dal mouse in modalità
    ' anteprima 3D al fine di poter apportare modifiche in tempo reale agli oggetti selezionati
    '---------------------------------------------------------------------------------------------
    'Dichiaro una variabile di appoggio che mi servirà per salvare il valore temporaneo
    'del cursore del mouse sull'asse delle X
    Dim TmpX As Long
    'Dichiaro una variabile di appoggio che mi servirà per salvare il valore temporaneo
    'del cursore del mouse sull'asse delle Y
    Dim TmpY As Long
    'Dichiaro una variabile che mi servirà per verificare se è stato premuto il tasto
    'sinistro del mouse
    Dim TmpB1 As Integer
    'Dichiaro una variabile che mi servirà per verificare se è stato premuto il tasto
    'destro del mouse
    Dim TmpB2 As Integer
    'Effettuo un'effetto di sfumatura in entrata in modo da rendere un pò più particolare
    'il programma
    Effetti.FadeIn
    'Disattivo l'eventuale immagine di sfondo dell'editor che si aveva selezionato
    Schermo.EnableBackground False
    'Cambio il colore di sfondo da grigio a nero
    Scena.SetSceneBackGround 0, 0, 0
    'Richiama dall'oggetto Telecamera il metodo che inizializza tutte le variabili necessarie
    'al corretto funzionamento della modalità Anteprima 3D
    Telecamera.Inizializzazione_Variabili_Di_Comando
    'Avvio un ciclo for che richiamerà da ogni oggetto creato il metodo Attiva_Oggetto,addetto
    'al corretto posizionamento dello stesso all'interno della mappa 3D
    For IOg = 0 To IOg
        'Se l'oggetto attualmente analizzato non possiede un valore salvato all'interno del
        'proprio campo chiave,allora questo verrà attivato.
        If Oggetto(IOg).Key <> "" Then
            'Chiamata al metodo Attiva_Oggetto
            Oggetto(IOg).Attiva_Oggetto
        End If
    'Si passa all'oggetto successivo
    Next
    'Da qui inizia il vero e proprio CicloAnteprima3d,ovvero quel ciclo che permetterà
    'all'utente,grazie all'aiuto delle classi ClsMappa3D,addetta alla conversione della mappa
    'attuale in 3D,e alla classe ClsTele,che permette il movimento all'interno della stessa,
    'di osservare la mappa appena costruita in modo assolutamente 3D,con tanto di movimento
    'e cambio di direzione di visuale al suo interno
    Do
        DoEvents
        'Cancello il contenuto dell'oggetto TV8
        TV8.Clear
        'Richiamo la funzione che allinea i form in caso questi venissero spostati
        Allinea_form
        'Richiamo la funzione che mi permetterà di far coincidere le coordinate del mouse
        'di windows con quelle dell'editor
        Setta_cursore
        'Richiamo dall'oggetto Telecamera,il quale è stato definito di tipo ClsTele,il metodo
        'che mi permetterà di verificare quale movimento dovrà compiere la Telecamera.
        'Questo avverrà mediante la pressione dei tasti sulla tastiera
        Telecamera.Controlla_Comandi
        'Richiamo sempre dallo stesso oggetto,il metodo che mi permetterà,in base ai tasti
        'premuti sulla tastiera,di muovermi all'interno della mappa3D appena costruita
        Telecamera.Aggiorna_Comandi
        'Renderizzo la mappa3D creata in modo che questa possa essere visibile
        Scena.RenderAllMeshes
        'Verifico se è stato premuto un pulsante del mouse,è nel caso in cui questo fosse accaduto
        'richiamo la funzione Verifica_Selezione_Oggetto,il cui scopo è quello di verificare
        'se è stato selezionato un oggetto
        Comandi.GetAbsMouseState TmpX, TmpY, TmpB1, TmpB2
        'Se è stato premuto il tasto sinistro del mouse richiamerò la funzione Verifica_Selezione_Oggetto
        If TmpB1 <> 0 Then Verifica_Selezione_Oggetto TmpX, TmpY
        'Se la variabile booleano MenùAiuto possiede il valore False,allora...
        If MenùAiuto = False Then
            'Richiamo la funzione che disegnerà semplicemente su schermo un piccolo box che informerà
            'l'utente che alla pressione del tasto H si attiverà appunto il menù di aiuto che guiderà
            'l'utente nelle operazioni di modifica dell'oggetto selezionato
            Disegna_Box_Modifica_Oggetto
        'Altrimenti...
        Else
            'Richiamo la funzione che disegnerà su schermo il menù che guiderà l'utente nelle operazioni
            'di modifica dell'oggetto selezionato
            Disegna_Menù_Modifica_Oggetto
            'Richiamo la funzione che permetterà all'utente,a seconda del tasto premuto sulla tastiera,
            'di apportere modifiche in tempo reale all'oggetto selezionato
            Ricevi_Modifiche_Oggetto
        End If
        'Renderizzo tutto il contenuto dell'oggetto TV8 su schermo
        TV8.RenderToScreen
    'Continuo a ciclare finchè la variabile Continua_Anteprima_3D non sarà = False
    Loop Until Continua_Anteprima_3d = False
    'Effettuo un'effetto di dissolvenza in uscita
    Effetti.FadeOut
    'Se è stata precedentemente selezionata un'immagine da applicare come fondale dell'editor
    'allora...
    If Form_Opzioni.FondaleStatico = True And ImmagineSfondo <> "Nessuna" Then
        'Ativo il fondale dell'editor
        Schermo.EnableBackground True
        'Carico come immagine di sfondo dell'editor,l'immagine che si aveva precedentemente
        'selezionato
        Schermo.LoadBackground ImmagineSfondo
    End If
    'Elimino tutte le Texture degli oggetti caricati all'interno della mappa 3D dall'oggetto
    'FabbricaTexture
    FabbricaTexture.DeleteAll
    'Impongo alla variabile MenùAiuto il valore boleano False in modo che la prossima volta
    'che si entrerà in modalità Antepima 3D venga disegnato solamente il Box che segnala che
    'premendo il tasto P si attiverà il menù di modifica dell'oggetto
    MenùAiuto = False
End Sub

Private Sub Carica_mappa_Click()
    'Avvio la funzione che mi permetterà di caricare una mappa precedentemente
    'creata e salvata
    Carica_mappa_salvata
End Sub


Private Sub Esci_Click()
    'Dichiaro una variabile che servirà al programma per capire quale bottone del messaggio viene
    'premuto dall'utente
    Dim Risposta As VbMsgBoxResult
    'Viene visualizzato un messaggio che chiede all'utente se si vuole uscire dal programma
    If LinguaS = "Italiano" Then Risposta = MsgBox("Sei sicuro di voler abbandonare il programma?", vbOKCancel, "Uscita dal programma")
    If LinguaS = "Inglese" Then Risposta = MsgBox("Do you really want exit the program?", vbOKCancel, "Exit program")
    'Se viene premuto il bottone OK...
    If Risposta = vbOK Then
        'Richiamo il metodo dalla Classe Sonora che riprodurra il file sonoro di uscita
        'dal programma
        Suoni.Esegui_suono_uscita
        'Metto in pausa l'oggetto TV8 in modo che aspetterà che la riproduzione del file
        'sonoro sia terminata prima di eseguire le istruzioni di uscita dal programma
        TV8.Pause 2000
        '...Pongo la variabile Continua e Continua_Animazione a stato false permettendomi così di uscire dal
        'ciclo principale o se in esecuzione da quello di animazione e di conseguenza anche dal programma
        Continua_Animazione = False
        Continua = False
        'Nel caso in cui si fosse in modalità Anteprima 3D il programma verrebbe terminato
        'direttamente
        If Continua_Anteprima_3d = True Then
            End
        End If
    End If
End Sub

Private Sub File_Click()
    'Alla pressione della voce di menù File richiamerò il metodo dalla classe sonora
    'che riprodurrà il suono di menù
    Suoni.Esegui_suono_menù
End Sub

Private Sub Form_Load()
    'Dichiaro una variabile booleana che mi servirà per capire se la risoluzione di schermo del computer su cui è stato avviato
    'il programma è errata (inferiore a 1024 X 768).
    'Il suo valore verrà assegnato dalla funzione Problemi_Risoluzione è se questa restituirà il valore False, allora
    'l'utente verrà informato della scorretta risoluzione di video è alla pressione del tasto OK, lo stesso programma
    'verrà terminato, altrimenti l'esecuzione procederà normalmente
    Dim Problemi_risoluzione_schermo As Boolean
    'Assegnazione del valore della variabile Problemi_risoluzione_schermo tramite la funzione Problemi_Risoluzione
    Problemi_risoluzione_schermo = (Problemi_Risoluzione)
    'Se la variabile Problemi_risoluzione_video avrà assunto il valore True dopo l'assegnazione, allora verrà visualizzato il messaggio
    'di errore
    If Problemi_risoluzione_schermo = True Then
        'Creazione del messaggio di erore dell'incorretta risoluzione video
        If LinguaS = "Italiano" Then MsgBox "L'attuale risoluzione di schermo è " + Str(Screen.Width / Screen.TwipsPerPixelX) + " X " + Str(Screen.Height / Screen.TwipsPerPixelY) + Chr(13) + ". L'esecuzione del programma invece ridchiede una risoluzione video non inferiore a 1024 X 768. " + Chr(13) + "Risetta le impostazioni dello schermo e riavvia il programma!", vbOKOnly, "Risoluzione video incorretta!"
        If LinguaS = "Inglese" Then MsgBox "The actual screen risolution is " + Str(Screen.Width / Screen.TwipsPerPixelX) + " X " + Str(Screen.Height / Screen.TwipsPerPixelY) + ". For Start the program you must set the screen resolution at 1024 X 768! " + Chr(13) + "Set the video and before start the program!", vbOKOnly, "Incorrect screen resolution"
        'Il programma verrà interrotto
        End
    End If
    'Richiamo la funzione per l'inizializzazione del motore 3D
    Inizializza_3d
    'Inizializzo con un valore iniziale tutte le variabili necessarie
    Inizializza_Variabili
    'Visualizza il form delle opzioni
    Form_Opzioni.Show
    'Mostro il form principale.Questa istruzione può risultare stupida,ma è assolutamente necessaria
    'per avviare la grafica
    Me.Show
    'Richiamo il metodo dalla Classe Sonora che riprodurrà il file di Avvio del programma
    Suoni.Esegui_suono_avvio
    'Avvia il ciclo dell'animazione
    Ciclo_Animazione
    'Richiama il ciclo precedentemente descritto
    Ciclo
    'Questa funzione serve per distruggere tutti gli oggetti che mi sono creato
    'per il funzionamento del programma
    Distruggi_oggetti
    'Esce dal programma
    End
End Sub

Sub Inizializza_3d()
    'Indico la directory iniziale da cui partire a ricercare tutti i file annessi al programma
    TV8.SetSearchDirectory App.Path
    'Avviamo il 3D in una finestra,e più precisamente all'interno di una picturebox
    'chiamata Editor
    TV8.Init3DWindowedMode Editor.hWnd
    'Creo dei tipi di caratteri personalizzati che mi serviranno per disegnare il menù informazioni sullo schermo
    Schermo.CreateUserFont "Carattere_personalizzato1", "Comic sans Ms", 10, True, False, False
    Schermo.CreateUserFont "Carattere_personalizzato2", "Arial", 8, True, False, False
    'Richiama il metodo dall'oggetto animazione che servirà ad evviare la sequenza
    'introduttiva del programma
    Animazione.Start
    'Richiamo dall'oggetto Suoni il metodo che caricherà all'interno di se stesso il
    'file sonoro da riprodurre quando ci si muoverà all'interno della mappa.
    'Questo farà sembrare il programma più realistico,simulando il rumore dei passi
    Suoni.Inizializza_Suoni
    'Carico all'interno della scena un cursore da me costruito
    Scena.LoadCursor "Images/pointer.bmp", TV_COLORKEY_BLACK, 16, 16
    'Attivo il MipMapping.
    'Il MipMapping è una particolare funzione 3D che mi permetterà di migliorare la grafica nella modalità
    'Anteprima3D
    Scena.EnableMipMapping True
End Sub

Sub Inizializza_Variabili()
    'Dichiaro un indice che mi tornerà utile nell'assegnazione ai quattro angoli di ogni
    'muro,pavimento e soffitto,dello stesso indice di solidità
    Dim K As Integer
    'Inizializziamo la variabile continua con un valore iniziale True
    Continua = True
    'Inizializziamo la variabile continua_animazione con un valore iniziale True
    Continua_Animazione = True
    'Dichiaro l'indice I con valore = 0
    I = 0
    'e lo stesso faccio anche per l'indice J
    J = 0
    'Dichiaro il numero massimo di linee = 0
    Max = 0
    'La somma dei soffitti + i pavimenti presenti all'interno della mappa attuale
    Max2 = 0
    'Il numero massimo di pavimenti
    Max3 = 0
    'E il numero massimo di soffitti
    Max4 = 0
    'Imposto lo stato iniziale della prima riga che verrà costruita
    Stato_riga = 0
    'Imposto lo stato iniziale del primo soffitto o pavimento che verrà creato
    Stato_sop = 0
    'Assegno alla variabile Larghezza il valore espresso in pixel della PictureBox in cui verrà
    'inizializzato il 3D,ovvero la PictureBox Editor
    Larghezza = Editor.ScaleWidth
    'Faccio la stessa cosa anche per la variabile Altezza,solo che questa volta gli assegnerò
    'intuitivamente il valore espresso in pixel dell'altezza della stessa PictureBox
    Altezza = Editor.ScaleHeight
    'Imposto un valore di partenza allo scale
    VScale = 1
    'Imposto la variabile ImmagineSfondo con il valore alfanumerico "Nessuna".
    'Questo farà in modo che all'avvio del programma non venga caricata appunto nessuna
    'immagine da applicare al fondale dell'editor
    ImmagineSfondo = "Nessuna"
    'Imposto dei valori di default a tutte e 10000 le righe dichiarate.Sarà compito dell'utente
    'modificare in seguito a suo piacimento questi valori per ogni muro
    For I = 0 To 10000
        With Riga(I)
            'Inizializzo l'altezza base del muro con un valore pari a 1000
            .Altezza = 1000
            'Inizializzo il numero di mattonelle base disposte in altezza sulla superfice
            'del muro corrente
            .NMattonelleALtezza = 10
            'Inizializzo il numero di mattonelle base disposte in larghezza sulla superfice
            'del muro corrente
            .NMAttonelleLarghezza = 10
            'Imposto a "Nessuna", l'immagine iniziale che dovrà essere disposta tante volte
            'quanto il numero di mattonelle,sulla superfice del muro corrente
            .Texture = "Nessuna"
            'Imposto la proprietà iniziale che dovrà adottare il muro corrente,ovvero dovrà
            'essere di tipo Normale
            .Proprietà = "Normale"
            'Assegno al muro corresnte il proprio nome iniziale,formato dalla parola
            '"Muro" più l'indice della sua posizione all'interno della tabella Coordinate_Riga
            .Nome = "Muro " + Str(I)
            'Imposto i quattro angoli del muro con un grado di solidità iniziale
            'pari a 0.5
            For K = 0 To 3
                With .ColVertici(K)
                    .A = 0.5
                End With
            'Passo ad esaminare il muro successivo
            Next
        End With
        'Richiamo la funzione pubblica addetta alla reimpostazione del materiale specifico
        'passato alla funzione stessa con il colore nullo (bianco)
        Reimposta_materiale Riga(I).Materiale
    Next
    'Come per le righe, imposto anche per tutti e 10000 SoP (Soffitti o Pavimenti) dei valori
    'di default.Anche qui sarà l'utente che potrà reimpostare questi valori a suo piacimento
    For J = 0 To 10000
        With SoP(J).CR
            'Inizializzo il numero di mattonelle base disposte in altezza sulla superfice
            'del pavimento / soffitto corrente
            .NMattonelleALtezza = 20
            'Inizializzo il numero di mattonelle base disposte in larghezza sulla superfice
            'del pavimento / soffitto corrente
            .NMAttonelleLarghezza = 20
            'Imposto a "Nessuna", l'immagine iniziale che dovrà essere disposta tante volte
            'quanto il numero di mattonelle,sulla superfice del pavimento / soffitto corrente
            .Texture = "Nessuna"
            'Imposto la proprietà iniziale che dovrà adottare il pavimento / soffitto corrente,
            'ovvero dovrà essere di tipo Normale
            .Proprietà = "Normale"
            'Imposto i quattro angoli del pavimento / soffitto con un grado di solidità iniziale
            'pari a 0.5
            For K = 0 To 3
                With .ColVertici(K)
                    .A = 0.5
                End With
            'Passo ad analizzare il pavimento / soffitto successivo
            Next
        End With
        'Richiamo la funzione pubblica addetta alla reimpostazione del materiale specifico
        'passato alla funzione stessa con il colore nullo (bianco)
        Reimposta_materiale SoP(J).CR.Materiale
    Next
    'Inizializzo le coordinate iniziali della telecamera
    With PosizioneTelecamera
        .X = 100
        .Y = 500
        .Z = 100
    End With
    'Inizializzo la variabile Molt con un valore pari a 1,in modo da far capire al programma che le linee
    'non sono state modificate dall'operazione di zoom
    Molt = 1
    'Inizializzo la variabile VCambiamentiGriglia con il valore 20,in modo che se non sono state ancora effettuate
    'operazioni di zoom,e il controllo GrigliaControllataDaZoom è attivato,i quadrati della griglia avranno tutti dimensioni
    '60 * 60
    VCambiamentiGriglia = 60
    'Richiamo la funzione addetta all'assegnazione iniziale dei colori dei vari componenti utilizzati dall'editor,
    'ovvero il colore delle linee che rappresentano i muri,il colore delle linee che rappresentano i pavimenti,
    'il colore delle linee che rappresentano i soffitti,ecc.
    Inizializza_colori
    'Inizializzo la variabile pubblica LinguaS ad "Italiano".
    'Il programma interpreterà questa istruzione come lingua in cui avviare il programma
    LinguaS = "Inglese"
    'Richiamo la funzione che mi permetterà di tradurre il programma nella lingua desiderata
    Traduci "Inglese"
End Sub

Sub Ciclo_Animazione()
    'Da qui comincia il Ciclo_Animazione, ovvero quel ciclo che mi permetterà di ricreare,grazie all'ausilio
    'della classe ClsIntroduzione,una piccola animazione che rappresenterà appunto la presentazione
    'del programma
    Do
        DoEvents
        'Richiama la funzione che allinea tutti i form
        Allinea_form
        'Richiamo la funzione che mi permetterà di allineare il cursore di window a quello
        'dell'editor
        Setta_cursore
        'Ripulisce il contenuto dell'oggetto TV8
        TV8.Clear
        'Richiama dall'interno della classe Animazione la funzione che ruoterà gli anelli che
        'compongono la sequenza introduttiva del programma
        Animazione.Ruota_anelli
        'Renderizza tutti il contenuto della scena corrente
        Scena.RenderAllMeshes
        'Questa istuzione mi servirà per renderizzare il terreno presente all'interno della classe ClsIntroduzione,
        'affinchè sia visibile l'effetto neve
        Animazione.Terreno.Render True
        'Renderizza tutto il contenuto dell'oggetto TV8 su schermo
        TV8.RenderToScreen
    'Tutto questo avviene finchè la variabile continua_animazione non assumerà il valore da false
    Loop Until Continua_Animazione = False
    'Richiama dall'interno della classe animazione,la funzione che distruggera tutti gli oggetti
    'presenti all'interno dell'animazione.Questo al fine di non lasciarne residui in memoria centrale
    Animazione.Distruggi_Animazione
    'L'istruzione che segue mi permetterà di avere una portata di angolo di visuale pari a PI * 20, e
    'una di distanza pari a 10000.
    'In poche parole mediante quaesta semplice funzione, è possibile vedere tutti quei muri,pavimenti,soffitti o
    'oggetti che sono posti ad una distanza massima di 10000 dalla nostra posizione
    Scena.SetViewFrustum 3.14159265359 * 20, 10000
End Sub
Sub Ciclo()
    'Effettuo un effetto di sfumatura in entrata
    Effetti.FadeIn 1000
    'Imposto un colore di sfondo alla scena e quindi anche all'editor
    Scena.SetSceneBackGround 0.5, 0.5, 0.5
    Do
        DoEvents
        'Richiama la funzione che allinea tutti i form
        Allinea_form
        'Ripulisce il contenuto dell'oggetto TV8
        TV8.Clear
        'Richiama la funzione che,se selezionata l'opzione di visualizzazione griglia
        'dal menù opzioni,disegnerà la griglia sullo schermo
        Controlla_griglia
        'Chiama la funzione che disegnerà le righe crete sullo schermo
        Disegna_righe
        'Richiama la funzione che verifica lo stato della tastiera e del mouse
        Controlla_input
        'Chiama la funzione che disegnerà una sorta di menù sullo schermo
        Crea_menù
        'Renderizza tutto il contenuto dell'oggetto TV8 su schermo
        TV8.RenderToScreen
    'Tutto questo avviene finchè la variabile continua non assumerà il valore False
    Loop Until Continua = False
End Sub

Sub Controlla_input()
    Dim Scroll As Boolean
    'Richiamo la funzione che mi permetterà di far coincidere le coordinate del mouse di
    'windows con quello dell'editor
    Setta_cursore
    'Se viene premuto sulla tastiera il pulsante S allora imporrò alla variabile Scroll
    'il valore True
    If Comandi.IsKeyPressed(TV_KEY_S) = True Then
        Scroll = True
    'Altrimenti questa avrà valore False
    Else
        Scroll = False
    End If
    'Se viene premuto il tasto Canc, quando è in costruzione un muro, un pavimento o un soffitto,
    'la costruzione dello stesso verrà annullata,in modo tale che si possa iniziare a crearne un altro
    'con nuove coordinate
    If Comandi.IsKeyPressed(TV_KEY_DELETE) = True Then
        'Se la variabile Stato_riga è diversa da 0, allora la reimposto in modo tale che si possa reiniziare
        'a settare le coordinate iniziali del nuovo muro
        If Stato_riga <> 0 Then Stato_riga = 0
        'Nel caso in cui questa volta la variabile Stato_Sop avesse valore diverso da 0,la si reimposterebbe
        'per risettare appunto le coordinate iniziali del nouvo pavimento o soffitto
        If Stato_sop <> 0 Then Stato_sop = 0
    End If
    'Se la variabile Scroll a il valore True,allora richiamerò la funzione che mi servirà per simulare
    'uno scrolling della mappa,ridimensionando tutte le righe presenti.
    'Notare che alla funzione gli passo due valori che sono rispettivamente i valori del mouse
    'al momento della pressione del tasto S
    If Scroll = True Then Scorri_Mappa MouseX, MouseY
    'Verifico se la variabile Scroll a valore False,lo stato attuale del mouse,cioè le sue coordinate e se viene premuto
    'o il pulsante sinistro (B1) o quello destro (B2)
    If Scroll = False Then
        Comandi.GetAbsMouseState MouseX, MouseY, B1, B2
        'Duplico le coordinate correnti del mouse da Long a Single
        SmouseX = CSng(MouseX)
        SmouseY = CSng(MouseY)
    End If
    'Queste istruzione mi permettono,se è stato selezionato dall'utente nel form delle opzioni,
    'di poter visualizzare delle linee che fungeranno da ulteriore aiuto per l'allineamento
    If Form_Opzioni.Linee_guida.Value = 1 Then
        'Disegno la linea orizzontale
        Schermo.DrawLine 0, SmouseY, Larghezza - 1, SmouseY, RGBA(CLG.R, CLG.G, CLG.B, 1)
        'Disegno la linea verticale
        Schermo.DrawLine SmouseX, 0, SmouseX, Altezza - 1, RGBA(CLG.R, CLG.G, CLG.B, 1)
    End If
    'Questo ciclo For è molto importante,ora vi spiegherò im perchè:
    'Pensate di dover impostare le coordinate di una nuova linea perfettamente allineate
    'a quelle di un'altra linea,che però è molto distante dall'altra.Come fare?
    'I metodi sono due.O avere un occhio da "lince" e riuscire a scrutare perfettamente l'allineatura:
    'cosa molto improbabile; oppure grazie al metodo sotto descritto,disegnare su schermo un piccolo rettangolo
    'blu,sia sulle coordinate finale della nuova linea che verrà creta,sia sulle coordinate di tutte le
    'linee a essa allineate,in modo da poter capire se le nuova linea e appunto,immaginariamente collegata.
    'ATTENZIONE: Questo avverrà solamente se la checkbox Rileva_Allineamento è selezionata.
    If Form_Opzioni.Rileva_allineamento.Value = 1 Then
        For I = 0 To Max
            If SmouseX = Riga(I).X1 Or SmouseY = Riga(I).Y1 Then
                Schermo.DrawFilledBox SmouseX - 3, SmouseY - 3, SmouseX + 3, SmouseY + 3, RGBA(CAM.R, CAM.G, CAM.B, 1)
                Schermo.DrawFilledBox Riga(I).X1 - 3, Riga(I).Y1 - 3, Riga(I).X1 + 3, Riga(I).Y1 + 3, RGBA(CAM.R, CAM.G, CAM.B, 1)
            End If
            If SmouseX = Riga(I).X2 Or SmouseY = Riga(I).Y2 Then
                Schermo.DrawFilledBox SmouseX - 3, SmouseY - 3, SmouseX + 3, SmouseY + 3, RGBA(CAM.R, CAM.G, CAM.B, 1)
                Schermo.DrawFilledBox Riga(I).X2 - 3, Riga(I).Y2 - 3, Riga(I).X2 + 3, Riga(I).Y2 + 3, RGBA(CAM.R, CAM.G, CAM.B, 1)
            End If
            If SmouseY = Riga(I).X1 Or SmouseY = Riga(I).Y1 Then
                Schermo.DrawFilledBox SmouseX - 3, SmouseY - 3, SmouseX + 3, SmouseY + 3, RGBA(CAM.R, CAM.G, CAM.B, 1)
                Schermo.DrawFilledBox Riga(I).X1 - 3, Riga(I).Y1 - 3, Riga(I).X1 + 3, Riga(I).Y1 + 3, RGBA(CAM.R, CAM.G, CAM.B, 1)
            End If
            If SmouseY = Riga(I).X2 Or SmouseY = Riga(I).Y2 Then
                Schermo.DrawFilledBox SmouseX - 3, SmouseY - 3, SmouseX + 3, SmouseY + 3, RGBA(CAM.R, CAM.G, CAM.B, 1)
                Schermo.DrawFilledBox Riga(I).X2 - 3, Riga(I).Y2 - 3, Riga(I).X2 + 3, Riga(I).Y2 + 3, RGBA(CAM.R, CAM.G, CAM.B, 1)
            End If
        Next
    End If
    'Invece, la funzione di quest'altro ciclo è la medesimo di quella sopra indicata,con la sola differenza che questa volta
    'l'allineamento verrà visualizzato di colore Verde e in corrispondenza delle coordinate di tutti i
    'pavimenti e soffitti creati.
    'Anche questa volta tutto questo si verificherà solamente se la checkbox Rileva_Allineamento2 presente
    'all'interno del Form_Opzioni, verrà selezionata
    If Form_Opzioni.Rileva_Allineamento2.Value = 1 Then
        For J = 0 To Max2
            If SmouseX = SoP(J).CR.X1 Or SmouseY = SoP(J).CR.Y1 Then
                Schermo.DrawFilledBox SmouseX - 3, SmouseY - 3, SmouseX + 3, SmouseY + 3, RGBA(CASOP.R, CASOP.G, CASOP.B, 1)
                Schermo.DrawFilledBox SoP(J).CR.X1 - 3, SoP(J).CR.Y1 - 3, SoP(J).CR.X1 + 3, SoP(J).CR.Y1 + 3, RGBA(CASOP.R, CASOP.G, CASOP.B, 1)
            End If
            If SmouseX = SoP(J).CR.X2 Or SmouseY = SoP(J).CR.Y2 Then
                Schermo.DrawFilledBox SmouseX - 3, SmouseY - 3, SmouseX + 3, SmouseY + 3, RGBA(CASOP.R, CASOP.G, CASOP.B, 1)
                Schermo.DrawFilledBox SoP(J).CR.X2 - 3, SoP(J).CR.Y2 - 3, SoP(J).CR.X2 + 3, SoP(J).CR.Y2 + 3, RGBA(CASOP.R, CASOP.G, CASOP.B, 1)
            End If
            If SmouseX = SoP(J).X3 Or SmouseY = SoP(J).Y3 Then
                Schermo.DrawFilledBox SmouseX - 3, SmouseY - 3, SmouseX + 3, SmouseY + 3, RGBA(CASOP.R, CASOP.G, CASOP.B, 1)
                Schermo.DrawFilledBox SoP(J).X3 - 3, SoP(J).Y3 - 3, SoP(J).X3 + 3, SoP(J).Y3 + 3, RGBA(CASOP.R, CASOP.G, CASOP.B, 1)
            End If
            If SmouseX = SoP(J).X4 Or SmouseY = SoP(J).Y4 Then
                Schermo.DrawFilledBox SmouseX - 3, SmouseY - 3, SmouseX + 3, SmouseY + 3, RGBA(CASOP.R, CASOP.G, CASOP.B, 1)
                Schermo.DrawFilledBox SoP(J).X4 - 3, SoP(J).Y4 - 3, SoP(J).X4 + 3, SoP(J).Y4 + 3, RGBA(CASOP.R, CASOP.G, CASOP.B, 1)
            End If
            If SmouseY = SoP(J).CR.X1 Or SmouseY = SoP(J).CR.Y1 Then
                Schermo.DrawFilledBox SmouseX - 3, SmouseY - 3, SmouseX + 3, SmouseY + 3, RGBA(CASOP.R, CASOP.G, CASOP.B, 1)
                Schermo.DrawFilledBox SoP(J).CR.X1 - 3, SoP(J).CR.Y1 - 3, SoP(J).CR.X1 + 3, SoP(J).CR.Y1 + 3, RGBA(CASOP.R, CASOP.G, CASOP.B, 1)
            End If
            If SmouseY = SoP(J).CR.X2 Or SmouseY = SoP(J).CR.Y2 Then
                Schermo.DrawFilledBox SmouseX - 3, SmouseY - 3, SmouseX + 3, SmouseY + 3, RGBA(CASOP.R, CASOP.G, CASOP.B, 1)
                Schermo.DrawFilledBox SoP(J).CR.X2 - 3, SoP(J).CR.Y2 - 3, SoP(J).CR.X2 + 3, SoP(J).CR.Y2 + 3, RGBA(CASOP.R, CASOP.G, CASOP.B, 1)
            End If
            If SmouseY = SoP(J).X3 Or SmouseY = SoP(J).Y3 Then
                Schermo.DrawFilledBox SmouseX - 3, SmouseY - 3, SmouseX + 3, SmouseY + 3, RGBA(CASOP.R, CASOP.G, CASOP.B, 1)
                Schermo.DrawFilledBox SoP(J).X3 - 3, SoP(J).Y3 - 3, SoP(J).X3 + 3, SoP(J).Y3 + 3, RGBA(CASOP.R, CASOP.G, CASOP.B, 1)
            End If
            If SmouseY = SoP(J).X4 Or SmouseY = SoP(J).Y4 Then
                Schermo.DrawFilledBox SmouseX - 3, SmouseY - 3, SmouseX + 3, SmouseY + 3, RGBA(CASOP.R, CASOP.G, CASOP.B, 1)
                Schermo.DrawFilledBox SoP(J).X4 - 3, SoP(J).Y4 - 3, SoP(J).X4 + 3, SoP(J).Y4 + 3, RGBA(CASOP.R, CASOP.G, CASOP.B, 1)
            End If
        Next
    End If
    'Salvo il valore degli indici correnti su due variabili di appoggio
    AppI = I
    AppJ = J
    'Se viene premuto il tasto sinistro del mouse e la variabile stato_riga=1 allora questo indicherà che la
    'costruzione della riga attuale è complettata,quindi assegno le coordinate finali di questa,rispettivamente
    'alla posizione attuale del mouse (valori X e Y),incremento il contatore per la costruzione di una nuova
    'riga e aumento anche il valore della variabile Max che,ricordo,tiene il numero di righe già costruite,
    'riporto la varibile stato_riga a valore 0,e infine,cosa più importante di tutte,eseguo un controllo
    'IMPORTANTISSIMO che ora spiegherò...
    'Se le coordinate di riga fossero impostate normalmente dove si trova il mouse,queste,
    'se si trovassero a coincidere o all'inizio o alla fine,queste,non sarebbero perfettamente
    'coincidenti,perchè ci sarebbe sempre un minimo errore,invece, adottando il metodo sotto
    'descritto si può fare in modo che le coordinate delle nuova riga siano impostate in
    '"Modo intelligente",assegnando le coordinate iniziali della nuova riga uguali
    'alle coordinate della riga più vicina.Questo fa in modo che quando si dovranno creare
    'i muri,si formino degli amgoli nella nostra mappa...
    'Tutto questo avverrà solamente nel caso in cui l'utente premerà il pulsante Muri dal form
    'delle opzioni
    If Scelta_Oggetto = "Muro" Then
        If B1 <> 0 And Stato_riga = 1 And SmouseX > 0 And SmouseX < Larghezza And SmouseY > 50 And SmouseY < Altezza Then
            Riga(I).X2 = SmouseX
            Riga(I).Y2 = SmouseY
            For I = 0 To Max
                TmpX1 = Riga(I).X1
                TmpX2 = Riga(I).X2
                TmpY1 = Riga(I).Y1
                TmpY2 = Riga(I).Y2
                If Riga(AppI).X2 > TmpX1 - 5 And Riga(AppI).X2 < TmpX1 + 5 And Riga(AppI).Y2 > TmpY1 - 5 And Riga(AppI).Y2 < TmpY1 + 5 Then
                    Riga(AppI).X2 = TmpX1
                    Riga(AppI).Y2 = TmpY1
                    Riga(AppI).SpigoloF = True
                    Riga(I).SpigoloI = True
                ElseIf Riga(AppI).X2 > TmpX2 - 5 And Riga(AppI).X2 < TmpX2 + 5 And Riga(AppI).Y2 > TmpY2 - 5 And Riga(AppI).Y2 < TmpY2 + 5 Then
                    Riga(AppI).X2 = TmpX2
                    Riga(AppI).Y2 = TmpY2
                    Riga(AppI).SpigoloF = True
                    Riga(I).SpigoloF = True
                End If
            Next
            'Ora aggiungo il muro appena creto all'oggetto Elenco_Muri presente all'interno del Form_Opzioni,
            'in modo tale che questo possa essere selezionato in seguito per effettuarvi tutte le eventuali
            'modifiche che l'utente riterrà opportune
            Form_Opzioni.ElencoMuri.AddItem Riga(I).Nome
            'Incemento la variabile I,in modo da passare al muro successivo di quello appena creato
            I = I + 1
            'Incremente la variabile Max,poichè il numero di muri preenti all'interno della mappa attuale
            'è aumentato di 1
            Max = Max + 1
            'Riporto la variabile Stato_Riga alla sua impostazione iniziale,in modo che si possa
            'ricominciare in seguito a costruire un nuovo muro
            Stato_riga = 0
        'Altrimenti se viene premuto il tasto destro del mouse, e lo stato_riga = 0,allora imposto le coordinate
        'iniziali della nuova riga,rispettivamente alla posizione del mouse,e, ATTENZIONE,eseguo nuovamente
        'il controllo precedentemente spiegato,però,questa volta,lo farò per le coordinate iniziali
        'della nuova riga
        ElseIf B2 <> 0 And Stato_riga = 0 And SmouseX > 0 And SmouseX < Larghezza And SmouseY > 50 And SmouseY < Altezza Then
            Riga(I).X1 = SmouseX
            Riga(I).Y1 = SmouseY
            For I = 0 To Max
                TmpX1 = Riga(I).X1
                TmpX2 = Riga(I).X2
                TmpY1 = Riga(I).Y1
                TmpY2 = Riga(I).Y2
                If Riga(AppI).X1 > TmpX1 - 5 And Riga(AppI).X1 < TmpX1 + 5 And Riga(AppI).Y1 > TmpY1 - 5 And Riga(AppI).Y1 < TmpY1 + 5 Then
                    Riga(AppI).X1 = TmpX1
                    Riga(AppI).Y1 = TmpY1
                    Riga(AppI).SpigoloI = True
                    Riga(I).SpigoloI = True
                ElseIf Riga(AppI).X1 > TmpX2 - 5 And Riga(AppI).X1 < TmpX2 + 5 And Riga(AppI).Y1 > TmpY2 - 5 And Riga(AppI).Y1 < TmpY2 + 5 Then
                    Riga(AppI).X1 = TmpX2
                    Riga(AppI).Y1 = TmpY2
                    Riga(AppI).SpigoloI = True
                    Riga(I).SpigoloF = True
                End If
            Next
            'Assegno alla variabile Stato_riga il valore 1,in modo da far capire al programma che sono già state
            'impostate le coordinate iniziali del nuovo muro,e che quindi si dovrà passare ad assegnare le coordinate finali
            Stato_riga = 1
        End If
        'Nel caso in cui invece l'utente premesse dal menù opzioni il pulsante per l'aggiunta di un pavimento allora
        'si avvierà la stessa funzione sopra descritta,solamente che questa volta le linee intelligenti opereranno sulle
        'coordinate dei muri, in modo da creare dei pavimenti assolutamente allineati con i limiti della stanza e quindi
        'dei muri che la compongono
        ElseIf Scelta_Oggetto = "Pavimento" Or Scelta_Oggetto = "Soffitto" Then
            If B1 <> 0 And Stato_sop = 3 And SmouseX > 0 And SmouseX < Larghezza And SmouseY > 50 And SmouseY < Altezza Then
                SoP(J).X4 = SmouseX
                SoP(J).Y4 = SmouseY
                For I = 0 To Max
                    TmpX1 = Riga(I).X1
                    TmpX2 = Riga(I).X2
                    TmpY1 = Riga(I).Y1
                    TmpY2 = Riga(I).Y2
                    If SoP(J).X4 > TmpX1 - 5 And SoP(J).X4 < TmpX1 + 5 And SoP(J).Y4 > TmpY1 - 5 And SoP(J).Y4 < TmpY1 + 5 Then
                        SoP(J).X4 = TmpX1
                        SoP(J).Y4 = TmpY1
                        SoP(J).CR.SpigoloI = True
                    ElseIf SoP(J).X4 > TmpX2 - 5 And SoP(J).X4 < TmpX2 + 5 And SoP(J).Y4 > TmpY2 - 5 And SoP(J).Y4 < TmpY2 + 5 Then
                        SoP(J).X4 = TmpX2
                        SoP(J).Y4 = TmpY2
                        SoP(J).CR.SpigoloF = True
                    End If
                Next
                'Se la costruzione appena creata è un Pavimento,allora...
                If Scelta_Oggetto = "Pavimento" Then
                    'Assegno al Pavimento appena creato il suo nome,formato dalla parola "Pavimento" più
                    'l'indice della sua posizione all'interno della tabella Coordinate_SoP
                    If LinguaS = "Italiano" Then SoP(J).CR.Nome = "Pavimento " + Str(Max3 + 1)
                    If LinguaS = "Inglese" Then SoP(J).CR.Nome = "Floor " + Str(Max3 + 1)
                    'Aggiungo il Pavimento appena creato all'oggetto Elenco_SoP presnte all'interno
                    'del Form_Opzioni,in modo tale che l'utente potrà effettuarvi tutte le modifiche
                    'necessarie
                    Form_Opzioni.ElencoSoP.AddItem RTrim(SoP(J).CR.Nome)
                    'Inizializzo l'altitudine di default di ogni Pavimento uguale a 0
                    SoP(J).CR.Altitudine = 0
                    'Dichiaro che la costruzione appena creata è un Pavimento
                    SoP(J).Tipo = "Pavimento"
                    'Incremento la variabile Max3 di 1,ovvero quella variabile che tiene conto del numero
                    'di Pavimenti presenti all'interno della mappa attuale
                    Max3 = Max3 + 1
                'Se invece la costruzione appena creata è un Soffitto,allora...
                ElseIf Scelta_Oggetto = "Soffitto" Then
                    'Assegno al pavimento appena creato il suo nome,formato dalla parola "Soffitto" più
                    'l'indice della sua posizione all'interno della tabella Coordinate_SoP
                    If LinguaS = "Italiano" Then SoP(J).CR.Nome = "Soffitto " + Str(Max4 + 1)
                    If LinguaS = "Inglese" Then SoP(J).CR.Nome = "Ceiling " + Str(Max4 + 1)
                    'Aggiungo il Soffitto appena creato all'oggetto Elenco_SoP presnte all'interno
                    'del Form_Opzioni,in modo tale che l'utente potrà effettuarvi tutte le modifiche
                    'necessarie
                    Form_Opzioni.ElencoSoP.AddItem RTrim(SoP(J).CR.Nome)
                    'Inizializzo l'altitudine di default di ogni Soffitto uguale a 1000
                    SoP(J).CR.Altitudine = 1000
                    'Dichiaro che la costruzione appena creata è un Soffitto
                    SoP(J).Tipo = "Soffitto"
                    'Incremento la variabile Max4 di 1,ovvero quella variabile che tiene conto del numero
                    'di Soffitti presenti all'interno della mappa attuale
                    Max4 = Max4 + 1
                End If
                'Riporto la variabile Stato_Sop al suo valore iniziale in modo che in seguito si possa
                'iniziare a costruire un nuovo Pavimento / Soffitto
                Stato_sop = 0
                'Incremento l'indice J,in modo che si possa passare al Pavimento / Soffitto successeivo
                'a quello appena creato
                J = J + 1
                'Incremento la variabile Max2 che tiene conto del numero di Pavimenti più il numero di Soffitti
                'presenti all'interno della mappa attuale
                Max2 = Max2 + 1
            '----------------------------
            ElseIf B2 <> 0 And Stato_sop = 2 And SmouseX > 0 And SmouseX < Larghezza And SmouseY > 50 And SmouseY < Altezza Then
                SoP(J).X3 = SmouseX
                SoP(J).Y3 = SmouseY
                For I = 0 To Max
                    TmpX1 = Riga(I).X1
                    TmpX2 = Riga(I).X2
                    TmpY1 = Riga(I).Y1
                    TmpY2 = Riga(I).Y2
                    If SoP(J).X3 > TmpX1 - 5 And SoP(J).X3 < TmpX1 + 5 And SoP(J).Y3 > TmpY1 - 5 And SoP(J).Y3 < TmpY1 + 5 Then
                        SoP(J).X3 = TmpX1
                        SoP(J).Y3 = TmpY1
                        SoP(J).CR.SpigoloI = True
                    ElseIf SoP(J).X3 > TmpX2 - 5 And SoP(J).X3 < TmpX2 + 5 And SoP(J).Y3 > TmpY2 - 5 And SoP(J).Y3 < TmpY2 + 5 Then
                        SoP(J).X3 = TmpX2
                        SoP(J).Y3 = TmpY2
                        SoP(J).CR.SpigoloF = True
                    End If
                Next
                Stato_sop = 3
            '---------------------------------------
            ElseIf B1 <> 0 And Stato_sop = 1 And SmouseX > 0 And SmouseX < Larghezza And SmouseY > 50 And SmouseY < Altezza Then
                SoP(J).CR.X2 = SmouseX
                SoP(J).CR.Y2 = SmouseY
                For I = 0 To Max
                    TmpX1 = Riga(I).X1
                    TmpX2 = Riga(I).X2
                    TmpY1 = Riga(I).Y1
                    TmpY2 = Riga(I).Y2
                    If SoP(J).CR.X2 > TmpX1 - 5 And SoP(J).CR.X2 < TmpX1 + 5 And SoP(J).CR.Y2 > TmpY1 - 5 And SoP(J).CR.Y2 < TmpY1 + 5 Then
                        SoP(J).CR.X2 = TmpX1
                        SoP(J).CR.Y2 = TmpY1
                        SoP(J).CR.SpigoloI = True
                    ElseIf SoP(J).CR.X2 > TmpX2 - 5 And SoP(J).CR.X2 < TmpX2 + 5 And SoP(J).CR.Y2 > TmpY2 - 5 And SoP(J).CR.Y2 < TmpY2 + 5 Then
                        SoP(J).CR.X2 = TmpX2
                        SoP(J).CR.Y2 = TmpY2
                        SoP(J).CR.SpigoloF = True
                    End If
                Next
                Stato_sop = 2
            '--------------
            ElseIf B2 <> 0 And Stato_sop = 0 And SmouseX > 0 And SmouseX < Larghezza And SmouseY > 50 And SmouseY < Altezza Then
                SoP(J).CR.X1 = SmouseX
                SoP(J).CR.Y1 = SmouseY
                For I = 0 To Max
                    TmpX1 = Riga(I).X1
                    TmpX2 = Riga(I).X2
                    TmpY1 = Riga(I).Y1
                    TmpY2 = Riga(I).Y2
                    If SoP(J).CR.X1 > TmpX1 - 5 And SoP(J).CR.X1 < TmpX1 + 5 And SoP(J).CR.Y1 > TmpY1 - 5 And SoP(J).CR.Y1 < TmpY1 + 5 Then
                        SoP(J).CR.X1 = TmpX1
                        SoP(J).CR.Y1 = TmpY1
                        SoP(J).CR.SpigoloI = True
                    ElseIf SoP(J).CR.X1 > TmpX2 - 5 And SoP(J).CR.X1 < TmpX2 + 5 And SoP(J).CR.Y1 > TmpY2 - 5 And SoP(J).CR.Y1 < TmpY2 + 5 Then
                        SoP(J).CR.X1 = TmpX2
                        SoP(J).CR.Y1 = TmpY2
                        SoP(J).CR.SpigoloF = True
                    End If
                Next
                Stato_sop = 1
            End If
        '----------------------------------------------
    ElseIf Scelta_Oggetto = "Telecamera" Then
        If B2 <> 0 Then
            With PosizioneTelecamera
                .X = SmouseX
                .Z = SmouseY
                Form_Opzioni.TelecameraX = Str(.X)
                Form_Opzioni.TelecameraZ = Str(.Z)
            End With
        End If
    End If
End Sub

Sub Disegna_righe()
        'Dichiaro una variabile booleana che mi servirà per capire se sono stati riscontrati
        'dei problemi nel disegnare i muri,pavimenti o soffitti su schermo.
        'Attenzione: il valore di questa variabile (True o False) verrà assegnato dalla funzione
        'Problemi_righe
        Dim Problemi As Boolean
        'Se è presente almeno un muro,pavimento o soffitto all'interno dell'editor allora verranno attivate
        'le voci di menù Salva_con_nome,Converti_mappa_in_3d e Stampa
        If Max <> 0 Or Max2 <> 0 Then
            Salva_con_nome.Enabled = True
            Converti_mappa_in_3D.Enabled = True
            Stampa.Enabled = True
            AvviaAnteprima.Enabled = True
        'Altrimenti queste vengono disattivate con in più anche la voce di menù Salva
        Else
            Salva_con_nome.Enabled = False
            Salva.Enabled = False
            Converti_mappa_in_3D.Enabled = False
            Stampa.Enabled = False
        End If
        'Se il controllo Controlla_muri all'interno del form_opzioni è selezionato, allora
        'disegna tutte le righe precedentemente create
        For I = 0 To Max
            If Form_Opzioni.Controlla_Muri.Value = 1 Then
                'La variabile Problemi,assumerà il valore ritornato dalla funzione Problemi_righe
                'Se sono stati individuati dei problemi nel disegnare la riga corrente allora
                'questa assumerà il valore True, in caso contrario assumerà il valore
                'opposto (False)
                Problemi = (Problemi_righe(I, "Muro"))
                'Se la funzione Problemi_righe non ha riscontrato problemi nel disegnare le righe e quindi
                'non ha assegnato il valore True alla variabile Problemi,allora possiamo procedere a
                'disegnare normalmente (con tutte le rispettive coordinate originali) la riga corrente
                If Problemi = False Then
                    Schermo.DrawLine Riga(I).X1, Riga(I).Y1, Riga(I).X2, Riga(I).Y2, RGBA(CM.R, CM.G, CM.B, 1)
                End If
                'Se non sono stati riscontrati problemi nel disegnare la riga corrente,allora,
                'disegno il muro selezionato con un colore differente in modo da individuarlo
                'all'interno della mappa
                If Form_Opzioni.ElencoMuri.Text <> "" And Problemi = False Then
                    Schermo.DrawLine Riga(IndiceLista).X1, Riga(IndiceLista).Y1, Riga(IndiceLista).X2, Riga(IndiceLista).Y2, RGBA(CMS.R, CMS.G, CMS.B, 1)
                End If
            End If
            'Se la controlbox Controlla_spigoli è attivata,allora verranno disegnati dei piccoli
            'quadratini in corrispondensa degli angoli dei muri (cioè dove uno o più muri hanno ho le stesse coordinate
            'iniziali,o quelle finali
            If Form_Opzioni.Controlla_spigoli = 1 Then
                If Riga(I).SpigoloI = True Then
                    Schermo.DrawFilledBox Riga(I).X1 - 2, Riga(I).Y1 - 2, Riga(I).X1 + 2, Riga(I).Y1 + 2, RGBA(CSM.R, CSM.G, CSM.B, 1)
                End If
                If Riga(I).SpigoloF = True Then
                    Schermo.DrawFilledBox Riga(I).X2 - 2, Riga(I).Y2 - 2, Riga(I).X2 + 2, Riga(I).Y2 + 2, RGBA(CSM.R, CSM.G, CSM.B, 1)
                End If
            End If
        Next
        'Se dal form delle opzioni è stato scelto di visualizzare i pavimenti che sono stati creati
        'all'interno della mappa attuale,allora questi verranno mostrati all'utente tramite due linee
        'che collegano i quattro vertici del pavimento
        If Form_Opzioni.Visualizza_Pavimenti = 1 Or Form_Opzioni.Visualizza_soffitti = 1 Then
            'Avvio il ciclo for il quale disegnerà il quadruplo delle righe rispetto ai pavimenti e soffitti creati
            'PS: Il quadruplo perchè ogni pavimento o soffitto viene creato tramite quattro linee linee:
            ' - Una che unisce le coordinate A1 con le coordinate A2
            ' - Un'altra che unisce le coordinate A2 con le coordinate A4
            ' - Un'altra ancora che unisce le coordinate A4 con le coordinate A3
            ' - E infine l'ultima che unisce le coordinate A3 con le coordinate A1
            'Queste quattro linne formaranno un quadrilatero al fine di mostrare all'utente
            'una rappresentazione grafica del soffitto o pavimento appena creato
            For J = 0 To Max2
                'Se l'elemento che si stà analizzando è di tipo "Pavimento",allora setto il colore
                'con il relativo impostato con cui le linee che lo formano dovranno essere rappresentate su schermo
                If RTrim(SoP(J).Tipo) = "Pavimento" Then
                    With Colore
                        .R = CP.R
                        .G = CP.G
                        .B = CP.B
                        .A = 1
                    End With
                'In tutti gli altri casi,quindi se l'elemento analizzato è di tipo "Soffitto",allora setto
                'il colore,con il relativo impostato,con cui le linee che lo rappresentano su schermo dovranno
                'essere disegnate
                Else
                    With Colore
                        .R = CS.R
                        .G = CS.G
                        .B = CS.B
                        .A = 1
                    End With
                End If
                'Invece,se l'elemento che si stà analizzando è uguale a quello selezionato dalla combobox ElencoSoP
                'dele Form_Opzioni, allora assegnerò al pavimento / soffitto corrente il relativo colore impostato,
                'in modo da far capire all'utente quale elemento è stato selezionato all'interno della mappa attuale
                If J = IndiceLista2 Then
                    With Colore
                        .R = CSOPS.R
                        .G = CSOPS.G
                        .B = CSOPS.B
                        .A = 1
                    End With
                End If
                'Assegno tramite la funzione Problemi_righe, il valore booleano alla variabile Problemi,la quale
                'assumerà True, in caso una delle coordinate delle righe che compongono il pavimento o soffitto corrente
                'uscirà al di fuori della superficie dell'editor,in caso contrario il suo valore sarà False
                Problemi = (Problemi_righe(J, "SoP"))
                'Se non sono stati rilevati problemi nel disegnare le righe correnti, cioè nessuna delle coordinate
                'del pavimento o soffitto esce al di fuori della superficie dell'editor,allora le righe che lo compongono
                'graficamente, verranno tracciate normalmente senza nessun problema
                If Problemi = False Then
                    'Se è stato scelto dal Form_Opzioni di rendere visibile sull'editor tutti i pavimenti creati,e l'elemento
                    'analizzato è appunto un pavimento,allora verranno disegnate le linee che lo compongono del colore precedentemente
                    'settato (Blu)
                    If Form_Opzioni.Visualizza_Pavimenti = 1 And RTrim(SoP(J).Tipo) = "Pavimento" Then
                        'Disegno su schermo (la superficie dell'editor la prima linea che collegherà le coordinata A1
                        'del pavimento con le sue rispettive A2
                        Schermo.DrawLine SoP(J).CR.X1, SoP(J).CR.Y1, SoP(J).CR.X2, SoP(J).CR.Y2, RGBA(Colore.R, Colore.G, Colore.B, Colore.A)
                        'Disegno la seconda linea che questa volta collegherà le coordinate A2 del pavimento
                        'con le sue rispettive A4
                        Schermo.DrawLine SoP(J).CR.X2, SoP(J).CR.Y2, SoP(J).X4, SoP(J).Y4, RGBA(Colore.R, Colore.G, Colore.B, Colore.A)
                        'Disegno le terza linea che collegherà le coordinata A4 del pavimento con le sue rispettive A3
                        Schermo.DrawLine SoP(J).X4, SoP(J).Y4, SoP(J).X3, SoP(J).Y3, RGBA(Colore.R, Colore.G, Colore.B, Colore.A)
                        'Ora disegno la quarta e ultima linea che questa volta collegherà le coordinate A3 del pavimento
                        'con le sue rispettive A1
                        Schermo.DrawLine SoP(J).X3, SoP(J).Y3, SoP(J).CR.X1, SoP(J).CR.Y1, RGBA(Colore.R, Colore.G, Colore.B, Colore.A)
                    'In quest'altro caso invece,se è stato scelto dal Form_Opzioni di visualizzare tutti i soffitti creati,e l'elemento analizzato è
                    'appunto un soffitto,allora verranno disegnate le linee che lo compongono del colore,anch'esso precedentemente settato,Verde
                    ElseIf Form_Opzioni.Visualizza_soffitti = 1 And RTrim(SoP(J).Tipo) = "Soffitto" Then
                        'Disegno su schermo (la superficie dell'editor la prima linea che collegherà le coordinata A1
                        'del pavimento con le sue rispettive A2
                        Schermo.DrawLine SoP(J).CR.X1, SoP(J).CR.Y1, SoP(J).CR.X2, SoP(J).CR.Y2, RGBA(Colore.R, Colore.G, Colore.B, Colore.A)
                        'Disegno la seconda linea che questa volta collegherà le coordinate A2 del pavimento
                        'con le sue rispettive A4
                        Schermo.DrawLine SoP(J).CR.X2, SoP(J).CR.Y2, SoP(J).X4, SoP(J).Y4, RGBA(Colore.R, Colore.G, Colore.B, Colore.A)
                        'Disegno le terza linea che collegherà le coordinata A4 del pavimento con le sue rispettive A3
                        Schermo.DrawLine SoP(J).X4, SoP(J).Y4, SoP(J).X3, SoP(J).Y3, RGBA(Colore.R, Colore.G, Colore.B, Colore.A)
                        'Ora disegno la quarta e ultima linea che questa volta collegherà le coordinate A3 del pavimento
                        'con le sue rispettive A1
                        Schermo.DrawLine SoP(J).X3, SoP(J).Y3, SoP(J).CR.X1, SoP(J).CR.Y1, RGBA(Colore.R, Colore.G, Colore.B, Colore.A)
                    End If
                End If
            Next
        End If
        'La struttura di controllo riportata qui di seguito, mi serviranno per mostrare all'utente lo stato
        'di costruzione del nuovo pavimento
        '-Se questo è uguale a 1 verrà disegnato un cerchio nella rispettive coordinate X1 e Y1 del pavimento
        'in fase di costruzione
        Select Case Stato_sop
        Case Is = 1
            'Viene disegnato il primo cerchio in corrispondenza delle coordinate A1
            Schermo.DrawCircle SoP(J).CR.X1, SoP(J).CR.Y1, 5, 20, RGBA(0, 0, 1, 1)
        '-Se è uguale a due, allora verra ridisegnato il cerchio precedente più il nuovo
        'in corrispondenza delle coordinate X2 e Y2 del pavimento in fase di costruzione
        Case Is = 2
            'Viene disegnato il primo cerchio vecchio (in coordinate A1)
            Schermo.DrawCircle SoP(J).CR.X1, SoP(J).CR.Y1, 5, 20, RGBA(0, 0, 1, 1)
            'Viene disegnato il nuovo cerchio (in coordinate A2)
            Schermo.DrawCircle SoP(J).CR.X2, SoP(J).CR.Y2, 5, 20, RGBA(0, 0, 1, 1)
        '-Se è uguale a tre,allora verranno ridisegnati i due cerchi precedentemente creati
        'più il nuovo in corrispondenza delle rispettive coordianate X3 e Y3
        Case Is = 3
            'Viene disegnato il primo cerchio vecchio (in coordinate A1)
            Schermo.DrawCircle SoP(J).CR.X1, SoP(J).CR.Y1, 5, 20, RGBA(0, 0, 1, 1)
            'Viene disegnato il secondo cerchio vecchio (in coordinate A2)
            Schermo.DrawCircle SoP(J).CR.X2, SoP(J).CR.Y2, 5, 20, RGBA(0, 0, 1, 1)
            'Viene disegnato il nuovo cerchio (in coordinate A3)
            Schermo.DrawCircle SoP(J).X3, SoP(J).Y3, 5, 20, RGBA(0, 0, 1, 1)
        End Select
        'Richiamo la funzione che risolverà i problemi delle coordinate del mouse se queste fuoriescono dalla superficie
        'dell'editor
        Problemi_coordinate_mouse
        'Se la variabile Stato_riga a valore pari a 1 allora verrà disegnata una riga con coordinate iniziali già impostate
        'alla pressione del tasto sinistro del mouse,e finale pari alle coordinate attuali del mouse
        If Stato_riga = 1 Then
            Schermo.DrawLine Riga(I).X1, Riga(I).Y1, SmouseX, SmouseY, RGBA(CM.R, CM.G, CM.B, 1)
        End If
        'Se viene scelto di visualizzare la telecamera dal form delle opzioni allora questa verrà disegnata
        'all'interno dell'editor
        If Form_Opzioni.Visualizza_telecamera.Value = 1 Then
            Schermo.DrawFilledBox Larghezza - 110, 40, Larghezza - 25, 65, RGBA(0.5, 0.5, 1, 1)
            If LinguaS = "Italiano" Then Schermo.DrawText "Telecamera", Larghezza - 100, 45, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
            If LinguaS = "Inglese" Then Schermo.DrawText "Camera", Larghezza - 85, 45, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
            Problemi = Problemi_telecamera
            If Problemi = False Then
                Schermo.DrawLine Larghezza - 112, 67, PosizioneTelecamera.X, PosizioneTelecamera.Z, RGBA(0, 0.8, 0.4, 1)
                Schermo.DrawFilledBox PosizioneTelecamera.X - 2, PosizioneTelecamera.Z - 2, PosizioneTelecamera.X + 2, PosizioneTelecamera.Z + 2, RGBA(0, 0.8, 0.4, 1)
            End If
        End If
        If Form_Opzioni.VisualizzaOggetto.Value = 1 Then
            Schermo.DrawFilledBox 10, Altezza - 35, 95, Altezza - 10, RGBA(1, 0, 0, 1)
            Schermo.DrawText Form_Opzioni.ElencoGruppiOggetti.SelectedItem.Text, 20, Altezza - 30, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
            Problemi = Problemi_telecamera
            Dim TmpX As Single
            Dim TmpY As Single
            Problemi = Problemi_Oggetto
            If Problemi = False Then
                TmpX = Oggetto(IndiceOggettoSelezionato).Ricava_Posizione.X
                TmpY = Oggetto(IndiceOggettoSelezionato).Ricava_Posizione.Z
                Schermo.DrawLine 52, Altezza - 38, TmpX, TmpY, RGBA(0, 0.8, 0.4, 1)
                Schermo.DrawFilledBox TmpX - 2, TmpY - 2, TmpX + 2, TmpY + 2, RGBA(0, 0.8, 0.4, 1)
            End If
        End If
End Sub

Sub Crea_menù()
    'Se dal mnù opzione è stato selezionato di visualizzare il menù di riferimento agli oggetti creati,allora...
    If Form_Opzioni.Mostra_Menù.Value = 1 Then
        'Viene disegnato un rettangolo largo tutta la superficie delll'editor,all'interno del quale verranno scitte tutte
        'le informazioni (numero di muri,numero di pavimenti e numero di soffitti)
        Schermo.DrawFilledBox 0, 0, Larghezza, 30, RGBA(CSFM.R, CSFM.G, CSFM.B, 1)
        'Vengono scritti all'interno del rettangolo precedentemente creato,il numero di muri presenti
        'all'interno della mappa attuale
        If LinguaS = "Italiano" Then
            Schermo.DrawText "Muri:", 10, 10, RGBA(C1M.R, C1M.G, C1M.B, 1), "carattere_personalizzato2"
            Schermo.DrawText "Pavimenti:", 130, 10, RGBA(C1M.R, C1M.G, C1M.B, 1), "carattere_personalizzato2"
            Schermo.DrawText "Soffitti:", 280, 10, RGBA(C1M.R, C1M.G, C1M.B, 1), "carattere_personalizzato2"
            Schermo.DrawText "Oggetti:", 430, 10, RGBA(C1M.R, C1M.G, C1M.B, 1), "carattere_personalizzato2"
        ElseIf LinguaS = "Inglese" Then
            Schermo.DrawText "Walls:", 10, 10, RGBA(C1M.R, C1M.G, C1M.B, 1), "carattere_personalizzato2"
            Schermo.DrawText "Floors:", 130, 10, RGBA(C1M.R, C1M.G, C1M.B, 1), "carattere_personalizzato2"
            Schermo.DrawText "Ceilings:", 270, 10, RGBA(C1M.R, C1M.G, C1M.B, 1), "carattere_personalizzato2"
            Schermo.DrawText "Objects:", 425, 10, RGBA(C1M.R, C1M.G, C1M.B, 1), "carattere_personalizzato2"
        End If
        Schermo.DrawText Str(Max), 45, 8, RGBA(C2M.R, C2M.G, C2M.B, 1), "carattere_personalizzato1"
        '...il numero di pavimenti
        Schermo.DrawText Str(Max3), 190, 8, RGBA(C2M.R, C2M.G, C2M.B, 1), "carattere_personalizzato1"
        '... il numero di soffitti
        Schermo.DrawText Str(Max4), 320, 8, RGBA(C2M.R, C2M.G, C2M.B, 1), "carattere_personalizzato1"
        '...e infine il numero di oggetti inseriti
        Schermo.DrawText Str(IOg), 470, 8, RGBA(C2M.R, C2M.G, C2M.B, 1), "carattere_personalizzato1"
    End If
End Sub

Private Sub Form_Terminate()
    'Elimino tutti gli oggetti precedentemente creati
    Distruggi_oggetti
    'Termino il programma
    End
End Sub

Private Sub Info_Click()
    'Alla pressione della voce di menù File richiamerò il metodo dalla classe sonora
    'che riprodurrà il suono di menù
    Suoni.Esegui_suono_menù
End Sub

Private Sub LInglese_Click()
    'Richiama il metodo dalla  Classe Sonora che riprodurrà il file sonoro Menù2
    Suoni.Esegui_suono_menù_2
    'Assegno il valore di proprietà Checked = True al bottone LInglese stesso.
    'Questo sarà di aiuto all'utente per capire con quale linguaggio è attualmete tradotto il programma
    LInglese.Checked = True
    'Assegno il valore di proprietà Checked = False al bottone LItaliano.
    LItaliano.Checked = False
    'Assegno alla variabile LinguaS la stringa Inglese,in modo da far capire al programma che la lingua
    'attualmente usata è appunto l'inglese
    LinguaS = "Inglese"
    'Richiamo la funzione che mi permetterà di tradurre il programma nella lingua desiderata,che in questo caso sarà l'inglese
    Traduci "Inglese"
    'Se il Form_Assegnazione_Multipla è aperto,allora richiamerò la funzione addetta alla sua
    'traduzione
    If Form_Assegnazione_Multipla.Visible = True Then Form_Assegnazione_Multipla.Traduci "Inglese"
    'Se il Form_Materiali è aperto,allora richiamerò la funzione addetta alla sua
    'traduzione
    If Form_Materiali.Visible = True Then Form_Materiali.Traduci "Inglese"
End Sub

Private Sub Lingua_Click()
    'Alla pressione della voce di menù File richiamerò il metodo dalla classe sonora
    'che riprodurrà il suono di menù
    Suoni.Esegui_suono_menù
End Sub

Private Sub LItaliano_Click()
    'Richiama il metodo dalla  Classe Sonora che riprodurrà il file sonoro Menù2
    Suoni.Esegui_suono_menù_2
    'Assegno alla variabile LinguaS la stringa Italiano,in modo da far capire al programma che la lingua
    'attualmente usata è appunto l'italiano
    LinguaS = "Italiano"
    'Richiamo la funzione che mi permetterà di tradurre il programma nella lingua desiderata,che in questo caso sarà l'italiano
    Traduci "Italiano"
    'Se il Form_Assegnazione_Multipla è aperto,allora richiamerò la funzione addetta alla sua
    'traduzione
    If Form_Assegnazione_Multipla.Visible = True Then Form_Assegnazione_Multipla.Traduci "Italiano"
    'Se il Form_Materiali è aperto,allora richiamerò la funzione addetta alla sua
    'traduzione
    If Form_Materiali.Visible = True Then Form_Materiali.Traduci "Italiano"
End Sub

Private Sub m3D_Click()
    'Alla pressione della voce di menù 3D richiamerò il metodo dalla classe sonora
    'che riprodurrà il suono di menù
    Suoni.Esegui_suono_menù
End Sub

Private Sub Modifica_Click()
    'Alla pressione della voce di menù Modifica richiamerò il metodo dalla classe sonora
    'che riprodurrà il suono di menù
    Suoni.Esegui_suono_menù
End Sub

Private Sub Nuovo_Click()
    'Richiama il metodo dalla  Classe Sonora che riprodurrà il file sonoro Menù2
    Suoni.Esegui_suono_menù_2
    Dim Risposta_Nuovo As VbMsgBoxResult
    'Nel caso in cui si stesse visualizzando l'animazione,alla pressione del tasto nuovo questa verrebbe
    'terminata
    If Continua_Animazione = True Then
        'Imposto alla variabile continua_animazione il valore booleano false,cosicchè alla pressione
        'della voce di menù verrà terminata la sequenza introduttiva e verrà introdotto il programma
        'vero e proprio
        Continua_Animazione = False
    ElseIf Max <> 0 Then
        'Viene visualizzato un messaggio che avviserà l'utente se si vorrà salvare la mappa corrente prima di procedere
        'e viene assegnato il valore del tasto premuto alla variabile Risposta_nuovo
        'If LinguaS = "Italiano" Then Risposta_Nuovo = MsgBox("La creazione di una nuova mappa cancellerà quella attuale! Vuoi salvare la mappa corrente prima di procedere?", vbYesNoCancel, "Creazione di una nuova mappa")
        'If LinguaS = "Inglese" Then Risposta_Nuovo = MsgBox("If you create a new map, the current map will be erased!Do you want save the current map befor proceding?", vbYesNoCancel, "New map creation")
        'Se viene premuto il tasto yes si procederà con le operazioni di salvataggio del file...
        'If Risposta_Nuovo = vbYes Then
            'Salva_mappa
        '...in caso contrario la mappa corrente non verrà salvata
        'If Risposta_Nuovo = vbNo Then
            'Chiama la funzione che reinizializzerà tutti i campi del tipo definito coordinate_riga e coordinate_sop e imposta
            'la variabile Max,Max2,Max3,Max4,I,J = 0
            Reinizializza_righe
        'End If
    End If
End Sub

Private Sub Opzioni_Click()
    'Questa funzione,permette alla pressione della voce di menù Visualizza/Opzioni
    'di nascondere o mostrare su schermo il form delle opzioni e contrassegnare nel menù
    'stesso se questo è visibile o meno
    If Form_Opzioni.Visible = True Then
        Form_Opzioni.Visible = False
        Map_Editor.Opzioni.Checked = False
    Else
        Form_Opzioni.Visible = True
        Map_Editor.Opzioni.Checked = True
    End If
End Sub

Sub Distruggi_oggetti()
    'Le istruzioni che seguono,servono per distruggere tutti gli oggetti che mi sono creato per
    'il funzionamento del programma,prima dell'uscita dallo stesso.Questa operazione è molto
    'importante,in quanto serve a non lasciare in memoria,residui dello stesso
    Set Comandi = Nothing
    Set Schermo = Nothing
    Set Scena = Nothing
    Set TV8 = Nothing
End Sub

Private Sub Converti_mappa_in_3D_Click()
    'Assegno alla variabile CM3D il valore True,in modo da far capire al compilatore che
    'è stata scelta la voce di menù Converti mappa in 3D
    CM3D = True
    'Richiamo la funzione che mi permetterà di convertire la mappa corrente in una mappa 3D
    Funzione_converti_mappa_in_3d
End Sub


Private Sub Salva_Click()
    'Richiamo il metodo dalla Classe Sonora che riprodurrà il terzo file sonoro di menù
    Suoni.Esegui_suono_menù_3
    'Dichiaro che è stata scelta la voce di menù salva
    SCN = False
    'Avvio la funzione che salverà la mappa corrente senza richiedere il file da salvare
    Salva_mappa
End Sub

Private Sub Salva_con_nome_Click()
    'Richiamo il metodo dalla Classe Sonora che riprodurrà il terzo file sonoro di menù
    Suoni.Esegui_suono_menù_3
    'Dichiaro che è stata scelta la voce di menù Salva_con_nome
    SCN = True
    'Avvio la funzione che mi permetterà di salvare la mappa corrente
    Salva_mappa
End Sub

Sub Allinea_form()
    'Questa funzione permette di allineare tutti e tre i form anche quando uno di questo viene spostato
    'al fine di farli sembrare "un blocco unico",solo se il form_opzioni è visibile
    If Form_Opzioni.Visible = True Then
        Form_Opzioni.Top = Me.Top
        Form_Opzioni.Left = Me.Left + Me.Width
    End If
End Sub

Sub Controlla_griglia()
    'Se viene scelto di visualizzare la griglia dal form delle opzioni e il controllo GrigliaControllataDaZoom è disattivato,
    'allora questa verrà creata e visualizzata secondo i valori specificati dall'utente
    If Form_Opzioni.Visualizza_griglia = 1 And Form_Opzioni.GrigliaControllataDaZoom = 0 Then
        Form_Opzioni.AltezzaGriglia.Enabled = True
        Form_Opzioni.LarghezzaGriglia.Enabled = True
        'Dichiaro una variabile che mi tornerà utile per disegnare la griglia
        Dim IGriglia As Single
        'Dichiaro due variabili di appoggio che conterranno il valore dell'altezza e della
        'larghezza di ogni quadrato della griglia
        Dim TmpLarghezza As Single
        Dim TmpAltezza As Single
        'Assegno i valori dei altezza e larghezza alle rispettive variabili
        TmpLarghezza = Val(Form_Opzioni.LarghezzaGriglia.Text)
        TmpAltezza = Val(Form_Opzioni.AltezzaGriglia.Text)
        'Le istruzioni che seguono servono al fine di reimpostare le due variabili con un valore
        'pari a 1,nel caso in cui le rispettive textbox fossero prive di valori.Questo eviterà
        'uno stallo da parte del programma
        If TmpLarghezza = 0 Then TmpLarghezza = 1
        If TmpAltezza = 0 Then TmpAltezza = 1
        'Disegno le righe orizzontali della griglia...
        For IGriglia = 0 To Larghezza - 1 Step TmpAltezza
            Schermo.DrawLine 0, IGriglia, Larghezza - 1, IGriglia, RGBA(0, 0, 0, Form_Opzioni.Luminosità_griglia.Value / 10)
        Next
        '...e quelle verticali
        For IGriglia = 0 To Larghezza - 1 Step TmpLarghezza
            Schermo.DrawLine IGriglia, 0, IGriglia, Altezza - 1, RGBA(0, 0, 0, Form_Opzioni.Luminosità_griglia.Value / 10)
        Next
    'Se viene scelto di visualizzare la griglia dal form delle opzioni e il controllo GrigliaControllataDaZoom è attivato,
    'allora questa verrà creata e visualizzata secondo i valori ridimensioti dalle operazioni di zoom
    ElseIf Form_Opzioni.Visualizza_griglia = 1 And Form_Opzioni.GrigliaControllataDaZoom = 1 Then
        'Disattivo le due textbox che permettono di modificare la grandezza dei quadrati della griglia
        Form_Opzioni.AltezzaGriglia.Enabled = False
        Form_Opzioni.LarghezzaGriglia.Enabled = False
        'Disegno le righe della griglia, secondo il valore contenuto in VCambiamentiGriglia
        For IGriglia = 0 To Larghezza - 1 Step VCambiamentiGriglia
            Schermo.DrawLine IGriglia, 0, IGriglia, Altezza - 1, RGBA(0, 0, 0, Form_Opzioni.Luminosità_griglia.Value / 10)
            Schermo.DrawLine 0, IGriglia, Larghezza - 1, IGriglia, RGBA(0, 0, 0, Form_Opzioni.Luminosità_griglia.Value / 10)
        Next
    '...altrimenti se l'opzione è deselezionata,ci si limiterà ad imporre alle due textbox
    'contenenti i valori di altezza e larghezza della griglia,uno stato di enabled = false,
    'in modo che non possano essere introdotti i rispettivi valori
    Else
        Form_Opzioni.AltezzaGriglia.Enabled = False
        Form_Opzioni.LarghezzaGriglia.Enabled = False
    End If
End Sub

Sub Reinizializza_righe()
        'Reinizializzo tutte le righe create sulla mappa corrente con dei valori di default
        For I = 0 To Max
            Riga(I).X1 = 0
            Riga(I).X2 = 0
            Riga(I).Y1 = 0
            Riga(I).Y2 = 0
            Riga(I).Altezza = 1000
            Riga(I).Altitudine = 0
            Riga(I).NMattonelleALtezza = 10
            Riga(I).NMAttonelleLarghezza = 10
            Riga(I).Texture = "Nessuna"
            Riga(I).Proprietà = "Normale"
            Riga(I).SpigoloI = False
            Riga(I).SpigoloF = False
            Reimposta_materiale Riga(I).Materiale
        Next
        'Reinizializzo il numero di righe presenti sullo schermo con un valore pari a 0...
        Max = 0
        '...e faccio la stessa cosa anche per l'indice e l'indice di lista
        I = 0
        IndiceLista = 0
        'Reinizializzo tutti i pavimenti o soffitti creati con dei valori di default
        For J = 0 To Max2
            SoP(J).CR.X1 = 0
            SoP(J).CR.X2 = 0
            SoP(J).X3 = 0
            SoP(J).X4 = 0
            SoP(J).CR.Y1 = 0
            SoP(J).CR.Y2 = 0
            SoP(J).Y3 = 0
            SoP(J).Y4 = 0
            SoP(J).CR.Altitudine = 0
            SoP(J).CR.NMattonelleALtezza = 20
            SoP(J).CR.NMAttonelleLarghezza = 20
            SoP(J).CR.Texture = "Nessuna"
            SoP(J).CR.SpigoloI = False
            SoP(J).CR.SpigoloI = False
            Reimposta_materiale SoP(J).CR.Materiale
        Next
        'Reinizializzo l'indice della lista presente nel Form_Opzioni, contenente l'elenco dei pavimenti
        'o soffitti creati
        IndiceLista2 = 0
        'Reinizializzo la varaibile che tiene conto del numero di pavimenti più il numero di soffitti
        'presenti all'interno della mappa attuale
        Max2 = 0
        'Reinizializzo la variabile che tiene conto del solo numero dei pavimenti
        Max3 = 0
        'Reinizializzo la variabile che tiene conto del solo numero dei soffitti creati
        Max4 = 0
        'Cancello tutti gli elementi presenti all'interno della combobox Elenco_muri
        Form_Opzioni.ElencoMuri.Clear
        'Cancello tutti gli elementi contenuti all'interno della combobox Elenco_SoP
        Form_Opzioni.ElencoSoP.Clear
        'Elimino l'immagine risidua all'interno del controllo image di selezione texture del muro
        Form_Opzioni.Texture_muro.Picture = Nothing
        'Elimino l'immagine residua all'interno del controllo image di selezion
        Form_Opzioni.Texture_SoP.Picture = Nothing
        'Azzero il contenuto di alcune Textbox presenti all'interno del form opzioni,in
        'modo che non rimangana residui per la nuova mappa che ci andremo a creare
        With Form_Opzioni
            'Reimposto la proprietà Text della TextBox contenuta nel Form_Opzioni Altezza_muro
            .Altezza_muro = ""
            'Reimposto la proprietà Text della TextBox contenuta nel Form_Opzioni Altitudine_muro
            .Altitudine_muro = ""
            'Reimposto la proprietà Text della TextBox contenuta nel Form_Opzioni Mattonelle_Altezza
            .MattonelleAltezza = ""
            'Reimposto la proprietà Text della TextBox contenuta nel Form_Opzioni Mattonelle_Larghezza
            .MattonelleLarghezza = ""
            'Reimposto la proprietà Text della TextBox contenuta nel Form_Opzioni Nomer
            .Nomer = ""
            'Reimposto la proprietà Text della TextBox contenuta nel Form_Opziini AltitudineSoP
            .AltitudineSoP = ""
            'Reimposto la proprietà Text della TextBox contenuta nel Form_Opzioni Mattonelle_Altezza2
            .MattonelleAltezza2 = ""
            'Reimposto la proprietà Text della TextBox contenuta nel Form_Opzioni Mattonelle_Larghezza2
            .MattonelleLarghezza2 = ""
            'Reimposto la proprietà Text della TextBox contenuta nel Form_Opzioni Nomer2
            .Nomer2 = ""
            'Reimposto la proprietà Text della TextBox contenuta nel Form_Opzioni TipoPavimento
            .TipoPavimento.Value = False
            'Reimposto la proprietà Text della TextBox contenuta nel Form_Opzioni TiposSoffitto
            .TipoSoffitto.Value = False
        End With
        'Dichairo un indice di appoggio per scorrere all'interno degli oggetti caricati
        Dim K As Integer
        'Scandisco tutti gli oggetti caricati al fine di...
        For K = 0 To IOg
            'Distruggerli, reinizializzando tutte le sue variabili private tramite la
            'chiamata appunto del metodo Distruggi_Oggetto
            Oggetto(K).Distruggi_Oggetto
        'Si passa ad esaminare l'oggetto successivo
        Next
        'Reinizializzo la variabile IOg, la quale tiene conto del numero di oggetti inseriti
        'all'interno della mappa 3D
        IOg = 0
        'Cancello tutti i nodi creati all'interno del controllo ELencoGruppiOggetti presente
        'nel Form_Opzioni
        Form_Opzioni.ElencoGruppiOggetti.Nodes.Clear
        'Riaggiungo il nodo primario,il quale servirà a contenere tutti quegli oggetti privi di gruppo
        With Form_Opzioni.ElencoGruppiOggetti.Nodes.Add(, , "OSG", "Oggetti senza gruppo")
            'Setto il colore blu al nodo appena aggiunto
            .ForeColor = vbBlue
    End With
End Sub

Sub Salva_mappa()
    'Imposto il titolo dell'oggetto operazioni,in modo che l'utente potrà capire quale operazione
    'verrà effettuata
    If LinguaS = "Italiano" Then Operazioni.DialogTitle = "Salvataggio mappa 2D"
    If LinguaS = "Inglese" Then Operazioni.DialogTitle = "2D map save"
    'Impongo un filtro all'oggetto operazione, in modo che la mappa verrà salvata con un'estensione
    'da me definita e diversa da quella delle mappe 3D
    If LinguaS = "Italiano" Then Operazioni.Filter = "Mappa 2D (*.Map2D)|*.Map2D"
    If LinguaS = "Inglese" Then Operazioni.Filter = "2D Map (*.Map2D)|*.Map2D"
    'Se è stato scelto Salva con nome allora verrà chiesto il file in cui salvare
    If SCN = True Then
        'Se viene premuto annulla si verrà reinderizzati al label Annulla1.
        On Error GoTo Annulla1
        'Visualizzo la finestra di dialogo di salvataggio del file
        Operazioni.ShowSave
        'Salvo il file selezionato nella variabile FileSalvato
        FileSalvato = Operazioni.FileName
    End If
    'impongo un limite di errore nel caso il file da eliminare non esistesse
    On Error Resume Next
    'Elimino (se esiste) il file selezionato.Questo mi servirà per eliminare completamente tutto
    'il contenuto precedentemente salvato del file,per ricostruirlo da 0
    Kill FileSalvato
    'Apro il file selezionato e comincio le operazioni di salvataggio
    Open FileSalvato For Random As #1 Len = Len(Muro)
    For I = 1 To Max
        '---------------------------------------------------------------------------------------
        'Operazione di ripristino delle righe, nel caso fossero state modificate dall'operazione
        'di Zoom
        '---------------------------------------------------------------------------------------
        Riga(I).X1 = Fix(Riga(I).X1 / Molt)
        Riga(I).X2 = Fix(Riga(I).X2 / Molt)
        Riga(I).Y1 = Fix(Riga(I).Y1 / Molt)
        Riga(I).Y2 = Fix(Riga(I).Y2 / Molt)
        '------------------------------------
        'Operazione di modifica
        '------------------------------------
        Muro.X1 = Riga(I).X1
        Muro.X2 = Riga(I).X2
        Muro.Y1 = Riga(I).Y1
        Muro.Y2 = Riga(I).Y2
        Muro.Altezza = Riga(I).Altezza
        Muro.Altitudine = Riga(I).Altitudine
        Muro.Texture = Riga(I).Texture
        Muro.Nome = Riga(I).Nome
        Muro.NMattonelleALtezza = Riga(I).NMattonelleALtezza
        Muro.NMAttonelleLarghezza = Riga(I).NMAttonelleLarghezza
        Muro.SpigoloI = Riga(I).SpigoloI
        Muro.SpigoloF = Riga(I).SpigoloF
        '------------------------------------
        'Operazione di scrittura del record
        '------------------------------------
        Put #1, , Muro
        '--------------------------------------------------------------------------
        'Operazione di ripristino Zoom (Mediante le istruzioni che seguono,dopo
        'il salvataggio il programma tornerà automaticamente allo Zoom selezionato
        'reimpostando nuovamente tutte le righe della mappa attuale
        '--------------------------------------------------------------------------
        Riga(I).X1 = Riga(I).X1 * Molt
        Riga(I).X2 = Riga(I).X2 * Molt
        Riga(I).Y1 = Riga(I).Y1 * Molt
        Riga(I).Y2 = Riga(I).Y2 * Molt
    Next
    'Chiudo il file appena aperto
    Close #1
    'Attivo la voce di menù Salva
    Salva.Enabled = True
    'Esco dalla funzione
    Exit Sub
Annulla1:
End Sub

Sub Carica_mappa_salvata()
    'Imposto il titolo dell'oggetto operazioni,in modo che l'utente potrà capire quale operazione
    'verrà effettuata
    If LinguaS = "Italiano" Then Operazioni.DialogTitle = "Apertura mappa 2D"
    If LinguaS = "Inglese" Then Operazioni.DialogTitle = "2D map open"
    'Impongo un filtro all'oggetto operazione, in modo che la mappa verrà salvata con un'estensione
    'da me definita e diversa da quella delle mappe 3D
    If LinguaS = "Italiano" Then Operazioni.Filter = "Mappa 2D (*.Map2D)|*.Map2D"
    If LinguaS = "Inglese" Then Operazioni.Filter = "2D Map (*.Map2D)|*.Map2D"
    'Dichiaro una variabile che mi servirà per capire quale bottone viene premuto
    'dall'utente come risposta al messaggio di salvataggio
    Dim Risposta3 As VbMsgBoxResult
    'Verifica il tastopremuto dall'utente e se viene premuto si alla domanda di salvataggio
    'allora si richiamerà la funzione che salverà la mappa corrente e in seguito si procederà
    'con il caricamento di una mappa già esistente
    If Max <> 0 Then
        If LinguaS = "Italiano" Then Risposta3 = MsgBox("La mappa corrente non è stata salvata!Si desidera salvare il progetto corrente prima di procedere con l'operazione di caricamento?", vbYesNoCancel, "Salvataggio mappa")
        If LinguaS = "Inglese" Then Risposta3 = MsgBox("The current map isn't saved!Do you want save the current map before proceding with load operation?", vbYesNoCancel, "Map save")
        'Se al messaggio appena visualizzato l'utente ha premuto il pulsante Yes,allora verrà richiamata
        'la funzione addetta al salvataggio della mappa attuale
        If Risposta3 = vbYes Then
            'Chiamata dell funzione addetta al salvataggio della mappa attuale
            Salva_mappa
        End If
        'Se invece è stato premuto il pulsante Cancel,allora si verrà reinderizzati verso il
        '"Label" Annulla
        If Risposta3 = vbCancel Then GoTo Annulla
        'Se viene premuto annulla si verrà reinderizzati al label Annulla2
        On Error GoTo Annulla
        'Visualizzo la finestra di dialogo di caricamento del file
        Operazioni.ShowOpen
        'Cancello tutti gli eventuali elementi presenti all'interno della combobox
        'Elenco_muri
        Form_Opzioni.ElencoMuri.Clear
        'Il pezzo di codice che segue serve per cancellare il contenuto di alcune textbox presenti all'interno del
        'Form_Opzioni
        With Form_Opzioni
            'Reimposto la proprietà Text della TextBox contenuta nel Form_Opzioni Altezza_muro
            .Altezza_muro = ""
            'Reimposto la proprietà Text della TextBox contenuta nel Form_Opzioni Altitudine_muro
            .Altitudine_muro = ""
            'Reimposto la proprietà Text della TextBox contenuta nel Form_Opzioni MattonelleAltezza
            .MattonelleAltezza = ""
            'Reimposto la proprietà Text della TextBox contenuta nel Form_Opzioni MattonelleLarghezza
            .MattonelleLarghezza = ""
            'Reimposto la proprietà Text della TextBox contenuta nel Form_Opzioni Nomer
            .Nomer = ""
        End With
    Else
        On Error GoTo Annulla
        Operazioni.ShowOpen
        'Nel caso in cui fosse in corso l'animazione introduttiva,questa verrà terminata e verrà caricato un
        'cursore da me costruito
        If Continua_Animazione = True Then
            Continua_Animazione = False
        End If
        'Attivo la voce di menù salva in modo che ogni eventuale modifica apportata alla nuova mappa potrà
        'essere salvata direttamente senza mostrare la finestra di dialogo di apertura file
        Salva.Enabled = True
    End If
    'Salvo in una variabile il file che si è deciso di aprire
    FileSalvato = Operazioni.FileName
    'Eseguo un effetto di sfumatura in uscita,in modo che la mappa attuale scompaia pian piano
    Effetti.FadeOut
    'Effettuo un effetto di sfumatura in entrata,in modo che la mappa caricata compaia graduatamente
    Effetti.FadeIn 1000
    'Apro il file selezionato e comincio le operazioni di salvataggio
    Open FileSalvato For Random As #1 Len = Len(Muro)
    'Inizializzo l'indice a 1
    I = 1
    'Finchè il file non è finito...
    While Not EOF(1)
        '...leggo un record alla volta dal file...
        Get #1, , Muro
        '...imposto la riga con tutti i valori appena prelevati
        Riga(I).X1 = Muro.X1
        Riga(I).X2 = Muro.X2
        Riga(I).Y1 = Muro.Y1
        Riga(I).Y2 = Muro.Y2
        Riga(I).Altezza = Muro.Altezza
        Riga(I).Altitudine = Muro.Altitudine
        Riga(I).Texture = Muro.Texture
        Riga(I).Nome = Muro.Nome
        Riga(I).NMattonelleALtezza = Muro.NMattonelleALtezza
        Riga(I).NMAttonelleLarghezza = Muro.NMAttonelleLarghezza
        Riga(I).SpigoloI = Muro.SpigoloI
        Riga(I).SpigoloF = Muro.SpigoloF
        Form_Opzioni.ElencoMuri.AddItem Muro.Nome
        'Increamento l'indice
        I = I + 1
    Wend
    'Attivo il primo elemento della lista
    Form_Opzioni.ElencoMuri.ListIndex = 0
    'Elimino l'ultimo elemento alla combobox elenco_muri,che viene aggiunto per errore
    Form_Opzioni.ElencoMuri.RemoveItem (Form_Opzioni.ElencoMuri.ListCount - 1)
    'Chiudo il file appena aperto
    Close #1
    Max = I - 2
    'Reinizializzo le righe rimanenti con dei valori di default
    For I = Max + 1 To 10000
        With Riga(I)
            .Nome = "Muro" + Str(I)
            .NMattonelleALtezza = 10
            .NMAttonelleLarghezza = 10
            .Altezza = 1000
            .Altitudine = 0
        End With
    Next
    'Esco dalla funzione
    Exit Sub
Annulla:
End Sub

Private Sub StoppaAnteprima_Click()
    'Dichiaro una variabile che mi servirà a contenere il valore del tasto premuto in
    'corrispondenza del messaggio visualizzato
    Dim RispostaUscitaModalità3D As VbMsgBoxResult
    'Assegno alla variabile sopra dichiarata il valore del tasto premuto
    If LinguaS = "Italiano" Then RispostaUscitaModalità3D = MsgBox("Vuoi Uscire dalla modalità Anteprima 3D e tornare all'editor?", vbYesNo, "Uscita dalla modalità Anteprima 3D")
    If LinguaS = "Inglese" Then RispostaUscitaModalità3D = MsgBox("Do you really want exit from 3D preview modaty and return to the editor?", vbYesNo, "3D preview modality exit")
    'Se l'utente ha scelto di abbandonare la suddetta modalità allora si tornerà al
    'ciclo principale,ovvero quello che consente di effettuare tutte le operazioni di
    'creazione e modifica della mappa corrente
    If RispostaUscitaModalità3D = vbYes Then
        'Richiamo il metodo dalla Classe Sonora che riprodurra il file sonoro di uscita dalla
        'modalità Anteprima 3D
        Suoni.Esegui_suono_StoppaAnteprima
        'Assegno alla variabile Continua_Animazione_3D il valore booleano = false.Questo
        'mi permetterà di uscire dal ciclo dall'anteprima della mappa 3d e tornare all'editor
        'Vero e proprio
        Continua_Anteprima_3d = False
        'Disabilito lui stesso (Rendo non cliccabile la voce di menù StoppaAnteprima)
        StoppaAnteprima.Enabled = False
        'Attivo la voce di menù Avvia anteprima,in modo che una volta terminata la modalità Anteprima 3D
        'sia possibile ritornarci in un secondo momento
        AvviaAnteprima.Enabled = True
        'Richiamo il metodo dell'oggetto Mappa3D che distruggerà la mappa 3D creata
        Mappa3D.Distruggi_Mappa_3D
    End If
End Sub

Sub Setta_cursore()
    'Controllo le coordinate del cursore
    GetCursorPos tmpWindowsMousePosition
    'Posiziono il cursore 3D nelle stesse coordinate del cursore di windows
    Comandi.SetMousePosition tmpWindowsMousePosition.X - ((Map_Editor.Left / 15) + 18), tmpWindowsMousePosition.Y - ((Map_Editor.Top / 15) + 63)
End Sub

Private Sub Visualizza_Click()
    'Alla pressione della voce di menù Visualizza richiamerò il metodo dalla classe sonora
    'che riprodurrà il suono di menù
    Suoni.Esegui_suono_menù
End Sub

Sub Funzione_converti_mappa_in_3d()
    'Le istruzioni che seguono servono ad avvisare l'utente che la scalatura selezionata non è una delle migliori
    'al fine di una corretta visualizzazione della mappa 3D in modalità Anteprima 3D,tuttavia egli potrà scegliere
    'se continuare con l'operazione di Conversione oppure annullare e quindi selezionare una scala migliore
    If VScale <= 9 Then
        Dim RispostaContinuaConversione As VbMsgBoxResult
        If LinguaS = "Italiano" Then RispostaContinuaConversione = MsgBox("Per una buona visualizzazione della mappa 3d in modalità Anteprima 3D e consigliabile effettuare una conversione con una scalatura maggiore o uguale a 10! " + Chr(13) + "Vuoi comunque convertire la mappa attuale in scala 1 :" + Str(VScale) + "? ", vbYesNo, "Conversione mappa 3D")
        If LinguaS = "Inglese" Then RispostaContinuaConversione = MsgBox("For a better view of the 3D map in 3D preview modality, you must do a conversion with scale major or equal than 0!" + Chr(13) + "You want aniway convert the map with 1 :" + Str(VScale) + " scale? ", vbYesNo, "3D map conversion")
        If RispostaContinuaConversione = vbNo Then GoTo Errore
    End If
    'Richiamo il metodo dalla Classe Sonora che riprodurrà il terzo file sonoro di menù
    Suoni.Esegui_suono_menù_3
    'Se è stata scelta la voce di menù Converti Mappa in 3D allora si mostrerà la finestra di
    'dialogo di scelta del file in cui verrà convertita in 3D la mappa corrente
    If CM3D = True Then
        'Definisco il titolo dell'oggetto operazioni.In tal modo l'utente capira quale operazione si
        'stà effettuando
        Operazioni.DialogTitle = "Conversione mappa 3D"
        'Impongo un filtro all'oggetto operazioni,in modo che si potrà convertire la mappa corrente
        'salvandola con un'estensione da me definita
        Operazioni.Filter = "Mappa 3D (*.Map3D)|*.Map3D"
        On Error GoTo Errore:
        'Apro la finestra di dialogo che mi permetterà di scegliere la directory e il nome di
        'salvataggio della mappa corrente
        Operazioni.ShowSave
        'Salvo il file selezionato all'interno della variabile FileSalvato
        FileConvertito = Operazioni.FileName
    End If
    'Elimino dal computer il file appena selezionato.
    'Ho adottato questo metodo perchè,qualora il file selezionato fosse già esistente
    'ne sarebbe rimasta traccia dei muri creti nella vecchia versione del file, invece
    'in questo modo,il file viene prima eliminato e poi ricreato,in modo da ricreare i muri
    'partendo da zero
    '(In caso il file non esistesse si procede normalmente senza eliminare il file)
    On Error Resume Next
    Kill FileConvertito
    'Attivo la voce di menù AvviaAnteprima.Questo mi permetterà,una volta convertita la
    'mappa corrente in 3D,di poter girare liberamente in modo assolutamente tridimensionale all'interno
    'dell'ambiente appena creato
    AvviaAnteprima.Enabled = True
    'Apro il file appena selezionato
    Open FileConvertito For Random As #1 Len = Len(Muro)
        'Avvio un ciclo For che salverà tutti i valori delle righe create,in un altra
        'variabile che conterrà i valori modificati secondo lo scale selezionato.
        'Questa operazione è importantissima,in quanto se avremmo modificato direttamente i
        'valori delle linee,dopo la conversione,queste non sarebbero più vivibili sullo schermo,in
        'quanto avrebbero valori molto grandi,che supererebbero la superficie dell'editor
        For I = 1 To Max
            '---------------------------------------------------------------------------------------
            'Operazione di ripristino delle righe, nel caso fossero state modificate dall'operazione
            'di Zoom
            '---------------------------------------------------------------------------------------
            Riga(I).X1 = Fix(Riga(I).X1 / Molt)
            Riga(I).X2 = Fix(Riga(I).X2 / Molt)
            Riga(I).Y1 = Fix(Riga(I).Y1 / Molt)
            Riga(I).Y2 = Fix(Riga(I).Y2 / Molt)
            '------------------------------------
            'Operazione di modifica
            '------------------------------------
            Muro.X1 = Riga(I).X1 * VScale
            Muro.X2 = Riga(I).X2 * VScale
            Muro.Y1 = -Riga(I).Y1 * VScale
            Muro.Y2 = -Riga(I).Y2 * VScale
            Muro.Altezza = Riga(I).Altezza
            Muro.Altitudine = Riga(I).Altitudine
            Muro.Texture = RTrim(Riga(I).Texture)
            Muro.Proprietà = RTrim(Riga(I).Proprietà)
            Muro.Nome = RTrim(Riga(I).Nome)
            Muro.NMattonelleALtezza = Riga(I).NMattonelleALtezza
            Muro.NMAttonelleLarghezza = Riga(I).NMAttonelleLarghezza
            '------------------------------------
            'Operazione di scrittura del record
            '------------------------------------
            Put #1, , Muro
            '--------------------------------------------------------------------------
            'Operazione di ripristino Zoom (Mediante le istruzioni che seguono,dopo
            'il salvataggio il programma tornerà automaticamente allo Zoom selezionato
            'reimpostando nuovamente tutte le righe della mappa attuale
            '--------------------------------------------------------------------------
            Riga(I).X1 = Riga(I).X1 * Molt
            Riga(I).X2 = Riga(I).X2 * Molt
            Riga(I).Y1 = Riga(I).Y1 * Molt
            Riga(I).Y2 = Riga(I).Y2 * Molt
        Next
    Close #1
    Open "SoP" + FileConvertito For Random As #1 Len = Len(SofPav)
    For J = 0 To Max2
        '---------------------------------------------------------------------------------------
        'Operazione di ripristino dei pavimenti, nel caso fossero state modificate dall'operazione
        'di Zoom
        '---------------------------------------------------------------------------------------
        SoP(J).CR.X1 = Fix(SoP(J).CR.X1 / Molt)
        SoP(J).CR.X2 = Fix(SoP(J).CR.X2 / Molt)
        SoP(J).X3 = Fix(SoP(J).X3 / Molt)
        SoP(J).X4 = Fix(SoP(J).X4 / Molt)
        SoP(J).CR.Y1 = Fix(SoP(J).CR.Y1 / Molt)
        SoP(J).CR.Y2 = Fix(SoP(J).CR.Y2 / Molt)
        SoP(J).Y3 = Fix(SoP(J).Y3 / Molt)
        SoP(J).Y4 = Fix(SoP(J).Y4 / Molt)
        '------------------------------------
        'Operazione di modifica
        '------------------------------------
        SofPav.CR.X1 = SoP(J).CR.X1 * VScale
        SofPav.CR.X2 = SoP(J).CR.X2 * VScale
        SofPav.X3 = SoP(J).X3 * VScale
        SofPav.X4 = SoP(J).X4 * VScale
        SofPav.CR.Y1 = -SoP(J).CR.Y1 * VScale
        SofPav.CR.Y2 = -SoP(J).CR.Y2 * VScale
        SofPav.Y3 = -SoP(J).Y3 * VScale
        SofPav.Y4 = -SoP(J).Y4 * VScale
        SofPav.Tipo = SoP(J).Tipo
        SofPav.CR.Altitudine = SoP(J).CR.Altitudine
        SofPav.CR.Texture = SoP(J).CR.Texture
        SofPav.CR.NMattonelleALtezza = SoP(J).CR.NMattonelleALtezza
        SofPav.CR.NMAttonelleLarghezza = SoP(J).CR.NMAttonelleLarghezza
        SofPav.CR.Nome = SoP(J).CR.Nome
        SofPav.CR.SpigoloF = True
        SofPav.CR.SpigoloI = True
        SofPav.CR.Altezza = 0
        '------------------------------------
        'Operazione di scrittura del record
        '------------------------------------
        Put #1, , SofPav
        '--------------------------------------------------------------------------
        'Operazione di ripristino Zoom (Mediante le istruzioni che seguono,dopo
        'il salvataggio il programma tornerà automaticamente allo Zoom selezionato
        'reimpostando nuovamente tutte le righe della mappa attuale
        '--------------------------------------------------------------------------
        SoP(J).CR.X1 = SoP(J).CR.X1 * Molt
        SoP(J).CR.X2 = SoP(J).CR.X2 * Molt
        SoP(J).X3 = SoP(J).X3 * Molt
        SoP(J).X4 = SoP(J).X4 * Molt
        SoP(J).CR.Y1 = SoP(J).CR.Y1 * Molt
        SoP(J).CR.Y2 = SoP(J).CR.Y2 * Molt
        SoP(J).Y3 = SoP(J).Y3 * Molt
        SoP(J).Y4 = SoP(J).Y4 * Molt
    Next
    Close #1
    'Esco dalla funzione
    Exit Sub
'Label in cui si sarà indirizzati nel caso in cui verrà scelto di annullare
'l'operazione di conversione del file
Errore:
End Sub


Sub Scorri_Mappa(VecchioXMouse As Long, VecchioYMouse As Long)
    'Dichiaro due variabili che mi serviranno per mantenere rispettivamente i due nuovi
    'valori del mouse dell'editor (NuovoXMouse per le nuova coordinata X del mouse,
    'NuovoYMouse per la nuova coordinata y del mouse
    Dim NuovoXMouse As Long
    Dim NuovoYMouse As Long
    'Dichiaro una variabile che mi servirà come indice nell'operazione di spostamento degli
    'oggetti caricati all'interno della mappa 3D
    Dim NOggetto As Integer
    'Ora dichiaro quattro variabili,le quali mi serviranno per capire in che direzione è stato effettuato
    'lo scrolling della mappa
    Dim Su As Boolean
    Dim Giù As Boolean
    Dim Sinistra As Boolean
    Dim Destra As Boolean
    'Quest'altra variabile mi servirà per capire se è stato premuto il bottone sinistro del mouse
    Dim TmpB1 As Integer
        'Prendo in imput lo stato del mouse e ne salvo tutti i valori nelle rispettive coordinate
        Comandi.GetAbsMouseState NuovoXMouse, NuovoYMouse, TmpB1
        'Verifico se è stato premuto il pulsante sinistro del mouse e in caso positivo...
        If TmpB1 <> 0 Then
            'Avvio un ciclo for che risetterà tutti i valori dele righe a seconda
            'se il valore della nuova coordinata X e Y del mouse sono maggiori o minori
            'dei loro vecchi valori
            For I = 0 To Max
                'Se la nuova coordinata X del mouse è minore del suo vecchio valore,allora
                'diminuisco a tutte le righe le coordinate X di 3
                If NuovoXMouse <> VecchioXMouse Then
                    If NuovoXMouse < VecchioXMouse Then
                        Riga(I).X1 = Riga(I).X1 - 3
                        Riga(I).X2 = Riga(I).X2 - 3
                        Sinistra = True
                'Se la nuova coordinata X del mouse è maggiore del suo vecchio valore,allora
                'aumento a tutte le righe le coordinate X di 3
                    ElseIf NuovoXMouse > VecchioXMouse Then
                        Riga(I).X1 = Riga(I).X1 + 3
                        Riga(I).X2 = Riga(I).X2 + 3
                        Destra = True
                    End If
                End If
                'Se la nuova coordinata Y del mouse è minore del suo vecchio valore,allora
                'diminuisco a tutte le righe le coordinate Y di 3
                If NuovoYMouse <> VecchioYMouse Then
                    If NuovoYMouse < VecchioYMouse Then
                        Riga(I).Y1 = Riga(I).Y1 - 3
                        Riga(I).Y2 = Riga(I).Y2 - 3
                        Su = True
                'Se la nuova coordinata Y del mouse è maggiore del suo vecchio valore,allora
                'aumento a tutte le righe le coordinate Y di 3
                    ElseIf NuovoYMouse > VecchioYMouse Then
                        Riga(I).Y1 = Riga(I).Y1 + 3
                        Riga(I).Y2 = Riga(I).Y2 + 3
                        Giù = True
                    End If
                End If
            'Incremento la variabile I
            Next
            'Ora,mediante i valori assunti dalle quattro variabili boleane di direzione,posso
            'effettuare anche uno scrolling della posizione della telecamera
            If Su = True Then
                PosizioneTelecamera.Z = PosizioneTelecamera.Z - 3
                'Ora sposto tutti gli oggetti caricati all'interno della mappa 3D indietro sull'asse delle Z
                For NOggetto = 0 To IOg
                    Oggetto(NOggetto).Setta_Posizione "Z", Oggetto(NOggetto).Ricava_Posizione.Z - 3
                Next
            ElseIf Giù = True Then
                'Ora sposto tutti gli oggetti caricati all'interno della mappa 3D avanti sull'asse delle Z
                PosizioneTelecamera.Z = PosizioneTelecamera.Z + 3
                For NOggetto = 0 To IOg
                    Oggetto(NOggetto).Setta_Posizione "Z", Oggetto(NOggetto).Ricava_Posizione.Z + 3
                Next
            End If
            If Sinistra = True Then
                'Ora sposto tutti gli oggetti caricati all'interno della mappa 3D indietro sull'asse delle X
                PosizioneTelecamera.X = PosizioneTelecamera.X - 3
                For NOggetto = 0 To IOg
                    Oggetto(NOggetto).Setta_Posizione "X", Oggetto(NOggetto).Ricava_Posizione.X - 3
                Next
            ElseIf Destra = True Then
                'Ora sposto tutti gli oggetti caricati all'interno della mappa 3D avanti sull'asse delle X
                PosizioneTelecamera.X = PosizioneTelecamera.X + 3
                For NOggetto = 0 To IOg
                    Oggetto(NOggetto).Setta_Posizione "X", Oggetto(NOggetto).Ricava_Posizione.X + 3
                Next
            End If
            Form_Opzioni.TelecameraX = Str(PosizioneTelecamera.X)
            Form_Opzioni.TelecameraZ = Str(PosizioneTelecamera.Z)
            For J = 0 To Max2
                If NuovoXMouse <> VecchioXMouse Then
                    If NuovoXMouse < VecchioXMouse Then
                        SoP(J).CR.X1 = SoP(J).CR.X1 - 3
                        SoP(J).CR.X2 = SoP(J).CR.X2 - 3
                        SoP(J).X3 = SoP(J).X3 - 3
                        SoP(J).X4 = SoP(J).X4 - 3
                    ElseIf NuovoXMouse > VecchioXMouse Then
                        SoP(J).CR.X1 = SoP(J).CR.X1 + 3
                        SoP(J).CR.X2 = SoP(J).CR.X2 + 3
                        SoP(J).X3 = SoP(J).X3 + 3
                        SoP(J).X4 = SoP(J).X4 + 3
                    End If
                End If
                If NuovoYMouse <> VecchioYMouse Then
                    If NuovoYMouse < VecchioYMouse Then
                        SoP(J).CR.Y1 = SoP(J).CR.Y1 - 3
                        SoP(J).CR.Y2 = SoP(J).CR.Y2 - 3
                        SoP(J).Y3 = SoP(J).Y3 - 3
                        SoP(J).Y4 = SoP(J).Y4 - 3
                    ElseIf NuovoYMouse > VecchioYMouse Then
                        SoP(J).CR.Y1 = SoP(J).CR.Y1 + 3
                        SoP(J).CR.Y2 = SoP(J).CR.Y2 + 3
                        SoP(J).Y3 = SoP(J).Y3 + 3
                        SoP(J).Y4 = SoP(J).Y4 + 3
                    End If
                End If
            Next
            'Aggiorno le vecchie coordinate del mouse
            Comandi.GetAbsMouseState MouseX, MouseY
        End If
End Sub

Function Problemi_righe(Indice As Integer, Oggetto As String) As Boolean
    'Dichiaro quattro variabili temporanee e ognuna di questa conterrà momentaneamente il valore della rispettiva
    'coordinata della linea analizzata.
    '------------------------------------------------------------------------------------------------------------
    'Tmpx1 conterrà il valore temporaneo di Riga(Indice).X1 oppure SoP(Indice).CR.X1
    Dim TmpX1 As Single
    'Tmpx2 conterrà il valore temporaneo di Riga(Indice).X2 oppure SoP(Indice).CR.X2
    Dim TmpX2 As Single
    'TmpX3 potrà contenere solamente il valore di SoP(Indice).X3
    Dim TmpX3 As Single
    'TmpX4 potrà contenere solamente il valore di SoP(Indice).X4
    Dim TmpX4 As Single
    'TmpY1 conterrà il valore temporaneo di Riga(Indice).Y1 oppure SoP(Indice).Y1
    Dim TmpY1 As Single
    'TmpY2 conterrà il valore temporaneo di Riga(Indice).Y2 oppure SoP(Indice).Y2
    Dim TmpY2 As Single
    'TmpY3 potrà contenere solamente il valore di SoP(Indice).Y3
    Dim TmpY3 As Single
    'TmpY4 potrà contenere solamente il valore di SoP(Indice).Y4
    Dim TmpY4 As Single
    '--------------------------------------------------------------------------------------------------
    'Dichiarate le variabili procedo con il passo successivo:
    'Se la variabile Oggetto passata nella chiamata della funzione ha il valore Muro,
    'assegno a tutte le variabili temporanee il valore della rispettiva coordinata di riga.
    'Questo mi serve in modo che se le coordinate della riga corrente non escono dalla superficie
    'dell'editor, ovvero "non creeranno problemi",queste manterranno il loro valore originale,senza
    'essere modificate dai limiti dell'editor
    '--------------------------------------------------------------------------------------------------
    If Oggetto = "Muro" Then
        TmpX1 = Riga(Indice).X1
        TmpX2 = Riga(Indice).X2
        TmpY1 = Riga(Indice).Y1
        TmpY2 = Riga(Indice).Y2
        'Se qualsiasi coordinata del muro che stiamo analizzando assume un valore che va al di fuori
        'della superficie dell'editor, tipo riga(I).x1 < 1 che è il primo punto orizzontale dell'editor, allora la funzione
        'assumenra il valore booleano True, che stà ad indicare che sono stati riscontrati dei problemi
        'nella creazione della riga su schermo.
        'Inoltre verranno avviate tutte quelle condizioni che permetteranno di risolvere tutti questi problemi.
        If Riga(Indice).X1 < 1 Or Riga(Indice).X1 > Larghezza - 1 Or Riga(Indice).X2 < 1 Or Riga(Indice).X2 > Larghezza - 1 Or Riga(Indice).Y1 < 1 Or Riga(Indice).Y1 > Altezza - 1 Or Riga(Indice).Y2 < 1 Or Riga(Indice).Y2 > Altezza - 1 Then
            'Se la coordinata X1 del muro corrente, assumerà un valore < 1, ovvero minore del limite sinistro
            'dell'editor, allora la rispettiva variabile temporanea assumera il valore 1
            If Riga(Indice).X1 < 1 Then TmpX1 = 1
            'Se la coordinata X1 del muro corrente, assumerà un valore >= Larghezza -1, ovvero maggiore del limite destro
            'dell'editor, allora la rispettiva variabile temporanea assumera il valore Larghezza -1
            If Riga(Indice).X1 >= Larghezza Then TmpX1 = Larghezza - 1
            'Se la coordinata X2 del muro corrente, assumerà un valore < 1, ovvero minore del limite sinistro
            'dell'editor, allora la rispettiva variabile temporanea assumera il valore 1
            If Riga(Indice).X2 < 1 Then TmpX2 = 1
            'Se la coordinata X2 del muro corrente, assumerà un valore < 1, ovvero minore del limite sinistro
            'dell'editor, allora la rispettiva variabile temporanea assumera il valore 1
            If Riga(Indice).X2 >= Larghezza Then TmpX2 = Larghezza - 1
            'Se la coordinata Y1 del muro corrente, assumerà un valore < 1, ovvero minore del limite alto
            'dell'editor, allora la rispettiva variabile temporanea assumera il valore 1
            If Riga(Indice).Y1 < 1 Then TmpY1 = 1
            'Se la coordinata Y1 del muro corrente, assumerà un valore >= Altezza - 1, ovvero minore del limite basso
            'dell'editor, allora la rispettiva variabile temporanea assumera il valore Altezza - 1
            If Riga(Indice).Y1 >= Altezza Then TmpY1 = Altezza - 1
            'Se la coordinata Y2 del muro corrente, assumerà un valore < 1, ovvero minore del limite alto
            'dell'editor, allora la rispettiva variabile temporanea assumera il valore 1
            If Riga(Indice).Y2 < 1 Then TmpY2 = 1
            'Se la coordinata Y2 del muro corrente, assumerà un valore >= Altezza - 1, ovvero minore del limite basso
            'dell'editor, allora la rispettiva variabile temporanea assumera il valore Altezza - 1
            If Riga(Indice).Y2 >= Altezza Then TmpY2 = Altezza - 1
            'Ora il programma è pronto a disegnare del muro corrente con tutti i valori contenuti nelle variabili
            'temporanee.
            'PS: Se il valore di Indice è diverso da quello di IndiceLista allora la riga corente verrà disegnata con il suo
            'normalissimo colore (Bianco), altrimenti verrà disegnata con il colore Giallo.
            'Spieghiamo meglio la funzione di questo controllo:
            'Se l'indice di riga è uguale al muro selezionato dall'ellenco dei muri esistenti presente nel
            'form delle opzioni allora disegna la riga corrente in giallo,in modo da far capire che questa
            'riga corrisponde a quella su cui dover effettuare eventuali modifiche come cambiare la Textures,
            'o l'altezza,o l'altitudine,etc.
            If Indice <> IndiceLista Then
                Schermo.DrawLine TmpX1, TmpY1, TmpX2, TmpY2, RGBA(CM.R, CM.G, CM.B, 1)
            Else
                Schermo.DrawLine TmpX1, TmpY1, TmpX2, TmpY2, RGBA(CMS.R, CMS.G, CMS.B, 1)
            End If
            'Assegnamo alla funzione il valore booleano True in modo da far capire al programma che sono stati riscontrati
            'dei problemi nella creazione della riga corrente
            Problemi_righe = True
            'Qui diciamo che finisce la funzione vera e proprio, le istruzioni sottostanti servono a modificare
            'due variabili globali che sarebbero le coordinate del mouse
    End If
    '--------------------------------------------------------------------------------------------------
    'Se la variabile Oggetto passata nella chiamata della funzione ha il valore "SoP",
    'assegno a tutte le variabili temporanee il valore della rispettiva coordinata di SoP.
    'Questo mi serve in modo che se le coordinate della riga corrente non escono dalla superficie
    'dell'editor, ovvero "non creeranno problemi",queste manterranno il loro valore originale,senza
    'essere modificate dai limiti dell'editor,insomma,svolgo la stessa funzione eseguita per tutte le
    'coordinate di riga,solo che questa volta verranno "messe sotto torchio" le coordinate di SoP
    '--------------------------------------------------------------------------------------------------
    ElseIf Oggetto = "SoP" Then
        TmpX1 = SoP(Indice).CR.X1
        TmpX2 = SoP(Indice).CR.X2
        TmpX3 = SoP(Indice).X3
        TmpX4 = SoP(Indice).X4
        TmpY1 = SoP(Indice).CR.Y1
        TmpY2 = SoP(Indice).CR.Y2
        TmpY3 = SoP(Indice).Y3
        TmpY4 = SoP(Indice).Y4
        'Se qualsiasi coordinata del pavimento o soffitto che stiamo analizzando assume un valore che va al di fuori
        'della superficie dell'editor, tipo sop(J).CR.X1 < 1 che è il primo punto orizzontale dell'editor, allora la funzione
        'assumenra il valore booleano True, che stà ad indicare che sono stati riscontrati dei problemi
        'nella creazione della riga su schermo.
        'Inoltre verranno avviate tutte quelle condizioni che permetteranno di risolvere tutti questi problemi.
        If SoP(Indice).CR.X1 < 1 Or SoP(Indice).CR.X1 > Larghezza - 1 Or SoP(Indice).CR.X2 < 1 Or SoP(Indice).CR.X2 > Larghezza - 1 Or SoP(Indice).X3 < 1 Or SoP(Indice).X3 > Larghezza - 1 Or SoP(Indice).X4 < 1 Or SoP(Indice).X4 > Larghezza - 1 Or SoP(Indice).CR.Y1 < 1 Or SoP(Indice).CR.Y1 > Altezza - 1 Or SoP(Indice).CR.Y2 < 1 Or SoP(Indice).CR.Y2 > Altezza - 1 Or SoP(Indice).Y3 < 1 Or SoP(Indice).Y3 > Altezza - 1 Or SoP(Indice).Y4 < 1 Or SoP(Indice).Y4 > Altezza - 1 Then
            'Se la coordinata X1 del pavimento o soffitto corrente, assumerà un valore < 1, ovvero minore del limite sinistro
            'dell'editor, allora la rispettiva variabile temporanea assumera il valore 1
            If SoP(Indice).CR.X1 < 1 Then TmpX1 = 1
            'Se la coordinata X1 del pavimento o soffitto corrente, assumerà un valore > Larghezza - 1, ovvero maggiore del limite destro
            'dell'editor, allora la rispettiva variabile temporanea assumera il valore Larghezza -1
            If SoP(Indice).CR.X1 > Larghezza - 1 Then TmpX1 = Larghezza - 1
            'Se la coordinata X2 del pavimento o soffitto corrente, assumerà un valore < 1, ovvero minore del limite sinistro
            'dell'editor, allora la rispettiva variabile temporanea assumera il valore 1
            If SoP(Indice).CR.X2 < 1 Then TmpX2 = 1
            'Se la coordinata X2 del pavimento o soffitto corrente, assumerà un valore > Larghezza - 1, ovvero maggiore del limite destro
            'dell'editor, allora la rispettiva variabile temporanea assumera il valore Larghezza -1
            If SoP(Indice).CR.X2 > Larghezza - 1 Then TmpX2 = Larghezza - 1
            'Se la coordinata X3 del pavimento o soffitto corrente, assumerà un valore < 1, ovvero minore del limite sinistro
            'dell'editor, allora la rispettiva variabile temporanea assumera il valore 1
            If SoP(Indice).X3 < 1 Then TmpX3 = 1
            'Se la coordinata X3 del pavimento o soffitto corrente, assumerà un valore > Larghezza - 1, ovvero maggiore del limite destro
            'dell'editor, allora la rispettiva variabile temporanea assumera il valore Larghezza -1
            If SoP(Indice).X3 > Larghezza - 1 Then TmpX3 = Larghezza - 1
            'Se la coordinata X4 del pavimento o soffitto corrente, assumerà un valore < 1, ovvero minore del limite sinistro
            'dell'editor, allora la rispettiva variabile temporanea assumera il valore 1
            If SoP(Indice).X4 < 1 Then TmpX4 = 1
            'Se la coordinata X4 del pavimento o soffitto corrente, assumerà un valore > Larghezza - 1, ovvero maggiore del limite destro
            'dell'editor, allora la rispettiva variabile temporanea assumera il valore Larghezza -1
            If SoP(Indice).X4 > Larghezza - 1 Then TmpX4 = Larghezza - 1
            'Se la coordinata Y1 del pavimento o soffitto corrente, assumerà un valore < 1, ovvero minore del limite alto
            'dell'editor, allora la rispettiva variabile temporanea assumera il valore 1
            If SoP(Indice).CR.Y1 < 1 Then TmpY1 = 1
            'Se la coordinata Y1 del pavimento o soffitto corrente, assumerà un valore > Altezza - 1, ovvero maggiore del limite basso
            'dell'editor, allora la rispettiva variabile temporanea assumera il valore Altezza -1
            If SoP(Indice).CR.Y1 > Altezza - 1 Then TmpY1 = Altezza - 1
            'Se la coordinata Y2 del pavimento o soffitto corrente, assumerà un valore < 1, ovvero minore del limite alto
            'dell'editor, allora la rispettiva variabile temporanea assumera il valore 1
            If SoP(Indice).CR.Y2 < 1 Then TmpY2 = 1
            'Se la coordinata Y2 del pavimento o soffitto corrente, assumerà un valore > Altezza - 1, ovvero maggiore del limite basso
            'dell'editor, allora la rispettiva variabile temporanea assumera il valore Altezza -1
            If SoP(Indice).CR.Y2 > Altezza - 1 Then TmpY2 = Altezza - 1
            'Se la coordinata Y3 del pavimento o soffitto corrente, assumerà un valore < 1, ovvero minore del limite alto
            'dell'editor, allora la rispettiva variabile temporanea assumera il valore 1
            If SoP(Indice).Y3 < 1 Then TmpY3 = 1
            'Se la coordinata Y3 del pavimento o soffitto corrente, assumerà un valore > Altezza - 1, ovvero maggiore del limite basso
            'dell'editor, allora la rispettiva variabile temporanea assumera il valore Altezza -1
            If SoP(Indice).Y3 > Altezza - 1 Then TmpY3 = Altezza - 1
            'Se la coordinata Y4 del pavimento o soffitto corrente, assumerà un valore < 1, ovvero minore del limite alto
            'dell'editor, allora la rispettiva variabile temporanea assumera il valore 1
            If SoP(Indice).Y4 < 1 Then TmpY4 = 1
            'Se la coordinata Y4 del pavimento o soffitto corrente, assumerà un valore > Altezza - 1, ovvero maggiore del limite basso
            'dell'editor, allora la rispettiva variabile temporanea assumera il valore Altezza -1
            If SoP(Indice).Y4 > Altezza - 1 Then TmpY4 = Altezza - 1
            'Se l'elemento che si stà analizzando è un pavimento,allora setto la variabile colore con i valori
            'necessari a formare appunto un colore Blu
            If SoP(Indice).Tipo = "Pavimento" Then
                With Colore
                    .R = CP.R
                    .G = CP.G
                    .B = CP.B
                    .A = 1
                End With
            'Se invece l'elemento che si stà analizzando è un soffitto,allora setterò la variabile colore con i valori
            'necessari a formare un colore Verde
            ElseIf SoP(Indice).Tipo = "Soffitto" Then
                With Colore
                    .R = CS.R
                    .G = CS.G
                    .B = CS.B
                    .A = 1
                End With
            End If
            'Se il valore (l'indice) passato alla funzione stessa è uguale a l'elemento selezionato dal Form_Opzioni
            'dall'oggetto ComboBox ElencoSop,allora setto la variabile Colore con i valori necessari a formare
            'un colore Rossaceo
            If Indice = IndiceLista2 Then
                With Colore
                    .R = CSOPS.R
                    .G = CSOPS.G
                    .B = CSOPS.B
                    .A = 1
                End With
            End If
            'Ora si è pronti a disegnare la quattro righe che formano il pavimento o soffitto analizzato
            'con tutte le coordinate temporanee e i colori appena settati
            Schermo.DrawLine TmpX1, TmpY1, TmpX2, TmpY2, RGBA(Colore.R, Colore.G, Colore.B, Colore.A)
            Schermo.DrawLine TmpX2, TmpY2, TmpX4, TmpY4, RGBA(Colore.R, Colore.G, Colore.B, Colore.A)
            Schermo.DrawLine TmpX4, TmpY4, TmpX3, TmpY3, RGBA(Colore.R, Colore.G, Colore.B, Colore.A)
            Schermo.DrawLine TmpX3, TmpY3, TmpX1, TmpY1, RGBA(Colore.R, Colore.G, Colore.B, Colore.A)
            'La funzione assume il valore boleano True,in modo da non fare ridisegnare le linee che compongono
            'questo pavimento o soffitto nella funzione Disegna_Righe
            Problemi_righe = True
        End If
    End If
End Function

Sub Problemi_coordinate_mouse()
    'Se la coordinata SmouseX, assumerà un valore < 1, ovvero minore del limite sinistro
    'dell'editor, allora imporrò alla stessa il valore uguale al limite sinistro dell'editor
    If SmouseX < 1 Then SmouseX = 1
    'Se la coordinata SmouseX, assumerà un valore >= Larghezza+1, ovvero maggiore del limite destro
    'dell'editor, allora imporrò alla stessa il valore uguale al limite destro dell'editor
    If SmouseX >= Larghezza - 1 Then SmouseX = Larghezza - 1
    'Se la coordinata SmouseY, assumerà un valore < 1, ovvero minore del limite alto
    'dell'editor, allora imporrò alla stessa il valore uguale al limite alto dell'editor
    If SmouseY < 1 Then SmouseY = 1
    'Se la coordinata SmouseY, assumerà un valore >= Altezza -1, ovvero maggiore del limite basso
    'dell'editor, allora imporrò alla stessa il valore uguale al limite basso dell'editor
    If SmouseY >= Altezza - 1 Then SmouseY = Altezza - 1
End Sub

Function Problemi_Risoluzione() As Boolean
    'Tramite questi piccoli passaggi è possibile ricavarsi la risoluzione corrente del video.
    'Tramite l'operazione Screen.Width / Screen.TwipsPerPixelX viene divisa la dimensione in larghezza dello
    'schermo per i pixel presenti in larghezza (Si ricaverà così la risoluzione appunto in larghezza dello schermo),
    'mentre tramite l'operazione Screen.Height / Screen.TwipsPerPixelY viene effettuato lo stesso procedimento, solo che
    'questa volta verrà effettuata per la risoluzione in altezza dello schermo, ottenendo così la rispettiva risoluzione
    'in altezza.
    'Se la risoluzione in larghezza è diversa da 1024 e quella in altezza è diversa da 768 allora la funzione restituirà
    'il valore booleano True che starà ad indicare che sono stati rilevati problemi nell'impostazione dello schermo.
    'Questo (come già descitto precedentemente) comporterà la visualizzazione di un messaggio al fine di avvisare l'utente
    'dell'errore e l'immediata uscita dal programma
    If Screen.Width / Screen.TwipsPerPixelX <> 1024 And Screen.Height / Screen.TwipsPerPixelY <> 768 Then
        Problemi_Risoluzione = True
    End If
End Function

Function Problemi_telecamera() As Boolean
    'Dichiaro la variabile di appoggio TmpX, la quale assumerà il valore temporaneo di posizionetelecamera.x
    Dim TmpX As Single
    'Dichiaro la variabile di appoggio Tmpz, la quale assumerà il valore temporaneo di posizionetelecamera.z
    Dim TmpZ As Single
    'Assegno alle due variabile temporanee,i rispettivi valori delle coordinate di telecamera
    With PosizioneTelecamera
        TmpX = .X
        TmpZ = .Z
    End With
    'Se una delle due coordinate,va al di fuori della superficie dell'editor,allora queste assumeranno
    'il valore del limite stesso
    If TmpX < 1 Or TmpX > Larghezza - 1 Or TmpZ < 1 Or TmpZ > Altezza - 1 Then
        'Se la coordinata X della telecamera è minore del limite sinistro dell'editor,
        'allora avverrà la sostituzione
        If TmpX < 1 Then TmpX = 1
        'Se la coordinata X della telecamera è maggiore del limite destro dell'editor,
        'allora avverrà la sostituzione
        If TmpX > Larghezza - 1 Then TmpX = Larghezza - 1
        'Se la coordinata Z della telecamera è minore del limite alto dell'editor,
        'allora avverrà la sostituzione
        If TmpZ < 1 Then TmpZ = 1
        'Se la coordinata z della telecamera è minore del limite basso dell'editor,
        'allora avverrà la sostituzione
        If TmpZ > Altezza - 1 Then TmpZ = Altezza - 1
        'Ora il programma è pronto a sisegnare la nuova linea secondo le coordinate reimpostate
        'della telecamera
        Schermo.DrawLine Larghezza - 112, 67, TmpX, TmpZ, RGBA(0, 0.8, 0.4, 1)
        'La funzione aasume il valore boolenao True,che starà a significare che sono stati riscontrati
        'dei problemi del disegnare la linea che segnala all'interno dell'editor,la posizione
        'della telecamera
        Problemi_telecamera = True
    End If
End Function

Sub Inizializza_colori()
    'Inizializzo il colore iniziale con cui verranno disegnate le linee che rappresenteranno i muri all'interno dell'editor
    With CM
        .R = 1
        .G = 1
        .B = 1
    End With
    'Imposto il colore del rispettivo pulsante presente all'interno del Form_Opzioni
    Form_Opzioni.ColoreMuri.BackColor = &HFFFFFF
    'Inizializzo il colore iniziale con cui verrà disegnata la linea che rappresenterà il muro
    'selezionato tramite l'oggetto Elenco_Muri presente nel Form_Opzioni
    With CMS
        .R = 173
        .G = 216
        .B = 238
    End With
    'Imposto il colore del rispettivo pulsante presente all'interno del Form_Opzioni
    Form_Opzioni.ColoreMuriSelezionati.BackColor = &HFFFF&
    'Inizializzo il colore iniziale con cui verranno disegnati i quadratini che rappresenteranno
    'gli spigoli,cioè i punti di intersezione dei muri all'interno dell'editor
    With CSM
        .R = 1
        .G = 0
        .B = 0
    End With
    'Imposto il colore del rispettivo pulsante presente all'interno del Form_Opzioni
    Form_Opzioni.ColoreSpigoliMuri.BackColor = &HFF&
    'Inizializzo il colore iniziale con cui verranno disegnate le linee che rappresenteranno i soffitti all'interno dell'editor
    With CS
        .R = 0
        .G = 1
        .B = 0
    End With
    'Imposto il colore del rispettivo pulsante presente all'interno del Form_Opzioni
    Form_Opzioni.ColoreSoffitti.BackColor = &HFF00&
    'Inizializzo il colore iniziale con cui verranno disegnate le linee che rappresenteranno i pavimenti all'interno dell'editor
    With CP
        .R = 0
        .G = 0
        .B = 1
    End With
    'Imposto il colore del rispettivo pulsante presente all'interno del Form_Opzioni
    Form_Opzioni.ColorePavimenti.BackColor = &HFF0000
    'Inizializzo il colore iniziale con cui verranno disegnati i quadratini che segnaleranno l'allineamento tra le coordinate del mouse
    'e i soffitti / pavimenti all'interno dell'editor
    With CASOP
        .R = 0
        .G = 1
        .B = 0
    End With
    'Imposto il colore del rispettivo pulsante presente all'interno del Form_Opzioni
    Form_Opzioni.ColoreAllineamentoSP.BackColor = &HFF00&
    'Inizializzo il colore iniziale con cui verranno disegnati i quadratini che segnaleranno l'allineamento tra le coordinate del mouse
    'e i muri all'interno dell'editor
    With CAM
        .R = 0
        .G = 0
        .B = 1
    End With
    'Imposto il colore del rispettivo pulsante presente all'interno del Form_Opzioni
    Form_Opzioni.ColoreAllineamentoMuri.BackColor = &HFF0000
    'Inizializzo il colore iniziale con cui verranno disegnate le quattro linee che rappresenteranno il pavimento / soffitto
    'selezionato all'interno dell'editor
    With CSOPS
        .R = 1
        .G = 0
        .B = 0
    End With
    'Imposto il colore del rispettivo pulsante presente all'interno del Form_Opzioni
    Form_Opzioni.ColoreSPSelezionati.BackColor = &HFF&
    'Inizializzo il colore di sfondo del menù che vaerrà creato nella parte superiore dell'editor
    With CSFM
        .R = 0.3
        .G = 0.4
        .B = 0.7
    End With
    'Imposto il colore del rispettivo pulsante presente all'interno del Form_Opzioni
    Form_Opzioni.ColoreSfondoMenù.BackColor = &HFF8080
    'Inizializzo il 1° colore con cui verrano create le scritte di informazione all'interno del menù
    'che verrà creato nella parte superiore dell'editor
    With C1M
        .R = 0
        .G = 1
        .B = 0
    End With
    'Imposto il colore del rispettivo pulsante presente all'interno del Form_Opzioni
    Form_Opzioni.Colore1Menù.BackColor = &HFF00&
    'Inizializzo il 2° colore con cui verrano create le scritte di informazione all'interno del menù
    'che verrà creato nella parte superiore dell'editor
    With C2M
        .R = 1
        .G = 1
        .B = 1
    End With
    'Imposto il colore del rispettivo pulsante presente all'interno del Form_Opzioni
    Form_Opzioni.Colore2Menù.BackColor = &HFFFFFF
    'Inizializzo il colore iniziale con cui verranno create le linee guida che faciliteranno l'allineamento
    'di muri,pavimenti o soffitti da parte dell'utente
    With CLG
        .R = 1
        .G = 1
        .B = 0.1
    End With
    'Imposto il colore del rispettivo pulsante presente all'interno del Form_Opzioni
    Form_Opzioni.ColoreLineeGuida.BackColor = &HFFFF&
End Sub

Sub Verifica_Selezione_Oggetto(CursoreX As Long, CursoreY As Long)
    'Dichiaro una variabile che mi tornerà utile nell'individuare,all'interno del ciclo For,
    'l'oggetto selezionato
    Dim NOggetto As Integer
    'La variabile Trovato,mi segnalerà invece,alla fine del ciclo For se è stato trovato
    'l'oggetto su cui è stato cliccato
    Dim Trovato As Boolean
    'Imposto l'oggetto collisione in modo che mi verificherà se è stato cliccato sulla superficie di un oggetto
    Set Collisione = Scena.MousePicking(CursoreX, CursoreY, TV_COLLIDE_MESH, TV_TESTTYPE_BOUNDINGBOX)
    'Se l'oggetto ha rilevato una collisione con il cursore del mouse è un oggetto,allora...
    If Collisione.IsCollision Then
        'Inizio a ricercare l'oggetto su cui si è cliccato
        For NOggetto = 0 To IOg
            'Se lo cordinate di impatto del mouse coincidono con quelle dell'oggetto preso attualmente in
            'considerazione dal ciclo For,allora...
            If Collisione.GetCollisionMesh.GetPosition.X = Oggetto(NOggetto).Scheletro.GetPosition.X And Collisione.GetCollisionMesh.GetPosition.Y = Oggetto(NOggetto).Scheletro.GetPosition.Y Then
                'Assegno alla variabile booleana Trovato il valore True,in modo da segnalare al programma stesso
                'chè è stato trovato l'oggetto richiesto
                Trovato = True
                'Salvo l'indice dell'oggetto selezionato all'interno della variabile pubblica
                'IndiceOggettoSelezionato
                IndiceOggettoSelezionato = NOggetto
                'Impongo la variabile NOggetto uguale a IOg in modo tale che si possa uscire
                'dal ciclo For
                NOggetto = IOg
            End If
            'Si passa ad analizzare l'oggetto successivo
        Next
    End If
    'Se la variabile Trovato possiede valore booleano uguale a True,allora...
    If Trovato = True Then
        If Collisione.GetCollisionMesh.GetMeshName <> "Struttura" And Collisione.GetCollisionMesh.GetMeshName <> "StrutturaMap" And Collisione.GetCollisionMesh.GetMeshName <> "Trasparenza" Then Suoni.Esegui_suono_menù
    End If
End Sub

Sub Disegna_Menù_Modifica_Oggetto()
    'Alla pressione del tasto L...
    If Comandi.IsKeyPressed(TV_KEY_L) = True Then
        'La variabile MenùAiuto assumerà valore booleano False,in modo da far capire
        'al programma che al prossimo ciclo di rendering non dovrà più eseguire la funzione
        'corrente
        MenùAiuto = False
    End If
    'Innanzitutto creo un rettangolo multicolore trasparente che sarà lo sfondo del menù che guiderà l'utente
    'nelle operazioni di modifica dell'oggetto selezionato
    Schermo.DrawFilledColorBox 10, 10, 210, Altezza - 10, RGBA(0, 0.4, 1, 0.8), RGBA(0.5, 0.2, 0.8, 0.8), RGBA(0, 0.8, 0.8, 0.8), RGBA(0.3, 0.9, 0.6, 0.8)
    'Ora inizio a scrivere sul rettangolo appena creato tutte le lettere dei tasti corrispondenti
    'che andranno premuti per spostare l'oggetto
    Schermo.DrawText "< Q >", 20, 40, RGBA(0, 1, 0.4, 1), "carattere_personalizzato2"
    Schermo.DrawText "< W >", 20, 60, RGBA(0, 1, 0.4, 1), "carattere_personalizzato2"
    Schermo.DrawText "< E >", 20, 80, RGBA(0, 1, 0.4, 1), "carattere_personalizzato2"
    Schermo.DrawText "< R >", 20, 100, RGBA(0, 1, 0.4, 1), "carattere_personalizzato2"
    Schermo.DrawText "< T >", 20, 120, RGBA(0, 1, 0.4, 1), "carattere_personalizzato2"
    Schermo.DrawText "< Y >", 20, 140, RGBA(0, 1, 0.4, 1), "carattere_personalizzato2"
    'Scrivo sul rettangolo appena creato tutte le lettere dei tasti corrispondenti
    'che andranno premuti per ridimensionare l'oggetto
    Schermo.DrawText "< A >", 20, 180, RGBA(0, 1, 0.4, 1), "carattere_personalizzato2"
    Schermo.DrawText "< S >", 20, 200, RGBA(0, 1, 0.4, 1), "carattere_personalizzato2"
    Schermo.DrawText "< D >", 20, 220, RGBA(0, 1, 0.4, 1), "carattere_personalizzato2"
    Schermo.DrawText "< F >", 20, 240, RGBA(0, 1, 0.4, 1), "carattere_personalizzato2"
    Schermo.DrawText "< G >", 20, 260, RGBA(0, 1, 0.4, 1), "carattere_personalizzato2"
    Schermo.DrawText "< H >", 20, 280, RGBA(0, 1, 0.4, 1), "carattere_personalizzato2"
    'Scrivo sul rettangolo appena creato tutte le lettere dei tasti corrispondenti
    'che andranno premuti per ruotare l'oggetto
    Schermo.DrawText "< Z >", 20, 320, RGBA(0, 1, 0.4, 1), "carattere_personalizzato2"
    Schermo.DrawText "< X >", 20, 340, RGBA(0, 1, 0.4, 1), "carattere_personalizzato2"
    Schermo.DrawText "< C >", 20, 360, RGBA(0, 1, 0.4, 1), "carattere_personalizzato2"
    Schermo.DrawText "< V >", 20, 380, RGBA(0, 1, 0.4, 1), "carattere_personalizzato2"
    Schermo.DrawText "< B >", 20, 400, RGBA(0, 1, 0.4, 1), "carattere_personalizzato2"
    Schermo.DrawText "< N >", 20, 420, RGBA(0, 1, 0.4, 1), "carattere_personalizzato2"
    'Schrivo il tasto che andrà premuto per chiudere il menù modifica oggetto
    Schermo.DrawText "< L >", 20, 460, RGBA(0, 1, 0.4, 1), "carattere_personalizzato2"
    If LinguaS = "Italiano" Then
        Schermo.DrawText "Ruota", 80, 300, RGBA(0.6, 0.9, 0.8, 1), "carattere_personalizzato2"
        Schermo.DrawText "Scala", 80, 160, RGBA(0.6, 0.9, 0.8, 1), "carattere_personalizzato2"
        Schermo.DrawText "Sposta", 80, 20, RGBA(0.6, 0.9, 0.8, 1), "carattere_personalizzato2"
        'Scrivo sul rettangolo appena creato tutte le indicazioni corrispondenti ai tasti
        'corrispondenti che quideranno l'utente nell'operazione di spostamento dell'oggetto
        Schermo.DrawText "Avanti sull'asse X", 50, 40, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        Schermo.DrawText "Indietro sull'asse X", 50, 60, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        Schermo.DrawText "Su sull'asse Y", 50, 80, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        Schermo.DrawText "Giu sull'asse Y", 50, 100, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        Schermo.DrawText "Avanti sull'asse Z", 50, 120, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        Schermo.DrawText "Indietro sull'asse Z", 50, 140, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        'Scrivo sul rettangolo appena creato tutte le indicazioni corrispondenti ai tasti
        'corrispondenti che quideranno l'utente nell'operazione di ridimensionamento dell'oggetto
        Schermo.DrawText "Aumenta larghezza", 50, 180, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        Schermo.DrawText "Diminuisci larghezza", 50, 200, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        Schermo.DrawText "Aumenta altezza", 50, 220, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        Schermo.DrawText "Diminuisci altezza", 50, 240, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        Schermo.DrawText "Aumenta spessore", 50, 260, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        Schermo.DrawText "Diminuisci spessore", 50, 280, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        'Scrivo sul rettangolo appena creato tutte le indicazioni corrispondenti ai tasti
        'corrispondenti che quideranno l'utente nell'operazione di rotazione dell'oggetto
        Schermo.DrawText "Orario sull'asse X", 50, 320, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        Schermo.DrawText "Antiorario sull'asse X", 50, 340, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        Schermo.DrawText "Orario sull'asse Y", 50, 360, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        Schermo.DrawText "Antiorario sull'asse Y", 50, 380, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        Schermo.DrawText "Orario sull'asse Z", 50, 400, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        Schermo.DrawText "Antiorario sull'asse Z", 50, 420, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        'Infine scrivo l'informazione corrispondente per il tasto L che servirà a chiudere il
        'menù di modifica oggetto e quindi a terminare la sua modifica
        Schermo.DrawText "Chiudi", 50, 460, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
    ElseIf LinguaS = "Inglese" Then
        Schermo.DrawText "Rotate", 80, 300, RGBA(0.6, 0.9, 0.8, 1), "carattere_personalizzato2"
        Schermo.DrawText "Scale", 80, 160, RGBA(0.6, 0.9, 0.8, 1), "carattere_personalizzato2"
        Schermo.DrawText "Move", 80, 20, RGBA(0.6, 0.9, 0.8, 1), "carattere_personalizzato2"
        'Scrivo sul rettangolo appena creato tutte le indicazioni corrispondenti ai tasti
        'corrispondenti che quideranno l'utente nell'operazione di spostamento dell'oggetto
        Schermo.DrawText "Forward on X axis", 50, 40, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        Schermo.DrawText "Backward on X axis", 50, 60, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        Schermo.DrawText "Up on Y axis", 50, 80, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        Schermo.DrawText "Down on Y axis", 50, 100, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        Schermo.DrawText "Forward on Z axis", 50, 120, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        Schermo.DrawText "Backward on Z axis", 50, 140, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        'Scrivo sul rettangolo appena creato tutte le indicazioni corrispondenti ai tasti
        'corrispondenti che quideranno l'utente nell'operazione di ridimensionamento dell'oggetto
        Schermo.DrawText "Increase width", 50, 180, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        Schermo.DrawText "Decrease width", 50, 200, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        Schermo.DrawText "Increase height", 50, 220, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        Schermo.DrawText "Decrease height", 50, 240, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        Schermo.DrawText "Increase thickness", 50, 260, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        Schermo.DrawText "Decrease thickness", 50, 280, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        'Scrivo sul rettangolo appena creato tutte le indicazioni corrispondenti ai tasti
        'corrispondenti che quideranno l'utente nell'operazione di rotazione dell'oggetto
        Schermo.DrawText "Clockwise on X axis", 50, 320, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        Schermo.DrawText "Anticlockwise on X axis", 50, 340, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        Schermo.DrawText "Clockwise on Y axis", 50, 360, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        Schermo.DrawText "Anticlockwise on Y axis", 50, 380, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        Schermo.DrawText "Clockwise on Z axis", 50, 400, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        Schermo.DrawText "Anticlockwise on Z axis", 50, 420, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
        'Infine scrivo l'informazione corrispondente per il tasto L che servirà a chiudere il
        'menù di modifica oggetto e quindi a terminare la sua modifica
        Schermo.DrawText "Close", 50, 460, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
    End If
End Sub

Sub Disegna_Box_Modifica_Oggetto()
    'Alla pressione del tasto P...
    If Comandi.IsKeyPressed(TV_KEY_P) = True Then
        'La variabile MenùAiuto assumerà valore booleano False,in modo da far capire
        'al programma che al prossimo ciclo di rendering non dovrà più eseguire la funzione
        'corrente
        MenùAiuto = True
    End If
    'Disegno un piccolo box verde sullo schermo
    Schermo.DrawFilledBox 10, 10, 150, 40, RGBA(0, 0.4, 1, 0.8)
    'Scrivo la lettera P sul box appena creato
    Schermo.DrawText "< P >", 20, 15, RGBA(0, 1, 0.4, 1), "carattere_personalizzato2"
    'Scrivo la descrizione dell'azione associata alla pressione del tasto P
    If LinguaS = "Italiano" Then Schermo.DrawText "Modifica oggetto", 50, 15, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
    If LinguaS = "Inglese" Then Schermo.DrawText "Object modify", 50, 15, RGBA(1, 1, 1, 1), "carattere_personalizzato2"
End Sub

Sub Ricevi_Modifiche_Oggetto()
    If Comandi.IsKeyPressed(TV_KEY_Q) = True Then
        Oggetto(IndiceOggettoSelezionato).Setta_Posizione "X", Oggetto(IndiceOggettoSelezionato).Ricava_Posizione.X + 5
    End If
    If Comandi.IsKeyPressed(TV_KEY_W) = True Then
        Oggetto(IndiceOggettoSelezionato).Setta_Posizione "X", Oggetto(IndiceOggettoSelezionato).Ricava_Posizione.X - 5
    End If
    If Comandi.IsKeyPressed(TV_KEY_E) = True Then
        Oggetto(IndiceOggettoSelezionato).Setta_Posizione "Y", Oggetto(IndiceOggettoSelezionato).Ricava_Posizione.Y + 5
    End If
    If Comandi.IsKeyPressed(TV_KEY_R) = True Then
        Oggetto(IndiceOggettoSelezionato).Setta_Posizione "Y", Oggetto(IndiceOggettoSelezionato).Ricava_Posizione.Y - 5
    End If
        If Comandi.IsKeyPressed(TV_KEY_T) = True Then
        Oggetto(IndiceOggettoSelezionato).Setta_Posizione "Z", Oggetto(IndiceOggettoSelezionato).Ricava_Posizione.Z + 5
    End If
    If Comandi.IsKeyPressed(TV_KEY_Y) = True Then
        Oggetto(IndiceOggettoSelezionato).Setta_Posizione "Z", Oggetto(IndiceOggettoSelezionato).Ricava_Posizione.Z - 5
    End If
    If Comandi.IsKeyPressed(TV_KEY_A) = True Then
        Oggetto(IndiceOggettoSelezionato).Setta_Dimensione "X", Oggetto(IndiceOggettoSelezionato).Ricava_Dimensione.X + 5
    End If
    If Comandi.IsKeyPressed(TV_KEY_S) = True Then
        Oggetto(IndiceOggettoSelezionato).Setta_Dimensione "X", Oggetto(IndiceOggettoSelezionato).Ricava_Dimensione.X - 5
    End If
    If Comandi.IsKeyPressed(TV_KEY_D) = True Then
        Oggetto(IndiceOggettoSelezionato).Setta_Dimensione "Y", Oggetto(IndiceOggettoSelezionato).Ricava_Dimensione.Y + 5
    End If
    If Comandi.IsKeyPressed(TV_KEY_F) = True Then
        Oggetto(IndiceOggettoSelezionato).Setta_Dimensione "Y", Oggetto(IndiceOggettoSelezionato).Ricava_Dimensione.Y - 5
    End If
        If Comandi.IsKeyPressed(TV_KEY_G) = True Then
        Oggetto(IndiceOggettoSelezionato).Setta_Dimensione "Z", Oggetto(IndiceOggettoSelezionato).Ricava_Dimensione.Z + 5
    End If
    If Comandi.IsKeyPressed(TV_KEY_H) = True Then
        Oggetto(IndiceOggettoSelezionato).Setta_Dimensione "Z", Oggetto(IndiceOggettoSelezionato).Ricava_Dimensione.Z - 5
    End If
        If Comandi.IsKeyPressed(TV_KEY_Z) = True Then
        Oggetto(IndiceOggettoSelezionato).Setta_Rotazione "X", Oggetto(IndiceOggettoSelezionato).Ricava_Rotazione.X + 5
    End If
    If Comandi.IsKeyPressed(TV_KEY_X) = True Then
        Oggetto(IndiceOggettoSelezionato).Setta_Rotazione "X", Oggetto(IndiceOggettoSelezionato).Ricava_Rotazione.X - 5
    End If
    If Comandi.IsKeyPressed(TV_KEY_C) = True Then
        Oggetto(IndiceOggettoSelezionato).Setta_Rotazione "Y", Oggetto(IndiceOggettoSelezionato).Ricava_Rotazione.Y + 5
    End If
    If Comandi.IsKeyPressed(TV_KEY_V) = True Then
        Oggetto(IndiceOggettoSelezionato).Setta_Rotazione "Y", Oggetto(IndiceOggettoSelezionato).Ricava_Rotazione.Y - 5
    End If
        If Comandi.IsKeyPressed(TV_KEY_B) = True Then
        Oggetto(IndiceOggettoSelezionato).Setta_Rotazione "Z", Oggetto(IndiceOggettoSelezionato).Ricava_Rotazione.Z + 5
    End If
    If Comandi.IsKeyPressed(TV_KEY_N) = True Then
        Oggetto(IndiceOggettoSelezionato).Setta_Rotazione "Z", Oggetto(IndiceOggettoSelezionato).Ricava_Rotazione.Z - 5
    End If
End Sub

Function Problemi_Oggetto() As Boolean
    'Dichiaro la variabile di appoggio TmpX, la quale assumerà il valore temporaneo di Oggetto(IndiceOggettoSelezionato).Ricava_posizione.X
    Dim TmpX As Single
    'Dichiaro la variabile di appoggio TmpZ, la quale assumerà il valore temporaneo di Oggetto(IndiceOggettoSelezionato).Ricava_posizione.Z
    Dim TmpZ As Single
    'Assegno alle due variabile temporanee,i rispettivi valori delle coordinate di oggetto
    With Oggetto(IndiceOggettoSelezionato).Ricava_Posizione
        TmpX = .X
        TmpZ = .Z
    End With
    'Se una delle due coordinate,va al di fuori della superficie dell'editor,allora queste assumeranno
    'il valore del limite stesso
    If TmpX < 1 Or TmpX > Larghezza - 1 Or TmpZ < 1 Or TmpZ > Altezza - 1 Then
        'Se la coordinata X dell'oggetto è minore del limite sinistro dell'editor,
        'allora avverrà la sostituzione
        If TmpX < 1 Then TmpX = 1
        'Se la coordinata X dell'oggetto è maggiore del limite destro dell'editor,
        'allora avverrà la sostituzione
        If TmpX > Larghezza - 1 Then TmpX = Larghezza - 1
        'Se la coordinata Z dell'oggetto è minore del limite alto dell'editor,
        'allora avverrà la sostituzione
        If TmpZ < 1 Then TmpZ = 1
        'Se la coordinata z dell'oggetto è minore del limite basso dell'editor,
        'allora avverrà la sostituzione
        If TmpZ > Altezza - 1 Then TmpZ = Altezza - 1
        'Ora il programma è pronto a disegnare la nuova linea secondo le coordinate reimpostate
        'dell'oggetto
        Schermo.DrawLine 52, Altezza - 38, TmpX, TmpZ, RGBA(0, 0.8, 0.4, 1)
        'La funzione aasume il valore boolenao True,che starà a significare che sono stati riscontrati
        'dei problemi del disegnare la linea che segnala all'interno dell'editor,la posizione
        'dell'oggetto
        Problemi_Oggetto = True
    End If
End Function

Sub Traduci(NuovaLingua As String)
    Dim Nodo As Node
    Select Case NuovaLingua
    Case Is = "Italiano"
        'Assegno il valore di proprietà Checked = True al bottone LItaliano.
        'Questo sarà di aiuto all'utente per capire con quale linguaggio è attualmete tradotto il programma
        LItaliano.Checked = True
        'Assegno il valore di proprietà Checked = False al bottone LInglese.
        LInglese.Checked = False
        'Incomincio con il tradurre il menù del Form Map_Editor in Italiano
        Nuovo.Caption = "Nuovo"
        Carica_mappa.Caption = "Carica Mappa"
        Salva.Caption = "Salva"
        Salva_con_nome.Caption = "Salva con nome"
        Converti_mappa_in_3D.Caption = "Converti mappa in 3D"
        Stampa.Caption = "Stampa"
        Esci.Caption = "Esci"
        Visualizza.Caption = "Visualizza"
        Opzioni.Caption = "Opzioni"
        AvviaAnteprima.Caption = "Avvia Anteprima"
        StoppaAnteprima.Caption = "Stoppa Anteprima"
        Lingua.Caption = "Lingua"
        LItaliano.Caption = "Italiano"
        LInglese.Caption = "Inglese"
        Registra.Caption = "Registra"
        InfoMapEditor.Caption = "Informazioni su Map Editor 1.0"
        With Form_Opzioni
            'Ora traduco tutto il contenuto del Form_Opzioni sempre in lingua Italiana
            .Caption = "Opzioni"
            .Tabella.Tabs.Item(1).Caption = "Gestione Costruzione"
            .Tabella.Tabs.Item(2).Caption = "Oggetti"
            .Tabella.Tabs.Item(3).Caption = "Opzioni Editor"
            'Traduco gli oggetti contenuti nel FrameOperazioni e il frame stesso
            .FrameOperazioni.Caption = "Esegui..."
            .CaricaOggetto.Caption = "Carica Oggetto"
            .EliminaOggetto.Caption = "Rimuovi Oggetto"
            .AggiungiGruppo.Caption = "Aggiungi Gruppo"
            .EliminaGruppo.Caption = "Rimuovi Gruppo"
            'Traduco gli oggetti presenti all'interno del FrameAssegnazioneGruppo e il
            'Frame stesso
            .FrameAssegnazioneGruppo.Caption = "Assegnazione Gruppo"
            .FrameElenco.Caption = "Elenco Oggetti"
            .SpostaOggettoInGruppo.Caption = "Sposta in..."
            .FrameInformazioniOggetto.Caption = "Informazioni Oggetto"
            .FrameDescrizioneOggetto.Caption = "Descrizione"
            .FrameOpzioniOggetto.Caption = "Opzioni"
            .CancellaDescrizione.Caption = "Cancella"
            .ConfermaDescrizione.Caption = "Conferma"
            .VisualizzaOggetto.Caption = "Visualizza nell'editor"
            .AttivaOggetto.Caption = "Attiva in M.A.3D"
            .CaricaTextureOggetto.Caption = "Cambia"
            .AnnullaTextureOggetto.Caption = "Annulla"
            .FrameModificaOggetto.Caption = "Modifica"
            .LabelOperazione = "Operazione:"
            .LabelAsse = "Asse:"
            .LabelNuovoValore = "Nuovo Valore:"
            .ConfermaNuovoValoreOggetto.Caption = "Conferma"
            'Traduco la lista delle operazioni presente all'interno del controllo OperazioneModificaOggetto
            .OperazioneModificaOggetto.List(0) = "Muovi"
            .OperazioneModificaOggetto.List(1) = "Scala"
            .OperazioneModificaOggetto.List(2) = "Ruota"
            'Traduco il FrameCostruisci e tutti gli oggetti contenuti al suo interno
            .FrameCostruisci.Caption = "Costruisci"
            .Pavimento.Caption = "Pavimenti"
            .Muri.Caption = "Muri"
            .Soffitto.Caption = "Soffitti"
            .FrameElencoCostruzioni.Caption = "Elenco Costruzioni"
            .Frame_muri.Caption = "Muri"
            .FramePavimenti = "Pavimenti / Soffitti"
            .Labmuri = "Muri esistenti:"
            .Rinomina.Caption = "Rinomina"
            .LabAltezzamuro = "Altezza:"
            .LabMatAltezza = "Mat. Altezza:"
            .LabMatLarghezza = "Mat.Larghezza:"
            .CambiaTexture.Caption = "Cambia"
            .NessunaTexture.Caption = "Annulla"
            .Modifica1.Caption = "Modifica"
            .Conferma1.Caption = "Conferma"
            .EliminaMuro.Caption = "Elimina"
            .AssegnazioneMultiplaMuri.Caption = "Assegnazione Multipla"
            .Materiale.Caption = "Materiale"
            .LabPavimenti = "Pavimenti / Soffitti esistenti:"
            .Rinomina2.Caption = "Rinomina"
            .LabMatAltezza2 = "Mat. Altezza:"
            .LabmatLarghezza2 = "Mat.Larghezza:"
            .TipoPavimento.Caption = "Pavimento"
            .TipoSoffitto.Caption = "Soffitto"
            .CambiaTexture2.Caption = "Cambia"
            .NessunaTexture2.Caption = "Annulla"
            .Modifica2.Caption = "Modifica"
            .Conferma2.Caption = "Conferma"
            .EliminaSoP.Caption = "Elimina"
            .AssegnazioneMultiplaSoP.Caption = "Assegnazione multipla"
            .Materiale2.Caption = "Materiale"
            'Traduco il FormPreferenze e tutti gli oggetti contenuti al suo interno
            .FramePreferenze = "Preferenze"
            .Linee_guida.Caption = "Mostra linee guida"
            .Mostra_Menù.Caption = "Mostra menù"
            .Rileva_allineamento.Caption = "Rileva allineamento muri"
            .Rileva_Allineamento2.Caption = "Rileva allineamento S/P"
            .Controlla_Muri.Caption = "Visualizza muri"
            .Visualizza_Pavimenti.Caption = "Visualizza Pavimenti"
            .Visualizza_soffitti.Caption = "Visualizza Soffitti"
            .Controlla_spigoli.Caption = "Visualizza spigoli"
            'Traduco il form Opzioni_Telecamera e tutti gli oggetti contenuti al suoi interno
            .Opzioni_telecamera = "Opzioni Telecamera"
            .Visualizza_telecamera.Caption = "Visualizza telecamera"
            .ModificaTelecamera.Caption = "Modifica"
            'Traduco il form Opzioni_griglia e tutti gli oggetti contenuti al suo interno
            .Opzioni_griglia = "Opzioni Griglia"
            .Visualizza_griglia.Caption = "Visualizza griglia"
            .GrigliaControllataDaZoom.Caption = "Ridimensiona griglia tramite zoom"
            .labAltezzaGriglia = "Altezza:"
            .LabLarghezzaGriglia = "Larghezza:"
            .LabLuminosità_griglia = "Luminosità:"
            'Traduco il form OpzioniScale e tutti gli oggetti contenuti al suo interno
            .OpzioniScale = "Opzioni Scale"
            .LabSelezionaScale = "Seleziona scale:"
            .Impostato.Caption = "Impostato:"
            .Personalizzato.Caption = "Personalizzato:"
            .SalvaScalePersonalizzato.Caption = "Salva"
            'Traduco il contenuto del FrameFondaleEditor e tutti gli oggetti contenuti al suo interno
            .FrameFondaleEitor = "Fondale"
            .DisattivaFondale.Caption = "Disattiva fondale"
            .FondaleStatico.Caption = "Attiva fondale"
            .CambiaImmagineDiSfondo.Caption = "Cambia"
            'Traduco il frame Opzioni Zoom e tutti gli oggetti contenuti al suo interno
            .OpzioniZoom = "Opzioni Zoom"
            .LabDiminuisciZoom = "Diminuisci"
            .LabAumentaZoom = "Aumenta"
            .LabeRipristinaZoom = "Reimposta"
            .RipristinaZoom.Caption = "Ripristina"
            'Traduco il frame PersonalizzaColori e tutti gli oggetti contenuti al suo interno
            .FramePersonalizzaColori = "Personalizza Colori"
            .ColoreMuri.Caption = "Muri"
            .ColoreMuriSelezionati.Caption = "Muri Selezionati"
            .ColoreSpigoliMuri.Caption = "Spigoli Muri"
            .ColoreSoffitti.Caption = "Soffitti"
            .ColorePavimenti.Caption = "Pavimenti"
            .ColoreAllineamentoSP.Caption = "Allineamento S/P"
            .ColoreSPSelezionati.Caption = "S/P Selezionati"
            .ColoreAllineamentoMuri.Caption = "Allineamento Muri"
            .ColoreLineeGuida.Caption = "Linee Guida"
            .ColoreSfondoMenù.Caption = "Sfondo Menù"
            .ElencoGruppiOggetti.Nodes.Item(1).Text = "Oggetti senza gruppo"
            For Each Nodo In .ElencoGruppiOggetti.Nodes
                If Nodo.Text = "New Group" Then Nodo.Text = "Nuovo Gruppo"
            Next
        End With
    Case Is = "Inglese"
        'Assegno il valore di proprietà Checked = True al bottone LInglese.
        'Questo sarà di aiuto all'utente per capire con quale linguaggio è attualmete tradotto il programma
        LInglese.Checked = True
        'Assegno il valore di proprietà Checked = False al bottone LItaliano.
        LItaliano.Checked = False
        'Incomincio con il tradurre il menù del Form Map_Editor in Inglese
        Nuovo.Caption = "New"
        Carica_mappa.Caption = "Load Map"
        Salva.Caption = "Save"
        Salva_con_nome.Caption = "Save with name"
        Converti_mappa_in_3D.Caption = "Generate 3D map"
        Stampa.Caption = "Print"
        Esci.Caption = "Exit"
        Visualizza.Caption = "Show"
        Opzioni.Caption = "Option"
        AvviaAnteprima.Caption = "Run Preview"
        StoppaAnteprima.Caption = "Stop Preview"
        Lingua.Caption = "Lenguage"
        LItaliano.Caption = "Italian"
        LInglese.Caption = "English"
        Registra.Caption = "Register"
        InfoMapEditor.Caption = "About Map Editor 1.0"
        'Ora traduco tutto il contenuto del Form_Opzioni sempre in lingua Inglese
        With Form_Opzioni
            .Caption = "Option"
            .Tabella.Tabs.Item(1).Caption = "Manage Construction"
            .Tabella.Tabs.Item(2).Caption = "Object"
            .Tabella.Tabs.Item(3).Caption = "Editor Option"
            'Traduco gli oggetti contenuti nel FrameOperazioni e il frame stesso
            .FrameOperazioni.Caption = "Execute..."
            .CaricaOggetto.Caption = "Load Object"
            .EliminaOggetto.Caption = "Remove Object"
            .AggiungiGruppo.Caption = "Add Group"
            .EliminaGruppo.Caption = "Remove Group"
            'Traduco gli oggetti presenti all'interno del FrameAssegnazioneGruppo e il
            'Frame stesso
            .FrameAssegnazioneGruppo.Caption = "Assign group"
            .FrameElenco.Caption = "Object List"
            .SpostaOggettoInGruppo.Caption = "Move to..."
            .FrameInformazioniOggetto = "Object Information"
            .FrameDescrizioneOggetto = "Description"
            .FrameOpzioniOggetto = "Option"
            .CancellaDescrizione.Caption = "Erase"
            .ConfermaDescrizione.Caption = "Confirm"
            .VisualizzaOggetto.Caption = "Show in editor"
            .AttivaOggetto.Caption = "Activate in 3D.P.M"
            .CaricaTextureOggetto.Caption = "Change"
            .AnnullaTextureOggetto.Caption = "Cancel"
            .FrameModificaOggetto.Caption = "Modify"
            .LabelOperazione = "Operation:"
            .LabelAsse = "Axis:"
            .LabelNuovoValore = "New Value:"
            .ConfermaNuovoValoreOggetto.Caption = "Confirm"
            'Traduco la lista delle operazioni presente all'interno del controllo OperazioneModificaOggetto
            .OperazioneModificaOggetto.List(0) = "Move"
            .OperazioneModificaOggetto.List(1) = "Scale"
            .OperazioneModificaOggetto.List(2) = "Rotate"
            'Traduco il FrameCostruisci e tutti gli oggetti contenuti al suo interno
            .FrameCostruisci = "Build"
            .Pavimento.Caption = "Floors"
            .Muri.Caption = "Walls"
            .Soffitto.Caption = "Ceilings"
            .FrameElencoCostruzioni = "Construction List"
            .Frame_muri = "Walls"
            .FramePavimenti = "Floor / Ceiling"
            .Labmuri = "Existen walls:"
            .Rinomina.Caption = "Rename"
            .LabAltezzamuro = " Height:"
            .LabMatAltezza = "          Tile H:"
            .LabMatLarghezza = "            Tile W:"
            .CambiaTexture.Caption = "Change"
            .NessunaTexture.Caption = "Cancel"
            .Modifica1.Caption = "Modify"
            .Conferma1.Caption = "Confirm"
            .EliminaMuro.Caption = "Delete"
            .AssegnazioneMultiplaMuri.Caption = "Multiple assegnation"
            .Materiale.Caption = "Material"
            .LabPavimenti = "Exist Floors / Ceilings:"
            .Rinomina2.Caption = "Rename"
            .LabMatAltezza2 = "          Tile H:"
            .LabmatLarghezza2 = "            Tile W:"
            .TipoPavimento.Caption = "Floor"
            .TipoSoffitto.Caption = "Ceiling"
            .CambiaTexture2.Caption = "Change"
            .NessunaTexture2.Caption = "Cancel"
            .Modifica2.Caption = "Modify"
            .Conferma2.Caption = "Confirm"
            .EliminaSoP.Caption = "Delete"
            .AssegnazioneMultiplaSoP.Caption = "Multiple assegnation"
            .Materiale2.Caption = "Material"
            'Traduco il FormPreferenze e tutti gli oggetti contenuti al suo interno
            .FramePreferenze = "Preference"
            .Linee_guida.Caption = "Show guid lines"
            .Mostra_Menù.Caption = "Show menù"
            .Rileva_allineamento.Caption = "Detect walls allign"
            .Rileva_Allineamento2.Caption = "Detect F / C align"
            .Controlla_Muri.Caption = "Show walls"
            .Visualizza_Pavimenti.Caption = "Show floors"
            .Visualizza_soffitti.Caption = "Show ceiling"
            .Controlla_spigoli.Caption = "Show walls corner"
            'Traduco il form Opzioni_Telecamera e tutti gli oggetti contenuti al suoi interno
            .Opzioni_telecamera = "Camera Options"
            .Visualizza_telecamera.Caption = "Show camera"
            .ModificaTelecamera.Caption = "Modify"
            'Traduco il form Opzioni_griglia e tutti gli oggetti contenuti al suo interno
            .Opzioni_griglia = "Grid Options"
            .Visualizza_griglia.Caption = "Show grid"
            .GrigliaControllataDaZoom.Caption = "Control grid by zoom"
            .labAltezzaGriglia = "Height:"
            .LabLarghezzaGriglia = "Width:"
            .LabLuminosità_griglia = "Brightness:"
            'Traduco il form OpzioniScale e tutti gli oggetti contenuti al suo interno
            .OpzioniScale = "Scale Options"
            .LabSelezionaScale = "Select scale:"
            .Impostato.Caption = "Impostated:"
            .Personalizzato.Caption = "Personal:"
            .SalvaScalePersonalizzato.Caption = "Save"
            'Traduco il contenuto del FrameFondaleEditor e tutti gli oggetti contenuti al suo interno
            .FrameFondaleEitor = "Background"
            .DisattivaFondale.Caption = "Deactivate background"
            .FondaleStatico.Caption = "Activate background"
            .CambiaImmagineDiSfondo.Caption = "Change"
            'Traduco il frame Opzioni Zoom e tutti gli oggetti contenuti al suo interno
            .OpzioniZoom = "Zoom Options"
            .LabDiminuisciZoom = "Decrease"
            .LabAumentaZoom = "Increase"
            .LabeRipristinaZoom = "Reset"
            .RipristinaZoom.Caption = "Reset"
            'Traduco il frame PersonalizzaColori e tutti gli oggetti contenuti al suo interno
            .FramePersonalizzaColori = "Colours"
            .ColoreMuri.Caption = "Walls"
            .ColoreMuriSelezionati.Caption = "Selected Walls"
            .ColoreSpigoliMuri.Caption = "Walls Corner"
            .ColoreSoffitti.Caption = "Ceilings"
            .ColorePavimenti.Caption = "Floors"
            .ColoreAllineamentoSP.Caption = "F/C Align"
            .ColoreSPSelezionati.Caption = "Selected F/C"
            .ColoreAllineamentoMuri.Caption = "Walls Align"
            .ColoreLineeGuida.Caption = "Guid Lines"
            .ColoreSfondoMenù.Caption = "Menù Back."
            .ElencoGruppiOggetti.Nodes.Item(1).Text = "Objects without group"
            For Each Nodo In .ElencoGruppiOggetti.Nodes
                If Nodo.Text = "Nuovo Gruppo" Then Nodo.Text = "New Group"
            Next
        End With
    End Select
End Sub

