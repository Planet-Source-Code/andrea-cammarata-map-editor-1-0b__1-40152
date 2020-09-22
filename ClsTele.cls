VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsTele"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Dichiaro un gestore di comandi per la telecamera.Ho scelto di dichiarare questo oggetto
'di tipo pubblico perchè mi servirà anche nella classe sonora,dove riprodurro un suono ogni qualvolta
'sposterò la telecamera in una direzione,per simulare i passi del personaggio
Public Comandi_Telecamera As New InputEngine8
'Dichiaro una telecamera
Private Telecamera As New Camera8
'Inizializzo una variabile che funzionerà da controllo per vedere se mi stò spostando
'avanti oppure indietro
Private Cammina As Single
'Inizializzo una variabile che funzionerà da controllo per vedere se mi stò spostando a
'destra oppure a sinistra
Private Gira As Single
'Ora inizializzo le tre variabili che avranno utilità di mantenere ognuna le coordinate
'spaziali in cui è ancorata la telecamera
Private Tele_posX As Single
Private Tele_posY As Single
Private Tele_posZ As Single
'Creo altre tre variabili che mi servono per verificare dope la telecamera viene puntata
Private Tele_guardaX As Single
Private Tele_guardaY As Single
Private Tele_guardaZ As Single
'Definisco l'angolazione orizzontale della telecamera e quella verticale
Private AngoloX As Single
Private AngoloY As Single
'Dichiaro due variabili che fungeranno a contenere il valore dello spostamento orizzontale del mouse
'e di quello verticale
Private TmpMouseX As Long
Private TmpMouseY As Long
'Dichiaro i tre bottoni del mouse per verificare se questi vengono premuti
Private MouseButton1 As Integer
Private MouseButton2 As Integer
Private MouseButton3 As Integer
'Dichiaro due variabili che mi servono a contenere una il valore vecchio dello scrolling del mouse
'e l'altro,al contrario,contiene il valore aggiornato
Private Vecchio_scroll As Long
Private Nuovo_scroll As Long
Public Sub Controlla_Comandi()
    'Controllo se viene premuto il tasto "Freccia SU", e , in caso positivo il motore 3d, grazie
    'alle altra funzioni create,capirà che ci stiamo spostando avanti...
    If Comandi_Telecamera.IsKeyPressed(TV_KEY_UP) = True Then
        Cammina = 1
        Suoni.Riproduci_Passi
    End If
    '...In caso contrario,premendo il tasto "Freccia GIU'" verrà riconosciuto che mi stò spostando
    'nella direzione opposta,ovvero indietro
    If Comandi_Telecamera.IsKeyPressed(TV_KEY_DOWN) = True Then
        Cammina = -1
        Suoni.Riproduci_Passi
    End If
    'Controllo se viene premuto il tasto "Freccia DESTRA", e , in caso positivo il motore 3d, grazie
    'alle altra funzioni create,capirà che ci stiamo spostando verso destra...
    If Comandi_Telecamera.IsKeyPressed(TV_KEY_LEFT) = True Then
        Gira = 1
        Suoni.Riproduci_Passi
    End If
    '...In caso contrario,premendo il tasto "Freccia GIU'" verrà riconosciuto che mi stò spostando
    'nella direzione opposta,ovvero verso sinistra
    If Comandi_Telecamera.IsKeyPressed(TV_KEY_RIGHT) = True Then
        Gira = -1
        Suoni.Riproduci_Passi
    End If
End Sub

Public Sub Aggiorna_Comandi()
    'Aggiorno il valore di scrolling del mouse.Lo scrolling mi permette di effettuare alcune operazioni
    'sulla telecamera,grazie ai mouse dotati dell'apposita rotellina appunto di scrolling
    Vecchio_scroll = Nuovo_scroll
    'Rilevo i movimenti del mouse con la rispettiva velocità.Se si muoverà ad esempio il mouse velocemente
    'in su, la telecamera verrà puntata con la stessa velocità verso l'alto
    Comandi_Telecamera.GetMouseState TmpMouseX, TmpMouseY, MouseButton1, MouseButton2, MouseButton3, Nuovo_scroll
    'Grazie al comando precedente,tutti i valori del mouse vengono aggionate nelle rispettive variabili.
    'Ora aggiorno l'angolazione della telecamera aggiungendo o sottraendo i valori fornitoci dal mouse
    AngoloX = AngoloX - (TmpMouseY / 100)
    AngoloY = AngoloY - (TmpMouseX / 100)
    'Pongo i limiti di angolazione di visuale a 80° e - 80°
    If AngoloX > 1.3 Then AngoloX = 1.3
    If AngoloX < -1.3 Then AngoloX = -1.3
    'Creo una struttura di controllo che mi servirà per aggiornare i valore della variabile Cammina.
    'Se Questa è > 0 allore gli aggiungeremo 0.05 (questo valore si riferisce ai "passi" che dovrà compiere
    'la telecamera) se invece sarà < 0 effettueremo l'operazione contraria
    Select Case Cammina
    Case Is > 0
        Cammina = Cammina - 0.05
        If Cammina < 0 Then Cammina = 0
    Case Is < 0
        Cammina = Cammina + 0.05
        If Cammina > 0 Then Cammina = 0
    End Select
    'Creo una struttura di controllo che mi servirà per aggiornare i valore della variabile Gira.
    'Se Questa è > 0 allore gli aggiungeremo 0.05 (questo valore si riferisce ai "passi" che dovrà compiere
    'la telecamera) se invece sarà < 0 effettueremo l'operazione contraria
    Select Case Gira
    Case Is > 0
        Gira = Gira - 0.05
        If Gira < 0 Then Gira = 0
    Case Is < 0
        Gira = Gira + 0.05
        If Gira > 0 Then Gira = 0
    End Select
        
    'Aggiorniamo i vettori grazie alle posizioni ricavate e gli angoli
    Tele_posX = Tele_posX + (Cos(AngoloY) * Cammina * TV8.TimeElapsed \ 2) + (Cos(AngoloY + 3.141596 / 2) * Gira * TV8.TimeElapsed \ 2)
    Tele_posZ = Tele_posZ + (Sin(AngoloY) * Cammina * TV8.TimeElapsed \ 2) + (Sin(AngoloY + 3.141596 / 2) * Gira * TV8.TimeElapsed \ 2)
    Tele_guardaX = Tele_posX + Cos(AngoloY)
    Tele_guardaY = Tele_posY + Tan(AngoloX)
    Tele_guardaZ = Tele_posZ + Sin(AngoloY)
    'Aggiorno la posizione di telecamera con i dati appena elaborati
    Telecamera.SetCameraPosition Tele_posX, Tele_posY, Tele_posZ
    'Aggiorno le coordinate di dove è puntata la telecammera
    Telecamera.SetCameraLookAt Tele_guardaX, Tele_guardaY, Tele_guardaZ
End Sub
Public Sub Inizializzazione_Variabili_Di_Comando()
    'Inizializzo tutte le variabili necessarie, con il proprio valore iniziale
    Cammina = 0
    Gira = 0
    Tele_posX = Fix(PosizioneTelecamera.X / Molt) * VScale
    Tele_posY = PosizioneTelecamera.Y
    Tele_posZ = Fix(-PosizioneTelecamera.Z / Molt) * VScale
    Tele_guardaX = 0
    Tele_guardaY = 500
    Tele_guardaZ = 0
    AngoloX = 0
    AngoloY = 0
 End Sub
