VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsClasse_sonora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Dichiaro il motore sonoro,cioè l'oggetto fondamentale affinchè vengano riprodotti
'i file sonori
Private Motore_Sonoro As SoundEngine8
'Ora dichiaro un'altro oggetto che mi servirà per aggiungere,rimuovere o riprodurre
'i file sonori
Private Av_Suoni As Sounds8

Public Sub Inizializza_Suoni()
    'Setto i due oggetti sopra dichiarati
    On Error GoTo Errore
    Set Motore_Sonoro = New SoundEngine8
    Set Av_Suoni = New Sounds8
    'Ora carico tutti i file che mi serviranno riprodurre all'interno del
    'programma
    Av_Suoni.AddFile "Sound\Cammina.wav", "Passi"
    Av_Suoni.AddFile "Sound\Menù.wav", "Menù"
    Av_Suoni.AddFile "Sound\Menù2.wav", "Menù2"
    Av_Suoni.AddFile "Sound\Menù3.wav", "Menù3"
    Av_Suoni.AddFile "Sound\Uscita.wav", "Uscita"
    Av_Suoni.AddFile "Sound\Avvio.wav", "Avvio"
    Av_Suoni.AddFile "Sound\AvviaAnteprima.wav", "AvviaAnteprima"
    Av_Suoni.AddFile "Sound\StoppaAnteprima.wav", "StoppaAnteprima"
    Exit Sub
Errore:
    'Dichiaro una variabile,che mi servirà per identificare la risposta da parte dell'utente,al messaggio di
    'errore che gli verrà posto
    Dim Risposta As VbMsgBoxResult
    'Visualizzo il messaggio di errore che informa l'utente che le risorse audio della scheda sonora sono attualmente in uso da un'altra applicazione
    'L'utente a questo punto potrà scegliere se eseguire comunque il programma senza suoni,oppure uscire dallo stesso
    Risposta = MsgBox("ATTENZIONE!Le risorse audio sono attualmente in uso da un'altra applicazione!Si desidera comunque continuare avviare il programma,il quale sarà privo di effetti sonori?", vbYesNo, "Risorsa occupata")
    'Se viene scelto di uscire dal programma,allora questo verrà terminato
    If Risposta = vbNo Then End
End Sub

Public Sub Riproduci_Passi()
    'Nel caso le risorse audio fossero utilizzate da un'altra applicazione,si procederebbe
    'senza che il suono venga riprodotto evitando così un errore di automazione
    On Error Resume Next
    'Facciamo riprodurre il file sonoro "Cammina.wav",dichiarato come "Passi"
    Av_Suoni("Passi").Play
End Sub

Public Sub Esegui_suono_menù()
    'Nel caso le risorse audio fossero utilizzate da un'altra applicazione,si procederebbe
    'senza che il suono venga riprodotto evitando così un errore di automazione
    On Error Resume Next
    'Quando viene premuta una qualsiasi voce di menù (Tranne quelle che avviano dell funzioni)
    'verrà riprodotto il seguente suono
    Av_Suoni("Menù").Play
End Sub

Public Sub Esegui_suono_menù_2()
    'Nel caso le risorse audio fossero utilizzate da un'altra applicazione,si procederebbe
    'senza che il suono venga riprodotto evitando così un errore di automazione
    On Error Resume Next
    'Quando viene premuta la voce di menù AvviaAnteprima verrà riprodotto il seguente suono
    Av_Suoni("Menù2").Play
End Sub

Public Sub Esegui_suono_menù_3()
    'Nel caso le risorse audio fossero utilizzate da un'altra applicazione,si procederebbe
    'senza che il suono venga riprodotto evitando così un errore di automazione
    On Error Resume Next
    'Quando vengono premute le voci di menù Salva,Salva Con Nome,Converti Mappa 3d e
    'Aggiorna mappa 3d,verrà riprodotto il seguente file sonoro
    Av_Suoni("Menù3").Play
End Sub

Public Sub Esegui_suono_AvviaAnteprima()
    'Nel caso le risorse audio fossero utilizzate da un'altra applicazione,si procederebbe
    'senza che il suono venga riprodotto evitando così un errore di automazione
    On Error Resume Next
    'Quando viene premuta la voce di menù AvviaAnteprima verrà riprodotto il seguente suono
    Av_Suoni("AvviaAnteprima").Play
End Sub

Public Sub Esegui_suono_StoppaAnteprima()
    'Nel caso le risorse audio fossero utilizzate da un'altra applicazione,si procederebbe
    'senza che il suono venga riprodotto evitando così un errore di automazione
    On Error Resume Next
    'Quando viene premuta la voce di menù StoppaAnteprima verrà riprodotto il seguente suono
    Av_Suoni("StoppaAnteprima").Play
End Sub

Public Sub Esegui_suono_uscita()
    'Nel caso le risorse audio fossero utilizzate da un'altra applicazione,si procederebbe
    'senza che il suono venga riprodotto evitando così un errore di automazione
    On Error Resume Next
    'Quando viene premuta la voce di menù Esci,ovvero quando si deciderà di uscire dal
    'programma verrà riprodotto il seguente suono
    Av_Suoni("Uscita").Play
End Sub

Public Sub Esegui_suono_avvio()
    'Nel caso le risorse audio fossero utilizzate da un'altra applicazione,si procederebbe
    'senza che il suono venga riprodotto evitando così un errore di automazione
    On Error Resume Next
    'Il seguente suono andrà in esecuzione solamente quando verrà avviato il programma
    Av_Suoni("Avvio").Play
End Sub

