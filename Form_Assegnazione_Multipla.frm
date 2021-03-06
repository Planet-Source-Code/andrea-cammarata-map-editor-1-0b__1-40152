VERSION 5.00
Begin VB.Form Form_Assegnazione_Multipla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assegnazione multipla muri"
   ClientHeight    =   2955
   ClientLeft      =   3390
   ClientTop       =   3060
   ClientWidth     =   6495
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrameParametri 
      Caption         =   "Parametri da assegnare:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   2295
      Begin VB.CheckBox AMColoriTrasparenza 
         Caption         =   "Colori Trasparenza"
         Enabled         =   0   'False
         Height          =   195
         Left            =   480
         TabIndex        =   15
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CheckBox AMMateriale 
         Caption         =   "Materiale"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox AMPropriet� 
         Caption         =   "Propriet�"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   1800
         Top             =   240
      End
      Begin VB.CommandButton Esci 
         Caption         =   "Esci"
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton Conferma 
         Caption         =   "Conferma"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CheckBox AMAltezza 
         Caption         =   "Altezza"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox AMAltitudine 
         Caption         =   "Altitudine"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox AMMatAltezza 
         Caption         =   "Mat. Altezza"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox AMMatLarghezza 
         Caption         =   "Mat. Larghezza"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CheckBox AMTexture 
         Caption         =   "Texture"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin VB.Line Line2 
         X1              =   360
         X2              =   240
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   240
         Y1              =   1920
         Y2              =   2160
      End
   End
   Begin VB.Frame FrameTipoElementi 
      Caption         =   "Muri esistenti :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.CommandButton DeselezionaTutti 
         Caption         =   "Nessuno"
         Height          =   255
         Left            =   2760
         TabIndex        =   11
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton SelezionaTutti 
         Caption         =   "Tutti"
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   2400
         Width           =   975
      End
      Begin VB.ListBox ElementiEsistenti 
         Height          =   1860
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label TotaleElementiSelezionati 
         Caption         =   "Totale: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form_Assegnazione_Multipla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Conferma_Click()
    'Dichiaro una variabile che mi servir� per capire se � stata selezionata almeno una propriet� del muro,
    'pavimento o soffitto da assegnare a pi� di essi
    Dim Selezionati As Boolean
    'Se il form � stato avviato al fine di effettuare un'assegnazione multipla delle propriet� dei muri,allora...
    If Modalit�AssegnazioneMultipla = "Muri" Then
        'Per un'assegnazione multipla dei valori dei muri � necessario che almeno uno di questi sia selezionato
        'quindi,in caso non ne fosse stato selezionato neanche uno, verrebbe visualizzato un messaggio di errore che avviserebbe l'utente
        'dell'errore
        If ElementiEsistenti.SelCount = 0 Then
            'Visualizzazione del messaggio di errore
            If LinguaS = "Italiano" Then MsgBox "Non � stato selezionato nessun muro su cui effettuare un'assegnazone multipla dei valori!" + Chr(13) + "Si prega di selezionare almeno un muro esistente all'inteno della mappa!", vbOKOnly, "Assegnazione multipla"
            If LinguaS = "Inglese" Then MsgBox "You haven't selected any walls for multipled parameters assign!" + Chr(13) + "Please select at last one wall", vbOKOnly, "Multipled assign"
            'Dopo il messaggio che avvisa l'utente,si viene reinderizzati verso il "label" Errore
            GoTo Errore
        End If
        'Alla pressione del tasto Conferma, tutti i muri selezionati dall'elenco, assumeranno gli stessi parametri richiesti
        'Questo avviene tramite un ciclo For che scandisce tutti gli elementi presenti all'interno della lista
        For I = 0 To ElementiEsistenti.ListCount - 1
            'Se l'elemento che il ciclo st� analizzando � selezionato allora...
            If ElementiEsistenti.Selected(I) = True Then
                'Se � stato richiesto un'assegnazione multipla del materiale, cio� il comando AMMateriale � selezionato,
                'allora assegner� al rispettivo muro la Texture comune
                If AMTexture.Value = 1 Then Riga(I + 1).Texture = Riga(IndiceLista).Texture: Selezionati = True
                'Se � stato richiesto un'assegnazione multipla della Texture, cio� il comando AMTexture � selezionato,
                'allora assegner� al rispettivo muro il materiale comune
                If AMPropriet�.Value = 1 Then Riga(I + 1).Propriet� = Riga(IndiceLista).Propriet�: Selezionati = True
                'Se � stato richiesto un'assegnazione multipla dell'altezza, cio� il comando AMAltezza � selezionato,
                'allora assegner� al rispettivo muro l'altezza comune
                If AMAltezza.Value = 1 Then Riga(I + 1).Altezza = Riga(IndiceLista).Altezza: Selezionati = True
                'Se � stato richiesto un'assegnazione multipla dell'altitudine, cio� il comando AMAltitudine � selezionato,
                'allora assegner� al rispettivo muro l'altezza comune
                If AMAltitudine.Value = 1 Then Riga(I + 1).Altitudine = Riga(IndiceLista).Altitudine: Selezionati = True
                'Se � stato richiesto un'assegnazione multipla del numero di mattonelle in altezza, cio� il comando AMMatAltezza � selezionato,
                'allora assegner� al rispettivo muro il numero di mattonelle in altezza comune
                If AMMatAltezza.Value = 1 Then Riga(I + 1).NMattonelleALtezza = Riga(IndiceLista).NMattonelleALtezza: Selezionati = True
                'Se � stato richiesto un'assegnazione multipla del numero di mattonelle in larghezza, cio� il comando AMMatLarghezza � selezionato,
                'allora assegner� al rispettivo muro il numero di mattonelle in larghezza comune
                If AMMatLarghezza.Value = 1 Then Riga(I + 1).NMAttonelleLarghezza = Riga(IndiceLista).NMAttonelleLarghezza: Selezionati = True
                'Se � stata richiesta un'assegnazione multipla dei colori di trasparenza dei muri,allora richiamer� la funzione
                'addetta ad avviare la multi assegnazione degli stessi
                If AMColoriTrasparenza.Value = 1 Then Assegna_colori_comuni: Selezionati = True
                'Se � stata richiesta un'assegnazione multipla del materiale,allora richiamer� la funzione addetta ad assegnare
                'i rispettivi valori del materiale del muro modello al muro analizzato
                If AMMateriale.Value = 1 Then Assegna_materiale_comune: Selezionati = True
            End If
        Next
        'Se nessuna delle CheckBox presenti all'interno del form � stata selezionata,quindi � stata richiesta una multi assegnazione a vuoto,allora
        'verr� visualizzato un messaggio che informer� l'utente dell'errore
        If Selezionati = False Then
            If LinguaS = "Italiano" Then MsgBox "Non � stata selezionata nessuna propriet� da applicare a pi� Muri!" + Chr(13) + "Si prega di selezionare almeno una propriet� per avviare una multi assegnazione!", vbOKOnly, "Assegnazione multipla"
            If LinguaS = "Inglese" Then MsgBox "You haven't selected any propriety for multipled assign!" + Chr(13) + "Please select at last one propriety!", vbOKOnly, "Multipled assign"
        End If
    'Nel caso in cui,invece il form sia stato avviato al fine di effettuare una multi assegnazione delle propriet� dei pavimenti o soffitti,allora...
    Else
        'Anche per un'assegnazione multipla dei valori dei soffitti o pavimenti � necessario che almeno uno di questi sia selezionato
        'quindi,in caso non ne fosse stato selezionato neanche uno, verrebbe visualizzato un messaggio di errore che avviserebbe l'utente
        'dell'errore
        If ElementiEsistenti.SelCount = 0 Then
            'Visualizzazione del messaggio di errore
            If LinguaS = "Italiano" Then MsgBox "Non � stato selezionato nessun pavimento o soffitto su cui effettuare un'assegnazone multipla dei valori!" + Chr(13) + "Si prega di selezionare almeno un pavimento o un soffitto esistente all'inteno della mappa!", vbOKOnly, "Assegnazione multipla"
            If LinguaS = "Inglese" Then MsgBox "You haven't selected any floor or ceiling for multipled parameters assign!" + Chr(13) + "Please select at last one floor or ceiling", vbOKOnly, "Multipled assign"
            'Dopo il messaggio che avvisa l'utente,si viene reinderizzati verso il "label" Errore
            GoTo Errore
        End If
        'Alla pressione del tasto Conferma, tutti i muri selezionati dall'elenco, assumeranno gli stessi parametri richiesti
        'Questo avviene tramite un ciclo For che scandisce tutti gli elementi presenti all'interno della lista
        For J = 0 To ElementiEsistenti.ListCount - 1
            'Se l'elemento che il ciclo st� analizzando � selezionato allora...
            If ElementiEsistenti.Selected(J) = True Then
                'Se � stato richiesto un'assegnazione multipla della Texture, cio� il comando AMTexture � selezionato,
                'allora assegner� al rispettivo pavimento o soffitto la Texture comune
                If AMTexture.Value = 1 Then SoP(J + 1).CR.Texture = SoP(IndiceLista2).CR.Texture: Selezionati = True
                'Se � stato richiesto un'assegnazione multipla della Texture, cio� il comando AMTexture � selezionato,
                'allora assegner� al rispettivo muro il materiale comune
                If AMPropriet�.Value = 1 Then SoP(J + 1).CR.Propriet� = SoP(IndiceLista2).CR.Propriet�: Selezionati = True
                'Se � stato richiesto un'assegnazione multipla dell'altezza, cio� il comando AMAltezza � selezionato,
                'allora assegner� al rispettivo pavimento o soffitto l'altezza comune
                If AMAltezza.Value = 1 Then SoP(J + 1).CR.Altezza = SoP(IndiceLista2).CR.Altezza: Selezionati = True
                'Se � stato richiesto un'assegnazione multipla dell'altitudine, cio� il comando AMAltitudine � selezionato,
                'allora assegner� al rispettivo pavimento o soffitto l'altezza comune
                If AMAltitudine.Value = 1 Then SoP(J + 1).CR.Altitudine = SoP(IndiceLista2).CR.Altitudine: Selezionati = True
                'Se � stato richiesto un'assegnazione multipla del numero di mattonelle in altezza, cio� il comando AMMatAltezza � selezionato,
                'allora assegner� al rispettivo pavimento o soffitto il numero di mattonelle in altezza comune
                If AMMatAltezza.Value = 1 Then SoP(J + 1).CR.NMattonelleALtezza = SoP(IndiceLista2).CR.NMattonelleALtezza: Selezionati = True
                'Se � stato richiesto un'assegnazione multipla del numero di mattonelle in larghezza, cio� il comando AMMatLarghezza � selezionato,
                'allora assegner� al rispettivo pavimento o soffitto il numero di mattonelle in larghezza comune
                If AMMatLarghezza.Value = 1 Then SoP(J + 1).CR.NMAttonelleLarghezza = SoP(IndiceLista2).CR.NMAttonelleLarghezza: Selezionati = True
                'Se � stata richiesta un'assegnazione multipla dei colori di trasparenza dei muri,allora richiamer� la funzione
                'addetta ad avviare la multi assegnazione degli stessi
                If AMColoriTrasparenza.Value = 1 Then Assegna_colori_comuni: Selezionati = True
                'Se � stata richiesta un'assegnazione multipla del materiale,allora richiamer� la funzione addetta ad assegnare
                'i rispettivi valori del materiale del pavimento o soffitto modello pavimento o soffitto analizzato
                If AMMateriale.Value = 1 Then Assegna_materiale_comune: Selezionati = True
            End If
        Next
        'Se nessuna delle CheckBox presenti all'interno del form � stata selezionata,quindi � stata richiesta una multi assegnazione a vuoto,allora
        'verr� visualizzato un messaggio che informer� l'utente dell'errore
        If Selezionati = False Then
            If LinguaS = "Italiano" Then MsgBox "Non � stata selezionata nessuna propriet� da applicare a pi� pavimenti o soffitti!" + Chr(13) + "Si prega di selezionare almeno una propriet� per avviare una multi assegnazione!", vbOKOnly, "Assegnazione multipla"
            If LinguaS = "Inglese" Then MsgBox "You haven't selected any propriety for multipled assign!" + Chr(13) + "Please select at last one propriety!", vbOKOnly, "Multipled assign"
        End If
    End If
'Label Errore al quale si viene reinderizzati in caso non venga selezionata nessuna propriet� per l'assegnazione multipla,oppure non � stato selezionato nessun
'muro,pavimento o soffitto
Errore:
End Sub

Private Sub DeselezionaTutti_Click()
    'Avvio un ciclo che deselezioner� tutti gli elementi presenti all'interno della lista MuriEsistenti
    For I = 0 To ElementiEsistenti.ListCount - 1
        ElementiEsistenti.Selected(I) = False
    Next
End Sub

Private Sub Esci_Click()
    'Scarica se stesso, cio� chiude il form di assegnazione multipla dei parametri dei muri
    Unload Me
End Sub

Private Sub Form_Load()
    'Richiamo la funzione che mi permetter� di tradurre il programma in una delle due lingue
    'desiderate
    Traduci LinguaS
    'Cancello il contenuto dell'oggetto MuriEsistenti.
    ElementiEsistenti.Clear
    'Se la modalit� con cui � stato avvaito il form � uguale a "Muri" allora..
    If Modalit�AssegnazioneMultipla = "Muri" Then
        'Aggiungo alla lista ElementiEsistenti, tutti i muri presenti all'interno della mappa attual
        For I = 1 To Max
            ElementiEsistenti.AddItem RTrim(Riga(I).Nome)
        Next
    Else
        'Aggiungo alla lista ElementiEsistenti, tutti i pavimenti / soffitti presenti all'interno della mappa attuale
        For J = 1 To Max2
            ElementiEsistenti.AddItem RTrim(SoP(J).CR.Nome)
        Next
    End If
End Sub

Private Sub SelezionaTutti_Click()
    'Avvio un ciclo che selezioner� tutti gli elementi presenti all'interno della lista MuriEsistenti
    For I = 0 To ElementiEsistenti.ListCount - 1
        ElementiEsistenti.Selected(I) = True
    Next
End Sub

Private Sub Timer1_Timer()
    'Aggiorno il Label che mostrer� all'utente il numero di elementi selezionati
    TotaleElementiSelezionati = "Totale: " + Str(ElementiEsistenti.SelCount)
    If Modalit�AssegnazioneMultipla = "Muri" Then
        'Se il muro selezionato come modello di assegnazione ha una propriet� di Trasparenza e viene richiesta
        'un'assegnazione multipla della Propriet�, allora la checkbox AMColoriTrasparenza (cio� quella che mi permetter� di assegnare
        'i quattro colori di trasparenza e del grado di solidit� dei rispettivi quattro vertici),verr� attivata
        If RTrim(Riga(IndiceLista).Propriet�) = "Trasparenza" And AMPropriet�.Value = 1 Then
            AMColoriTrasparenza.Enabled = True
        Else
            '...altrimenti questa verrebbe disattivata.
            'Non avrebbe senza assegnare i colori di trasparenza se un muro non � appunto trasparente
            AMColoriTrasparenza.Enabled = False
        End If
    Else
        'Se il pavimento o soffitto selezionato come modello di assegnazione ha una propriet� di Trasparenza e viene richiesta
        'un'assegnazione multipla della Propriet�, allora la checkbox AMColoriTrasparenza (cio� quella che mi permetter� di assegnare
        'i quattro colori di trasparenza e del grado di solidit� dei rispettivi quattro vertici),verr� attivata
        If RTrim(SoP(IndiceLista2).CR.Propriet�) = "Trasparenza" And AMPropriet�.Value = 1 Then
            AMColoriTrasparenza.Enabled = True
        Else
            '...altrimenti questa verrebbe disattivata.
            'Non avrebbe senza assegnare i colori di trasparenza se un pavimento o soffitto non � appunto trasparente
            AMColoriTrasparenza.Enabled = False
        End If
    End If
End Sub

Sub Assegna_colori_comuni()
    'Dichiaro un indice per assegnare molto pi� rapidamente i colori e il grado di solidit� dei rispettivi
    'quattro vertici
    Dim K As Integer
    If Modalit�AssegnazioneMultipla = "Muri" Then
        'Avvio un ciclo che assegner� ai quattro vertici,i rispettivi colori di trasparenza e i gradi di solidit�
        For K = 0 To 3
            With Riga(I + 1).ColVertici(K)
                'Assegno la stessa quantit� di Rosso presente nel rispettivo vertice del muro modello, al
                'muro che si st� analizzando
                .R = Riga(IndiceLista).ColVertici(K).R
                'Assegno la stessa quantit� di Verde presente nel rispettivo vertice del muro modello, al
                'muro che si st� analizzando
                .G = Riga(IndiceLista).ColVertici(K).G
                'Assegno la stessa quantit� di Blu presente nel rispettivo vertice del muro modello, al
                'muro che si st� analizzando
                .B = Riga(IndiceLista).ColVertici(K).B
                'Assegno il grado di solidit� presente nel rispettivo vertice del muro modello, al
                'muro che si st� analizzando
                .A = Riga(IndiceLista).ColVertici(K).A
            End With
        Next
    Else
        'Avvio un ciclo che assegner� ai quattro vertici,i rispettivi colori di trasparenza e i gradi di solidit�
        For K = 0 To 3
            With SoP(J + 1).CR.ColVertici(K)
                'Assegno la stessa quantit� di Rosso presente nel rispettivo vertice del pavimento o soffitto modello, al
                'pavimento o soffitto che si st� analizzando
                .R = SoP(IndiceLista2).CR.ColVertici(K).R
                'Assegno la stessa quantit� di Verde presente nel rispettivo vertice del pavimento o soffitto modello, al
                'pavimento o soffitto che si st� analizzando
                .G = SoP(IndiceLista2).CR.ColVertici(K).G
                'Assegno la stessa quantit� di Blu presente nel rispettivo vertice del pavimento o soffitto modello, al
                'pavimento o soffitto che si st� analizzando
                .B = SoP(IndiceLista2).CR.ColVertici(K).B
                'Assegno il grado di solidit� presente nel rispettivo vertice del pavimento o soffitto modello, al
                'pavimento o soffitto che si st� analizzando
                .A = SoP(IndiceLista2).CR.ColVertici(K).A
            End With
        Next
    End If
End Sub
Sub Assegna_materiale_comune()
    If Modalit�AssegnazioneMultipla = "Muri" Then
        'Assegno al muro analizzato,gli stessi parametri della voce Ambiente del materiale del muro
        'selezionato come modello
        With Riga(I + 1).Materiale.Ambiente
            'Assegno al muro analizzato la stessa quantit� di Rosso assegnata alla voce Ambiente del
            'muro selezionato come modello
            .R = Riga(IndiceLista).Materiale.Ambiente.R
            'Assegno al muro analizzato la stessa quantit� di Verde assegnata alla voce Ambiente del
            'muro selezionato come modello
            .G = Riga(IndiceLista).Materiale.Ambiente.G
            'Assegno al muro analizzato la stessa quantit� di Blu assegnata alla voce Ambiente del
            'muro selezionato come modello
            .B = Riga(IndiceLista).Materiale.Ambiente.B
            'Assegno al muro analizzato lo stesso valore di AlphaBleding assegnata alla voce Ambiente del
            'muro selezionato come modello
            .A = Riga(IndiceLista).Materiale.Ambiente.A
        End With
        'Assegno al muro analizzato,gli stessi parametri della voce Diffusa del materiale del muro
        'selezionato come modello
        With Riga(I + 1).Materiale.Diffusa
            'Assegno al muro analizzato la stessa quantit� di Rosso assegnata alla voce Diffusa del
            'muro selezionato come modello
            .R = Riga(IndiceLista).Materiale.Diffusa.R
            'Assegno al muro analizzato la stessa quantit� di Verde assegnata alla voce Diffusa del
            'muro selezionato come modello
            .G = Riga(IndiceLista).Materiale.Diffusa.G
            'Assegno al muro analizzato la stessa quantit� di Blu assegnata alla voce Diffusa del
            'muro selezionato come modello
            .B = Riga(IndiceLista).Materiale.Diffusa.B
            'Assegno al muro analizzato lo stesso valore di AlphaBleding assegnata alla voce Diffusa del
            'muro selezionato come modello
            .A = Riga(IndiceLista).Materiale.Diffusa.A
        End With
        'Assegno al muro analizzato,gli stessi parametri della voce Emissiva del materiale del muro
        'selezionato come modello
        With Riga(I + 1).Materiale.Emissiva
            'Assegno al muro analizzato la stessa quantit� di Rosso assegnata alla voce Emissiva del
            'muro selezionato come modello
            .R = Riga(IndiceLista).Materiale.Emissiva.R
            'Assegno al muro analizzato la stessa quantit� di Verde assegnata alla voce Emissiva del
            'muro selezionato come modello
            .G = Riga(IndiceLista).Materiale.Emissiva.G
            'Assegno al muro analizzato la stessa quantit� di Blu assegnata alla voce Emissiva del
            'muro selezionato come modello
            .B = Riga(IndiceLista).Materiale.Emissiva.B
            'Assegno al muro analizzato lo stesso valore di AlphaBleding assegnata alla voce Emissiva del
            'muro selezionato come modello
            .A = Riga(IndiceLista).Materiale.Emissiva.A
        End With
        'Assegno al muro analizzato,gli stessi parametri della voce Speculare del materiale del muro
        'selezionato come modello
        With Riga(I + 1).Materiale.Speculare
            'Assegno al muro analizzato la stessa quantit� di Rosso assegnata alla voce Speculare del
            'muro selezionato come modello
            .R = Riga(IndiceLista).Materiale.Speculare.R
            'Assegno al muro analizzato la stessa quantit� di Verde assegnata alla voce Speculare del
            'muro selezionato come modello
            .G = Riga(IndiceLista).Materiale.Speculare.G
            'Assegno al muro analizzato la stessa quantit� di Blu assegnata alla voce Speculare del
            'muro selezionato come modello
            .B = Riga(IndiceLista).Materiale.Speculare.B
            'Assegno al muro analizzato lo stesso valore di AlphaBleding assegnata alla voce Speculare del
            'muro selezionato come modello
            .A = Riga(IndiceLista).Materiale.Speculare.A
        End With
        'Assegno lo stesso valore della voce Potenza del muro selezionato come modello almuro Analizzato
        Riga(I + 1).Materiale.Potenza = Riga(IndiceLista).Materiale.Potenza
    Else
        'Assegno al pavimento o soffitto analizzato,gli stessi parametri della voce Ambiente del materiale del pavimento o soffitto
        'selezionato come modello
        With SoP(J + 1).CR.Materiale.Ambiente
            'Assegno al pavimento o soffitto analizzato la stessa quantit� di Rosso assegnata alla voce Ambiente del
            'pavimento o soffitto selezionato come modello
            .R = SoP(IndiceLista2).CR.Materiale.Ambiente.R
            'Assegno al pavimento o soffitto analizzato la stessa quantit� di Verde assegnata alla voce Ambiente del
            'pavimento o soffitto selezionato come modello
            .G = SoP(IndiceLista2).CR.Materiale.Ambiente.G
            'Assegno al pavimento o soffitto analizzato la stessa quantit� di Blu assegnata alla voce Ambiente del
            'pavimento o soffitto selezionato come modello
            .B = SoP(IndiceLista2).CR.Materiale.Ambiente.B
            'Assegno al pavimento o soffitto analizzato lo stesso valore di AlphaBleding assegnata alla voce Ambiente del
            'pavimento o soffitto selezionato come modello
            .A = SoP(IndiceLista2).CR.Materiale.Ambiente.A
        End With
        'Assegno al pavimento o soffitto analizzato,gli stessi parametri della voce Diffusa del materiale del pavimento o soffitto
        'selezionato come modello
        With SoP(J + 1).CR.Materiale.Diffusa
            'Assegno al pavimento o soffitto analizzato la stessa quantit� di Rosso assegnata alla voce Diffusa del
            'pavimento o soffitto selezionato come modello
            .R = SoP(IndiceLista2).CR.Materiale.Diffusa.R
            'Assegno al pavimento o soffitto analizzato la stessa quantit� di Verde assegnata alla voce Diffusa del
            'pavimento o soffitto selezionato come modello
            .G = SoP(IndiceLista2).CR.Materiale.Diffusa.G
            'Assegno al pavimento o soffitto analizzato la stessa quantit� di Blu assegnata alla voce Diffusa del
            'pavimento o soffitto selezionato come modello
            .B = SoP(IndiceLista2).CR.Materiale.Diffusa.B
            'Assegno al muro analizzato lo stesso valore di AlphaBleding assegnata alla voce Diffusa del
            'muro selezionato come modello
            .A = SoP(IndiceLista2).CR.Materiale.Diffusa.A
        End With
        'Assegno al pavimento o soffitto analizzato,gli stessi parametri della voce Emissiva del materiale del pavimento o soffitto
        'selezionato come modello
        With SoP(J + 1).CR.Materiale.Emissiva
            'Assegno al pavimento o soffitto analizzato la stessa quantit� di Rosso assegnata alla voce Emissiva del
            'pavimento o soffitto selezionato come modello
            .R = SoP(IndiceLista2).CR.Materiale.Emissiva.R
            'Assegno al pavimento o soffitto analizzato la stessa quantit� di Verde assegnata alla voce Emissiva del
            'pavimento o soffitto selezionato come modello
            .G = SoP(IndiceLista2).CR.Materiale.Emissiva.G
            'Assegno al pavimento o soffitto analizzato la stessa quantit� di Blu assegnata alla voce Emissiva del
            'pavimento o soffitto selezionato come modello
            .B = SoP(IndiceLista2).CR.Materiale.Emissiva.B
            'Assegno al pavimento o soffitto analizzato lo stesso valore di AlphaBleding assegnata alla voce Emissiva del
            'pavimento o soffitto selezionato come modello
            .A = SoP(IndiceLista2).CR.Materiale.Emissiva.A
        End With
        'Assegno al pavimento o soffitto analizzato,gli stessi parametri della voce Speculare del materiale del muro
        'selezionato come modello
        With SoP(J + 1).CR.Materiale.Speculare
            'Assegno al pavimento o soffitto analizzato la stessa quantit� di Rosso assegnata alla voce Speculare del
            'pavimento o soffitto selezionato come modello
            .R = SoP(IndiceLista2).CR.Materiale.Speculare.R
            'Assegno al pavimento o soffitto analizzato la stessa quantit� di Verde assegnata alla voce Speculare del
            'pavimento o soffitto selezionato come modello
            .G = SoP(IndiceLista2).CR.Materiale.Speculare.G
            'Assegno al pavimento o soffitto analizzato la stessa quantit� di Blu assegnata alla voce Speculare del
            'pavimento o soffitto selezionato come modello
            .B = SoP(IndiceLista2).CR.Materiale.Speculare.B
            'Assegno al pavimento o soffitto analizzato lo stesso valore di AlphaBleding assegnata alla voce Speculare del
            'pavimento o soffitto selezionato come modello
            .A = SoP(IndiceLista2).CR.Materiale.Speculare.A
        End With
        'Assegno lo stesso valore della voce Potenza del pavimento o soffitto selezionato come modello almuro Analizzato
        SoP(J + 1).CR.Materiale.Potenza = SoP(IndiceLista2).CR.Materiale.Potenza
    End If
End Sub

Sub Traduci(NuovaLingua As String)
    'Verifico in quale lingua dovr� essere tradotto il programma, in base al parametro passato alla
    'funzione stessa
    Select Case NuovaLingua
    'Se il programma dovr� essere tradotto in italiano...
    Case Is = "Italiano"
        With Form_Assegnazione_Multipla
            'Verifico prima di tutto in che modalit� � stato avviato il form:
            'Se � stato avviato al fine di effettuare un'assegnazione multipla dei parametri
            'dei muri, allora...
            If Modalit�AssegnazioneMultipla = "Muri" Then
                'Imposto il titolo del form in modo da far capire all'utente che si st�
                'effettuando un'assegnazione multipla dei parametri dei muri
                .Caption = "Assegnazione multipla muri"
                'Faccio la stessa cosa anche per il FrameTipoElementi
                .FrameTipoElementi = "Muri esistenti:"
            'In tutti gli altri casi...
            Else
                'Imposto il titolo del form in modo da far capire all'utente che si st�
                'effettuando un'assegnazione multipla dei parametri dei Pavimenti / Soffitti
                .Caption = "Assegnazione multipla Pavimenti / Soffitti"
                'Faccio la stessa cosa anche per il FrameTipoElementi
                .FrameTipoElementi = "Pavimenti / Soffitti esistenti:"
            End If
            'Inizio la traduzione di tutti gli oggetti presenti nel Form_Assegnazione_Multipla
            .TotaleElementiSelezionati = "Totale:"
            .SelezionaTutti.Caption = "Tutti"
            .DeselezionaTutti.Caption = "Nessuno"
            'Traduco il FrameParametri e tutti gli oggetti contenuti al suo interno
            .FrameParametri = "Parametri da assegnare:"
            .AMAltezza.Caption = "Altezza"
            .AMMatAltezza.Caption = "Mat. Altezza"
            .AMMatLarghezza.Caption = "Mat. Larghezza"
            .AMMateriale.Caption = "Materiale"
            .AMPropriet�.Caption = "Propriet�"
            .AMColoriTrasparenza.Caption = "Colori Trasparenza"
            .Conferma.Caption = "Conferma"
            .Esci.Caption = "Esci"
        End With
    'Se invece il programma dovr� essere tradotto in lingua Inglese...
    Case Is = "Inglese"
        With Form_Assegnazione_Multipla
            'Verifico la modalit� con la quale � stato avviato il form
            If Modalit�AssegnazioneMultipla = "Muri" Then
                'Imposto il titolo del form in modo da far capire all'utente che si st�
                'effettuando un'assegnazione multipla dei parametri dei muri
                .Caption = "Multiple walls assegnation"
                'Faccio la stessa cosa anche per il FrameTipoElementi
                .FrameTipoElementi = "Existen walls:"
            Else
                'Imposto il titolo del form in modo da far capire all'utente che si st�
                'effettuando un'assegnazione multipla dei parametri dei Pavimenti / Soffitti
                .Caption = "Multiple floors / ceilings assegnation"
                'Faccio la stessa cosa anche per il FrameTipoElementi
                .FrameTipoElementi = "Existen Floors / Ceilings"
            End If
            'Inizio la traduzione di tutti gli oggetti presenti nel Form_Assegnazione_Multipla
            .TotaleElementiSelezionati = "Total:"
            .SelezionaTutti.Caption = "All"
            .DeselezionaTutti.Caption = "Nothing"
            'Traduco il FrameParametri e tutti gli oggetti contenuti al suo interno
            .FrameParametri = "Assegnation parameters:"
            .AMAltezza.Caption = "Height"
            .AMMatAltezza.Caption = "Tile Height"
            .AMMatLarghezza.Caption = "Tile Width"
            .AMMateriale.Caption = "Material"
            .AMPropriet�.Caption = "Propriety"
            .AMColoriTrasparenza.Caption = "Trasparency Col."
            .Conferma.Caption = "Confirm"
            .Esci.Caption = "Exit"
        End With
    End Select
End Sub


