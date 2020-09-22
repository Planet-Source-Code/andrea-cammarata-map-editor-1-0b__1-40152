VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_Materiali 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestione materiale"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameTipoMateriali 
      Caption         =   "Tipo Materiale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   3015
      Begin VB.Frame FrameOpzioniMateriale 
         Caption         =   "Opzioni Materiale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   2775
         Begin VB.CommandButton AnnullaTPM 
            Caption         =   "Annulla tutto"
            Height          =   255
            Left            =   1440
            TabIndex        =   36
            Top             =   2760
            Width           =   1215
         End
         Begin MSComctlLib.Slider AlphaBAnbiente 
            Height          =   255
            Left            =   1680
            TabIndex        =   32
            Top             =   360
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            SelStart        =   10
            Value           =   10
         End
         Begin VB.PictureBox RapSpeculare 
            BackColor       =   &H00000000&
            Height          =   375
            Left            =   1200
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   31
            Top             =   1680
            Width           =   375
         End
         Begin VB.PictureBox RapEmissiva 
            BackColor       =   &H00000000&
            Height          =   375
            Left            =   1200
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   30
            Top             =   1200
            Width           =   375
         End
         Begin VB.PictureBox RapDiffusa 
            BackColor       =   &H00000000&
            Height          =   375
            Left            =   1200
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   29
            Top             =   720
            Width           =   375
         End
         Begin VB.PictureBox RapAmbiente 
            BackColor       =   &H00000000&
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   1200
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   28
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox ValorePotenza 
            Height          =   285
            Left            =   1200
            TabIndex        =   27
            Top             =   2280
            Width           =   1215
         End
         Begin MSComctlLib.Slider AlphaBDiffusa 
            Height          =   255
            Left            =   1680
            TabIndex        =   33
            Top             =   840
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            SelStart        =   10
            Value           =   10
         End
         Begin MSComctlLib.Slider AlphaBEmissiva 
            Height          =   255
            Left            =   1680
            TabIndex        =   34
            Top             =   1320
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            SelStart        =   10
            Value           =   10
         End
         Begin MSComctlLib.Slider AlphaBSpeculare 
            Height          =   255
            Left            =   1680
            TabIndex        =   35
            Top             =   1800
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            SelStart        =   10
            Value           =   10
         End
         Begin VB.Label LabelAmbiente 
            Caption         =   "Ambiente:"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   735
         End
         Begin VB.Label LabelDiffusa 
            Caption         =   "Diffusa:"
            Height          =   255
            Left            =   360
            TabIndex        =   25
            Top             =   840
            Width           =   615
         End
         Begin VB.Label LabelEmissiva 
            Caption         =   "Emissiva:"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label LabelPotenza 
            Caption         =   "Potenza:"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   2280
            Width           =   735
         End
         Begin VB.Label LabelSpeculare 
            Caption         =   "Speculare:"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   1800
            Width           =   855
         End
      End
      Begin VB.CommandButton Conferma 
         Caption         =   "Conferma"
         Height          =   255
         Left            =   1800
         TabIndex        =   20
         Top             =   4440
         Width           =   1095
      End
      Begin VB.OptionButton SphereMapping 
         Caption         =   "Falsa Riflessione"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton Normale 
         Caption         =   "Normale"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Trasparenza 
         Caption         =   "Trasparenza"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   600
      Top             =   4440
   End
   Begin MSComDlg.CommonDialog ControlloSceltaColori 
      Left            =   120
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame FrameOpzioniTrasparenza 
      Caption         =   "Opzioni Colori Multipli Trasparenza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   3615
      Begin VB.PictureBox RapMPS 
         Height          =   2175
         Left            =   1200
         ScaleHeight     =   2115
         ScaleWidth      =   1155
         TabIndex        =   37
         Top             =   1440
         Width           =   1215
         Begin VB.Shape AngoloNO 
            BorderStyle     =   0  'Transparent
            FillStyle       =   0  'Solid
            Height          =   1095
            Left            =   0
            Top             =   0
            Width           =   615
         End
         Begin VB.Shape AngoloNE 
            BorderStyle     =   0  'Transparent
            FillStyle       =   0  'Solid
            Height          =   1095
            Left            =   600
            Top             =   0
            Width           =   615
         End
         Begin VB.Shape AngoloSO 
            BorderStyle     =   0  'Transparent
            FillStyle       =   0  'Solid
            Height          =   1095
            Left            =   0
            Top             =   1080
            Width           =   615
         End
         Begin VB.Shape AngoloSE 
            BorderStyle     =   0  'Transparent
            FillStyle       =   0  'Solid
            Height          =   1095
            Left            =   600
            Top             =   1080
            Width           =   615
         End
      End
      Begin MSComctlLib.Slider SoliditàNO 
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   840
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         LargeChange     =   1
         SelStart        =   5
         Value           =   5
      End
      Begin VB.CommandButton AnnullaAssegnazioneColori 
         Caption         =   "Annulla tutto"
         Height          =   255
         Left            =   1200
         TabIndex        =   11
         Top             =   4440
         Width           =   1215
      End
      Begin VB.CommandButton CancellaSE 
         Caption         =   "Cancella"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton CancellaSO 
         Caption         =   "Cancella"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton CancellaNE 
         Caption         =   "Cancella"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton CancellaNO 
         Caption         =   "Cancella"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton CambiaSE 
         Caption         =   "Cambia"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   3240
         Width           =   855
      End
      Begin VB.CommandButton CambiaSO 
         Caption         =   "Cambia"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   3240
         Width           =   855
      End
      Begin VB.CommandButton CambiaNE 
         Caption         =   "Cambia"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2640
         TabIndex        =   3
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton CambiaNO 
         Caption         =   "Cambia"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   855
      End
      Begin MSComctlLib.Slider SoliditàNE 
         Height          =   255
         Left            =   1800
         TabIndex        =   13
         Top             =   840
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         LargeChange     =   1
         SelStart        =   5
         Value           =   5
      End
      Begin MSComctlLib.Slider SoliditàSO 
         Height          =   255
         Left            =   1080
         TabIndex        =   14
         Top             =   3960
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         LargeChange     =   1
         SelStart        =   5
         Value           =   5
      End
      Begin MSComctlLib.Slider SoliditàSE 
         Height          =   255
         Left            =   1800
         TabIndex        =   15
         Top             =   3960
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         LargeChange     =   1
         SelStart        =   5
         Value           =   5
      End
      Begin VB.Line Line16 
         X1              =   3000
         X2              =   3000
         Y1              =   3840
         Y2              =   4080
      End
      Begin VB.Line Line15 
         X1              =   2520
         X2              =   3000
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Line Line14 
         X1              =   600
         X2              =   1080
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Line Line13 
         X1              =   600
         X2              =   600
         Y1              =   3840
         Y2              =   4080
      End
      Begin VB.Line Line12 
         X1              =   3000
         X2              =   2520
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line11 
         X1              =   3000
         X2              =   3000
         Y1              =   1200
         Y2              =   960
      End
      Begin VB.Line Line10 
         X1              =   600
         X2              =   1080
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line9 
         X1              =   600
         X2              =   600
         Y1              =   960
         Y2              =   1200
      End
      Begin VB.Line Line8 
         X1              =   3000
         X2              =   3000
         Y1              =   2760
         Y2              =   3120
      End
      Begin VB.Line Line7 
         X1              =   2520
         X2              =   3000
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line6 
         X1              =   600
         X2              =   1080
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line5 
         X1              =   600
         X2              =   600
         Y1              =   2760
         Y2              =   3120
      End
      Begin VB.Line Line4 
         X1              =   600
         X2              =   1080
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line3 
         X1              =   600
         X2              =   600
         Y1              =   1920
         Y2              =   2280
      End
      Begin VB.Line Line2 
         X1              =   2520
         X2              =   3000
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line1 
         X1              =   3000
         X2              =   3000
         Y1              =   1920
         Y2              =   2280
      End
      Begin VB.Label LabeLInformazioni 
         Caption         =   "Seleziona i colori da assegnare ai rispettivi vertici:"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.CommandButton Esci 
      Caption         =   "Esci"
      Height          =   255
      Left            =   5880
      TabIndex        =   0
      Top             =   5040
      Width           =   975
   End
End
Attribute VB_Name = "Form_Materiali"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CambiamentoSolidità As Boolean
Dim CambiamentoAlpha As Boolean
Dim Attivati As Boolean
'Dichiaro sempre le stesse Tre variabili addette a contenere la quantità di
'Rosso, Verde e Blu degli spigoli del muro, pavimento o soffitto selezionato
Dim R As Integer
Dim G As Integer
Dim B As Integer
Dim A As Integer

Private Sub AlphaBAnbiente_Scroll()
    'Assegno alla variabile CambiamentoAlpha il valore booleano True, in modo da far capire
    'al programma che dovrà aggiornare in tempo reale tutti i rispettivi valori Alpha
    CambiamentoAlpha = True
End Sub

Private Sub AlphaBDiffusa_Scroll()
    'Assegno alla variabile CambiamentoAlpha il valore booleano True, in modo da far capire
    'al programma che dovrà aggiornare in tempo reale tutti i rispettivi valori Alpha
    CambiamentoAlpha = True
End Sub

Private Sub AlphaBEmissiva_Scroll()
    'Assegno alla variabile CambiamentoAlpha il valore booleano True, in modo da far capire
    'al programma che dovrà aggiornare in tempo reale tutti i rispettivi valori Alpha
    CambiamentoAlpha = True
End Sub

Private Sub AlphaBSpeculare_Scroll()
    'Assegno alla variabile CambiamentoAlpha il valore booleano True, in modo da far capire
    'al programma che dovrà aggiornare in tempo reale tutti i rispettivi valori Alpha
    CambiamentoAlpha = True
End Sub

Private Sub AnnullaAssegnazioneColori_Click()
    'Richiamo la funzione che mi permetterà di reimpostare a nulli (neri) i colori di tutti i vertici del muro,pavimento o soffitto
    'selezionato
    Reimposta_tutto
End Sub

Private Sub AnnullaTPM_Click()
    'Se il form è stato avviato al fine di settare il materiale di un muro,allora
    If ModalitàGestioneMateriale = "Muri" Then
        'Richiamo la funzione pubblica Reimposta_materiale, in modo che tutti i parametri di
        'materiale del muro corrente,vengano risettati a nulli
        Reimposta_materiale Riga(IndiceLista).Materiale
        'Richiamo la funzione addetta al caricamento del materiale del muro selezionato,al fine di mostrare all'utente
        'che la reimpostazione è avvenuta con successo
        CaricaMateriale 1
    Else
        'Richiamo la funzione pubblica Reimposta_materiale, in modo che tutti i parametri di
        'materiale del pavimento o soffitto corrente,vengano risettati a nulli
        Reimposta_materiale SoP(IndiceLista2).CR.Materiale
        'Richiamo la funzione addetta al caricamento del materiale del pavimento o soffitto selezionato,al fine di mostrare all'utente
        'che la reimpostazione è avvenuta con successo
        CaricaMateriale 2
    End If
End Sub

Private Sub CambiaNE_Click()
    'Apro la finestra di dilaogo dal componente ControlloSceltaColori che mi permetterà appunto
    'di selezionare il colore da ssegnare all'angolo posto a Nord - Ovest del muro, pavimento o soffitto selezionato
    On Error GoTo Annulla
    ControlloSceltaColori.ShowColor
    If ModalitàGestioneMateriale = "Muri" Then
        'Richiamo la funzione che mi servirà per convertire il colore appena selezionato in valori RGB, passandogli
        'come valore appunto il vertice selezionato
        Preleva_RGB Riga(IndiceLista).ColVertici(1)
    Else
        'Richiamo la funzione che mi servirà per convertire il colore appena selezionato in valori RGB, passandogli
        'come valore appunto il vertice selezionato
        Preleva_RGB SoP(IndiceLista2).CR.ColVertici(1)
    End If
    'Assegno alla rispettiva Shape i rispettivi valori RGB del colore appena selezionato
    AngoloNE.FillColor = ControlloSceltaColori.Color
Annulla:
End Sub

Private Sub CambiaNO_Click()
    'Apro la finestra di dilaogo dal componente ControlloSceltaColori che mi permetterà appunto
    'di selezionare il colore da ssegnare all'angolo posto a Nord - Ovest del muro, pavimento o soffitto selezionato
    On Error GoTo Annulla
    ControlloSceltaColori.ShowColor
    If ModalitàGestioneMateriale = "Muri" Then
        'Richiamo la funzione che mi servirà per convertire il colore appena selezionato in valori RGB, passandogli
        'come valore appunto il vertice selezionato
        Preleva_RGB Riga(IndiceLista).ColVertici(0)
    Else
        'Richiamo la funzione che mi servirà per convertire il colore appena selezionato in valori RGB, passandogli
        'come valore appunto il vertice selezionato
        Preleva_RGB SoP(IndiceLista2).CR.ColVertici(0)
    End If
    'Assegno alla rispettiva Shape i rispettivi valori RGB del colore appena selezionato
    AngoloNO.FillColor = ControlloSceltaColori.Color
Annulla:
End Sub

Private Sub CambiaSE_Click()
    'Apro la finestra di dilaogo dal componente ControlloSceltaColori che mi permetterà appunto
    'di selezionare il colore da ssegnare all'angolo posto a Nord - Ovest del muro, pavimento o soffitto selezionato
    On Error GoTo Annulla
    ControlloSceltaColori.ShowColor
    If ModalitàGestioneMateriale = "Muri" Then
        'Richiamo la funzione che mi servirà per convertire il colore appena selezionato in valori RGB, passandogli
        'come valore appunto il vertice selezionato
        Preleva_RGB Riga(IndiceLista).ColVertici(3)
    Else
        'Richiamo la funzione che mi servirà per convertire il colore appena selezionato in valori RGB, passandogli
        'come valore appunto il vertice selezionato
        Preleva_RGB SoP(IndiceLista2).CR.ColVertici(3)
    End If
    'Assegno alla rispettiva Shape i rispettivi valori RGB del colore appena selezionato
    AngoloSE.FillColor = ControlloSceltaColori.Color
Annulla:
End Sub

Private Sub CambiaSO_Click()
    'Apro la finestra di dilaogo dal componente ControlloSceltaColori che mi permetterà appunto
    'di selezionare il colore da ssegnare all'angolo posto a Nord - Ovest del muro, pavimento o soffitto selezionato
    On Error GoTo Annulla
    ControlloSceltaColori.ShowColor
    If ModalitàGestioneMateriale = "Muri" Then
        'Richiamo la funzione che mi servirà per convertire il colore appena selezionato in valori RGB, passandogli
        'come valore appunto il vertice selezionato
        Preleva_RGB Riga(IndiceLista).ColVertici(2)
    Else
        'Richiamo la funzione che mi servirà per convertire il colore appena selezionato in valori RGB, passandogli
        'come valore appunto il vertice selezionato
        Preleva_RGB SoP(IndiceLista2).CR.ColVertici(2)
    End If
    'Assegno alla rispettiva Shape i rispettivi valori RGB del colore appena selezionato
    AngoloSO.FillColor = ControlloSceltaColori.Color
Annulla:
End Sub

Private Sub CancellaNE_Click()
    'Chiamo la funzione che mi permette di reimpostare, in base al valore che le viene passato, il
    'colore del vertice del muro,pavimento o soffitto desiderato
    Reimposta_Colore 1
End Sub

Private Sub CancellaNO_Click()
    'Chiamo la funzione che mi permette di reimpostare, in base al valore che le viene passato, il
    'colore del vertice del muro,pavimento o soffitto desiderato
    Reimposta_Colore 0
End Sub

Private Sub CancellaSE_Click()
    'Chiamo la funzione che mi permette di reimpostare, in base al valore che le viene passato, il
    'colore del vertice del muro,pavimento o soffitto desiderato
    Reimposta_Colore 3
End Sub

Private Sub CancellaSO_Click()
    'Chiamo la funzione che mi permette di reimpostare, in base al valore che le viene passato, il
    'colore del vertice del muro,pavimento o soffitto desiderato
    Reimposta_Colore 2
End Sub

Private Sub Conferma_Click()
    If ModalitàGestioneMateriale = "Muri" Then
        If SphereMapping.Value = True Then
            Riga(IndiceLista).Proprietà = "SphereMapping"
            Reimposta_tutto
        ElseIf Trasparenza.Value = True Then
            Riga(IndiceLista).Proprietà = "Trasparenza"
        ElseIf Normale.Value = True Then
            Riga(IndiceLista).Proprietà = "Normale"
            Reimposta_tutto
        End If
    Else
        If SphereMapping.Value = True Then
            SoP(IndiceLista2).CR.Proprietà = "SphereMapping"
            Reimposta_tutto
        ElseIf Trasparenza.Value = True Then
            SoP(IndiceLista2).CR.Proprietà = "Trasparenza"
        ElseIf Normale.Value = True Then
            SoP(IndiceLista2).CR.Proprietà = "Normale"
            Reimposta_tutto
        End If
    End If
End Sub

Private Sub Esci_Click()
    'Scarica se stesso, cioè chiude il form di gestione dei materiali (Form_Materiali)
    Unload Me
End Sub

Private Sub Form_Load()
    'Richiamo la funzione addetta alla traduzione nella lingua selezionata dall'utente
    'del form stesso
    Traduci LinguaS
    'Inizializzo la variabile CambiaMentoSolidità un valore iniziale uguale a False
    CambiamentoSolidità = False
    'Richiamo la funzione addetta al caricamento delle proprietà
    Carica_proprietà
    'Se il form è stato avviato in modalità "Muri" allora...
    If ModalitàGestioneMateriale = "Muri" Then
        If RTrim(Riga(IndiceLista).Proprietà) = "Trasparenza" Then
            CaricaColoriTrasparenza 1
        Else: Reimposta_tutto
        End If
        'Richiamo la funzione addetta al caricamento del materiale.
        'Il parametro passato è uguale a 1 per fare in modo che venga caricato solo il
        'materiale del muro selezionato
        CaricaMateriale 1
    Else
        If RTrim(SoP(IndiceLista2).CR.Proprietà) = "Trasparenza" Then
            CaricaColoriTrasparenza 2
        Else: Reimposta_tutto
        End If
        'Richiamo la funzione addetta al caricamento del materiale.
        'Il parametro passato è uguale a 2 per fare in modo che venga caricato solo il
        'materiale del pavimento o soffitto selezionato
        CaricaMateriale 2
    End If
End Sub

Sub Attiva_Elementi_Trasparenza()
    'Attivo il pulsante per la modifica del colore del vertice posto a Nord - Ovest
    CambiaNO.Enabled = True
    'Attivo il pulsante per la modifica del colore del vertice posto a Nord - Est
    CambiaNE.Enabled = True
    'Attivo il pulsante per la modifica del colore del vertice posto a Sud - Ovest
    CambiaSO.Enabled = True
    'Attivo il pulsante per la modifica del colore del vertice posto a Sud - Est
    CambiaSE.Enabled = True
    'Attivo il pulsante addetto alla cancellazione del colore assegnato al vertice posto a Nord - Ovest
    CancellaNO.Enabled = True
    'Attivo il pulsante addetto alla cancellazione del colore assegnato al vertice posto a Nord - Est
    CancellaNE.Enabled = True
    'Attivo il pulsante addetto alla cancellazione del colore assegnato al vertice posto a Sud - Ovest
    CancellaSO.Enabled = True
    'Attivo il pulsante addetto alla cancellazione del colore assegnato al vertice posto a Sud - Est
    CancellaSE.Enabled = True
    'Attivo il pulsante addetto a richiamare la funzione per il ripristino totale dei colori di tutti i vertici
    'del muro, pavimento o soffitto selezionato
    AnnullaAssegnazioneColori.Enabled = True
    'Attivo lo slider che addetto ad impostare il grado di solidità del vertice posto a Nord - Ovest
    SoliditàNO.Enabled = True
    'Attivo lo slider che addetto ad impostare il grado di solidità del vertice posto a Nord - Est
    SoliditàNE.Enabled = True
    'Attivo lo slider che addetto ad impostare il grado di solidità del vertice posto a Sud - Ovest
    SoliditàSO.Enabled = True
    'Attivo lo slider che addetto ad impostare il grado di solidità del vertice posto a Sud - Est
    SoliditàSE.Enabled = True
    'Attivo l'oggetto PictureBox rappresentante il muro,pavimento o soffitto selezionato, contenente
    'le quattro shape che rappresentano a loro volta ognuna il vertice corrispondente
    RapMPS.Enabled = True
    'Assegno alla variabile Attivati il valore booleano uguale a True in modo da far capire al programma che
    'tutti gli elementi sopraindicati sono già stati attivati
    Attivati = True
End Sub

Sub Disattiva_Elementi_Trasparenza()
    'Disattivo il pulsante per la modifica del colore del vertice posto a Nord - Ovest
    CambiaNO.Enabled = False
    'Disattivo il pulsante per la modifica del colore del vertice posto a Nord - Est
    CambiaNE.Enabled = False
    'Disattivo il pulsante per la modifica del colore del vertice posto a Sud - Ovest
    CambiaSO.Enabled = False
    'Disttivo il pulsante per la modifica del colore del vertice posto a Sud - Est
    CambiaSE.Enabled = False
    'Disattivo il pulsante addetto alla cancellazione del colore assegnato al vertice posto a Nord - Ovest
    CancellaNO.Enabled = False
    'Disattivo il pulsante addetto alla cancellazione del colore assegnato al vertice posto a Nord - Est
    CancellaNE.Enabled = False
    'Disattivo il pulsante addetto alla cancellazione del colore assegnato al vertice posto a Sud - Ovest
    CancellaSO.Enabled = False
    'Disattivo il pulsante addetto alla cancellazione del colore assegnato al vertice posto a Sud - Est
    CancellaSE.Enabled = False
    'Disattivo il pulsante addetto a richiamare la funzione per il ripristino totale dei colori di tutti i vertici
    'del muro, pavimento o soffitto selezionato
    AnnullaAssegnazioneColori.Enabled = False
    'Disattivo lo slider che addetto ad impostare il grado di solidità del vertice posto a Nord - Ovest
    SoliditàNO.Enabled = False
    'Disattivo lo slider che addetto ad impostare il grado di solidità del vertice posto a Nord - Est
    SoliditàNE.Enabled = False
    'Disattivo lo slider che addetto ad impostare il grado di solidità del vertice posto a Sud - Ovest
    SoliditàSO.Enabled = False
    'Disattivo lo slider che addetto ad impostare il grado di solidità del vertice posto a Sud - Est
    SoliditàSE.Enabled = False
    'Disattivo l'oggetto PictureBox rappresentante il muro,pavimento o soffitto selezionato, contenente
    'le quattro shape che rappresentano a loro volta ognuna il vertice corrispondente
    RapMPS.Enabled = False
    'Assegno alla variabile Attivati il valore booleano uguale a False in modo da far capire al programma che
    'tutti gli elementi sopraindicati sono già stati disattivati
    Attivati = False
End Sub

Private Sub RapAmbiente_Click()
    'Apro il componente che mi permetterà di selezionare il colore da assegnare alla
    'voce Materiale / Ambiente del muro,pavimento o soffitto selezionato
    On Error GoTo Annulla
    ControlloSceltaColori.ShowColor
    If ModalitàGestioneMateriale = "Muri" Then
        'Richiamo la funzione Preleva_RGB passandogli come unico valore 0.
        'Questo parametro, elaborato dalla funzione stessa mi permetterà di impostare le
        'rispettive quantità di Rosso, Verde e Blu del colore selezionato alla voce Materiale / Ambiente
        Preleva_RGB Riga(IndiceLista).Materiale.Ambiente
    Else
        'Richiamo la funzione Preleva_RGB passandogli come unico valore 0.
        'Questo parametro, elaborato dalla funzione stessa mi permetterà di impostare le
        'rispettive quantità di Rosso, Verde e Blu del colore selezionato alla voce Materiale / Ambiente
        Preleva_RGB SoP(IndiceLista2).CR.Materiale.Ambiente
    End If
    'Infine assegno alla PictureBox che rappresenta il colore Ambientale, le rispettive quantità di colore
    'appena assegnate
    RapAmbiente.BackColor = ControlloSceltaColori.Color
Annulla:
End Sub

Private Sub RapDiffusa_Click()
    'Apro il componente che mi permetterà di selezionare il colore da assegnare alla
    'voce Materiale / Diffusa del muro,pavimento o soffitto selezionato
    On Error GoTo Annulla
    ControlloSceltaColori.ShowColor
    If ModalitàGestioneMateriale = "Muri" Then
        'Richiamo la funzione Preleva_RGB passandogli come unico valore 1.
        'Questo parametro, elaborato dalla funzione stessa mi permetterà di impostare le
        'rispettive quantità di Rosso, Verde e Blu del colore selezionato alla voce Materiale / Diffusa
        Preleva_RGB Riga(IndiceLista).Materiale.Diffusa
    Else
        'Richiamo la funzione Preleva_RGB passandogli come unico valore 0.
        'Questo parametro, elaborato dalla funzione stessa mi permetterà di impostare le
        'rispettive quantità di Rosso, Verde e Blu del colore selezionato alla voce Materiale / Ambiente
        Preleva_RGB SoP(IndiceLista2).CR.Materiale.Diffusa
    End If
    'Infine assegno alla PictureBox che rappresenta il colore Diffuso, le rispettive quantità di colore
    'appena assegnate
    RapDiffusa.BackColor = ControlloSceltaColori.Color
Annulla:
End Sub


Private Sub RapEmissiva_Click()
    'Apro il componente che mi permetterà di selezionare il colore da assegnare alla
    'voce Materiale / Emissiva del muro,pavimento o soffitto selezionato
    On Error GoTo Annulla
    ControlloSceltaColori.ShowColor
    If ModalitàGestioneMateriale = "Muri" Then
        'Richiamo la funzione Preleva_RGB passandogli come unico valore 0.
        'Questo parametro, elaborato dalla funzione stessa mi permetterà di impostare le
        'rispettive quantità di Rosso, Verde e Blu del colore selezionato alla voce Materiale / Emissiva
        Preleva_RGB Riga(IndiceLista).Materiale.Emissiva
    Else
        'Richiamo la funzione Preleva_RGB passandogli come unico valore 0.
        'Questo parametro, elaborato dalla funzione stessa mi permetterà di impostare le
        'rispettive quantità di Rosso, Verde e Blu del colore selezionato alla voce Materiale / Ambiente
        Preleva_RGB SoP(IndiceLista2).CR.Materiale.Emissiva
    End If
    'Infine assegno alla PictureBox che rappresenta il colore Emissivo, le rispettive quantità di colore
    'appena assegnate
    RapEmissiva.BackColor = ControlloSceltaColori.Color
Annulla:
End Sub

Private Sub RapSpeculare_Click()
    'Apro il componente che mi permetterà di selezionare il colore da assegnare alla
    'voce Materiale / Speculare del muro,pavimento o soffitto selezionato
    On Error GoTo Annulla
    ControlloSceltaColori.ShowColor
    If ModalitàGestioneMateriale = "Muri" Then
        'Richiamo la funzione Preleva_RGB passandogli come unico valore 3.
        'Questo parametro, elaborato dalla funzione stessa mi permetterà di impostare le
        'rispettive quantità di Rosso, Verde e Blu del colore selezionato alla voce Materiale / Speculare
        Preleva_RGB Riga(IndiceLista).Materiale.Speculare
    Else
        'Richiamo la funzione Preleva_RGB passandogli come unico valore 0.
        'Questo parametro, elaborato dalla funzione stessa mi permetterà di impostare le
        'rispettive quantità di Rosso, Verde e Blu del colore selezionato alla voce Materiale / Ambiente
        Preleva_RGB SoP(IndiceLista2).CR.Materiale.Speculare
    End If
    'Infine assegno alla PictureBox che rappresenta il colore Speculare, le rispettive quantità di colore
    'appena assegnate
    RapSpeculare.BackColor = ControlloSceltaColori.Color
Annulla:
End Sub

Private Sub SoliditàNE_Scroll()
    'Assegno alla variabile CambiamentoSolidità il valore boolenao True, in modo da far capire al programma
    'che dovrà aggiornare il grado di solidità del rispettivo vertice
    CambiamentoSolidità = True
End Sub

Private Sub SoliditàNO_Scroll()
    'Assegno alla variabile CambiamentoSolidità il valore boolenao True, in modo da far capire al programma
    'che dovrà aggiornare il grado di solidità del rispettivo vertice
    CambiamentoSolidità = True
End Sub

Private Sub SoliditàSE_Scroll()
    'Assegno alla variabile CambiamentoSolidità il valore boolenao True, in modo da far capire al programma
    'che dovrà aggiornare il grado di solidità del rispettivo vertice
    CambiamentoSolidità = True
End Sub

Private Sub SoliditàSO_Scroll()
    'Assegno alla variabile CambiamentoSolidità il valore boolenao True, in modo da far capire al programma
    'che dovrà aggiornare il grado di solidità del rispettivo vertice
    CambiamentoSolidità = True
End Sub

Private Sub Timer1_Timer()
    'Se l'option button trasparenza selezionato, cioè se è stata assegnata al muro, pavimento o soffitto
    'la proprietà di trasparenza,allora richiamerò la funzione Attiva_elementi_trasparenza che attiverà tutti quegli
    'oggetti che mi serviranno per modificare i colori dei vertici dei muri,pavimenti o soffitti con
    'proprietà Trasparenza
    If Trasparenza.Value = True And Attivati = False Then
        Attiva_Elementi_Trasparenza
    'Se invece questà opzione non è selezionato, richiamerò al contrario quella funzione Disattiva_elementi_trasparenza che
    'disattiverà gli stessi oggetti.
    'Sarebbe inutile assegnare i colori ad un muro che non è trasparente
    ElseIf Trasparenza.Value = False And Attivati = True Then
        Disattiva_Elementi_Trasparenza
    End If
    'Se il valore di qualche slider indicate il grado di solidità del rispettivo vertice è cambiato,allora
    'Verrà richiamata la funzione che aggiornerà in tempo reale il valore di solidità del rispettivo vertice
    'del muro,pavimento o soffitto selezionato.
    'Infine verrà riportata la variabile CambiamentoSolidità al suo valore originario (False)
    If CambiamentoSolidità = True Then
        Assegna_solidità
        CambiamentoSolidità = False
    End If
    If CambiamentoAlpha = True Then
        Assegna_alpha
        CambiamentoAlpha = False
    End If
End Sub

Sub Reimposta_Colore(Angolo As Integer)
    If ModalitàGestioneMateriale = "Muri" Then
        'Questa funzione mi permette, in base al parametro che le viene passato, di reimpostare a nullo (nero),
        'il colore del vertice desiderato del muro selezionato
        With Riga(IndiceLista).ColVertici(Angolo)
            'Reimposto la quantità di colore Rosso a 0
            .R = 0
            'Reimposto la quantità di colore Verde a 0
            .G = 0
            'Reimposto la quantità di colore Blu a 0
            .B = 0
        End With
        Select Case Angolo
        Case Is = 0
            AngoloNO.FillColor = RGB(0, 0, 0)
        Case Is = 1
            AngoloNE.FillColor = RGB(0, 0, 0)
        Case Is = 2
            AngoloSO.FillColor = RGB(0, 0, 0)
        Case Is = 3
            AngoloSE.FillColor = RGB(0, 0, 0)
        End Select
    Else
        'Questa funzione mi permette, in base al parametro che le viene passato, di reimpostare a nullo (nero),
        'il colore del vertice desiderato del pavimento o soffitto selezionato
        With SoP(IndiceLista2).CR.ColVertici(Angolo)
            'Reimposto la quantità di colore Rosso a 0
            .R = 0
            'Reimposto la quantità di colore Verde a 0
            .G = 0
            'Reimposto la quantità di colore Blu a 0
            .B = 0
        End With
        Select Case Angolo
        Case Is = 0
            AngoloNO.FillColor = RGB(0, 0, 0)
        Case Is = 1
            AngoloNE.FillColor = RGB(0, 0, 0)
        Case Is = 2
            AngoloSO.FillColor = RGB(0, 0, 0)
        Case Is = 3
            AngoloSE.FillColor = RGB(0, 0, 0)
        End Select
    End If
End Sub

Sub Reimposta_tutto()
    'Richiamo 4 volte la stessa funzione (Reimposta_Colore), passandogli però ogni volta un valore diverso:
    ' - La prima volta gli passerò il valore 0 per reimpostare a nullo (nero) il colore del vertice posto a Nord - Ovest del muro,
    '   pavimento o soffitto slezionato e la rispettiva Shape che lo rappresenta
    ' - La seconda volta gli passerò il valore 1 per reimpostare a nullo (nero) il colore del vertice posto a Nord - Est del muro,
    '   pavimento o soffitto slezionato e la rispettiva Shape che lo rappresenta
    ' - La terza volta gli passerò il valore 2 per reimpostare a nullo (nero) il colore del vertice posto a Sud - Ovest del muro,
    '   pavimento o soffitto slezionato e la rispettiva Shape che lo rappresenta
    ' - La seconda volta gli passerò il valore 3 per reimpostare a nullo (nero) il colore del vertice posto a Sud - Est del muro,
    '   pavimento o soffitto slezionato e la rispettiva Shape che lo rappresenta
    Reimposta_Colore 0
    Reimposta_Colore 1
    Reimposta_Colore 2
    Reimposta_Colore 3
End Sub

Sub Assegna_solidità()
    If ModalitàGestioneMateriale = "Muri" Then
        'Assegno al vertice posto a Nord - Ovest il livello di solidità selezionato dal rispettivo slider
        Riga(IndiceLista).ColVertici(0).A = SoliditàNO.Value / 10
        'Assegno al vertice posto a Nord - Est il livello di solidità selezionato dal rispettivo slider
        Riga(IndiceLista).ColVertici(1).A = SoliditàNE.Value / 10
        'Assegno al vertice posto a Sud - Ovest il livello di solidità selezionato dal rispettivo slider
        Riga(IndiceLista).ColVertici(2).A = SoliditàSO.Value / 10
        'Assegno al vertice posto a Sud - Est il livello di solidità selezionato dal rispettivo slider
        Riga(IndiceLista).ColVertici(3).A = SoliditàSE.Value / 10
    Else
        'Assegno al vertice posto a Nord - Ovest il livello di solidità selezionato dal rispettivo slider
        SoP(IndiceLista2).CR.ColVertici(0).A = SoliditàNO.Value / 10
        'Assegno al vertice posto a Nord - Est il livello di solidità selezionato dal rispettivo slider
        SoP(IndiceLista2).CR.ColVertici(1).A = SoliditàNE.Value / 10
        'Assegno al vertice posto a Sud - Ovest il livello di solidità selezionato dal rispettivo slider
        SoP(IndiceLista2).CR.ColVertici(2).A = SoliditàSO.Value / 10
        'Assegno al vertice posto a Sud - Est il livello di solidità selezionato dal rispettivo slider
        SoP(IndiceLista2).CR.ColVertici(3).A = SoliditàSE.Value / 10
    End If
End Sub

Sub Assegna_alpha()
    If ModalitàGestioneMateriale = "Muri" Then
        With Riga(IndiceLista).Materiale
            'Assegno al parametro del materiale Ambiente creato per il muro selezionato, un valore di AlphaBleding
            'pari al valore del rispettivo slider che lo rappresenta diviso 10
            .Ambiente.A = AlphaBAnbiente.Value / 10
            'Assegno al parametro del materiale Diffuso creato per il muro selezionato, un valore di AlphaBleding
            'pari al valore del rispettivo slider che lo rappresenta diviso 10
            .Diffusa.A = AlphaBDiffusa.Value / 10
            'Assegno al parametro del materiale Emissivo creato per il muro selezionato, un valore di AlphaBleding
            'pari al valore del rispettivo slider che lo rappresenta diviso 10
            .Emissiva.A = AlphaBEmissiva.Value / 10
            'Assegno al parametro del materiale Speculare creato per il muro selezionato, un valore di AlphaBleding
            'pari al valore del rispettivo slider che lo rappresenta diviso 10
            .Speculare.A = AlphaBSpeculare.Value / 10
            'Assegno al rispettivo parametro del materiale assegnato del muro selezionato, il rispettivo
            'valore di Potenza
            .Potenza = Val(ValorePotenza)
        End With
    Else
        With SoP(IndiceLista2).CR.Materiale
            'Assegno al parametro del materiale Ambiente creato per il pavimento o soffitto selezionato, un valore di AlphaBleding
            'pari al valore del rispettivo slider che lo rappresenta diviso 10
            .Ambiente.A = AlphaBAnbiente.Value / 10
            'Assegno al parametro del materiale Diffuso creato per il muro selezionato, un valore di AlphaBleding
            'pari al valore del rispettivo slider che lo rappresenta diviso 10
            .Diffusa.A = AlphaBDiffusa.Value / 10
            'Assegno al parametro del materiale Emissivo creato per il pavimento o soffitto selezionato, un valore di AlphaBleding
            'pari al valore del rispettivo slider che lo rappresenta diviso 10
            .Emissiva.A = AlphaBEmissiva.Value / 10
            'Assegno al parametro del materiale Speculare creato per il pavimento o soffitto selezionato, un valore di AlphaBleding
            'pari al valore del rispettivo slider che lo rappresenta diviso 10
            .Speculare.A = AlphaBSpeculare.Value / 10
            'Assegno al rispettivo parametro del materiale assegnato del pavimento o soffitto selezionato, il rispettivo
            'valore di Potenza
            .Potenza = Val(ValorePotenza)
        End With
    End If
End Sub

Private Sub ValorePotenza_Change()
    'Assegno alla variabile CambiamentoAlpha il valore booleano True, in modo da far capire
    'al programma che dovrà aggiornare in tempo reale il rispettivo parametro di Potenza
    'del materiale assegnato del muro selezionato
    CambiamentoAlpha = True
End Sub

Sub CaricaMateriale(TipoCaricamento As Integer)
    If TipoCaricamento = 1 Then
        'Ora Caricaro all'interno delle PictureBox che lo rappresentano, il colore
        'dei vari parametri del materiale assegnato
        With Riga(IndiceLista).Materiale.Ambiente
            'Setto le variabili R G B con le rispettive quantità di colore presenti
            'nel parametro Ambiente del materiale del muro selezionato
            R = .R
            G = .G
            B = .B
            A = .A * 10
        End With
        'Carico all'interno della PictureBox RapAmbiente, la quale rappresenta il colore
        'ambientale,i valori RGB appena settati
        RapAmbiente.BackColor = RGB(R, G, B)
        'Assegno allo slider che rappresenta il valore di luminosità del parametro Ambiente
        'presente nel materiale assegnato al muro selezionato
        AlphaBAnbiente.Value = A
        'Setto le variabili R G B con le rispettive quantità di colore presenti
        'nel parametro Diffusa del materiale del muro selezionato
        With Riga(IndiceLista).Materiale.Diffusa
            R = .R
            G = .G
            B = .B
            A = .A * 10
        End With
        'Carico all'interno della PictureBox RapDiffusa, la quale rappresenta il colore
        'Diffuso,i valori RGB appena settati
        RapDiffusa.BackColor = RGB(R, G, B)
        'Assegno allo slider che rappresenta il valore di luminosità del parametro Diffusa
        'presente nel materiale assegnato al muro selezionato
        AlphaBDiffusa.Value = A
        'Setto le variabili R G B con le rispettive quantità di colore presenti
        'nel parametro Emissiva del materiale del muro selezionato
        With Riga(IndiceLista).Materiale.Emissiva
            R = .R
            G = .G
            B = .B
            A = .A * 10
        End With
        'Carico all'interno della PictureBox RapEmissiva, la quale rappresenta il colore
        'Emissivo,i valori RGB appena settati
        RapEmissiva.BackColor = RGB(R, G, B)
        'Assegno allo slider che rappresenta il valore di luminosità del parametro Emissiva
        'presente nel materiale assegnato al muro selezionato
        AlphaBEmissiva.Value = A
        'Setto le variabili R G B con le rispettive quantità di colore presenti
        'nel parametro Speculare del materiale del muro selezionato
        With Riga(IndiceLista).Materiale.Speculare
            R = .R
            G = .G
            B = .B
            A = .A * 10
        End With
        'Carico all'interno della PictureBox RapSpeculare, la quale rappresenta il colore
        'Speculare,i valori RGB appena settati
        RapSpeculare.BackColor = RGB(R, G, B)
        'Assegno allo slider che rappresenta il valore di luminosità del parametro Speculare
        'presente nel materiale assegnato al muro selezionato
        AlphaBSpeculare.Value = A
        'Assegno alla Textbox ValorePotenza il rispettivo valore di potenza del materiale del
        'muro selezionato
        ValorePotenza = Str(Riga(IndiceLista).Materiale.Potenza)
    'Nel caso in cui invece il TipoCaricamento è diverso da 1,e quindi la ModalitàGestioneMateriali
    'è diverso da "Muri",verranno caricati tutti i valori di tutte le voci che compongono il materiale
    'del pavimento o soffitto selezionato
    Else
        'Ora Caricaro all'interno delle PictureBox che lo rappresentano, il colore
        'dei vari parametri del materiale assegnato
        With SoP(IndiceLista2).CR.Materiale.Ambiente
            'Setto le variabili R G B con le rispettive quantità di colore presenti
            'nel parametro Ambiente del materiale del pavimento o soffitto selezionato
            R = .R
            G = .G
            B = .B
            A = .A * 10
        End With
        'Carico all'interno della PictureBox RapAmbiente, la quale rappresenta il colore
        'ambientale,i valori RGB appena settati
        RapAmbiente.BackColor = RGB(R, G, B)
        'Assegno allo slider che rappresenta il valore di luminosità del parametro Ambiente
        'presente nel materiale assegnato al pavimento o soffitto selezionato
        AlphaBAnbiente.Value = A
        'Setto le variabili R G B con le rispettive quantità di colore presenti
        'nel parametro Diffusa del materiale del pavimento o soffitto selezionato
        With SoP(IndiceLista2).CR.Materiale.Diffusa
            R = .R
            G = .G
            B = .B
            A = .A * 10
        End With
        'Carico all'interno della PictureBox RapDiffusa, la quale rappresenta il colore
        'Diffuso,i valori RGB appena settati
        RapDiffusa.BackColor = RGB(R, G, B)
        'Assegno allo slider che rappresenta il valore di luminosità del parametro Diffusa
        'presente nel materiale assegnato al pavimento o soffitto selezionato
        AlphaBDiffusa.Value = A
        'Setto le variabili R G B con le rispettive quantità di colore presenti
        'nel parametro Emissiva del materiale del pavimento o sofitto selezionato
        With SoP(IndiceLista2).CR.Materiale.Emissiva
            R = .R
            G = .G
            B = .B
            A = .A * 10
        End With
        'Carico all'interno della PictureBox RapEmissiva, la quale rappresenta il colore
        'Emissivo,i valori RGB appena settati
        RapEmissiva.BackColor = RGB(R, G, B)
        'Assegno allo slider che rappresenta il valore di luminosità del parametro Emissiva
        'presente nel materiale assegnato al pavimento o soffitto selezionato
        AlphaBEmissiva.Value = A
        'Setto le variabili R G B con le rispettive quantità di colore presenti
        'nel parametro Speculare del materiale del pavimento o soffitto selezionato
        With SoP(IndiceLista2).CR.Materiale.Speculare
            R = .R
            G = .G
            B = .B
            A = .A * 10
        End With
        'Carico all'interno della PictureBox RapSpeculare, la quale rappresenta il colore
        'Speculare,i valori RGB appena settati
        RapSpeculare.BackColor = RGB(R, G, B)
        'Assegno allo slider che rappresenta il valore di luminosità del parametro Speculare
        'presente nel materiale assegnato al muro selezionato
        AlphaBSpeculare.Value = A
        'Assegno alla Textbox ValorePotenza il rispettivo valore di potenza del materiale del
        'pavimento o soffitto selezionato
        ValorePotenza = Str(SoP(IndiceLista2).CR.Materiale.Potenza)
    End If
End Sub

Sub CaricaColoriTrasparenza(TipoCaricamento As Integer)
    'Dichiaro un indice che mi faciliterà le operazioni di assegnazione dei colori
    'alle shape che rappresentano i rispettivi vertici
    Dim K As Integer
    'Se il TipoCaricamento è uguale a 1,e quindi la ModalitàGestioneMateriali
    'è uguale a "Muri",verranno caricati tutti i valori di tutte le voci che compongono il materiale
    'del muro selezionato
    If TipoCaricamento = 1 Then
        'Avvio un ciclo che assegnerà alle shape i colori dei rispettivi vertici
        'dei muri che rappresentano
        For K = 0 To 3
            R = Riga(IndiceLista).ColVertici(K).R
            G = Riga(IndiceLista).ColVertici(K).G
            B = Riga(IndiceLista).ColVertici(K).B
            A = Riga(IndiceLista).ColVertici(K).A * 10
        Select Case K
        Case Is = 0
            AngoloNO.FillColor = RGB(R, G, B)
            SoliditàNO.Value = A
        Case Is = 1
            AngoloNE.FillColor = RGB(R, G, B)
            SoliditàNE.Value = A
        Case Is = 2
            AngoloSO.FillColor = RGB(R, G, B)
            SoliditàSO.Value = A
        Case Is = 3
            AngoloSE.FillColor = RGB(R, G, B)
            SoliditàSE.Value = A
        End Select
        Next
    'Nel caso in cui invece il TipoCaricamento è diverso da 1,e quindi la ModalitàGestioneMateriali
    'è diverso da "Muri",verranno caricati tutti i colori di trasparenza dei rispettivi vertici
    'del pavimento o soffitto selezionato
    Else
        'Avvio un ciclo che assegnerà alle shape i colori dei rispettivi vertici
        'dei pavimenti o soffitti che rappresentano
        For K = 0 To 3
            R = SoP(IndiceLista2).CR.ColVertici(K).R
            G = SoP(IndiceLista2).CR.ColVertici(K).G
            B = SoP(IndiceLista2).CR.ColVertici(K).B
            A = SoP(IndiceLista2).CR.ColVertici(K).A * 10
        Select Case K
        Case Is = 0
            AngoloNO.FillColor = RGB(R, G, B)
            SoliditàNO.Value = A
        Case Is = 1
            AngoloNE.FillColor = RGB(R, G, B)
            SoliditàNE.Value = A
        Case Is = 2
            AngoloSO.FillColor = RGB(R, G, B)
            SoliditàSO.Value = A
        Case Is = 3
            AngoloSE.FillColor = RGB(R, G, B)
            SoliditàSE.Value = A
        End Select
        Next
    End If
End Sub

Sub Assegna_titolo()
    If ModalitàGestioneMateriale = "Muri" Then
        'Imposto il titolo del form, in modo da far capire all'utente a quale muro
        'verrà applicato il materiale che verrà creato e anche la scritta del LabelInformazioni
        If LinguaS = "Italiano" Then
            Me.Caption = "Gestione materiale del muro: " + Riga(IndiceLista).Nome
            LabeLInformazioni.Caption = "Seleziona i colori da assegnare ai rispettivi vertici del muro:"
        ElseIf LinguaS = "Inglese" Then
            Me.Caption = "Management of wall: " + Riga(IndiceLista).Nome
            LabeLInformazioni.Caption = "Select the colours to assign at the rispective wall corners:"
        End If
    ElseIf ModalitàGestioneMateriale = "Pavimento" Then
        If LinguaS = "Italiano" Then
            Me.Caption = "Gestione materiale del pavimento: " + RTrim(SoP(IndiceLista2).CR.Nome)
            LabeLInformazioni.Caption = "Seleziona i colori da assegnare ai rispettivi vertici del pavimento:"
        ElseIf LinguaS = "Inglese" Then
            Me.Caption = "Management of floor: " + RTrim(SoP(IndiceLista2).CR.Nome)
            LabeLInformazioni.Caption = "Select the colours to assign at the rispective floor corners:"
        End If
    Else
        If LinguaS = "Italiano" Then
            Me.Caption = "Gestione materiale del soffitto: " + RTrim(SoP(IndiceLista2).CR.Nome)
            LabeLInformazioni.Caption = "Seleziona i colori da assegnare ai rispettivi vertici del soffitto:"
        ElseIf LinguaS = "Inglese" Then
            Me.Caption = "Management of ceiling: " + RTrim(SoP(IndiceLista2).CR.Nome)
            LabeLInformazioni.Caption = "Select the colours to assign at the rispective ceiling corners:"
        End If
    End If
End Sub

Sub Carica_proprietà()
    If ModalitàGestioneMateriale = "Muri" Then
        If RTrim(Riga(IndiceLista).Proprietà) = "SphereMapping" Then
            SphereMapping.Value = True
        ElseIf RTrim(Riga(IndiceLista).Proprietà) = "Trasparenza" Then
            Trasparenza.Value = True
            'Richiamo la funzione che attiverà tutti quegli oggetti, come i bottoni
            'che mi permetteranno di modificare le opzioni di trasparenza del muro,
            'pavimento o soffitto selezionato
            Attiva_Elementi_Trasparenza
        ElseIf RTrim(Riga(IndiceLista).Proprietà) = "Normale" Then
            Normale.Value = True
        End If
    Else
        If RTrim(SoP(IndiceLista2).CR.Proprietà) = "SphereMapping" Then
            SphereMapping.Value = True
        ElseIf RTrim(SoP(IndiceLista2).CR.Proprietà) = "Trasparenza" Then
            Trasparenza.Value = True
            'Richiamo la funzione che attiverà tutti quegli oggetti, come i bottoni
            'che mi permetteranno di modificare le opzioni di trasparenza del muro,
            'pavimento o soffitto selezionato
            Attiva_Elementi_Trasparenza
        ElseIf RTrim(SoP(IndiceLista2).CR.Proprietà) = "Normale" Then
            Normale.Value = True
        End If
    End If
End Sub

Sub Traduci(NuovaLingua As String)
    'Richiamo la funzione addetta ad assegnare il titolo del form stesso e del LabelInformazione
    'in base alla modalità per cui questo è stato avviato
    Assegna_titolo
    'Verifico la lingua con cui dovrà essere tradotto il Form_Materiali
    Select Case NuovaLingua
    Case Is = "Italiano"
        'Traduco il FrameTipoMateriali e tutti gli oggetti comntenuti al suo interno
        With Me
            .FrameTipoMateriali = "Tipo Materiale:"
            .Normale.Caption = "Normale"
            .SphereMapping.Caption = "Falsa Riflessione"
            .Trasparenza.Caption = "Trasparenza"
            'Traduco il sotto frame FrameOpzioniMateriale e tutti gli oggetti contenuti al suo interno
            .FrameOpzioniMateriale = "Opzioni Materiale"
            .LabelAmbiente = "Ambiente:"
            .LabelDiffusa = "Diffusa:"
            .LabelEmissiva = "Emissiva:"
            .LabelSpeculare = "Speculare:"
            .LabelPotenza = "Potenza:"
            .AnnullaTPM.Caption = "Annulla tutto"
            .Conferma.Caption = "Conferma"
            'Traduco il FrameOpzioniTrasparenza e tutti gli oggetti contenuti al suo interno
            .FrameOpzioniTrasparenza = "Opzioni Colori Multipli Trasparenza"
            .CambiaNO.Caption = "Cambia"
            .CambiaNE.Caption = "Cambia"
            .CambiaSO.Caption = "Cambia"
            .CambiaSE.Caption = "Cambia"
            .CancellaNO.Caption = "Cancella"
            .CancellaNE.Caption = "Cancella"
            .CancellaSO.Caption = "Cancella"
            .CancellaSE.Caption = "Cancella"
            .AnnullaAssegnazioneColori.Caption = "Annulla tutto"
            Esci.Caption = "Esci"
        End With
    Case Is = "Inglese"
        'Traduco il FrameTipoMateriali e tutti gli oggetti comntenuti al suo interno
        With Me
            .FrameTipoMateriali = "Material Type:"
            .Normale.Caption = "Normal"
            .SphereMapping.Caption = "False Reflection"
            .Trasparenza.Caption = "Trasparency"
            'Traduco il sotto frame FrameOpzioniMateriale e tutti gli oggetti contenuti al suo interno
            .FrameOpzioniMateriale = "Material Option"
            .LabelAmbiente = "Ambient:"
            .LabelDiffusa = "Diffuse:"
            .LabelEmissiva = "Emissive:"
            .LabelSpeculare = "Specular:"
            .LabelPotenza = "Power:"
            .AnnullaTPM.Caption = "Cancel all"
            .Conferma.Caption = "Confirm"
            'Traduco il FrameOpzioniTrasparenza e tutti gli oggetti contenuti al suo interno
            .FrameOpzioniTrasparenza = "Multi Colours Trasparency Option"
            .CambiaNO.Caption = "Change"
            .CambiaNE.Caption = "Change"
            .CambiaSO.Caption = "Change"
            .CambiaSE.Caption = "Change"
            .CancellaNO.Caption = "Cancel"
            .CancellaNE.Caption = "Cancel"
            .CancellaSO.Caption = "Cancel"
            .CancellaSE.Caption = "Cancel"
            .AnnullaAssegnazioneColori.Caption = "Cancel all"
            Esci.Caption = "Exit"
        End With
    End Select
End Sub
