VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form_Opzioni 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opzioni"
   ClientHeight    =   8340
   ClientLeft      =   8850
   ClientTop       =   1485
   ClientWidth     =   5250
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrameGestioneOggetti 
      Height          =   7815
      Left            =   240
      TabIndex        =   110
      Top             =   360
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Frame FrameAssegnazioneGruppo 
         Caption         =   "Assegnazione Gruppo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6735
         Left            =   120
         TabIndex        =   116
         Top             =   960
         Width           =   4575
         Begin VB.Frame FrameInformazioniOggetto 
            Caption         =   "Informazioni Oggetto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3735
            Left            =   120
            TabIndex        =   121
            Top             =   2880
            Width           =   4335
            Begin VB.Frame FrameDescrizioneOggetto 
               Caption         =   "Descrizione"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1335
               Left            =   120
               TabIndex        =   136
               Top             =   240
               Width           =   2175
               Begin VB.CommandButton CancellaDescrizione 
                  Caption         =   "Cancella"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   138
                  Top             =   960
                  Width           =   855
               End
               Begin VB.CommandButton ConfermaDescrizione 
                  Caption         =   "Conferma"
                  Height          =   255
                  Left            =   1200
                  TabIndex        =   137
                  Top             =   960
                  Width           =   855
               End
               Begin RichTextLib.RichTextBox DescrizioneOggetto 
                  Height          =   735
                  Left            =   120
                  TabIndex        =   139
                  Top             =   240
                  Width           =   1935
                  _ExtentX        =   3413
                  _ExtentY        =   1296
                  _Version        =   393217
                  Enabled         =   -1  'True
                  TextRTF         =   $"Form_Opzioni.frx":0000
               End
            End
            Begin VB.Frame FrameOpzioniOggetto 
               Caption         =   "Opzioni"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   120
               TabIndex        =   133
               Top             =   1560
               Width           =   2175
               Begin VB.CheckBox AttivaOggetto 
                  Caption         =   "Attiva in M.A.3D"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   135
                  Top             =   480
                  Width           =   1815
               End
               Begin VB.CheckBox VisualizzaOggetto 
                  Caption         =   "Visualizza nell'editor"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   134
                  Top             =   240
                  Width           =   1935
               End
            End
            Begin VB.Frame FrameModificaOggetto 
               Caption         =   "Modifica"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1215
               Left            =   120
               TabIndex        =   125
               Top             =   2400
               Width           =   4095
               Begin VB.CommandButton ConfermaNuovoValoreOggetto 
                  Caption         =   "Conferma"
                  Height          =   255
                  Left            =   3000
                  TabIndex        =   129
                  Top             =   840
                  Width           =   975
               End
               Begin VB.TextBox NuovoValoreOggetto 
                  Height          =   285
                  Left            =   2520
                  TabIndex        =   128
                  Top             =   480
                  Width           =   1455
               End
               Begin VB.ComboBox AsseModificaOggetto 
                  Height          =   315
                  ItemData        =   "Form_Opzioni.frx":0082
                  Left            =   1560
                  List            =   "Form_Opzioni.frx":008F
                  TabIndex        =   127
                  Text            =   "X"
                  Top             =   480
                  Width           =   735
               End
               Begin VB.ComboBox OperazioneModificaOggetto 
                  Height          =   315
                  ItemData        =   "Form_Opzioni.frx":009C
                  Left            =   120
                  List            =   "Form_Opzioni.frx":00A9
                  TabIndex        =   126
                  Text            =   "Sposta"
                  Top             =   480
                  Width           =   1215
               End
               Begin VB.Label LabelNuovoValore 
                  Caption         =   "  Nuovo valore:"
                  Height          =   255
                  Left            =   2640
                  TabIndex        =   132
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.Label LabelAsse 
                  Caption         =   "Asse:"
                  Height          =   255
                  Left            =   1680
                  TabIndex        =   131
                  Top             =   240
                  Width           =   975
               End
               Begin VB.Label LabelOperazione 
                  Caption         =   "Operazione:"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   130
                  Top             =   240
                  Width           =   975
               End
            End
            Begin VB.Frame FrameTextureOggetto 
               Caption         =   "Texture"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2295
               Left            =   2400
               TabIndex        =   122
               Top             =   120
               Width           =   1815
               Begin VB.CommandButton AnnullaTextureOggetto 
                  Caption         =   "Annulla"
                  Height          =   255
                  Left            =   960
                  TabIndex        =   124
                  Top             =   1920
                  Width           =   735
               End
               Begin VB.CommandButton CaricaTextureOggetto 
                  Caption         =   "Cambia"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   123
                  Top             =   1920
                  Width           =   735
               End
               Begin VB.Image TextureOggetto 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   1575
                  Left            =   120
                  Stretch         =   -1  'True
                  Top             =   240
                  Width           =   1575
               End
            End
         End
         Begin VB.Frame FrameElenco 
            Caption         =   "Elenco Oggetti"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Left            =   120
            TabIndex        =   117
            Top             =   240
            Width           =   4335
            Begin VB.TextBox NomeGruppo 
               Height          =   285
               Left            =   2640
               TabIndex        =   119
               Top             =   2280
               Width           =   1575
            End
            Begin VB.CommandButton SpostaOggettoInGruppo 
               Caption         =   "Sposta in..."
               Height          =   255
               Left            =   1320
               TabIndex        =   118
               Top             =   2280
               Width           =   1215
            End
            Begin MSComctlLib.TreeView ElencoGruppiOggetti 
               Height          =   1935
               Left            =   120
               TabIndex        =   120
               Top             =   240
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   3413
               _Version        =   393217
               LineStyle       =   1
               Sorted          =   -1  'True
               Style           =   7
               FullRowSelect   =   -1  'True
               Appearance      =   1
               OLEDropMode     =   1
            End
         End
      End
      Begin VB.Frame FrameOperazioni 
         Caption         =   "Esegui..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   111
         Top             =   120
         Width           =   4575
         Begin VB.CommandButton EliminaGruppo 
            Caption         =   "Elimina Gruppo"
            Height          =   255
            Left            =   2280
            TabIndex        =   115
            Top             =   480
            Width           =   1935
         End
         Begin VB.CommandButton AggiungiGruppo 
            Caption         =   "Crea Nuovo Gruppo"
            Height          =   255
            Left            =   360
            TabIndex        =   114
            Top             =   480
            Width           =   1935
         End
         Begin VB.CommandButton EliminaOggetto 
            Caption         =   "Elimina Oggetto"
            Height          =   255
            Left            =   2280
            TabIndex        =   113
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton CaricaOggetto 
            Caption         =   "Carica Oggetto"
            Height          =   255
            Left            =   720
            TabIndex        =   112
            Top             =   240
            Width           =   1575
         End
      End
      Begin MSComDlg.CommonDialog CD2 
         Left            =   4200
         Top             =   7320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
   End
   Begin VB.Frame FrameCostruisci 
      Caption         =   "Costruisci"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   94
      Top             =   360
      Width           =   4815
      Begin VB.CommandButton Pavimento 
         Caption         =   "Pavimento"
         BeginProperty Font 
            Name            =   "Futura Md BT"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   97
         ToolTipText     =   "Se premuto si potrà aggiungere un pavimento alla mappa corrente"
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Soffitto 
         Caption         =   "Soffitto"
         BeginProperty Font 
            Name            =   "Futura Md BT"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   96
         ToolTipText     =   "Se premuto si potrà aggiungere un soffitto alla mappa corrente"
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Muri 
         Caption         =   "Muro"
         BeginProperty Font 
            Name            =   "Futura Md BT"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   95
         ToolTipText     =   "Se premuto si potrà aggiungere un muro alla mappa corrente"
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame FrameElencoCostruzioni 
      Caption         =   "Elenco Costruzioni"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   240
      TabIndex        =   51
      Top             =   1200
      Width           =   4815
      Begin VB.Frame FramePavimenti 
         Caption         =   "Pavimenti / Soffitti"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   120
         TabIndex        =   73
         Top             =   3600
         Width           =   4575
         Begin VB.CommandButton Materiale2 
            Caption         =   "Materiale"
            Height          =   255
            Left            =   1920
            TabIndex        =   89
            Top             =   2880
            Width           =   855
         End
         Begin VB.CommandButton AssegnazioneMultiplaSoP 
            Caption         =   "Assegnazione multipla"
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   2880
            Width           =   1815
         End
         Begin VB.CommandButton EliminaSoP 
            Caption         =   "Elimina"
            Height          =   255
            Left            =   1920
            TabIndex        =   87
            Top             =   2640
            Width           =   855
         End
         Begin VB.CommandButton Conferma2 
            Caption         =   "Conferma"
            Height          =   255
            Left            =   960
            TabIndex        =   86
            Top             =   2640
            Width           =   975
         End
         Begin VB.CommandButton Modifica2 
            Caption         =   "Modifica"
            Height          =   255
            Left            =   120
            TabIndex        =   85
            Top             =   2640
            Width           =   855
         End
         Begin VB.Frame FrameTextureSoP 
            Caption         =   "Texture"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Left            =   2640
            TabIndex        =   82
            Top             =   240
            Width           =   1815
            Begin VB.CommandButton CambiaTexture2 
               Caption         =   "Cambia"
               Height          =   255
               Left            =   120
               TabIndex        =   84
               Top             =   1920
               Width           =   735
            End
            Begin VB.CommandButton NessunaTexture2 
               Caption         =   "Annulla"
               Height          =   255
               Left            =   960
               TabIndex        =   83
               Top             =   1920
               Width           =   735
            End
            Begin VB.Image Texture_SoP 
               BorderStyle     =   1  'Fixed Single
               Height          =   1575
               Left            =   120
               Stretch         =   -1  'True
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.CommandButton Rinomina2 
            Caption         =   "Rinomina"
            Height          =   255
            Left            =   120
            TabIndex        =   81
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox Nomer2 
            Height          =   285
            Left            =   1200
            TabIndex        =   80
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox MattonelleAltezza2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   79
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox MattonelleLarghezza2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   78
            Top             =   1920
            Width           =   1335
         End
         Begin VB.TextBox AltitudineSoP 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   77
            Top             =   1200
            Width           =   1335
         End
         Begin VB.ComboBox ElencoSoP 
            Height          =   315
            Left            =   120
            TabIndex        =   76
            Top             =   480
            Width           =   2295
         End
         Begin VB.OptionButton TipoPavimento 
            Caption         =   "Pavimento"
            Height          =   255
            Left            =   120
            TabIndex        =   75
            Top             =   2280
            Width           =   1095
         End
         Begin VB.OptionButton TipoSoffitto 
            Caption         =   "Soffitto"
            Height          =   255
            Left            =   1320
            TabIndex        =   74
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label LabAltitudine2 
            Caption         =   "Altitudine:"
            Height          =   255
            Left            =   360
            TabIndex        =   93
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label LabMatAltezza2 
            Caption         =   "Mat. Altezza:"
            Height          =   255
            Left            =   120
            TabIndex        =   92
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label LabmatLarghezza2 
            Caption         =   "Mat.Larghezza:"
            Height          =   255
            Left            =   0
            TabIndex        =   91
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label LabPavimenti 
            Caption         =   "Pavimenti / Soffitti esistenti:"
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame_muri 
         Caption         =   "Muri"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   4575
         Begin VB.CommandButton Materiale 
            Caption         =   "Materiale"
            Height          =   255
            Left            =   1920
            TabIndex        =   67
            Top             =   2880
            Width           =   855
         End
         Begin VB.CommandButton AssegnazioneMultiplaMuri 
            Caption         =   "Assegnazione multipla"
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   2880
            Width           =   1815
         End
         Begin VB.CommandButton EliminaMuro 
            Caption         =   "Elimina"
            Height          =   255
            Left            =   1920
            TabIndex        =   65
            Top             =   2640
            Width           =   855
         End
         Begin VB.CommandButton Conferma1 
            Caption         =   "Conferma"
            Height          =   255
            Left            =   960
            TabIndex        =   64
            Top             =   2640
            Width           =   975
         End
         Begin VB.CommandButton Modifica1 
            Caption         =   "Modifica"
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   2640
            Width           =   855
         End
         Begin VB.ComboBox ElencoMuri 
            Height          =   315
            Left            =   120
            TabIndex        =   62
            Top             =   480
            Width           =   2295
         End
         Begin VB.TextBox Altitudine_muro 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   61
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox Altezza_muro 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   60
            Top             =   1200
            Width           =   1335
         End
         Begin VB.TextBox MattonelleLarghezza 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   59
            Top             =   2280
            Width           =   1335
         End
         Begin VB.TextBox MattonelleAltezza 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   58
            Top             =   1920
            Width           =   1335
         End
         Begin VB.TextBox Nomer 
            Height          =   285
            Left            =   1200
            TabIndex        =   57
            Top             =   840
            Width           =   1335
         End
         Begin VB.CommandButton Rinomina 
            Caption         =   "Rinomina"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   840
            Width           =   855
         End
         Begin VB.Frame FrameTextureMuri 
            Caption         =   "Texture"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Left            =   2640
            TabIndex        =   53
            Top             =   240
            Width           =   1815
            Begin VB.CommandButton CambiaTexture 
               Caption         =   "Cambia"
               Height          =   255
               Left            =   120
               TabIndex        =   55
               Top             =   1920
               Width           =   735
            End
            Begin VB.CommandButton NessunaTexture 
               Caption         =   "Annulla"
               Height          =   255
               Left            =   960
               TabIndex        =   54
               Top             =   1920
               Width           =   735
            End
            Begin VB.Image Texture_muro 
               BorderStyle     =   1  'Fixed Single
               Height          =   1575
               Left            =   120
               Stretch         =   -1  'True
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.Label LabMatLarghezza 
            Caption         =   "Mat.Larghezza:"
            Height          =   255
            Left            =   0
            TabIndex        =   72
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label LabMatAltezza 
            Caption         =   "Mat. Altezza:"
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label Labmuri 
            Caption         =   "Muri esistenti:"
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label LabAltezzamuro 
            Caption         =   "Altezza:"
            Height          =   255
            Left            =   480
            TabIndex        =   69
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label LabAltitudine 
            Caption         =   "Altitudine:"
            Height          =   255
            Left            =   360
            TabIndex        =   68
            Top             =   1560
            Width           =   735
         End
      End
   End
   Begin VB.Frame Preferenze 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7815
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   4815
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   4320
         Top             =   7320
      End
      Begin VB.Frame FramePersonalizzaColori 
         Caption         =   "Personalizza Colori"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   50
         Top             =   6240
         Width           =   4575
         Begin VB.CommandButton Colore2Menù 
            Caption         =   "2° Menù"
            Height          =   255
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   109
            Top             =   1080
            Width           =   1575
         End
         Begin VB.CommandButton Colore1Menù 
            Caption         =   "1° Menù"
            Height          =   255
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   108
            Top             =   1080
            Width           =   1455
         End
         Begin VB.CommandButton ColoreSfondoMenù 
            Caption         =   "Sfondo Menù"
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   107
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CommandButton ColoreLineeGuida 
            Caption         =   "Linee Guida"
            Height          =   255
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   106
            Top             =   840
            Width           =   1575
         End
         Begin VB.CommandButton ColoreAllineamentoMuri 
            Caption         =   "Allineamento Muri"
            Height          =   255
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   105
            Top             =   840
            Width           =   1455
         End
         Begin VB.CommandButton ColoreAllineamentoSP 
            Caption         =   "Allineamento S/P"
            Height          =   255
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   104
            Top             =   600
            Width           =   1575
         End
         Begin VB.CommandButton ColoreSPSelezionati 
            Caption         =   "S/P Selezionati"
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   103
            Top             =   840
            Width           =   1335
         End
         Begin VB.CommandButton ColorePavimenti 
            Caption         =   "Pavimenti"
            Height          =   255
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   102
            Top             =   600
            Width           =   1455
         End
         Begin VB.CommandButton ColoreSoffitti 
            Caption         =   "Sofitti"
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   101
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton ColoreSpigoliMuri 
            Caption         =   "Spigoli Muri"
            Height          =   255
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   100
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton ColoreMuriSelezionati 
            Caption         =   "Muri Selezionati"
            Height          =   255
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   99
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton ColoreMuri 
            Caption         =   "Muri"
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   98
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame FrameFondaleEitor 
         Caption         =   "Fondale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   2280
         TabIndex        =   46
         Top             =   3720
         Width           =   2415
         Begin VB.CommandButton CambiaImmagineDiSfondo 
            Caption         =   "Cambia"
            Height          =   255
            Left            =   1200
            TabIndex        =   49
            Top             =   840
            Width           =   975
         End
         Begin VB.OptionButton FondaleStatico 
            Caption         =   "Attiva fondale"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   480
            Width           =   2200
         End
         Begin VB.OptionButton DisattivaFondale 
            Caption         =   "Disattiva fondale"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Value           =   -1  'True
            Width           =   2200
         End
      End
      Begin VB.Frame FramePreferenze 
         Caption         =   "Preferenze"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   37
         Top             =   120
         Width           =   4575
         Begin VB.CheckBox Visualizza_soffitti 
            Caption         =   "Visualizza soffitti"
            Height          =   255
            Left            =   2280
            TabIndex        =   45
            Top             =   720
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox Visualizza_Pavimenti 
            Caption         =   "Visualizza pavimenti"
            Height          =   255
            Left            =   2280
            TabIndex        =   44
            Top             =   480
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox Mostra_Menù 
            Caption         =   "Mostra menù"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   480
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox Controlla_spigoli 
            Caption         =   "Visualizza spigoli"
            Height          =   255
            Left            =   2280
            TabIndex        =   42
            Top             =   960
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox Controlla_Muri 
            Caption         =   "Visualizza muri"
            Height          =   255
            Left            =   2280
            TabIndex        =   41
            Top             =   240
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox Rileva_Allineamento2 
            Caption         =   "Rileva allineamento S/ P"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   960
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.CheckBox Rileva_allineamento 
            Caption         =   "Rileva allineamento muri"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   720
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.CheckBox Linee_guida 
            Caption         =   "Mostra linee guida"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Value           =   1  'Checked
            Width           =   1815
         End
      End
      Begin VB.Frame OpzioniZoom 
         Caption         =   "Opzioni Zoom"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   2280
         TabIndex        =   30
         Top             =   4920
         Width           =   2415
         Begin VB.CommandButton RipristinaZoom 
            Caption         =   "Ripristina"
            Height          =   255
            Left            =   1200
            TabIndex        =   35
            Top             =   840
            Width           =   975
         End
         Begin VB.CommandButton ZoomP 
            Caption         =   "Zoom +"
            Height          =   255
            Left            =   1200
            TabIndex        =   32
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton ZoomM 
            Caption         =   "Zoom -"
            Height          =   255
            Left            =   1200
            TabIndex        =   31
            Top             =   360
            Width           =   975
         End
         Begin VB.Label LabeRipristinaZoom 
            Caption         =   "Reimposta"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   840
            Width           =   855
         End
         Begin VB.Label LabAumentaZoom 
            Caption         =   "Aumenta"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   600
            Width           =   855
         End
         Begin VB.Label LabDiminuisciZoom 
            Caption         =   "Diminuisci"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame OpzioniScale 
         Caption         =   "Opzioni Scale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   120
         TabIndex        =   14
         Top             =   3720
         Width           =   2175
         Begin VB.CommandButton SalvaScalePersonalizzato 
            Caption         =   "Salva"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1200
            TabIndex        =   21
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox ValoreScalePersonalizzato 
            Enabled         =   0   'False
            Height          =   315
            Left            =   480
            TabIndex        =   20
            Text            =   "1"
            Top             =   1800
            Width           =   1335
         End
         Begin VB.OptionButton Personalizzato 
            Caption         =   "Personalizzato"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1440
            Width           =   1575
         End
         Begin VB.OptionButton Impostato 
            Caption         =   "Impostato"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.ComboBox ValoreScale 
            Height          =   315
            ItemData        =   "Form_Opzioni.frx":00C3
            Left            =   120
            List            =   "Form_Opzioni.frx":00D6
            TabIndex        =   15
            Text            =   "1 : 1"
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label LabelScalePersonalizzato 
            Caption         =   "1 : "
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   1850
            Width           =   615
         End
         Begin VB.Label LabSelezionaScale 
            Caption         =   "Seleziona scale:"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Opzioni_griglia 
         Caption         =   "Opzioni Griglia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   2280
         TabIndex        =   4
         Top             =   1440
         Width           =   2415
         Begin VB.CheckBox Visualizza_griglia 
            Caption         =   "Visualizza griglia"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.TextBox AltezzaGriglia 
            Height          =   285
            Left            =   120
            TabIndex        =   9
            Text            =   "10"
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox LarghezzaGriglia 
            Height          =   285
            Left            =   1320
            TabIndex        =   8
            Text            =   "10"
            Top             =   1320
            Width           =   855
         End
         Begin VB.HScrollBar Luminosità_griglia 
            Height          =   255
            Left            =   120
            Max             =   10
            TabIndex        =   7
            Top             =   1920
            Value           =   3
            Width           =   1455
         End
         Begin VB.CheckBox GrigliaControllataDaZoom 
            Caption         =   "Ridimensiona griglia tramite zoom"
            Height          =   495
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.TextBox Valore_luminosità 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   1920
            Width           =   375
         End
         Begin VB.Label labAltezzaGriglia 
            Caption         =   "Altezza:"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label LabLarghezzaGriglia 
            Caption         =   "Larghezza:"
            Height          =   255
            Left            =   1320
            TabIndex        =   12
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label LabLuminosità_griglia 
            Caption         =   "Luminosità:"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1680
            Width           =   855
         End
      End
      Begin VB.Frame Opzioni_telecamera 
         Caption         =   "Opzioni Telecamera"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   2175
         Begin VB.CommandButton ConfermaTelecamera 
            Caption         =   "OK!"
            Height          =   255
            Left            =   1080
            TabIndex        =   29
            Top             =   1920
            Width           =   855
         End
         Begin VB.CommandButton ModificaTelecamera 
            Caption         =   "Modifica"
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   1920
            Width           =   855
         End
         Begin VB.TextBox TelecameraZ 
            Height          =   285
            Left            =   840
            TabIndex        =   24
            Text            =   "100"
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox TelecameraY 
            Height          =   285
            Left            =   840
            TabIndex        =   23
            Text            =   "500"
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox TelecameraX 
            Height          =   285
            Left            =   840
            TabIndex        =   22
            Text            =   "100"
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox Visualizza_telecamera 
            Caption         =   "Visualizza Telecamera"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label LabeltelecameraPosZ 
            Caption         =   "Pos. Z:"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label LabelTelecameraPosY 
            Caption         =   "Pos. Y:"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label LabelTelecameraPosX 
            Caption         =   "Pos. X:"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   735
         End
      End
   End
   Begin MSComctlLib.TabStrip Tabella 
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   14631
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Gestione Costruzione"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Oggetti"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Opzioni Editor"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form_Opzioni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dichiaro una variabile che terrà conto di quante volte la mappa è stata o ingrandita o rimpicciolita
'dalle operazioni di Zoom questo mi servirà per abilitare o disattivare i relativi bottoni di Zoom
Dim Cont As Integer
'Dichiaro una variabile che mi servirà per capire quale oggetto è stato selezionato in modalità
'Anteprima 3D,salvando in essa la propria "Chiave" univoca
Dim OggettoSelezionato As String

Private Sub AggiungiGruppo_Click()
    'Dichiaro la variabile Doppio, la quale mi servirà per contenere il valore booleano restituito dalla
    'funzione Verifica_Doppio.
    Dim Doppio As Boolean
    'La variabile Nome invece mi servirà per contenere la nuova chiave restituita dalla funzione
    'Nuovo_Nome in caso di nodi aventi la stessa chiave
    Dim Nome As String
    Dim TmpNome As String
    'Assegno il valore della variabile Doppio tramite la funzione Verifica_doppio
    If LinguaS = "Italiano" Then Doppio = Verifica_Doppio("Nuovo Gruppo")
    If LinguaS = "Inglese" Then Doppio = Verifica_Doppio("New Group")
    'Se la variabile Doppio ha assunto il valore True,cioè sono stati trovati due nodi con la stessa
    'chiave,allora...
    If Doppio = True Then
        'Richiamo la funzione Nuovo_nome con lo scopo di ricercare una nuova chiave per il gruppo
        'da aggiungere.
        'La nuova chiave restituita verrà salvata all'interno della variabile Nome
        If LinguaS = "Italiano" Then Nome = Nuovo_Nome("Nuovo Gruppo"): TmpNome = "Nuovo Gruppo"
        If LinguaS = "Inglese" Then Nome = Nuovo_Nome("New Group"): TmpNome = "New Group"
    'Nel caso in cui "sia andato tutto liscio" cioè la chiave è univoca all'interno dell'insieme,allora...
    Else
        'Verrà assegnata alla variabile Nome la stringa "Nuovo Nome"
        If LinguaS = "Italiano" Then Nome = "Nuovo Gruppo": TmpNome = "Nuovo Gruppo"
        If LinguaS = "Inglese" Then Nome = "New Group": TmpNome = "New Group"
    End If
    'Aggiungo al controllo TreeView ElencoGruppiOggetti un nuovo nodo che conterrà tutti gli oggetti
    'specificati dall'utente
    With ElencoGruppiOggetti.Nodes.Add(, , Nome, TmpNome)
        'Setto il colore blu al gruppo appena aggiunto
        .ForeColor = vbBlue
    End With
End Sub

Private Sub AnnullaTextureOggetto_Click()
    'Dichiaro una variabile di appoggio che mi servirà per contenere l'indice dell'oggetto trovato
    'dalla funzione Trova_Oggetto
    Dim OggettoTrovato As Integer
    'Assegno alla variabile OggettoTrovato,il valore restituito dalla funzione Trova_Oggetto.
    'Questo mi servirà per ricavare l'indice all'interno dell'array Oggetti,appunto dell'oggetto
    'corrispondente a quello selezionato
    OggettoTrovato = Trova_Oggetto(ElencoGruppiOggetti.SelectedItem.Key)
    'Cancello la Texture precedentemente settata all'oggetto appena trovato
    Oggetto(OggettoTrovato).Texture = ""
    'Elimino l'immagine residua dall'interno del controllo Image TextureOggetto,al fine di far notare
    'all'utente che l'immagine è stata cancellata con successo
    Set TextureOggetto.Picture = Nothing
End Sub

Private Sub AssegnazioneMultiplaMuri_Click()
    'Se sono presenti meno di due muri all'interno della mappa attuale, verrà visualizzato un messaggio che avviserà l'utente dell'errore
    If Max < 2 Then
        If LinguaS = "Italiano" Then MsgBox "Per effettuare un'assegnazione multipla dei parametri dei muri, devono esserne presenti almeno due all'interno della mappa attuale!", vbOKOnly, "Assegnazione multipla parametri muri"
        If LinguaS = "Inglese" Then MsgBox "For a multipled walls parameters assign , there must present at last two in the actual map!", vbInformation, "Mulpitped walls parameters assign"
    'Se invece non è stato selezionato nessun muro come modello per la multi assegnazione,allora verrà visualizzato un messaggio di errore che informerà
    'l'utente dell'errore
    ElseIf ElencoMuri.ListIndex = -1 Then
        If LinguaS = "Italiano" Then MsgBox "Non è stato selezionato nessun muro da utilizzare come modello per la multi assegnazione. Si prega di selezionare almeno un muro disponibile!", vbOKOnly, "Multi assegnazione"
        If LinguaS = "Inglese" Then MsgBox "You haven't selected any model wall!" + Chr(13) + "Please select at least one wall first!", vbOKOnly, "Multi assign"
    'Nel caso in cui invece sia tutto apposto,verrà assegnata la modalità di assegnazione multipla che indicherà che dovrà essere effettuata un'assegnazione multipla
    'delle proprietà dei pavimenti o soffitti,e verrà appunto visualizzato il form nella modalità appena settata
    Else
        'Assegno alla variabile ModalitàAssegnazioneMultipla il valore della rispettiva modalità
        ModalitàAssegnazioneMultipla = "Muri"
        'Avvio il form di assegnazione multipla dei muri.
        'Questo speciale form mi permetterà di selezionare più muri contemporaneamente in modo
        'da poter assegnare ad ognuno di questi uno o più parametri uguali.
        'Questo permetterà all'utente di risparmiare molto tempo nella costruzione della mappa
        'perchè permetterà di non doverli selezionare uno per uno
        Form_Assegnazione_Multipla.Show
    End If
End Sub

Private Sub AssegnazioneMultiplaSoP_Click()
    'Se sono presenti meno di due pavimenti o soffitti all'interno della mappa attuale, verrà visualizzato un messaggio che avviserà l'utente dell'errore
    If Max2 < 2 Then
        If LinguaS = "Italiano" Then MsgBox "Per effettuare un'assegnazione multipla dei parametri dei pavimenti o soffitti, devono esserne presenti almeno due all'interno della mappa attuale!", vbInformation, "Assegnazione multipla parametri pavimenti / soffitti"
        If LinguaS = "Inglese" Then MsgBox "For a multipled floors / ceilings parameters, there must present at last two in the actual map!", vbInformation, "Mulpitped Floors / Ceilings parameters assign"
    'Se invece non è stato selezionato nessun pavimento o soffitto come modello per la multi assegnazione,allora verrà visualizzato un messaggio di errore che informerà
    'l'utente dell'errore
    ElseIf ElencoSoP.ListIndex = -1 Then
        If LinguaS = "Italiano" Then MsgBox "Non è stato selezionato nessun pavimento o soffitto da utilizzare come modello per la multi assegnazione." + Chr(13) + " Si prega di selezionare un pavimento o soffitto disponibile!", vbInformation, "Multi assegnazione"
        If LinguaS = "Inglese" Then MsgBox "You haven't selected any model floor or ceiling for the multipled assegnation!" + Chr(13) + "Please select an existn floor or ceiling!", vbInformation, "Multipled assegnation"
    'Nel caso in cui invece sia tutto apposto,verrà assegnata la modalità di assegnazione multipla che indicherà che dovrà essere effettuata un'assegnazione multipla
    'delle proprietà dei muri,e verrà appunto visualizzato il form nella modalità appena settata
    Else
        'Assegno alla variabile ModalitàAssegnazioneMultipla, il valore della rispettiva modalità
        ModalitàAssegnazioneMultipla = "SoP"
        'Avvio il form di assegnazione multipla dei muri.
        'Questo speciale form mi permetterà di selezionare più pavimenti / soffitti contemporaneamente in modo
        'da poter assegnare ad ognuno di questi uno o più parametri uguali.
        'Questo permetterà all'utente di risparmiare molto tempo nella costruzione della mappa
        'perchè permetterà di non doverli selezionare uno per uno
        Form_Assegnazione_Multipla.Show
    End If
End Sub

Private Sub CambiaTexture2_Click()
    'Se non viene selezionato nessun pavimento o soffitto su cui applicare la Texture allora viene
    'visualizzato un messaggio che informa l'utente di selezionarne prima uno
    If ElencoSoP.Text = "" Then
        If LinguaS = "Italiano" Then MsgBox "Per caricare una texture devi prima selezionare un pavimento o un soffitto esistente!", vbInformation, "Avviso!"
        If LinguaS = "Inglese" Then MsgBox "You must choose a floor or ceiling for load a texture", vbInformation, "Advise!"
    Else
        'Imposto il titolo dell'oggetto operazioni,in modo che l'utente potrà capire quale operazione
        'verrà effettuata
        If LinguaS = "Italiano" Then CD2.DialogTitle = "Caricamento Texture"
        If LinguaS = "Inglese" Then CD2.DialogTitle = "Texture Load"
        'Impongo una serie di filtri che corrispondono ai più comuni formati di immagine
        If LinguaS = "Italiano" Then CD2.Filter = "Tutti i formati |*.bmp;*.jpg|BMP (Bitmap di Windows)|*.bmp|JPG (Formato di interscambio dei file JPEG)|*.jpg"
        If LinguaS = "Inglese" Then CD2.Filter = "All the formats |*.bmp;*.jpg|BMP (Bitmap di Windows)|*.bmp|JPG (Interchange format of JPEG files)|*.jpg"
        On Error GoTo Annulla
        'Apro il componente che mi permetterà di selezionare un'immagine da settare come
        'textures per il pavimento o soffitto corrente
        CD2.ShowOpen
        'Assegno al pavimento o soffitto corrente l'immagine selezionata e la applico come textures
        SoP(IndiceLista2).CR.Texture = CD2.FileName
        'Carico l'immagine appena selezionata all'interno dell'oggetto Texture_SoP
        Texture_SoP.Picture = LoadPicture(RTrim(SoP(IndiceLista2).CR.Texture))
    End If
    Exit Sub
Annulla:
End Sub

Private Sub CancellaDescrizione_Click()
    'Cancello la descrizione dell'oggetto selezionato
    Oggetto(IndiceOggettoSelezionato).Decrizione = ""
    'Cancello il contenuto del controllo DescrizioneOggetto
    DescrizioneOggetto = ""
End Sub

Private Sub CaricaOggetto_Click()
    Dim Doppio As Boolean
    Dim New_Name As String
    'In caso l'utente decidesse di annullare l'operazione di caricamento dell'oggetto,il programma
    'effettuerà un salto fino al label Annulla
    On Error GoTo Annulla
    'Imposto il titolo dell'oggetto operazioni,in modo che l'utente potrà capire quale operazione
    'verrà effettuata
    If LinguaS = "Italiano" Then CD2.DialogTitle = "Caricamento Oggetto"
    If LinguaS = "Inglese" Then CD2.DialogTitle = "Object Load"
    'Impongo una serie di filtri che corrispondono ai più comuni formati di immagine
    If LinguaS = "Italiano" Then CD2.Filter = "Tutti i formati |*.3ds;*.x|X (DirectX Files)|*.X|3ds (3D Studio Files)|*.3ds"
    If LinguaS = "Inglese" Then CD2.Filter = "All the formats |*.3ds;*.x|X (DirectX Files)|*.bmp|JPG (3D Studio Files)|*.3ds"
    'Apro il componente che mi permetterà di selezionare un'oggetto da caricare all'interno della mappa 3d
    'da una qualsiasi fonte di origine
    CD2.ShowOpen
    'Richiamo la funzione addetta alla ricerca di un possibile oggetto duplicato all'interno del controllo
    'ElencoGruppiOggetti.
    'Il valore restituito da questa funzione verrà salvato all'interno della variabile booleana Doppio
    Doppio = Verifica_Doppio(CD2.FileTitle)
    'Se la variabile Doppio ha assunto il valore booleano True,ovvero esiste già un nodo all'interno del
    'controllo ElencoGruppiOggetti con la stessa chiave, verrà avviata la funzione addetta alla creazione di
    'una nuova chiave in modo che questa possa essere univoca all'interno dell'insieme
    If Doppio = True Then
        New_Name = Nuovo_Nome(CD2.FileTitle)
    'In caso contrario potrà essere usata benissimo come chiave il nome del file stesso
    Else
        New_Name = CD2.FileTitle
    End If
    'Richiamo la funzione Carica_Oggetto dalla classe ClsOggetti,passando al metodo
    'il file appena selezionato,in modo che questo possa essere caricato all'interno della mappa 3D,
    'e la chiave appena generata dalla funzioneNew_Name
    Oggetto(IOg).Carica_Oggetto CD2.FileName, New_Name
    'Creo all'interno della scena "l'involucro" necessario a contenere l'oggetto
    Set Oggetto(IOg).Scheletro = Scena.CreateMeshBuilder
    'Incremento il contatore che tiene conto di quanti oggetti sono stati inseriti
    'all'interno della mappa attuale
    IOg = IOg + 1
    'Aggiungo l'oggetto appena caricato all'interno del controllo Treeview ElencoGruppiOggetti,in modo
    'che l'utente possa recuperarlo facilmente per effettuarvi eventuali modifiche
    ElencoGruppiOggetti.Nodes.Add "OSG", tvwChild, New_Name, CD2.FileTitle
'Label al quale si verrà reinderizzati in caso di annullamento dell'operazione di caricamento dell'oggetto
'all'interno della mappa 3D
Annulla:
End Sub

Private Sub CaricaTextureOggetto_Click()
    'Se è stato selezionato un oggetto si cui caricare una Texture allora...
    If OggettoSelezionato <> "" Then
        'Dichiaro una variabile che mi servirà per contenere il valore restituito dalla funzione
        'Trova_Oggetto.Questo valore verrà in seguito usato come indice al fine di caricare
        'la Texture del corrispondente Oggetto selezionato dal comntrollo Treeview ElencoGruppiOggetti
        Dim OggettoTrovato As Integer
        'Imposto il titolo dell'oggetto operazioni,in modo che l'utente potrà capire quale operazione
        'verrà effettuata
        If LinguaS = "Italiano" Then CD2.DialogTitle = "Caricamento Texture"
        If LinguaS = "Inglese" Then CD2.DialogTitle = "Texture Load"
        'Impongo una serie di filtri che corrispondono ai più comuni formati di immagine
        If LinguaS = "Italiano" Then CD2.Filter = "Tutti i formati |*.bmp;*.jpg|BMP (Bitmap di Windows)|*.bmp|JPG (Formato di interscambio dei file JPEG)|*.jpg"
        If LinguaS = "Inglese" Then CD2.Filter = "All the formats |*.bmp;*.jpg|BMP (Bitmap di Windows)|*.bmp|JPG (Interchange format of JPEG files)|*.jpg"
        'In caso l'utente volesse annullare l'operazione di caricamento di un nuovo oggetto,
        'il programma effettuerà un salto fino al label Annulla
        On Error GoTo Annulla
        'Apro il componente che mi permetterà di selezionare un oggetto da caricare
        CD2.ShowOpen
        'Assegno alla variabile OggettoTrovato,l'indice restituito dalla funzione Trova_Oggetto
        OggettoTrovato = Trova_Oggetto(ElencoGruppiOggetti.SelectedItem)
        'Richiamo il metodo Carica_Texture dalla classe ClsOggetti,il quale mi permetterà
        'di impostare "la pelle" dell'oggetto caricato
        Oggetto(OggettoTrovato).Carica_Texture CD2.FileName
        'Carico la sua nuova Texture all'interno del controllo image TextureOggetto
        TextureOggetto.Picture = LoadPicture(Oggetto(OggettoTrovato).Texture)
    'Altrimenti, se non è stato selezionato nessun oggetto su cui caricare la Texture, verrà visualizzato
    'un messaggio di errore che informerà l'utente dell'errore
    Else
        If LinguaS = "Italiano" Then MsgBox "Attenzione! Prima di caricare una texture, devi prima selezionare l'oggetto desiderato!", vbInformation, "Caricamento Texture"
        If LinguaS = "Inglese" Then MsgBox "Attention!You must choose an object first to load a texture!", vbInformation, "Texture Load"
    End If
'Label al quale il programma salterà in caso in cui l'utente decidesse di annullare
'l'operazione di caricamento di un nuovo oggetto
Annulla:
End Sub

Private Sub Colore1Menù_Click()
    'Se viene scelto di annullare l'operazione di selezione colore,allora si verrà reinderizzati
    'verso il "Label" Annulla
    On Error GoTo Annulla
    'Apro il componente che mi permetterà di selezionare il colore desiderato da assegnare
    'al 1° colore utilizzato dal menù posto nella parte alta dell'editor,per creare le scritte
    'di informazione verso l'utente
    Form_Materiali.ControlloSceltaColori.ShowColor
    'Imposto il colore di sfondo del pulsante che lo rappresenta il colore appena selezionato
    Colore1Menù.BackColor = Form_Materiali.ControlloSceltaColori.Color
    'Richinamo la funzione pubblica che mi permetterà di estrarre le relative quantità di Rosso,
    'Verde e Blu dal colore appena selezionato
    Preleva_RGB C1M
    'Richiamo la funzione pubblica addetta a convertire le rispettive quantità di Rosso,
    'Vetrde e Blu da volori RGB in RGBA
    RGB_To_RGBA C1M
'Label al quale si verrà reinderizzati in caso di annullamento di selezione del colore
Annulla:
End Sub

Private Sub Colore2Menù_Click()
    'Se viene scelto di annullare l'operazione di selezione colore,allora si verrà reinderizzati
    'verso il "Label" Annulla
    On Error GoTo Annulla
    'Apro il componente che mi permetterà di selezionare il colore desiderato da assegnare
    'al 2° colore utilizzato dal menù posto nella parte alta dell'editor,per creare le scritte
    'di informazione verso l'utente
    Form_Materiali.ControlloSceltaColori.ShowColor
    'Imposto il colore di sfondo del pulsante che lo rappresenta il colore appena selezionato
    Colore2Menù.BackColor = Form_Materiali.ControlloSceltaColori.Color
    'Richinamo la funzione pubblica che mi permetterà di estrarre le relative quantità di Rosso,
    'Verde e Blu dal colore appena selezionato
    Preleva_RGB C2M
    'Richiamo la funzione pubblica addetta a convertire le rispettive quantità di Rosso,
    'Vetrde e Blu da volori RGB in RGBA
    RGB_To_RGBA C2M
'Label al quale si verrà reinderizzati in caso di annullamento di selezione del colore
Annulla:
End Sub

Private Sub ColoreAllineamentoMuri_Click()
    'Se viene scelto di annullare l'operazione di selezione colore,allora si verrà reinderizzati
    'verso il "Label" Annulla
    On Error GoTo Annulla
    'Apro il componente che mi permetterà di selezionare il colore desiderato da assegnare
    'ai quadratini che si formeranno in seguito al perfetto allineamento tra le coordinate
    'del mouse e i muri
    Form_Materiali.ControlloSceltaColori.ShowColor
    'Imposto il colore di sfondo del pulsante che lo rappresenta il colore appena selezionato
    ColoreAllineamentoMuri.BackColor = Form_Materiali.ControlloSceltaColori.Color
    'Richinamo la funzione pubblica che mi permetterà di estrarre le relative quantità di Rosso,
    'Verde e Blu dal colore appena selezionato
    Preleva_RGB CAM
    'Richiamo la funzione pubblica addetta a convertire le rispettive quantità di Rosso,
    'Vetrde e Blu da volori RGB in RGBA
    RGB_To_RGBA CAM
'Label al quale si verrà reinderizzati in caso di annullamento di selezione del colore
Annulla:
End Sub

Private Sub ColoreAllineamentoSP_Click()
    'Se viene scelto di annullare l'operazione di selezione colore,allora si verrà reinderizzati
    'verso il "Label" Annulla
    On Error GoTo Annulla
    'Apro il componente che mi permetterà di selezionare il colore desiderato da assegnare
    'ai quadratini che si formeranno in seguito al perfetto allineamento tra le coordinate
    'del mouse e i pavimenti / soffitti
    Form_Materiali.ControlloSceltaColori.ShowColor
    'Imposto il colore di sfondo del pulsante che lo rappresenta il colore appena selezionato
    ColoreAllineamentoSP.BackColor = Form_Materiali.ControlloSceltaColori.Color
    'Richinamo la funzione pubblica che mi permetterà di estrarre le relative quantità di Rosso,
    'Verde e Blu dal colore appena selezionato
    Preleva_RGB CASOP
    'Richiamo la funzione pubblica addetta a convertire le rispettive quantità di Rosso,
    'Vetrde e Blu da volori RGB in RGBA
    RGB_To_RGBA CASOP
'Label al quale si verrà reinderizzati in caso di annullamento di selezione del colore
Annulla:
End Sub

Private Sub ColoreLineeGuida_Click()
    'Se viene scelto di annullare l'operazione di selezione colore,allora si verrà reinderizzati
    'verso il "Label" Annulla
    On Error GoTo Annulla
    'Apro il componente che mi permetterà di selezionare il colore desiderato da assegnare
    'alle linee guida
    Form_Materiali.ControlloSceltaColori.ShowColor
    'Imposto il colore di sfondo del pulsante che lo rappresenta il colore appena selezionato
    ColoreLineeGuida.BackColor = Form_Materiali.ControlloSceltaColori.Color
    'Richinamo la funzione pubblica che mi permetterà di estrarre le relative quantità di Rosso,
    'Verde e Blu dal colore appena selezionato
    Preleva_RGB CLG
    'Richiamo la funzione pubblica addetta a convertire le rispettive quantità di Rosso,
    'Vetrde e Blu da volori RGB in RGBA
    RGB_To_RGBA CLG
'Label al quale si verrà reinderizzati in caso di annullamento di selezione del colore
Annulla:
End Sub

Private Sub ColoreMuri_Click()
    'Se viene scelto di annullare l'operazione di selezione colore,allora si verrà reinderizzati
    'verso il "Label" Annulla
    On Error GoTo Annulla
    'Apro il componente che mi permetterà di selezionare il colore desiderato da assegnare
    'alle righe che rappresentano i muri all'interno dell'editor
    Form_Materiali.ControlloSceltaColori.ShowColor
    'Imposto il colore di sfondo del pulsante che lo rappresenta il colore appena selezionato
    ColoreMuri.BackColor = Form_Materiali.ControlloSceltaColori.Color
    'Richinamo la funzione pubblica che mi permetterà di estrarre le relative quantità di Rosso,
    'Verde e Blu dal colore appena selezionato
    Preleva_RGB CM
    'Richiamo la funzione pubblica addetta a convertire le rispettive quantità di Rosso,
    'Vetrde e Blu da volori RGB in RGBA
    RGB_To_RGBA CM
'Label al quale si verrà reinderizzati in caso di annullamento di selezione del colore
Annulla:
End Sub

Private Sub ColoreMuriSelezionati_Click()
    'Se viene scelto di annullare l'operazione di selezione colore,allora si verrà reinderizzati
    'verso il "Label" Annulla
    On Error GoTo Annulla
    'Apro il componente che mi permetterà di selezionare il colore desiderato da assegnare
    'alla riga che rappresenta il muro selezionato tramite l'oggetto Elenco_Muri
    Form_Materiali.ControlloSceltaColori.ShowColor
    'Imposto il colore di sfondo del pulsante che lo rappresenta il colore appena selezionato
    ColoreMuriSelezionati.BackColor = Form_Materiali.ControlloSceltaColori.Color
    'Richiamo la funzione pubblica che mi permetterà di estrarre le relative quantità di Rosso,
    'Verde e Blu dal colore appena selezionato
    Preleva_RGB CMS
    'Richiamo la funzione pubblica addetta a convertire le rispettive quantità di Rosso,
    'Vetrde e Blu da volori RGB in RGBA
    RGB_To_RGBA CMS
'Label al quale si verrà reinderizzati in caso di annullamento di selezione del colore
Annulla:
End Sub

Private Sub ColorePavimenti_Click()
    'Se viene scelto di annullare l'operazione di selezione colore,allora si verrà reinderizzati
    'verso il "Label" Annulla
    On Error GoTo Annulla
    'Apro il componente che mi permetterà di selezionare il colore desiderato da assegnare
    'alle righe che rappresentano i pavimenti all'interno dell'editor
    Form_Materiali.ControlloSceltaColori.ShowColor
    'Imposto il colore di sfondo del pulsante che lo rappresenta il colore appena selezionato
    ColorePavimenti.BackColor = Form_Materiali.ControlloSceltaColori.Color
    'Richiamo la funzione pubblica che mi permetterà di estrarre le relative quantità di Rosso,
    'Verde e Blu dal colore appena selezionato
    Preleva_RGB CP
    'Richiamo la funzione pubblica addetta a convertire le rispettive quantità di Rosso,
    'Vetrde e Blu da volori RGB in RGBA
    RGB_To_RGBA CP
'Label al quale si verrà reinderizzati in caso di annullamento di selezione del colore
Annulla:
End Sub

Private Sub ColoreSfondoMenù_Click()
    'Se viene scelto di annullare l'operazione di selezione colore,allora si verrà reinderizzati
    'verso il "Label" Annulla
    On Error GoTo Annulla
    'Apro il componente che mi permetterà di selezionare il colore desiderato da assegnare
    'allo sfondo del menù che verrà creato sulla parte alta dell'editor
    Form_Materiali.ControlloSceltaColori.ShowColor
    'Imposto il colore di sfondo del pulsante che lo rappresenta il colore appena selezionato
    ColoreSfondoMenù.BackColor = Form_Materiali.ControlloSceltaColori.Color
    'Richinamo la funzione pubblica che mi permetterà di estrarre le relative quantità di Rosso,
    'Verde e Blu dal colore appena selezionato
    Preleva_RGB CSFM
    'Richiamo la funzione pubblica addetta a convertire le rispettive quantità di Rosso,
    'Vetrde e Blu da volori RGB in RGBA
    RGB_To_RGBA CSFM
'Label al quale si verrà reinderizzati in caso di annullamento di selezione del colore
Annulla:
End Sub

Private Sub ColoreSoffitti_Click()
    'Se viene scelto di annullare l'operazione di selezione colore,allora si verrà reinderizzati
    'verso il "Label" Annulla
    On Error GoTo Annulla
    'Apro il componente che mi permetterà di selezionare il colore desiderato da assegnare
    'alle righe che rappresentano i soffitti all'interno dell'editor
    Form_Materiali.ControlloSceltaColori.ShowColor
    'Imposto il colore di sfondo del pulsante che lo rappresenta il colore appena selezionato
    ColoreSoffitti.BackColor = Form_Materiali.ControlloSceltaColori.Color
    'Richinamo la funzione pubblica che mi permetterà di estrarre le relative quantità di Rosso,
    'Verde e Blu dal colore appena selezionato
    Preleva_RGB CS
    'Richiamo la funzione pubblica addetta a convertire le rispettive quantità di Rosso,
    'Vetrde e Blu da volori RGB in RGBA
    RGB_To_RGBA CS
'Label al quale si verrà reinderizzati in caso di annullamento di selezione del colore
Annulla:
End Sub

Private Sub ColoreSpigoliMuri_Click()
    'Se viene scelto di annullare l'operazione di selezione colore,allora si verrà reinderizzati
    'verso il "Label" Annulla
    On Error GoTo Annulla
    'Apro il componente che mi permetterà di selezionare il colore desiderato da assegnare
    'alle righe che rappresentano i muri all'interno dell'editor
    Form_Materiali.ControlloSceltaColori.ShowColor
    'Imposto il colore di sfondo del pulsante che lo rappresenta il colore appena selezionato
    ColoreSpigoliMuri.BackColor = Form_Materiali.ControlloSceltaColori.Color
    'Richinamo la funzione pubblica che mi permetterà di estrarre le relative quantità di Rosso,
    'Verde e Blu dal colore appena selezionato
    Preleva_RGB CSM
    'Richiamo la funzione pubblica addetta a convertire le rispettive quantità di Rosso,
    'Vetrde e Blu da volori RGB in RGBA
    RGB_To_RGBA CSM
'Label al quale si verrà reinderizzati in caso di annullamento di selezione del colore
Annulla:
End Sub

Private Sub ColoreSPSelezionati_Click()
    'Se viene scelto di annullare l'operazione di selezione colore,allora si verrà reinderizzati
    'verso il "Label" Annulla
    On Error GoTo Annulla
    'Apro il componente che mi permetterà di selezionare il colore desiderato da assegnare
    'alle 4 righe che rappresentano il pavimento / soffitto selezionato tramite l'oggetto Elenco_SoP
    Form_Materiali.ControlloSceltaColori.ShowColor
    'Imposto il colore di sfondo del pulsante che lo rappresenta il colore appena selezionato
    ColoreSPSelezionati.BackColor = Form_Materiali.ControlloSceltaColori.Color
    'Richinamo la funzione pubblica che mi permetterà di estrarre le relative quantità di Rosso,
    'Verde e Blu dal colore appena selezionato
    Preleva_RGB CSOPS
    'Richiamo la funzione pubblica addetta a convertire le rispettive quantità di Rosso,
    'Vetrde e Blu da volori RGB in RGBA
    RGB_To_RGBA CSOPS
'Label al quale si verrà reinderizzati in caso di annullamento di selezione del colore
Annulla:
End Sub

Private Sub Conferma1_Click()
    'Assegno al muro corrente il valore indicato nella Textbox Altezza_muro.
    'Questo mi permetterà di assegnare ad ogni muro un altezza differente
    Riga(IndiceLista).Altezza = Val(Altezza_muro.Text)
    'Assegno al muro corrente il valore indicato nella Textbox Altitudine_muro.
    'Anche qui sarà possibile assegnare ad ogni muro uno spessore differente
    Riga(IndiceLista).Altitudine = Val(Altitudine_muro.Text)
    'Assegno al muro selezionato un determinato numero di mattonelle poste in altezza
    Riga(IndiceLista).NMattonelleALtezza = Val(MattonelleAltezza.Text)
    'Assegno al muro selezionato un determinato numero di mattonelle poste in larghezza
    Riga(IndiceLista).NMAttonelleLarghezza = Val(MattonelleLarghezza.Text)
End Sub

Private Sub Conferma2_Click()
    'Assegno al pavimento o soffitto corrente il valore indicato nella Textbox Altitudine_muro.
    'Anche qui sarà possibile assegnare ad ogni pavimento o soffitto un'altitudine differente
    SoP(IndiceLista2).CR.Altitudine = Val(AltitudineSoP.Text)
    'Assegno al pavimento o soffitto selezionato un determinato numero di mattonelle poste in altezza
    SoP(IndiceLista2).CR.NMattonelleALtezza = Val(MattonelleAltezza2.Text)
    'Assegno al pavimento o soffitto selezionato un determinato numero di mattonelle poste in larghezza
    SoP(IndiceLista2).CR.NMAttonelleLarghezza = Val(MattonelleLarghezza2.Text)
End Sub

Private Sub ConfermaDescrizione_Click()
    'Salvo la nuova descrizione dell'oggetto corrispondente a quello selezionato dal controllo
    'ElencoGruppiOggetti
    Oggetto(IndiceOggettoSelezionato).Decrizione = DescrizioneOggetto.Text
End Sub

Private Sub ConfermaNuovoValoreOggetto_Click()
    Dim IOperazione As Integer
    Dim Asse As String
    Dim NuovoValore As Single
    IOperazione = OperazioneModificaOggetto.ListIndex
    Asse = AsseModificaOggetto
    NuovoValore = Val(NuovoValoreOggetto)
    Select Case IOperazione
    Case Is = 0
        Oggetto(IndiceOggettoSelezionato).Setta_Posizione Asse, NuovoValore
    Case Is = 1
        Oggetto(IndiceOggettoSelezionato).Setta_Dimensione Asse, NuovoValore
    Case Is = 2
        Oggetto(IndiceOggettoSelezionato).Setta_Rotazione Asse, NuovoValore
    End Select
End Sub

Private Sub DescrizioneOggetto_KeyPress(KeyAscii As Integer)
    'Dichiaro una variabile che limiterà la descrizione dell'oggetto selezionato a 200 caratteri
    Dim LunghezzaDescrizione As Integer
    'Dichiaro una variabile che mi servirà per contenere la descrizione abbreviata dell'oggetto
    Dim DescrizioneRistretta As String
    'Verifico la lunghezza della descrizione contenuta all'interno del controllo DescrizioneOggetto
    LunghezzaDescrizione = Len(DescrizioneOggetto)
    'Se tale lunghezza è > 200 allora l'utente verrà avvisato di abbreviare la descrizione
    If LunghezzaDescrizione > 200 Then
        'Visualizzazione del messaggio di errore
        MsgBox "Attenzione!La descrizione di questo oggetto supera la lunghezza massima di 200 caratteri!" + Chr(13) + "Si prega di abbreviare la descrizione!", vbInformation, "Lunghezza massima raggiunta!"
        'Salvo all'interno della variabile DescrizioneRistretta la stringa contenuta all'interno del controllo
        'DescrizioneOggetto escludendo però l'ultimo carattere
        DescrizioneRistretta = Mid(DescrizioneOggetto.Text, 1, 198)
        'Cancello il contenuto del controllo Descrizione Oggetto
        DescrizioneOggetto.Text = ""
        'Infine aggiorno il suo contenuto con la stessa descrizione escludendo però l'ultimo carattere
        DescrizioneOggetto = DescrizioneRistretta
    End If
End Sub

Private Sub DisattivaFondale_Click()
    'Disattivo il fondale dell'editor
    Schermo.EnableBackground False
End Sub

Private Sub ElencoGruppiOggetti_AfterLabelEdit(Cancel As Integer, NewString As String)
    'Dichiaro una variabile booleano che servirà al programma per capire se vi è un'altro
    'Node all'interno del controllo ElencoGruppiOggetti avente chiave uguale alla nuova stringa
    'introdotta
    Dim Doppio As Boolean
    'Dichiaro la variabile Nome la quale mi servirà per contenere il nome restituito dalla funzione
    'Nuovo_Nome in caso di nodi aventi la stessa chiave
    Dim Nome As String
    'Se il nodo selezionato all'interno del controllo GruppiOggetti non è un gruppo,allora...
    If ElencoGruppiOggetti.SelectedItem.ForeColor <> vbBlue Then
        'Dichiaro una variabile di appoggio che mi servirà per contenere l'indice dell'oggetto trovato
        'dalla funzione Trova_Oggetto
        Dim OggettoTrovato As Integer
        'Assegno alla variabile OggettoTrovato,il valore restituito dalla funzione Trova_Oggetto.
        'Questo mi servirà per ricavare l'indice all'interno dell'array Oggetti,appunto dell'oggetto
        'corrispondente a quello selezionato
        OggettoTrovato = Trova_Oggetto(ElencoGruppiOggetti.SelectedItem.Key)
        'Assegno alla variabile Doppio il valore restituito dalla funzione Verifica_Doppio.
        Doppio = Verifica_Doppio(NewString)
        'Se è già stato trovato un node avente per chiave la nuova stringa inserita,allora
        'richiamerò la funzione Nuovo_Nodo con il compito di trovare appunto uns nuova chiave
        'per il nodo selezionato
        If Doppio = True Then
            'Salvo all'interno della variabile Nome,la nuova chiave restituita dalla funzione
            'Nuovo_Nome
            Nome = Nuovo_Nome(NewString)
        'Nel caso in cui non siano stati trovati altri nodi aventi la stessa chiave,allora si
        'potrà utilizzare tranquillamente la nuova stringa inserita
        Else
            'Salvo all'interno della variabile Nome la nuova stringa inserita
            Nome = NewString
        End If
        'Assegno anche al corrispondente Node il valore di proprietà Key con il valore contenuto nella
        'variabile Nome
        ElencoGruppiOggetti.SelectedItem.Key = Nome
        'Assegno anche al corrispondente Node il valore di proprietà Text con la nuova stringa
        ElencoGruppiOggetti.SelectedItem.Text = NewString
        'Assegno alla variabile Key dell'oggetto trovato,il nuovo nome appena settato
        Oggetto(OggettoTrovato).Key = Nome
    End If
End Sub

Private Sub ElencoGruppiOggetti_NodeClick(ByVal Node As MSComctlLib.Node)
    'Se il nodo selezionato all'interno del controllo GruppiOggetti non è un gruppo,allora...
    If Node.ForeColor <> vbBlue Then
        'Assegno alla variabile pubblica IndiceOggettoSelezionato,il valore restituito dalla funzione Trova_Oggetto.
        'Questo mi servirà per ricavare l'indice all'interno dell'array Oggetti,appunto dell'oggetto
        'corrispondente a quello selezionato
        IndiceOggettoSelezionato = Trova_Oggetto(Node.Key)
        'Assegno alla varaibile pubblica OggettoSelezionato,la chiave univoca dell'oggetto trovato.
        'Questo mi permetterà di visualizzare la sua posizione all'interno dell'editor
        OggettoSelezionato = Oggetto(IndiceOggettoSelezionato).Key
        'Ora carico all'interno del controllo image TextureOggetto,la Texture relativa all'oggetto trovato
        TextureOggetto.Picture = LoadPicture(Oggetto(IndiceOggettoSelezionato).Texture)
        'Infine carico la sua breve descrizione all'interno del controllo DescrizioneOggetto
        DescrizioneOggetto.Text = Oggetto(IndiceOggettoSelezionato).Decrizione
    End If
End Sub

Private Sub ElencoMuri_Click()
    'Assegno alle quattro seguenti textbox il valore Enabled = False
    Altezza_muro.Enabled = False
    Altitudine_muro.Enabled = False
    MattonelleAltezza.Enabled = False
    MattonelleLarghezza.Enabled = False
    'Assegno all'IndiceLista il valore dell'elemento selezionato dall'oggetto ComoBox
    'ElencoMuri
    IndiceLista = ElencoMuri.ListIndex + 1
    'Le operazioni sottoindicate si verificheranno solo se l'indice della Combobox ElencoMuri
    'sarà > -1
    If ElencoMuri.ListIndex > -1 Then
        'Assegno alla Textbox NomeR lo stesso valore alfanumerico dell'oggetto
        'elencoMuri.text.Questo mi permetterà in seguito di poterlo rinominare per
        'riuscire meglio ad individuare ogni singolo elemento
        Nomer.Text = ElencoMuri.Text
        'Assegno alla Textbox Altezza_Muro il valore dell'altezza della corrispondente
        'linea selezionata convertito in stringa
        Altezza_muro.Text = Str(Riga(IndiceLista).Altezza)
        'Assegno alla Textbox Alitudine_Muro il valore dell'altitudine della corrispondente
        'linea selezionata convertito in stringa
        Altitudine_muro.Text = Str(Riga(IndiceLista).Altitudine)
        'Assegno la Texture scelta alla rispettiva linea selezionata solo se il suo valore è
        'diverso da "Nessuna",altrimenti non viene assegnata nessuna immagine
        If Trim(Riga(IndiceLista).Texture) <> "Nessuna" Then
            Texture_muro.Picture = LoadPicture(Trim(Riga(IndiceLista).Texture))
        Else
            Texture_muro.Picture = Nothing
        End If
        'Assegno alle textbox MattonelleAltezza il rispettivo numero di mattonelle (convertito in stringa)
        'presenti in altezza
        MattonelleAltezza.Text = Str(Riga(IndiceLista).NMattonelleALtezza)
        'Assegno alle textbox MattonelleLarghezza il rispettivo numero di mattonelle (convertito in stringa)
        'presenti in larghezza
        MattonelleLarghezza.Text = Str(Riga(IndiceLista).NMAttonelleLarghezza)
    End If
End Sub

Private Sub ElencoSoP_Click()
    'Assegno alle quattro seguenti textbox il valore Enabled = False
    AltitudineSoP.Enabled = False
    MattonelleAltezza2.Enabled = False
    MattonelleLarghezza2.Enabled = False
    'Assegno alla variabile IndiceLista2 il valore dell'elemento selezionato dall'oggetto ComboBox
    'ElencoSoP
    IndiceLista2 = ElencoSoP.ListIndex + 1
    'Le operazioni sottoindicate si verificheranno solo se l'indice della Combobox ElencoMuri
    'sarà > -1
    If ElencoSoP.ListIndex > -1 Then
        'Assegno alla Textbox NomeR lo stesso valore alfanumerico dell'oggetto
        'elencosop.text.Questo mi permetterà in seguito di poterlo rinominare per
        'riuscire meglio ad individuare ogni singolo elemento
        Nomer2.Text = ElencoSoP.Text
        'Assegno alla Textbox Alitudine_sop il valore dell'altitudine del corrispondente pavimento o soffitto
        'selezionato convertito in stringa
        AltitudineSoP.Text = Str(SoP(IndiceLista2).CR.Altitudine)
        'Assegno la Texture scelta al rispettivo pavimento o soffitto selezionato solo se il suo valore è
        'diverso da "Nessuna",altrimenti non viene assegnata nessuna immagine
        If Trim(SoP(IndiceLista2).CR.Texture) <> "Nessuna" Then
            Texture_SoP.Picture = LoadPicture(Trim(SoP(IndiceLista2).CR.Texture))
        Else
            Texture_SoP.Picture = Nothing
        End If
        'Assegno alle textbox MattonelleAltezza il rispettivo numero di mattonelle (convertito in stringa)
        'presenti in altezza
        MattonelleAltezza2.Text = Str(SoP(IndiceLista2).CR.NMattonelleALtezza)
        'Assegno alle textbox MattonelleLarghezza il rispettivo numero di mattonelle (convertito in stringa)
        'presenti in larghezza
        MattonelleLarghezza2.Text = Str(SoP(IndiceLista2).CR.NMAttonelleLarghezza)
        'Se l'oggetto analizzato è un pavimento allora viene selezionato l'option button
        'TipoPavimento,altrimenti viene selezionato l'option button TipoSoffitto.
        'Questo permetterà all'utente,in caso l'oggetto venisse rinominato,di capire se l'oggetto
        'in questione è appunto un pavimento oppure un soffitto
        If RTrim(SoP(IndiceLista2).Tipo) = "Pavimento" Then
            TipoPavimento.Value = True
        Else
            TipoSoffitto.Value = True
        End If
    End If
End Sub

Private Sub EliminaGruppo_Click()
    '
End Sub

Private Sub EliminaMuro_Click()
    'Dichiaro un indice che mi servirà per identificare ogni singolo muro
    Dim K As Integer
    'Inizio con il rimuovere il muro selezionato dalla lista dei muri presenti
    ElencoMuri.RemoveItem (IndiceLista - 1)
    'Ora avvio una ricerca che,in caso questo muro fosse collegato ad altri mediante spigoli,
    'questi assumerebbero i valori false,in modo che non siano più visibili all'interno
    'della mappa corrente
    'Controllo gli spigoli iniziali dei muri
    For K = 0 To Max
        If Riga(K).X2 = Riga(IndiceLista).X1 And Riga(K).Y2 = Riga(IndiceLista).Y1 And Riga(K).SpigoloF = True Then
                Riga(K).SpigoloF = False
        ElseIf Riga(K).X2 = Riga(IndiceLista).X2 And Riga(K).Y2 = Riga(IndiceLista).Y2 And Riga(K).SpigoloF = True Then
                Riga(K).SpigoloF = False
        End If
        If Riga(K).X1 = Riga(IndiceLista).X1 And Riga(K).Y1 = Riga(IndiceLista).Y1 And Riga(K).SpigoloI = True Then
                Riga(K).SpigoloI = False
        ElseIf Riga(K).X1 = Riga(IndiceLista).X2 And Riga(K).Y1 = Riga(IndiceLista).Y2 And Riga(K).SpigoloI = True Then
                Riga(K).SpigoloI = False
        End If
    Next
    'Ora avvio l'operazione di Slide dei muri,ovvero tutti i muri che erano posti in una posizione
    'superiore rispetto al muro eliminato dovranno scalare indietro di una posizione
    For K = IndiceLista + 1 To Max
        Riga(K - 1) = Riga(K)
    Next
    'Aggiorno la variabile Max che tiene conto del numero di muri presenti
    Max = Max - 1
    'Aggiorno la variabile che mi permetterà di aggiungere un muro nella sua giusta
    'collocazione all'interno della mappa
    I = I - 1
    'Reinizializzo il muro subito dopo a quello appena cancellato con i valori di default
    With Riga(Max + 1)
        .X1 = 0
        .X2 = 0
        .Y1 = 0
        .Y2 = 0
        .Altezza = 1000
        .Altitudine = 0
        .Nome = "Muro" + Str(Max + 1)
        .Proprietà = "Normale"
        For K = 0 To 3
            With .ColVertici(K)
                .R = 0
                .G = 0
                .B = 0
                .A = 0.5
            End With
        Next
        .Texture = "Nessuna"
        .NMattonelleALtezza = 10
        .NMAttonelleLarghezza = 10
        .SpigoloI = False
        .SpigoloF = False
    End With
    'Richiamo la funzione pubblica addetta alla reimpostazione dei parametri di materiale
    'del muro corrente.
    'Il valore passato alla stessa è uguale a 6,in modo che il materiale reimpostato sia
    'quello del muro subito dopo a quello appena cancellato
    Reimposta_materiale Riga(Max + 1).Materiale
End Sub

Private Sub EliminaOggetto_Click()
    'Dichiaro una variabile che fungerà da indice all'interno dell'operazione di slide
    'degli oggetti
    Dim NOggetto As Integer
    'Se non è stato selezionato nessun oggetto si verrà reinderizzati al Label Annulla,nel quale
    'si provvederà ad avvisare l'utente dell'errore commesso
    On Error GoTo Annulla
    'Se l'utente avesse selezionato per errore un gruppo al posto di un oggetto,verrà avvertito
    'dell'errore tramite un messaggio
    If ElencoGruppiOggetti.SelectedItem.ForeColor = vbBlue Then
        'Visualizzazione del messaggio di errore
        If LinguaS = "Italiano" Then MsgBox "L'elemento selezionato è un gruppo!" + Chr(13) + "Si prega di selezionare un oggetto!", vbInformation, "Attenzione!"
        If LinguaS = "Inglese" Then MsgBox "The selected element is a group!" + Chr(13) + "Please select an object!", vbInformation, "Attention!"
    'Se invece il colore della scritta del nodo selezionato è diverso da blu,questo sarà un oggetto e si potrà
    'proseguire con la procedura di eliminazione
    Else
        'Eliminiamo il nodo selezionato dal controllo ElencoGruppiOggetti
        ElencoGruppiOggetti.Nodes.Remove (ElencoGruppiOggetti.SelectedItem.Index)
        'Ora devo scalare tutti gli oggetti posti dopo l'oggetto eliminato in modo da non lasciare buchi
        'all'interno dell'array Oggetto
        For NOggetto = IndiceOggettoSelezionato To IOg
            'Questa istruzione mi serve al fine di richiamare la funzione friend addetta a ricopiare
            'il secondo oggetto passato alla funzione,all'interno del primo sempre passato alla funzione,in
            'modo da permettermi di effettuare uno slide di tutti quegli oggetti posti dopo quello eliminato
            'all'interno dell'Array Oggetto
            Copia_Oggetto Oggetto(NOggetto), Oggetto(NOggetto + 1)
        Next
        'Richiamo il metodo Distruggi_Oggetto che avrà la funzione di reinizializzare tutte le variabili dell'ultimo
        'oggetto dell'Array Oggetto,in quanto questo non serve più
        Oggetto(IOg).Distruggi_Oggetto
        'Diminuisco il contatore contenente il numero di oggetti caricati all'interno della mappa 3D
        IOg = IOg - 1
    End If
    'Si esce dalla funzione
    Exit Sub
'Label al quale si verrà reinderizzati in caso di mancato selezionamento di un oggetto
Annulla:
    'Visualizzazione del messaggio di errore che avviserà l'utente che non è stato selezionato nessun oggetto
    If LinguaS = "Italiano" Then MsgBox "Non è stato selezionato nessun oggetto da eliminare!" + Chr(13) + "Si prega si selezionare un oggetto dall'elenco!", vbInformation, "Attenzione!"
    If LinguaS = "Inglese" Then MsgBox "You haven't selected any object to remove!" + Chr(13) + "Please select an object!", vbInformation, "Attention"
End Sub

Private Sub EliminaSoP_Click()
    'Dichiaro un indice che mi servirà per identificare ogni singolo pavimento o soffitto
    Dim K As Integer
    'Ora effettuo un piccolo controllo:
    'Se l'elemento appena eliminato era un pavimento,allora dovrò decrementare la variabile che tiene
    'conto del numero di pavimenti presenti all'interno della mappa attuale...
    If RTrim(SoP(IndiceLista2).Tipo) = "Pavimento" Then
        Max3 = Max3 - 1
    '...invece,nel caso in cui l'elemento appena eliminato era un soffitto, questa volta verrà decrementata la
    'variabile che tiene conto del numero di soffitti
    ElseIf RTrim(SoP(IndiceLista2).Tipo) = "Soffitto" Then
        Max4 = Max4 - 1
    End If
    'Inizio con il rimuovere il muro selezionato dalla lista dei muri presenti
    ElencoSoP.RemoveItem (IndiceLista2 - 1)
    'Ora avvio l'operazione di Slide dei pavimenti o soffitti,ovvero tutti i pavimenti o soffitti che erano posti in una posizione
    'superiore rispetto al "SoP" eliminato dovranno scalare indietro di una posizione
    For K = IndiceLista2 + 1 To Max2
        SoP(K - 1) = SoP(K)
    Next
    'Aggiorno la variabile Max2 che tiene conto del numero dei pavimenti + i soffitti presenti
    Max2 = Max2 - 1
    'Aggiorno la variabile che mi permetterà di aggiungere un pavimento o soffitto nella sua giusta
    'collocazione all'interno della mappa
    J = J - 1
    'Reinizializzo il pavimento o soffitto subito dopo a quello appena cancellato con i valori di default
    With SoP(Max2 + 1)
        .CR.X1 = 0
        .CR.X2 = 0
        .X3 = 0
        .X4 = 0
        .CR.Y1 = 0
        .CR.Y2 = 0
        .CR.Altezza = 0
        .CR.Altitudine = 0
        .CR.Nome = ""
        .CR.Proprietà = "Normale"
        .Tipo = ""
        For K = 0 To 3
            With .CR.ColVertici(K)
                .R = 0
                .G = 0
                .B = 0
                .A = 0.5
            End With
        Next
        .CR.Texture = "Nessuna"
        .CR.NMattonelleALtezza = 20
        .CR.NMAttonelleLarghezza = 20
        .CR.SpigoloI = False
        .CR.SpigoloF = False
    End With
    'Richiamo la funzione pubblica addetta alla reimpostazione dei parametri di materiale
    'del pavimento o soffitto corrente.
    'Il valore passato alla stessa è uguale a 6,in modo che il materiale reimpostato sia
    'quello del pavimento o soffitto subito dopo a quello appena cancellato
    Reimposta_materiale SoP(Max2 + 1).CR.Materiale
End Sub



Private Sub Materiale_Click()
    'Se non è stato selezionato nessun muro dall'elenco di quelli disponibili dall'oggetto combobox ElencoMuri
    'allora verrà visualizzato un messaggio che avviserà l'utente dell'errore
    If Form_Opzioni.ElencoMuri.Text = "" Then
        If LinguaS = "Italiano" Then MsgBox "Devi prima selezionare un muro disponibile dall'elenco per assegnare un materiale!", vbInformation, "Assegnazione materiale muro"
        If LinguaS = "Inglese" Then MsgBox "You must choose an available wall from the list first", vbInformation, "Wall material assegnation"
    Else
        'Assegno alla variabile ModalitàGestioneMateriale il valore alfanumerico Muri, in modo da far capire al programma
        'che dovrà avviare il sottoscritto form al fine di settare appunto tutte le proprietà del
        'materiale del muro selezionato
        ModalitàGestioneMateriale = "Muri"
        'Attiva su schermo il form per la gestione dei materiali nella modalità settata
        Form_Materiali.Show
    End If
End Sub

Private Sub Materiale2_Click()
    'Se non è stato selezionato nessun pavimento o soffitto dall'elenco di quelli disponibili dall'oggetto combobox ElencoSoP
    'allora verrà visualizzato un messaggio che avviserà l'utente dell'errore
    If Form_Opzioni.ElencoSoP.Text = "" Then
        If LinguaS = "Italiano" Then MsgBox "Devi prima selezionare un pavimento o soffitto disponibile dall'elenco per assegnare un materiale!", vbInformation, "Assegnazione materiale pavimento / soffitto"
        If LinguaS = "Inglese" Then MsgBox "You must choose an available floor or ceiling from the list first", vbInformation, "Floor / ceiling material assegnation"
    Else
        'Se l'elemento selezionato è un pavimento,allora la variabile ModalitàGestioneMateriale assumerà il valore alfanumerico
        'Pavimento,in modo da far capire al programma che si stà assegnando il materiale di un pavimento
        If RTrim(SoP(IndiceLista2).Tipo) = "Pavimento" Then
            ModalitàGestioneMateriale = "Pavimento"
        'Nel caso contrario fosse un soffitto, allora la variabile ModalitàGestioneMateriale assumerà il valore alfanumerico
        'soffitto,in modo da far capire al programma che si stà assegnando il materiale di un soffitto
        ElseIf RTrim(SoP(IndiceLista2).Tipo) = "Soffitto" Then
            ModalitàGestioneMateriale = "Soffitto"
        End If
        'Attiva su schermo il form per la gestione dei materiali nella modalità settata
        Form_Materiali.Show
    End If
End Sub

Private Sub Modifica1_Click()
    'Assegno alle quattro seguenti textbox il valore Enabled = True in modo da
    'poter applicare eventuali cambiamenti
    Altezza_muro.Enabled = True
    Altitudine_muro.Enabled = True
    MattonelleAltezza.Enabled = True
    MattonelleLarghezza.Enabled = True
End Sub

Private Sub Modifica2_Click()
    'Assegno alle quattro seguenti textbox il valore Enabled = True in modo da
    'poter applicare eventuali cambiamenti
    AltitudineSoP.Enabled = True
    MattonelleAltezza2.Enabled = True
    MattonelleLarghezza2.Enabled = True
End Sub


Private Sub NessunaTexture_Click()
    'Elimino dal muro corrente la texture che si aveva selezionato
    Riga(IndiceLista).Texture = "Nessuna"
    'Cancello l'immagine residua dall'oggetto image Texture_muro
    Texture_muro.Picture = Nothing
End Sub

Private Sub NessunaTexture2_Click()
    'Elimino dal pavimento / soffitto corrente la texture che si aveva selezionato
    SoP(IndiceLista2).CR.Texture = "Nessuna"
    'Cancello l'immagine residua dall'oggetto image Texture_SoP
    Texture_SoP.Picture = Nothing
End Sub


Private Sub Rinomina_Click()
    'Rimuovo l'elemento selezionato dall'oggetto elencoMuri...
    ElencoMuri.RemoveItem (IndiceLista - 1)
    '...per poi riinserirlo con il suo nuova nome
    ElencoMuri.AddItem Nomer.Text, IndiceLista - 1
    'Infine visualizzo il nuovo nome all'interno della ComboBox ElencoMuri
    ElencoMuri.Text = Nomer.Text
    'Assegnpo il nuovo nome inserito alla riga corrispondente
    Riga(IndiceLista).Nome = Nomer.Text
End Sub

Private Sub CambiaImmagineDiSfondo_Click()
    'Imposto il titolo dell'oggetto operazioni,in modo che l'utente potrà capire quale operazione
    'verrà effettuata
    If LinguaS = "Italiano" Then CD2.DialogTitle = "Caricamento fondale editor"
    If LinguaS = "Inglese" Then CD2.DialogTitle = "Load editor background"
    'Impongo una serie di filtri che corrispondono ai più comuni formati di immagine
    If LinguaS = "Italiano" Then CD2.Filter = "Tutti i formati |*.bmp;*.jpg|BMP (Bitmap di Windows)|*.bmp|JPG (Formato di interscambio dei file JPEG)|*.jpg"
    If LinguaS = "Inglese" Then CD2.Filter = "All the formats |*.bmp;*.jpg|BMP (Bitmap of Windows)|*.bmp|JPG (Interchange file of JPEG format)|*.jpg"
    'Apro dal Map_editor il componente che mi permetterà di selezionare un immagine da applicare come
    'fondale dell'editor
    On Error GoTo Annulla
    CD2.ShowOpen
    'Assegno alla variabile ImmagineSfondo il percorso e l'immagine appena selezionata
    ImmagineSfondo = CD2.FileName
    'Attivo il fondale dell'editor
    Schermo.EnableBackground True
    'Carico l'immagine appena selezionata come fondale dell'editor
    Schermo.LoadBackground ImmagineSfondo
Annulla:
End Sub

Private Sub ConfermaTelecamera_Click()
    'Imposto il vettore di coordinate telecamera con i valori aggiornati
    With PosizioneTelecamera
        .X = Val(TelecameraX.Text)
        .Y = Val(TelecameraY.Text)
        .Z = Val(TelecameraZ.Text)
    End With
End Sub

Private Sub FondaleStatico_Click()
    'Se non è stata selezionata nessuna immagine di sfondo,allora aprirò dal Map_editor il componente
    'che mi permetterà di selezionare un'immagine da applicare come fondale dell'editor
    If ImmagineSfondo = "Nessuna" Then
        On Error GoTo Annulla
        'Imposto il titolo dell'oggetto operazioni,in modo che l'utente potrà capire quale operazione
        'verrà effettuata
        If LinguaS = "Italiano" Then CD2.DialogTitle = "Caricamento fondale editor"
        If LinguaS = "Inglese" Then CD2.DialogTitle = "Load editor background"
        'Impongo una serie di filtri che corrispondono ai più comuni formati di immagine
        If LinguaS = "Italiano" Then CD2.Filter = "Tutti i formati |*.bmp;*.jpg|BMP (Bitmap di Windows)|*.bmp|JPG (Formato di interscambio dei file JPEG)|*.jpg"
        If LinguaS = "Inglese" Then CD2.Filter = "All the formats |*.bmp;*.jpg|BMP (Bitmap di Windows)|*.bmp|JPG (Interchange format of JPEG files)|*.jpg"
        'Apro il componente per selezionare l'immagine
        CD2.ShowOpen
        'Assegno alla variabile ImmagineSfondo,il percorso e l'immagine appena selezionata
        ImmagineSfondo = CD2.FileName
        'Attivo il fondale dell'editor
        Schermo.EnableBackground True
        'Carico l'immagine appena selezionata come fondale dell'editor
        Schermo.LoadBackground ImmagineSfondo
    'Altrimenti se l'immagine è già stata precedentemente selezionata...
    Else
        'Attivo solamente il fondale dell'editor...
        Schermo.EnableBackground True
        'Carico l'immagine già precedentemente selezionata,all'interno dell'editor
        Schermo.LoadBackground ImmagineSfondo
    End If
Annulla:
End Sub

Private Sub Impostato_Click()
    'Disattivo la textbox di personalizzazione dello scale,e il pulsante di comando Salva.
    'Inoltre attivo la combobox di selezione scale impostato
    ValoreScalePersonalizzato.Enabled = False
    SalvaScalePersonalizzato.Enabled = False
    ValoreScale.Enabled = True
End Sub

Private Sub ModificaTelecamera_Click()
    Scelta_Oggetto = "Telecamera"
    Controlla_pulsanti
End Sub

Private Sub Personalizzato_Click()
    'Attivo la textbox in cui sarà possibile definire uno scale personalizzato e il pulsante
    'di comando Salva.Inoltre disattivero la combobox di seleziona scale impostato
    ValoreScalePersonalizzato.Enabled = True
    SalvaScalePersonalizzato.Enabled = True
    ValoreScale.Enabled = False
End Sub

Private Sub Rinomina2_Click()
    'Rimuovo l'elemento selezionato dall'oggetto ElencoSoP...
    ElencoSoP.RemoveItem (IndiceLista2 - 1)
    '...per poi riinserirlo con il suo nuova nome
    ElencoSoP.AddItem Nomer2.Text, IndiceLista2 - 1
    'Infine visualizzo il nuovo nome all'interno della ComboBox ElencoMuri
    ElencoSoP.Text = Nomer2.Text
    'Assegnpo il nuovo nome inserito alla riga corrispondente
    SoP(IndiceLista2).CR.Nome = Nomer2.Text
End Sub

Private Sub RipristinaZoom_Click()
    'Dichiaro un indice per individuare singolarmente ogni lìriga della mappa attuale
    Dim K As Integer
    'Avvio un ciclo for per reimpostare ai loro valori originali tutti i muri che erano stati modificati dalle operazioni
    'di zoom
    For K = 0 To Max
        Riga(K).X1 = Fix(Riga(K).X1 / Molt)
        Riga(K).X2 = Fix(Riga(K).X2 / Molt)
        Riga(K).Y1 = Fix(Riga(K).Y1 / Molt)
        Riga(K).Y2 = Fix(Riga(K).Y2 / Molt)
    Next
    'Ora svolgo la stessa funzione anche per i valori di tutti i pavimenti e soffitti creati
    For K = 0 To Max2
        SoP(K).CR.X1 = Fix(SoP(K).CR.X1 / Molt)
        SoP(K).CR.X2 = Fix(SoP(K).CR.X2 / Molt)
        SoP(K).X3 = Fix(SoP(K).X3 / Molt)
        SoP(K).X4 = Fix(SoP(K).X4 / Molt)
        SoP(K).CR.Y1 = Fix(SoP(K).CR.Y1 / Molt)
        SoP(K).CR.Y2 = Fix(SoP(K).CR.Y2 / Molt)
        SoP(K).Y3 = Fix(SoP(K).Y3 / Molt)
        SoP(K).Y4 = Fix(SoP(K).Y4 / Molt)
    Next
    'Ora,la stessa funzione,la applico anche per le coordinate di posizione di tutti gli oggetti caricati all'interno
    'della mappa 3D
    For K = 0 To IOg
        With Oggetto(K)
            .Setta_Posizione "X", Fix(.Ricava_Posizione.X / Molt)
            .Setta_Posizione "Z", Fix(.Ricava_Posizione.Z / Molt)
        End With
    Next
    'Ora,la stessa funzione,la applico anche per le coordinate di posizione della telecamera
    With PosizioneTelecamera
        .X = Fix(.X / Molt)
        .Z = Fix(.Z / Molt)
    End With
    'Reimposto il valore della variabile Molt a 1
    Molt = 1
    'Reinizializzo la variabile VCambiamentiGriglia con il valore 60,in modo che se il controllo GrigliaControllataDaZoom è
    'attivato,allora anche la griglia tornerà alla sua reale dimensione
    VCambiamentiGriglia = 60
    'Reinizializzo la variabile Cont
    Cont = 0
End Sub

Private Sub SalvaScalePersonalizzato_Click()
    '"Prendo" il valore scritto nella textbox ValoreScalePersonalizzato e lo salvo
    'nella variabile VScale,solo se il valore immesso non è minore di 0 o superiore 32767,
    'cioè il valore massimo positivo che una variabile integer può assumere.Questa
    'condizione mi servirà per evitare un blocco da parte del programma.
    'Dichiaro una variabile temporanea che conterrà il valore introdotto nella textbox
    'ValoreScalePersonalizzato
    Dim TmpScale As Long
    'Assegno il valore introdotto alla variabile temporanea corrispondente
    TmpScale = Val(ValoreScalePersonalizzato.Text)
    Select Case TmpScale
    Case 1 To 32767
        VScale = Val(ValoreScalePersonalizzato.Text)
    Case Is <= 0
        If LinguaS = "Italiano" Then MsgBox "Attenzione!Il valore dello scale non può avere un valore minore o uguale a 0!", vbInformation, "Impostazioni di scale"
        If LinguaS = "Inglese" Then MsgBox "Attention!The scale value can't be lower or equal than 0!", vbInformation, "Scale impostation"
        ValoreScalePersonalizzato.Text = Str(VScale)
    Case Is > 32767
        If LinguaS = "Italiano" Then MsgBox "Attenzione!Il valore introdotto supera il limite concesso!", vbInformation, "Impostazioni di scale"
        If LinguaS = "Inglese" Then MsgBox "Attention!The value is over concess limit!", vbInformation, "Scale impostation"
        ValoreScalePersonalizzato.Text = Str(VScale)
    End Select
End Sub

Private Sub SpostaOggettoInGruppo_Click()
    'Dichiaro una variabile che mi servirà per capire se è stato selezionato un oggetto
    'dal controllo ElencoGruppiOggetti
    Dim Trovato As Boolean
    'Dichiaro una variabile che mi servirà per capire se il gruppo in cui spostare l'oggetto
    'selezionato esiste
    Dim GruppoTrovato As Boolean
    'Dichiaro un indice di appoggio
    Dim M As Integer
    'Dichiaro una variabile temporanea in cui andrò a salvare la chiave univoca dell'oggetto
    'selezionato
    Dim TmpKey As String
    'Dichiaro una variabile temporanea in cui andrò a salvare la chiave univoca del gruppo selezionato
    'selezionato
    Dim Tmpkey2 As String
    'La variabile Nome mi servirà per contenere la proprietà Text dell'oggetto selezionato
    Dim Nome As String
    'Ricerco il gruppo del quale è stato immesso il nome all'interno del controllo ElencoGruppiOggetti
    'presente nel Form_Opzioni
    With Form_Opzioni.ElencoGruppiOggetti
        'Per prima cosa inizio a ricercare l'oggetto selezionato partendo da 0 e arrivando al numero
        'di nodi totali presenti all'interno del controllo ElencoGruppiOggetti
        For M = 1 To .Nodes.Count
            'Se il colore del nodo attualmente analizzato è diverso da blu(ciò vuol dire che non è un
            'nono padre quindi un gruppo),e la sua chiave coincide con quella dell'oggetto da ricercare allora...
            'Se il colore del nodo attualmente analizzato è blu e la sua chiave coincide con quella del
            'gruppo richiesto allora...
            If .Nodes.Item(M).ForeColor = vbBlue And .Nodes.Item(M).Text = NomeGruppo.Text Then
                'Estendo il gruppo trovato per far vedere all'utente che l'operazione di spostamento è
                'avvenuta con successo
                .Nodes.Item(M).Expanded = True
                'Setto la variabile GruppoTrovato uguale a True che starà ad indicare che il gruppo
                'richiesto esiste realmente ed è stato trovato
                GruppoTrovato = True
                'Salvo la chiave dell'oggetto trovato all'interno della variabile TmpKey2
                Tmpkey2 = .Nodes.Item(M).Key
            End If
        Next
        If GruppoTrovato = True Then
            'Riporto la variabile M uguale a 1 per iniziare la nuova ricerca
            M = 1
            'Ora inizio La ricerca per trovare il gruppo richiesto
            Do
                If .Nodes.Item(M).ForeColor <> vbBlue And .Nodes.Item(M).Key = .SelectedItem.Key Then
                    'Assegno alla variabile TmpKey la chiave univoca dell'oggetto trovato
                    TmpKey = .SelectedItem.Key
                    'Salvo all'interno della variabile Nome la proprietà Text dell'oggetto selezionato
                    Nome = .SelectedItem.Text
                    'Rimuovo il nodo trovato,raffigurante l'oggetto selezionato...
                    .Nodes.Remove (M)
                    'lo riaggiungo all'interno del suo nuovo gruppo
                    .Nodes.Add Tmpkey2, tvwChild, TmpKey, Nome
                    'Pongo la variabile Trovato uguale a True in modo da far capire che è stato trovato
                    'l'oggetto selezionato
                    Trovato = True
                End If
                'Incremento la variabile M in modo tale che si possa passare ad analizzare il nodo successivo
                'del controllo ElencoGruppiOggetti
                M = M + 1
                'Continua a ciclare finchè non è stato trovato il gruppo richiesto oppure si è già arrivati ad
                'analizzare l'ultimo nodo
            Loop Until Trovato = True Or M > .Nodes.Count
        End If
    End With
    'Se non è stato trovato l'oggetto selezionato,l'utente verrà avvertito tramite un messaggio che segnalerà l'errore
    If Trovato = False Then
        If LinguaS = "Italiano" Then MsgBox "Devi selezionare prima un oggetto per poterlo spostare di gruppo!", vbInformation, "Attenzione!"
        If LinguaS = "Inglese" Then MsgBox "You must choose an object first for move it in a group!", vbInformation, "Attention"
        'Se non è stato trovato il gruppo richiesto,l'utente verrà avvertito tramite un messaggio che segnalerà l'errore
        If GruppoTrovato = False Then
            If LinguaS = "Italiano" Then MsgBox "Il gruppo richiesto non esiste!", vbInformation, "Errore!"
            If LinguaS = "Inglese" Then MsgBox "The request group is non-existen!", vbInformation, "Error!"
        End If
    End If
End Sub

Private Sub Tabella_Click()
    If Tabella.SelectedItem.Index = 3 Then
        Preferenze.Visible = True
        FrameCostruisci.Visible = False
        FrameElencoCostruzioni.Visible = False
        FrameGestioneOggetti.Visible = False
    ElseIf Tabella.SelectedItem.Index = 1 Then
        Preferenze.Visible = False
        FrameCostruisci.Visible = True
        FrameElencoCostruzioni.Visible = True
        FrameGestioneOggetti.Visible = False
    ElseIf Tabella.SelectedItem.Index = 2 Then
        Preferenze.Visible = False
        FrameCostruisci.Visible = False
        FrameElencoCostruzioni.Visible = False
        FrameGestioneOggetti.Visible = True
    End If
End Sub

Private Sub CambiaTexture_Click()
    'Se non viene selezionato nessun muro su cui applicare la Texture allora viene
    'visualizzato un messaggio che informa l'utente di scegliere prima un muro
    If ElencoMuri.Text = "" Then
        If LinguaS = "Italiano" Then MsgBox "Per caricare una texture devi prima selezionare un muro esistente!", vbInformation, "Avviso!"
        If LinguaS = "Inglese" Then MsgBox "You must choose an existen wall first for load a texture!", vbInformation, "Advise!"
    Else
        'Imposto il titolo dell'oggetto operazioni,in modo che l'utente potrà capire quale operazione
        'verrà effettuata
        If LinguaS = "Italiano" Then CD2.DialogTitle = "Caricamento Texture"
        If LinguaS = "Inglese" Then CD2.DialogTitle = "Texture Load"
        'Impongo una serie di filtri che corrispondono ai più comuni formati di immagine
        If LinguaS = "Italiano" Then CD2.Filter = "Tutti i formati |*.bmp;*.jpg|BMP (Bitmap di Windows)|*.bmp|JPG (Formato di interscambio dei file JPEG)|*.jpg"
        If LinguaS = "Inglese" Then CD2.Filter = "All the formats |*.bmp;*.jpg|BMP (Bitmap di Windows)|*.bmp|JPG (Interchange format of JPEG files)|*.jpg"
        On Error GoTo Annulla
        'Apro il componente che mi permetterà di selezionare un'immagine da settare come
        'textures per il muro corrente
        CD2.ShowOpen
        'Assegno al muro corrente l'immagine selezionata e la applico come textures
        Riga(IndiceLista).Texture = CD2.FileName
        'Carico l'immagine appena selezionata all'interno dell'oggetto Texture_Muro
        Texture_muro.Picture = LoadPicture(RTrim(Riga(IndiceLista).Texture))
    End If
    Exit Sub
Annulla:
End Sub

Private Sub Form_Load()
    Scelta_Oggetto = "Muro"
    'Richiama la funzione che assegnerà uno stile diverso al bottone cliccato,in modo da avere
    'una rapida visualizzazione dell'oggetto che si stà inserendo
    Controlla_pulsanti
    With ElencoGruppiOggetti.Nodes.Add(, , "OSG", "Oggetti senza gruppo")
        'Setto il colore blu al nodo appena aggiunto
        .ForeColor = vbBlue
    End With
End Sub

Private Sub Muri_Click()
    'Riferiamo al programma che l'oggetto che si vorrà inserire sarà un muro
    Scelta_Oggetto = "Muro"
    'Richiama la funzione che assegnerà uno stile diverso al bottone cliccato,in modo da avere
    'una rapida visualizzazione dell'oggetto che si stà inserendo
    Controlla_pulsanti
End Sub

Private Sub Pavimento_Click()
    'Riferiamo al programma che l'oggetto che si vorrà inserire sarà un pavimento
    Scelta_Oggetto = "Pavimento"
    'Richiama la funzione che assegnerà uno stile diverso al bottone cliccato,in modo da avere
    'una rapida visualizzazione dell'oggetto che si stà inserendo
    Controlla_pulsanti
End Sub

Private Sub Soffitto_Click()
    'Riferiamo al programma che l'oggetto che si vorrà inserire sarà un soffitto
    Scelta_Oggetto = "Soffitto"
    'Richiama la funzione che assegnerà uno stile diverso al bottone cliccato,in modo da avere
    'una rapida visualizzazione dell'oggetto che si stà inserendo
    Controlla_pulsanti
End Sub

Sub Controlla_pulsanti()
    'La struttura di controllo che segue, esegue delle piccole istruzioni che assegneranno hai pulsanti,
    'a seconda del valore della variabile Scelta_Oggetto,il colore Oro,oppure grigio.
    'Il colore Oro verrà assgnato al pulsante selezionato,e quello grigio,per tutti gli altri
    Select Case Scelta_Oggetto
    Case Is = "Muro"
        Muri.BackColor = RGB(255, 215, 0)
        Pavimento.BackColor = &H8000000F
        Soffitto.BackColor = &H8000000F
        ModificaTelecamera.BackColor = &H8000000F
    Case Is = "Pavimento"
        Pavimento.BackColor = RGB(255, 215, 0)
        Muri.BackColor = &H8000000F
        Soffitto.BackColor = &H8000000F
        ModificaTelecamera.BackColor = &H8000000F
    Case Is = "Soffitto"
        Soffitto.BackColor = RGB(255, 215, 0)
        Muri.BackColor = &H8000000F
        Pavimento.BackColor = &H8000000F
        ModificaTelecamera.BackColor = &H8000000F
    Case Is = "Telecamera"
        ModificaTelecamera.BackColor = RGB(255, 215, 0)
        Soffitto.BackColor = &H8000000F
        Muri.BackColor = &H8000000F
        Pavimento.BackColor = &H8000000F
    End Select
End Sub




Private Sub Timer1_Timer()
    'Imposto alla textbox contenente il valore di luminosità il valore dello scroller
    'orizzontale Luminosità_griglia
    Valore_luminosità.Text = Str(Luminosità_griglia.Value)
    'Se la variabile Cont ha assunto il valore 11,ciò vuol dire che la mappa è stata ingrandita 11 volte
    'e quindi si dovrà disabilitare il bottone Zoom+ per impedire all'utente di effettuare un ulteriore
    'ingrandimento
    If Cont = 11 Then
        ZoomP.Enabled = False
    'Per qualsiasi altro valore di Cont il bottone Zoom+ verrà abilitato
    Else
        ZoomP.Enabled = True
    End If
     'Se la variabile Cont ha assunto il valore -11,ciò vuol dire che la mappa è stata rimpicciolita 11 volte
    'e quindi si dovrà disabilitare il bottone Zoom- per impedire all'utente di effettuare un ulteriore
    'rimpicciolimento
    If Cont = -11 Then
        ZoomM.Enabled = False
    'Per qualsiasi altro valore di Cont il bottone Zoom- verrà abilitato
    Else
        ZoomM.Enabled = True
    End If
    'Se il bottone di opzione DisattivaFondale è selezionato,allora disabiliterò il pulsante che mi permette di
    'selezionare un'immagine da applicare come fondale dell'editor
    If DisattivaFondale.Value = True Then
        'Disattivo il pulsante CambiaImmagineDiSfondo
        CambiaImmagineDiSfondo.Enabled = False
    'Diversamente,se lo stesso bottone è deselezionato,allora questa volta attiverò sempre lo stesso pulsante
    Else
    'Attivo il pulsante CambiaImmagineDiSfondo
        CambiaImmagineDiSfondo.Enabled = True
    End If
End Sub

Private Sub TipoPavimento_Click()
    'Se l'elemento selezionato dalla ComboBox è un soffitto,allora
    'questo divernterà un pavimento
    If RTrim(SoP(IndiceLista2).Tipo) = "Soffitto" Then
        'Cambio del tipo all'interno del record
        SoP(IndiceLista2).Tipo = "Pavimento"
        'Decrementiamo la variabile che tiene conto del numero di soffitti...
        Max4 = Max4 - 1
        '...e incrementiamo quella che tiene conto dei pavimenti
        Max3 = Max3 + 1
    End If
End Sub

Private Sub TipoSoffitto_Click()
    'Se l'elemento selezionato dalla ComboBox è un pavimento,allora
    'questo divernterà un soffitto
    If RTrim(SoP(IndiceLista2).Tipo) = "Pavimento" Then
        'Cambio del tipo all'interno del record
        SoP(IndiceLista2).Tipo = "Soffitto"
        'Decrementiamo la variabile che tiene conto del numero di pavimenti...
        Max3 = Max3 - 1
        '...e incrementiamo quella che tiene conto dei soffitti
        Max4 = Max4 + 1
    End If
End Sub

Private Sub ValoreScale_click()
    'Imposto una struttura di controllo per verificare quale scale impostato è stato
    'scelto,in modo da capire quale valore assegnare alla variabile VScale
    Select Case ValoreScale.ListIndex
    Case Is = 0
        VScale = 1
    Case Is = 1
        VScale = 10
    Case Is = 2
        VScale = 100
    Case Is = 3
        VScale = 1000
    Case Is = 4
        VScale = 10000
    End Select
End Sub

Private Sub ZoomM_Click()
    Dim Z As Boolean
    'Richiamo la funzione Zoom passondogli come valore 0.9,in modo che tutte le righe presenti all'interno
    'della mappa attuale vengono rimpicciolite appunto di 0.9 rispetto alla loro dimensione originale
    Zoom 0.9, False
    'La mappa è stata rimpicciolita ulteriormente,quindi decremento la variabile Cont
    Cont = Cont - 1
End Sub

Private Sub ZoomP_Click()
    Dim Z As Boolean
    'Anche qui richiamo la funzione Zoom passandogli come valore 1.1,in modo che tutte le righe presenti
    'all'interno della mappa attuale vengano ingrandite appunto di 1.1 rispetto alla loro dimensione originale
    Zoom 1.1, True
    'La mappa è stata ingrandita ulteriormente,quindi incremento la variabile Cont
    Cont = Cont + 1
End Sub

Sub Zoom(ValoreZoom As Single, TipoDiZoom As Boolean)
    'Dichiaro un indice che mi servirà per identificare singolarmente ogni singola riga della mappa attuale
    Dim K As Integer
    Dim Cont As Integer
    If TipoDiZoom = True Then
        Molt = Molt * 1.1
        'Riferisco al programma che lo zoom è cambiato e quindi si dovrà assegnare il nuovo valore alla
        'variabileVCambiamentiGriglia per fare in modo che i quadrati verranno o rimpiccioliti
        'o ingranditi
        VCambiamentiGriglia = VCambiamentiGriglia + 5
    Else
        Molt = Molt * 0.9
         'Riferisco al programma che lo zoom è cambiato e quindi si dovrà assegnare il nuovo valore alla
        'variabileVCambiamentiGriglia per fare in modo che i quadrati verranno o rimpiccioliti
        'o ingranditi
        VCambiamentiGriglia = VCambiamentiGriglia - 5
    End If
    'Moltiplico tutti i valori dei muri presenti all'interno della mappa attuale per il valore passato
    'alla funzione stessa in modo che tutte le righe vengono reimpostate sencondo il loro nuovo valore
    'Usando questo metodo però si riscontravano problemi nel rilevare l'allineamento delle linne guida
    'con i valori delle riga presenti all'interno della mappa attuale.
    'Per risolverlo ho dovuto aggiungere la funzione Fix per troncare la parte decimale del valore
    For K = 0 To Max
        Riga(K).X1 = Fix(Riga(K).X1 * ValoreZoom)
        Riga(K).X2 = Fix(Riga(K).X2 * ValoreZoom)
        Riga(K).Y1 = Fix(Riga(K).Y1 * ValoreZoom)
        Riga(K).Y2 = Fix(Riga(K).Y2 * ValoreZoom)
    Next
    'Ora Svolgo le stessa operazione effettuata con i muri, per tutti i valori dei pavimenti e soffitti
    'presenti all'interno della mappa attuale
    For K = 0 To Max2
        SoP(K).CR.X1 = Fix(SoP(K).CR.X1 * ValoreZoom)
        SoP(K).CR.X2 = Fix(SoP(K).CR.X2 * ValoreZoom)
        SoP(K).X3 = Fix(SoP(K).X3 * ValoreZoom)
        SoP(K).X4 = Fix(SoP(K).X4 * ValoreZoom)
        SoP(K).CR.Y1 = Fix(SoP(K).CR.Y1 * ValoreZoom)
        SoP(K).CR.Y2 = Fix(SoP(K).CR.Y2 * ValoreZoom)
        SoP(K).Y3 = Fix(SoP(K).Y3 * ValoreZoom)
        SoP(K).Y4 = Fix(SoP(K).Y4 * ValoreZoom)
    Next
    'Ora Svolgo le stessa operazione effettuata per i muri,i pavimenti e soffitti, per tutti gli oggetti
    'presenti all'interno della mappa attuale
    For K = 0 To IOg
        With Oggetto(K)
            .Setta_Posizione "X", Fix(.Ricava_Posizione.X) * ValoreZoom
            .Setta_Posizione "Z", Fix(.Ricava_Posizione.Z) * ValoreZoom
        End With
    Next
    'Aggiorno le coordinate di telecamera
    With PosizioneTelecamera
        .X = Fix(.X * ValoreZoom)
        .Z = Fix(.Z * ValoreZoom)
    End With
End Sub

Function Trova_Oggetto(Chiave As String) As Integer
    'Dichiaro una variabile che servirà per capire se l'oggetto ricercato esiste realmente,
    'oppure non è presente all'interno dell'array Oggetto
    Dim Trovato As Boolean
    'Dichiaro una variabile di appoggio che mi faciliterà l'operazione di ricerca dell'oggetto
    'desiderato
    Dim NOggetto As Integer
    'Avvio una scansione di tutti gli oggetti caricati all'interno della mappa,al fine
    'di ricercare quello avente la propria variabile Key uguale al parametro passato alla
    'funzione stessa
    For NOggetto = 0 To IOg
        'Se l'oggetto ricerecato è stato trovato,allora...
        If Oggetto(NOggetto).Key = Chiave Then
            'La funzione assumerà il valore dell'indice dell'oggetto trovato dall'array contenente
            'tutti gli oggetti caricati
            Trova_Oggetto = NOggetto
            'A questo punto si può uscire dalla funzione
            Exit Function
        End If
    'Si passa ad analizzare l'oggetto successivo
    Next
End Function

Friend Function Copia_Oggetto(Oggetto1 As ClsOggetti, Oggetto2 As ClsOggetti)
    'Ricopio la chiave del secondo oggetto all'interno del primo oggetto passato alla funzione stessa
    Oggetto1.Key = Oggetto2.Key
    'Ricopio il percorso del secondo oggetto all'interno del primo oggetto passato alla funzione stessa
    Oggetto1.Percorso = Oggetto2.Percorso
    'Ricopio la texture del secondo oggetto all'interno del primo oggetto
    Oggetto1.Texture = Oggetto2.Texture
    'Ricopio la descrizione del secondo oggetto all'interno del primo oggetto passato alla funzione stessa
    Oggetto1.Decrizione = Oggetto2.Decrizione
    'Ora incomincio a ricopiare tutte le variabili del secondo oggetto addette al suo posizionamento
    'all'interno del primo oggetto passato alla funzione stessa
    Oggetto1.Setta_Posizione "X", Oggetto2.Ricava_Posizione.X
    Oggetto1.Setta_Posizione "Y", Oggetto2.Ricava_Posizione.Y
    Oggetto1.Setta_Posizione "Z", Oggetto2.Ricava_Posizione.Z
    'Ricopio anche tutte le variabili del secondo oggetto addette al ridimensionamento all'interno
    'del primo oggetto passato alla funzione stessa
    Oggetto1.Setta_Dimensione "X", Oggetto2.Ricava_Dimensione.X
    Oggetto1.Setta_Dimensione "Y", Oggetto2.Ricava_Dimensione.Y
    Oggetto1.Setta_Dimensione "Z", Oggetto2.Ricava_Dimensione.Z
    'Infine ricopio anche tutte le variabili del secondo oggetto addette alla rotazione sempre all'interno del primo
    'oggetto passato alla funzione stessa
    Oggetto1.Setta_Rotazione "X", Oggetto2.Ricava_Rotazione.X
    Oggetto1.Setta_Rotazione "Y", Oggetto2.Ricava_Rotazione.Y
    Oggetto1.Setta_Rotazione "Z", Oggetto2.Ricava_Rotazione.Z
End Function

Function Verifica_Doppio(Nome As String) As Boolean
    'Dichiaro un nodo generale che mi servirà per verificare tutti quelli all'interno del
    'controllo ElencoGruppiOggetti
    Dim Nodo As Node
    'Avvio una ricerca per ogni nodo contenuto all'interno del controllo ElencoGruppiOggetti
    For Each Nodo In ElencoGruppiOggetti.Nodes
        'Se la chiave del nodo esaminato è uguale al valore alfanumerico contenuto nel
        'parametro Nome passato alla funzione stessa, cioè, se vi è un altro nodo avente
        'questa chiave,allora
        If Nodo.Key = Nome Then
            'La funzione restituirà valore booleano uguale a True
            Verifica_Doppio = True
            'Si può uscire dalla funzione
            Exit Function
        End If
    'Si esamina il nodo successivo
    Next
    'Se la funzione è arrivato fino a questo punto, allora questa restituirà il valore booleano False
    Verifica_Doppio = False
End Function

Function Nuovo_Nome(Nome As String) As String
    'Dichiaro la variabile NNome, la quale conterrà il nuovo Nome generato dalla
    'funzione stessa
    Dim NNome As String
    'La variabile Doppio mi servirà per capire se il nuovo nome generato è nuovamente
    'doppio all'interno del controllo ElencoGruppiOggetti
    Dim Doppio As Boolean
    'Assegno alla variabile NNome il valore del parametro Nome passato alla funzione stessa
    NNome = Nome
    'Avvio un ciclo per generare un nuovo nome finchè questo non sarà più doppio
    Do
        'Genero il nuovo Nome
        NNome = NNome + "."
        'Assegno il valore della variabile Doppio tramite la funzione Verifica_Doppio,passandogli
        'come valore il nuovo nome generato
        Doppio = Verifica_Doppio(NNome)
    'Finchè il nome generato non è doppio all'interno del controllo ElencoGruppiOggetti
    Loop Until Doppio = False
    'La funzione restituisce il nuovo nome generato
    Nuovo_Nome = NNome
End Function
