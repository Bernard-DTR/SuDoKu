Option Strict On
Option Explicit On

Module M02_Menu_Management
  '-------------------------------------------------------------------------------
  ' Gestion des menus contextuels pour les modes Nrm-Sas
  '                                              Edi
  '-------------------------------------------------------------------------------

  Sub Mnu_EDI(Action As String, sender As Object, e As EventArgs)
    'Alternance de Swt_ModeEdition permettant de Lancer puis d'Arrêter le Mode Edition
    'Arrêt du Mode de clignotement qui "fatigue visuellement"
    'Suppression de Supprimer Edi_Timer et Edi_Clignotant
    'Résultat identique
    'Le menu apparait, la boite de dialogue apparait et disparait, il faut enfoncer Escape
    'Les autres options fonctionnent correctement

    'Le menu Edition permet les actions suivantes:
    '  Placer  une valeur dans une cellule
    '  Gérer les candidats dans une cellule
    '  Définir une cellule comme valeur initiale
    '  Définir une cellule comme valeur normale
    '  Effacer une valeur dans une cellule
    'Le mode EDI n'utilise pas les fonctions Cell_Val|Cdd_Insert|Suppr 
    Dim XPos As Integer
    Dim YPos As Integer
    Dim Titre As String
    Dim Texte As String
    Dim Cellule As Integer = Pbl_Cell_Select
    Dim Candidat As String = Cnddts_Blancs
    ' Val, Cdd, Eff, Ini, Nrm, Maj
    Select Case Action
      Case "Val"
        'Ajouter une valeur à UNE cellule
        XPos = Frm_SDK.Left + WH * (U_Col(Cellule) + 1)
        YPos = Frm_SDK.Top + WH * (U_Row(Cellule) + 1)
        Titre = "Saisir une valeur"
        Texte = "Valeur souhaitée en " & U_Coord(Cellule) & ":" & vbCrLf
        Texte &= "La valeur est comprise entre 1 et 9."
        Dim DftValue As String = U(Cellule, 2)
        Dim Valeur As String
        Valeur = InputBox(Texte, Titre, DftValue, XPos, YPos)
        If Valeur Is "" Then Exit Sub 'Cancel enfoncé
        If (Valeur.Length = 1 And Valeur >= "1" And Valeur <= "9") Then
          U(Cellule, 1) = " "
          U(Cellule, 2) = Valeur
          U(Cellule, 3) = Cnddts_Blancs
        End If
      Case "Cdd"
        'Placer des candidats dans une cellule
        XPos = Frm_SDK.Left + WH * (U_Col(Cellule) + 1)
        YPos = Frm_SDK.Top + WH * (U_Row(Cellule) + 1)
        Titre = "Saisir un ou plusieurs candidats"
        Texte = "Candidat(s) souhaité(s) en " & U_Coord(Cellule) & ":" & vbCrLf
        Texte &= "La chaîne de caractères ne doit pas dépasser 9."
        Texte &= "Les candidats peuvent être présentés dans le désordre."
        Dim DftValue As String = U(Cellule, 3)
        Dim Valeur As String = InputBox(Texte, Titre, DftValue, XPos, YPos)
        If Valeur Is "" Then Exit Sub 'Cancel enfoncé
        Dim l As Integer = Valeur.Length
        If l > 9 Then l = 9
        For i As Integer = 1 To l
          Dim Cdd As String = Valeur.Substring(i - 1, 1)
          If (Cdd >= "1" And Cdd <= "9") Then Mid$(Candidat, CInt(Cdd), 1) = Cdd
        Next i
        U(Cellule, 1) = " "
        U(Cellule, 2) = " "
        U(Cellule, 3) = Candidat
      Case "Eff"
        'Effacer la valeur dans une cellule
        U(Cellule, 1) = " "
        U(Cellule, 2) = " "
        U(Cellule, 3) = Cnddts
      Case "Ini"
        'Définir une cellule en Valeur Initiale
        U(Cellule, 1) = U(Cellule, 2)
      Case "Nrm"
        'Définir une cellule qui est en VI en valeur normale
        U(Cellule, 1) = " "
      Case Else
        Jrn_Add("ERR_00000", {sender.ToString() & " Action : " & Action & " en: " & U_Coord(Cellule)}, "Erreur")
    End Select

    Event_OnPaint = "Global"
    Frm_SDK.Invalidate()
  End Sub

  Sub Mnu_Mngt_Conditions_Générales(Cellule As Integer)
    ' Afin de faciliter les traitements suivants, mise en place de Plcy_* spécifiques
    Plcy_Nrm = False
    Plcy_Sas = False
    If Plcy_Gnrl = "Nrm" Then Plcy_Nrm = True
    If Plcy_Gnrl = "Sas" Then Plcy_Sas = True
    Plcy_Typ_I = False
    Plcy_Typ_R = False
    Plcy_Typ_V_sans_Cdd = False
    Plcy_Typ_V_avec_Cdd = False
    If (U(Cellule, 1) <> " ") Then Plcy_Typ_I = True
    If (U(Cellule, 1) = " " And U(Cellule, 2) <> " ") Then Plcy_Typ_R = True

    If (U(Cellule, 1) = " " And U(Cellule, 2) = " ") Then
      If Plcy_Nrm = True Then
        If (Plcy_Strg = "   ") _
        Or (Stg_Get(Plcy_Strg).Type = "I" And Not Plcy_AideGraphique) _
        Or (Stg_Get(Plcy_Strg).Type = "E" And Not Plcy_AideGraphique) _
        Or (Mid$(Plcy_Strg, 1, 1) = "F" And Not Plcy_AideGraphique) Then
          Plcy_Typ_V_sans_Cdd = True
        End If

        'Le menu contextuel est replacé pour les filtres
        If (Plcy_Strg = "Cdd") _
        Or (Stg_Get(Plcy_Strg).Type = "I" And Plcy_AideGraphique) _
        Or (Stg_Get(Plcy_Strg).Type = "E" And Plcy_AideGraphique) _
        Or (Mid$(Plcy_Strg, 1, 1) = "F" And Plcy_AideGraphique) Then
          Plcy_Typ_V_avec_Cdd = True
        End If

      End If
      If Plcy_Sas Then
        If (U(Cellule, 3) = Cnddts_Blancs) Then Plcy_Typ_V_sans_Cdd = True
        If (U(Cellule, 3) <> Cnddts_Blancs) Then Plcy_Typ_V_avec_Cdd = True
      End If
    End If
  End Sub

  Sub Mnu_Mngt_Barre_de_Menu()
    'Menu Effacer de la barre de Menu
    Frm_SDK.Mnu02_Effacer.Enabled = False
    If Plcy_Nrm And Plcy_Typ_R Then
      Frm_SDK.Mnu02_Effacer.Enabled = True
    End If
  End Sub

  Sub Mnu_Mngt_Barre_Outils_Filtres()
    ' Seuls les Boutons Filtres de la Barre d'Outils sont traités
    Dim n() As Integer = Wh_Nb_Cell(U).Val_Nb
    Dim Btn As System.Windows.Forms.ToolStripItem
    'Les ToolStripSeparator sont correctement placés
    'Les Btn Filtre n'ont pas de texte, uniquement une Image
    If Not Plcy_Nrm Then Exit Sub
    For Each Btn In Frm_SDK.BarreOutils.Items
      If Btn.GetType().ToString() <> "System.Windows.Forms.ToolStripButton" Then Continue For
      Btn.Visible = True
      'Remise à l'état standard, ne concerne que les filtres 1 à 9
      Dim Flt As String = Mid$(Btn.Name, 4, 1)
      'les boutons des filtre comporte une valeur de 1 à 9 pour une police non fantaisiste
      '                       comporte une image contenue dans Sqr_Fantasy(Flt) pour les polices fantaisistes
      If Not (Flt >= "1" And Flt <= "9") Then Continue For
      Btn.Enabled = True

      'Uniquement les boutons de filtre
      Select Case Plcy_Fantasy
        Case True
          Btn.DisplayStyle = ToolStripItemDisplayStyle.Image
          Btn.Image = Sqr_Fantasy(CInt(Flt))
          'Propriétés identiques aux autres boutons de Barre_Outils
          Btn.ImageScaling = ToolStripItemImageScaling.SizeToFit
          Btn.ImageAlign = ContentAlignment.MiddleCenter
          Btn.ToolTipText = "Filtrer les valeurs ou candidats " & Flt
          For i As Integer = 1 To 9
            If Flt = CStr(i) And n(i) = 9 Then
              Btn.Image = Sqr_Fantasy(0)     'Donc Subst_Police(Cstr(0)) retourne vide et le bouton n'affiche rien
              Btn.Enabled = False
              Btn.ToolTipText = "Les valeurs " & Flt & " sont toutes présentes."
            End If
          Next i
        Case False
          'Utilisation idifférente de "Arial" ou de "Segoe UI"
          Btn.DisplayStyle = ToolStripItemDisplayStyle.Text
          Btn.Text = Flt
          Btn.Font = New Font("Segoe UI", 8, FontStyle.Bold)
          Btn.ToolTipText = "Filtrer les valeurs ou candidats " & Flt
          For i As Integer = 1 To 9
            If Flt = CStr(i) And n(i) = 9 Then
              Btn.Font = New Font("Segoe UI", 8, FontStyle.Italic)
              Btn.Enabled = False
              Btn.ToolTipText = "Les valeurs " & Flt & " sont toutes présentes."
            End If
          Next i
      End Select
    Next Btn
  End Sub

  Sub Mnu_Mngt(Cellule As Integer)
    Dim Ligne As ToolStripItem
    Dim Plcy_Mnu_Groupe As Boolean
    Dim Lig As String = ""
    Dim Opt As String = ""

    Mnu_Mngt_Conditions_Générales(Cellule)
    Mnu_Mngt_Barre_de_Menu()
    Mnu_Mngt_Barre_Outils_Filtres()

    Plcy_Mnu_Groupe = False
    For Each Ligne In Frm_SDK.Mnu_Cel.Items
      Try
        'Le menu est systématiquement effacé, rendu invisible
        Ligne.Visible = False
        Lig = Ligne.Name.Substring(8, 8)
        Plcy_Mnu_Item = False : Plcy_Mnu_Sep = False
        Select Case Ligne.GetType().ToString()
          Case "System.Windows.Forms.ToolStripMenuItem" : Plcy_Mnu_Item = True
          Case "System.Windows.Forms.ToolStripSeparator" : Plcy_Mnu_Sep = True
        End Select

        Select Case Lig

          Case "Val_Ins_" 'Insérer la valeur 1 à 9
            If Plcy_Typ_V_sans_Cdd Then
              If Plcy_Mnu_Sep And Plcy_Mnu_Groupe Then Ligne.Visible = True
              If Plcy_Mnu_Item Then
                Opt = Ligne.Name.Substring(16, 1)
                If Plcy_Fantasy Then
                  Ligne.Text = "Insérer la valeur " & Mnu_Ctxt_Fantaisy(Opt)
                Else
                  Ligne.Text = "Insérer la valeur " & Opt
                End If
                Ligne.Visible = True : Plcy_Mnu_Groupe = True
                If Plcy_Fantasy Then Ligne.Image = Sqr_Fantasy(CInt(Opt))
                If Not Plcy_Fantasy Then Ligne.Image = Nothing
              End If
            End If
            'Le candidat CdU ou les candidats CdO, Flt sont affichés
            If Plcy_Typ_V_avec_Cdd Then
              Dim Candidats As String = U(Cellule, 3)
              Opt = Ligne.Name.Substring(16, 1)
              If Plcy_Mnu_Sep And Plcy_Mnu_Groupe Then Ligne.Visible = True
              If Plcy_Mnu_Item Then
                If Opt = Candidats.Substring(CInt(Opt) - 1, 1) Then
                  Ligne.BackColor = Control.DefaultBackColor
                  'Menu contextuel présenté lors d'une stratégie CdU ou CdO en Aide Graphique
                  If Plcy_Gnrl = "Nrm" AndAlso
                      Stg_Get(Plcy_Strg).Type = "I" AndAlso
                     Plcy_AideGraphique AndAlso
                     U_Strg_Val_Ins(Cellule) = Opt Then
                    Ligne.BackColor = Color_Cdd_Insérer
                  End If
                  If Plcy_Fantasy Then
                    Ligne.Text = "Insérer la Valeur " & Mnu_Ctxt_Fantaisy(Opt)
                  Else
                    Ligne.Text = "Insérer la Valeur " & Opt
                  End If
                  Ligne.Visible = True : Plcy_Mnu_Groupe = True
                Else
                  Ligne.Visible = False
                End If
              End If
            End If
          Case "Cdd_Exc_" 'Exclure le candidat 1 à 9
            If Plcy_Typ_V_avec_Cdd Then
              Dim Candidats As String = U(Cellule, 3)
              If (Plcy_Sas) _
              Or (Plcy_Nrm And Wh_Cell_Nb_Candidats(U, Cellule) > 1) Then
                Opt = Ligne.Name.Substring(16, 1)
                If Plcy_Mnu_Sep And Plcy_Mnu_Groupe Then Ligne.Visible = True
                'les stratégies proposent d'exclure un candidat.
                If Plcy_Mnu_Item Then
                  If Opt = Candidats.Substring(CInt(Opt) - 1, 1) Then
                    If Plcy_Fantasy Then Ligne.Image = Sqr_Fantasy(CInt(Opt))
                    If Not Plcy_Fantasy Then Ligne.Image = Nothing
                    Ligne.BackColor = Control.DefaultBackColor
                    If Plcy_Gnrl = "Nrm" AndAlso
                      Stg_Get(Plcy_Strg).Type = "E" AndAlso
                     Plcy_AideGraphique AndAlso
                     U_Strg_Cdd_Exc(Cellule).Contains(Opt) Then
                      'Le menu est colorisé rouge pour toutes les stratégies
                      Ligne.BackColor = Color_Cdd_Exclure
                    End If
                    If Plcy_Fantasy Then
                      Ligne.Text = "Exclure le candidat " & Mnu_Ctxt_Fantaisy(Opt)
                    Else
                      Ligne.Text = "Exclure le candidat " & Opt
                    End If
                    Ligne.Visible = True : Plcy_Mnu_Groupe = True
                  Else
                    Ligne.Visible = False
                  End If
                End If
              End If
            End If

          Case "Cdd_Exd_", "Cdd_Exe_" 'Exclure les candidats..., tous les candidats
            'If Plcy_Sas And Plcy_Typ_V_avec_Cdd Then
            If Plcy_Typ_V_avec_Cdd Then
              If Wh_Cell_Nb_Candidats(U, Cellule) > 1 Then
                Ligne.Visible = True
                If Plcy_Fantasy Then Ligne.Visible = False
              End If
            End If

          Case "Cdd_Ins_" 'Insérer les Candidats ....
            If Plcy_Sas And Plcy_Typ_V_sans_Cdd Then
              If Plcy_Mnu_Sep And Plcy_Mnu_Groupe Then Ligne.Visible = True
              If Plcy_Mnu_Item Then
                Opt = Ligne.Name.Substring(16, 1)
                If Plcy_Fantasy Then
                  Ligne.Text = "Insérer le candidat " & Mnu_Ctxt_Fantaisy(Opt)
                Else
                  Ligne.Text = "Insérer le candidat " & Opt
                End If
                Ligne.Visible = True : Plcy_Mnu_Groupe = True
                If Plcy_Fantasy Then Ligne.Image = Sqr_Fantasy(CInt(Opt))
                If Not Plcy_Fantasy Then Ligne.Image = Nothing
              End If
            End If
            If Plcy_Sas And Plcy_Typ_V_avec_Cdd Then
              Dim Candidats As String = U(Cellule, 3)
              Opt = Ligne.Name.Substring(16, 1)
              If Plcy_Mnu_Sep And Plcy_Mnu_Groupe Then Ligne.Visible = True
              If Plcy_Mnu_Item Then
                If Opt = Candidats.Substring(CInt(Opt) - 1, 1) Then
                  Ligne.Visible = False
                Else
                  Ligne.Visible = True : Plcy_Mnu_Groupe = True
                End If
              End If
            End If

          Case "Val_Eff_"
            If Plcy_Typ_R Then
              If Plcy_Mnu_Sep And Plcy_Mnu_Groupe Then Ligne.Visible = True
              If Plcy_Mnu_Item Then
                Ligne.Visible = True : Plcy_Mnu_Groupe = True
              End If
            End If

          Case Else
            Jrn_Add(, {"A " & Procédure_Name_Get() & " " & Ligne.Name}, "Erreur")
        End Select

      Catch ex As Exception
        Jrn_Add("ERR_00000", {"B " & Procédure_Name_Get()})
        Jrn_Add("ERR_00000", {Ligne.Name}, "Erreur")
        Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
        Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
        Jrn_Add("ERR_00000", {"Item    : " & Ligne.ToString()}, "Erreur")
        Jrn_Add("ERR_00000", {"Lig      : " & Lig}, "Erreur")
        Jrn_Add("ERR_00000", {"Opt      : " & Opt}, "Erreur")
        Jrn_Add("ERR_00000", {"Cellule  : " & U_Coord(Cellule) & " V : " & U(Cellule, 2) & " Candidats :" & U(Cellule, 3) & "."}, "Erreur")
      End Try
    Next Ligne
  End Sub

  Public Function Mnu_Ctxt_Fantaisy(Valeur As String) As String
    Dim Txt As String = ""
    Select Case Valeur
      Case "1" : Txt = "du haut à gauche"
      Case "2" : Txt = "du haut au centre"
      Case "3" : Txt = "du haut à droite"
      Case "4" : Txt = "du centre à gauche"
      Case "5" : Txt = "du centre au centre"
      Case "6" : Txt = "du centre à droite"
      Case "7" : Txt = "du bas à gauche"
      Case "8" : Txt = "du bas au centre"
      Case "9" : Txt = "du bas à droite"
    End Select
    Return Txt
  End Function
End Module