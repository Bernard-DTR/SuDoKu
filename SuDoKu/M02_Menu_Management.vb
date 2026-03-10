Option Strict On
Option Explicit On

Module M02_Menu_Management

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
    Application.DoEvents()
  End Sub

  Sub Mnu_Mngt_Barre_Outils_Filtres()
    ' Seuls les Boutons Filtres de la Barre d'Outils sont traités
    Dim n() As Integer = Wh_Nb_Cell(U).Val_Nb
    Dim Btn As System.Windows.Forms.ToolStripItem
    'Les ToolStripSeparator sont correctement placés
    'Les Btn Filtre n'ont pas de texte, uniquement une Image
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
          'Utilisation indifférente de "Arial" ou de "Segoe UI"
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
    Dim Mnu_Item As Boolean, Mnu_Sep As Boolean, Mnu_Grp As Boolean
    Dim Ligne As ToolStripItem
    Dim Candidats As String
    Dim Opt As String = ""
    Dim Opt_int As Integer

    Mnu_Mngt_Barre_Outils_Filtres()

    Mnu_Grp = False
    For Each Ligne In Frm_SDK.Mnu_Cel.Items
      Try
        'Le menu est systématiquement effacé, rendu invisible
        Ligne.Visible = False
        Mnu_Item = False : Mnu_Sep = False
        Select Case Ligne.GetType().ToString()
          Case "System.Windows.Forms.ToolStripMenuItem" : Mnu_Item = True
          Case "System.Windows.Forms.ToolStripSeparator" : Mnu_Sep = True
        End Select
        Select Case Ligne.Name.Substring(0, 16)

          Case "Mnu_Cel_Val_Ins_" 'Insérer la valeur 1 à 9
            If Mnu_Item AndAlso Stg_Get(Plcy_Strg).Family = 3 AndAlso Stg_Get(Plcy_Strg).Type = "I" AndAlso U(Cellule, 2) = " " Then
              Candidats = U(Cellule, 3)
              Opt = Ligne.Name(16)
              Opt_int = CInt(Opt) - 1
              If Opt = Candidats(Opt_int) AndAlso U_Strg_Val_Ins(Cellule) = Opt Then
                Ligne.Visible = True
                Ligne.BackColor = Color_Cdd_Insérer
              End If
            End If

          Case "Mnu_Cel_Cdd_Exc_" 'Exclure le candidat 1 à 9
            ' TODO Traitement insuffisant pour la stratégie Unq
            If Mnu_Item AndAlso Stg_Get(Plcy_Strg).Family = 3 AndAlso Stg_Get(Plcy_Strg).Type = "E" AndAlso U(Cellule, 2) = " " Then
              Candidats = U(Cellule, 3)
              Opt = Ligne.Name(16)
              Opt_int = CInt(Opt) - 1
              If Opt = Candidats(Opt_int) AndAlso U_Strg_Cdd_Exc(Cellule) = Opt Then
                Ligne.Visible = True
                Ligne.BackColor = Color_Cdd_Exclure
              End If
            End If

          Case "Mnu_Cel_Cdd_Ins_" 'Insérer les Candidats ....
            'If U(Cellule, 2) = " " Then

            '  If Stg_Get(Plcy_Strg).Family = 3 Then
            '    Dim Candidats_Excl As String = U_CddExc(Cellule)
            '    If Wh_Cell_Nb_Candidats(U, Cellule) > 1 Then
            '      Opt = Ligne.Name.Substring(16, 1)
            '      If Mnu_Sep And Mnu_Grp Then Ligne.Visible = True
            '      'les stratégies proposent d'exclure un candidat.
            '      If Mnu_Item Then
            '        If Opt = Candidats_Excl.Substring(CInt(Opt) - 1, 1) Then
            '          Ligne.BackColor = Control.DefaultBackColor
            '          Ligne.Visible = True : Mnu_Grp = True
            '        Else
            '          Ligne.Visible = False
            '        End If
            '      End If
            '    End If
            '  End If
            'End If

          Case "Mnu_Cel_Val_Eff_"
            ' Option Effacer la valeur dans une cellule remplie
            If (U(Cellule, 1) = " " AndAlso U(Cellule, 2) <> " ") Then
              If Mnu_Sep And Mnu_Grp Then Ligne.Visible = True
              If Mnu_Item Then Ligne.Visible = True : Mnu_Grp = True
            End If

          Case Else
            Jrn_Add(, {"Case:  " & Proc_Name_Get() & " " & Ligne.Name}, "Erreur")
        End Select

      Catch ex As Exception
        Jrn_Add("ERR_00000", {"Exception: " & Proc_Name_Get()})
        Jrn_Add("ERR_00000", {Ligne.Name}, "Erreur")
        Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
        Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
      End Try
    Next Ligne
  End Sub

End Module