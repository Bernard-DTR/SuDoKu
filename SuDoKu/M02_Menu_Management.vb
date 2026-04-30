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

    Frm_SDK.Invalidate()
  End Sub

  Sub Mnu_Mngt_Barre_Outils_Filtres_Enabled()
    For Each Btn As ToolStripItem In Frm_SDK.BarreOutils.Items
      ' Seuls les Boutons Filtres de la Barre d'Outils sont traités
      If Not TypeOf Btn Is ToolStripButton Then Continue For
      Dim Flt As String = Mid$(Btn.Name, 4, 1)
      If Not (Flt >= "1" AndAlso Flt <= "9") Then Continue For
      Btn.Visible = True
      Btn.Enabled = True
      Btn.ToolTipText = $"Filtrer les valeurs ou candidats {Flt}"
      If U_nb(CInt(Flt)) = 9 Then
        'Jrn_Add_Yellow(Proc_Name_Get() & " " & Btn.Name)
        Btn.Enabled = False
        Btn.ToolTipText = $"Les valeurs {Flt} sont toutes présentes."
      End If
    Next Btn
  End Sub

  Sub Mnu_Mngt_Barre_Outils_Filtres()
    ' Seuls les Boutons Filtres de la Barre d'Outils sont traités
    Dim n() As Integer = Wh_Nb_Cell(U).Val_Nb
    'Dim Btn As System.Windows.Forms.ToolStripItem
    'Les ToolStripSeparator sont correctement placés
    'Les Btn Filtre n'ont pas de texte, uniquement une Image
    For Each Btn As System.Windows.Forms.ToolStripItem In Frm_SDK.BarreOutils.Items
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
    Dim Mnu_Item As Boolean, Mnu_Sep As Boolean
    Dim Ligne As ToolStripItem
    Dim Candidats As String
    Dim Opt As String

    Mnu_Mngt_Barre_Outils_Filtres()

    For Each Ligne In Frm_SDK.Mnu_Cel.Items
      Try
        'Le menu est systématiquement effacé, rendu invisible
        Ligne.Visible = False
        Mnu_Item = False : Mnu_Sep = False
        ' la ligne est soit un menu_item, soit un separateur
        Select Case Ligne.GetType().ToString()
          Case "System.Windows.Forms.ToolStripMenuItem" : Mnu_Item = True
          Case "System.Windows.Forms.ToolStripSeparator" : Mnu_Sep = True
        End Select

        Select Case Ligne.Name.Substring(0, 16)

          Case "Mnu_Cel_Val_Ins_" 'Insérer la valeur 1 à 9  en JAUNE
            If Mnu_Item AndAlso Stg_Get(Plcy_Strg).Family = 3 AndAlso Stg_Get(Plcy_Strg).Type = "I" AndAlso U(Cellule, 2) = " " Then
              Candidats = U(Cellule, 3)
              Opt = Ligne.Name(16)
              If Opt = Candidats(CInt(Opt) - 1) AndAlso U_Strg_Val_Ins(Cellule) = Opt Then
                Ligne.Visible = True : Ligne.BackColor = Color_Cdd_Insérer
              End If
            End If

          Case "Mnu_Cel_Cdd_Exc_" 'Exclure le candidat 1 à 9
            '                     'Exclure le candidat 1 à 9 en Normal
            If Mnu_Item AndAlso Stg_Get(Plcy_Strg).Family = 1 AndAlso U(Cellule, 2) = " " Then
              Candidats = U(Cellule, 3)
              Opt = Ligne.Name(16)
              If Opt = Candidats(CInt(Opt) - 1) Then
                Ligne.Visible = True : Ligne.BackColor = Control.DefaultBackColor
              End If
            End If
            '                     'Exclure le candidat 1 à 9 en ROUGE
            If Mnu_Item AndAlso Stg_Get(Plcy_Strg).Family = 3 AndAlso Stg_Get(Plcy_Strg).Type = "E" AndAlso U(Cellule, 2) = " " Then
              Candidats = U(Cellule, 3)
              Opt = Ligne.Name(16)
              If Opt = Candidats(CInt(Opt) - 1) AndAlso U_Strg_Cdd_Exc(Cellule) = Opt Then
                Ligne.Visible = True : Ligne.BackColor = Color_Cdd_Exclure
              End If
            End If
            If Mnu_Item AndAlso Stg_Get(Plcy_Strg).Family = 4 AndAlso Stg_Get(Plcy_Strg).Type = "E" AndAlso U(Cellule, 2) = " " Then
              If GRslt.CelExcl.Count > 0 Then
                For Each XCel As GCel_Excl_Cls In GRslt.CelExcl
                  If XCel.Cel = Cellule Then
                    Opt = Ligne.Name(16)
                    If Opt = XCel.Cdd And U(Cellule, 3).Contains(XCel.Cdd) Then
                      Ligne.Visible = True : Ligne.BackColor = Color_Cdd_Exclure
                    End If
                  End If
                Next XCel
              End If
            End If
            If Mnu_Item AndAlso Stg_Get(Plcy_Strg).Family = 5 AndAlso Stg_Get(Plcy_Strg).Type = "E" AndAlso U(Cellule, 2) = " " Then
              If XRslt.CelExcl.Count > 0 Then
                For Each XCel As XCel_Excl_Cls In XRslt.CelExcl
                  If XCel.Cel = Cellule Then
                    Opt = Ligne.Name(16)
                    If Opt = XCel.Cdd And U(Cellule, 3).Contains(XCel.Cdd) Then
                      Ligne.Visible = True : Ligne.BackColor = Color_Cdd_Exclure
                    End If
                  End If
                Next XCel
              End If
            End If
          Case "Mnu_Cel_Cdd_Ins_" 'Insérer les Candidats .... en NORMAL
            If Mnu_Sep AndAlso Stg_Get(Plcy_Strg).Family = 1 AndAlso U(Cellule, 2) = " " Then
              Candidats = U_CddExc(Cellule)
              If Candidats <> Cnddts_Blancs Then
                Ligne.Visible = True : Ligne.BackColor = Control.DefaultBackColor
              End If
            End If
            If Mnu_Item AndAlso Stg_Get(Plcy_Strg).Family = 1 AndAlso U(Cellule, 2) = " " Then
              Candidats = U_CddExc(Cellule)
              Opt = Ligne.Name(16)
              If Opt = Candidats(CInt(Opt) - 1) Then
                Ligne.Visible = True : Ligne.BackColor = Control.DefaultBackColor
              End If
            End If
          Case "Mnu_Cel_Val_Eff_" 'On efface la valeur de la cellule remplie
            If (U(Cellule, 1) = " " AndAlso U(Cellule, 2) <> " ") Then
              If Mnu_Item Then
                Ligne.Visible = True
                Ligne.BackColor = Control.DefaultBackColor
              End If
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