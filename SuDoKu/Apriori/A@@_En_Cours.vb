Module A_En_Cours
#Region "Menus Test A à J"
  Public Sub TestA()
    Jrn_Add(, {Proc_Name_Get()})
    Pzzl_Crt_Interactif_81("P")
  End Sub
  Public Sub TestB()
    Jrn_Add(, {Proc_Name_Get()})
    If Batch_en_Cours Then Exit Sub
    If Plcy_Generate_Batch Then
      Jrn_Add(, {"Lancement du traitement de création de grille en arrière-plan"})
      Frm_SDK.Batch_Timer.Interval = 5000
      'Process_16x est un png 16x16 32 bits
      'Me.Mnu08.Size = New System.Drawing.Size(98, 32)
      Frm_SDK.Mnu08.Image = SuDoKu.My.Resources.Resources.Process_16x
      Frm_SDK.Mnu08.Font = New Font(Frm_SDK.Mnu08.Font, FontStyle.Italic)
      Batch_Initial()
    Else
      Jrn_Add(, {"L'option Plcy_Generate_Batch est unabled."})
    End If
  End Sub
  Public Sub TestC()
    Jrn_Add(, {Proc_Name_Get()})
  End Sub
  Public Sub TestD()
    Jrn_Add(, {Proc_Name_Get()})
  End Sub
  Public Sub TestE()
    Jrn_Add(, {Proc_Name_Get()})
  End Sub
  Public Sub TestF()
    Jrn_Add(, {Proc_Name_Get()})
  End Sub
  Public Sub TestG()
    Jrn_Add(, {Proc_Name_Get()})
  End Sub
  Public Sub TestH()
    Jrn_Add(, {Proc_Name_Get()})
  End Sub
  Public Sub TestI()
    Jrn_Add(, {Proc_Name_Get()})
  End Sub
  Public Sub TestJ()
    Jrn_Add(, {Proc_Name_Get()})
    Plcy_Strg = "CaG"
    Jrn_Add(, {Stg_Get(Plcy_Strg).Texte & " " & Stg_Get(Plcy_Strg).Family})
    ' Liste des règles Msg_Read("MNU_0800A")
    Jrn_Add(, {Msg_Read("CAG_00000")})
    Jrn_Add(, {Msg_Read("CAG_00010")})
    Jrn_Add(, {Msg_Read("CAG_00020")})
    Jrn_Add(, {Msg_Read("CAG_00030")})
    Jrn_Add(, {Msg_Read("CAG_00040")})
    Jrn_Add(, {Msg_Read("CAG_00050")})
    Jrn_Add(, {Msg_Read("CAG_00060")})
    Jrn_Add(, {Msg_Read("CAG_00070")})
    Jrn_Add(, {Msg_Read("CAG_00080")})
    Jrn_Add(, {Msg_Read("CAG_00090")})
    Jrn_Add(, {Msg_Read("CAG_00100")})
    Jrn_Add(, {Msg_Read("CAG_00110")})


    ' Prise en Compte de la Grille
    Frm_SDK.B_Famille.Text = Stg_Get(Plcy_Strg).Family.ToString()
    Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte
    Frm_SDK.B_Info.BackColor = Color.Orange
    For i As Integer = 0 To 80
      U(i, 3) = Cnddts_Blancs
      U_CddExc(i) = Cnddts_Blancs
      U_Sol(i) = "0"
    Next
    Frm_SDK.Invalidate()
  End Sub
#End Region
End Module