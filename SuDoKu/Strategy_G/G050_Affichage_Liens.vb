'-------------------------------------------------------------------------------------------
' GLk G Link Affichage Liens
' Mise en place 08/01/2025 
'-------------------------------------------------------------------------------------------
Module G050_Affichage_Liens

  Sub Strategy_GLk(U_temp(,) As String)
    Plcy_Strg = "GLk"

    Dim Titre As String = "Stratégie GLk"
    Dim Texte As String = "Entrez le candidat à traiter pour afficher les liens (ou 0 pour les traiter tous) :"
    Dim Candidat_IB As String = InputBox(Texte, Titre)
    If Candidat_IB Is "" Then Exit Sub 'Cancel enfoncé

    GRslt_Init()
    GLinks_Build(U_temp, Candidat_IB)
    GRslt.Candidat = {Candidat_IB, "0"}
    GRslt.Nb_Liens = GLinks.Count

    GLinks_OrderBy()
    GLinks_Display()
    GRslt_Display()
    Event_OnPaint = "Global"
    Frm_SDK.Invalidate()
    Application.DoEvents()
  End Sub

End Module
