Option Strict On
Option Explicit On

'-------------------------------------------------------------------------------------------
' Stratégie des CdU avec Strategy_Rslt
'
' La stratégie CdU est la plus simple, il faut que le candidat soit unique, le plus simple est de tester
'    la longueur de Trim(U_temp(Cellule, 3)).Length = 1
'-------------------------------------------------------------------------------------------
Friend Module P10_Strategy_CdU
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  'Il est impératif d'utiliser U_temp, copie de U, pour utiliser la stratégie interactivement ET en arrière-plan
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

  Function Strategy_CdU(U_temp(,) As String) As String(,)
    Dim Stratégie As String = "CdU"
    Dim Sous_Stratégie As String = "CdU"

    Dim Strategy_Rslt(99, 0) As String 'initialisé  
    Strategy_Rslt_Init(Strategy_Rslt, Stratégie, Sous_Stratégie) 'Initialisation 

    Dim Cel45(0 To 44) As String : For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
    Dim Exc45(0 To 44) As String : For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k
    Dim Code_LCR As String = "#"
    For Cellule As Integer = 0 To 80
      'La cellule doit être vide et ne comporter qu'un seul candidat
      ' La solution ci-dessous est préférée à Wh_Cell_Nb_Candidats(U, Cellule) = 1
      If U_temp(Cellule, 2) = " " And Trim(U_temp(Cellule, 3)).Length = 1 Then
        Dim Candidat As String = Trim(U_temp(Cellule, 3))
        'Seule la Cel45(0) comporte le CdU
        Cel45(0) = CStr(Cellule)
        Strategy_Rslt_Add(Strategy_Rslt, Stratégie, Sous_Stratégie, Code_LCR, "0", Candidat, Cel45, Exc45)
      End If
    Next Cellule
    Return Strategy_Rslt
  End Function

End Module