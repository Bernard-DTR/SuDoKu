Imports System.IO

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
    Dim Application_SDK As String = Path_SDK & "SuDoKu\SuDoKu"
    Dim File_nb As Integer

    Dim files As IEnumerable(Of String) = From file In Directory.EnumerateFiles(Application_SDK, "*.*", SearchOption.AllDirectories)
      Jrn_Add(, {Path_SDK & "SuDoKu\SuDoKu" & " : " & files.Count})

      For Each File As String In files
        Dim d As Integer = Application_SDK.Length
        Dim l As Integer = File.Length
        Dim File_Sht As String = Mid$(File, d + 2, l - d + 2)
        If Mid$(File_Sht, 1, 3) = ".vs" Then Continue For
        If Mid$(File_Sht, 1, 3) = "bin" Then Continue For
        If Mid$(File_Sht, 1, 11) = "My Project\" Then Continue For
        If Mid$(File_Sht, 1, 3) = "obj" Then Continue For
        File_nb += 1
        Jrn_Add(, {CStr(File_nb).PadLeft(3) & " " & File_Sht})
      Next File

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
  End Sub
#End Region
End Module