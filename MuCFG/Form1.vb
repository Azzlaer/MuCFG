Imports System.IO
Public Class Form1
    Dim objShell = CreateObject("Wscript.Shell")

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\", "ID", TextBox1.Text, Microsoft.Win32.RegistryValueKind.Unknown)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\", "UserID", TextBox1.Text, Microsoft.Win32.RegistryValueKind.Unknown)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\", "Exe", TextBox2.Text, Microsoft.Win32.RegistryValueKind.Unknown)
        ListBox1.Items.Add("Acceso rapido de USUARIO ha sido establecida")
        MsgBox("Cambios efectuados correctamente.", vbInformation)
        Select Case MsgBox("¿Desea Ejecutar el Main.exe desde esta aplicacion?", MsgBoxStyle.YesNoCancel, "caption")
            Case MsgBoxResult.Yes
                Process.Start(Application.StartupPath & "\main.exe")
            Case MsgBoxResult.Cancel
                Process.Start("www.facebook.com/latinbattle")
            Case MsgBoxResult.No
                Close()
        End Select
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        Select Case ComboBox3.SelectedItem
            Case "Volumen 1"
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\VolumeLevel", "1", "REG_DWORD")
                ListBox1.Items.Add("[Multimedia] El volumen de juego ha sido establecido a 1")
            Case "Volumen 2"
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\VolumeLevel", "2", "REG_DWORD")
                ListBox1.Items.Add("[Multimedia] El volumen de juego ha sido establecido a 2")
            Case "Volumen 3"
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\VolumeLevel", "3", "REG_DWORD")
                ListBox1.Items.Add("[Multimedia] El volumen de juego ha sido establecido a 3")
            Case "Volumen 4"
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\VolumeLevel", "4", "REG_DWORD")
                ListBox1.Items.Add("[Multimedia] El volumen de juego ha sido establecido a 4")
            Case "Volumen 5"
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\VolumeLevel", "5", "REG_DWORD")
                ListBox1.Items.Add("[Multimedia] El volumen de juego ha sido establecido a 5")
        End Select
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Select Case ComboBox1.SelectedItem
            Case "Spanish"
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\LangSelection", "SPN", "REG_SZ")
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\LanguageEx", "SPN", "REG_SZ")
                ListBox1.Items.Add("[Configuracion] Se ha establecido el idioma ESPAÑOL")
            Case "English"
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\LangSelection", "ENG", "REG_SZ")
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\LanguageEx", "ENG", "REG_SZ")
                ListBox1.Items.Add("[Configuracion] Se ha establecido el idioma INGLES")
            Case "Portuguese"
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\LangSelection", "POR", "REG_SZ")
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\LanguageEx", "POR", "REG_SZ")
                ListBox1.Items.Add("[Configuracion] Se ha establecido el idioma PORTUGUES")
        End Select
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Select Case ComboBox2.SelectedItem
            Case "640x480"
                On Error Resume Next
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\Resolution", "00000000", "REG_DWORD")
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\ResolutionA", "00000000", "REG_DWORD")
                ListBox1.Items.Add("[Graficos] Se ha establecido la resolucion 640x480")

            Case "800x600"
                On Error Resume Next
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\Resolution", "00000002", "REG_DWORD")
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\Resolution", "00000002", "REG_DWORD")
                ListBox1.Items.Add("[Graficos] Se ha establecido la resolucion 800x600")
            Case "1024x768"
                On Error Resume Next
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\Resolution", "00000001", "REG_DWORD")
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\ResolutionA", "00000001", "REG_DWORD")
                ListBox1.Items.Add("[Graficos] Se ha establecido la resolucion 1024x768")
            Case "1200x1024"
                On Error Resume Next
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\Resolution", "00000003", "REG_DWORD")
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\ResolutionA", "00000003", "REG_DWORD")
                ListBox1.Items.Add("[Graficos] Se ha establecido la resolucion 1200x1024")
            Case "1280x720"
                On Error Resume Next
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\Resolution", "00000005", "REG_DWORD")
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\ResolutionA", "00000005", "REG_DWORD")
                ListBox1.Items.Add("[Graficos] Se ha establecido la resolucion 1280x720")
            Case "1280x800"
                On Error Resume Next
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\Resolution", "00000006", "REG_DWORD")
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\ResolutionA", "00000006", "REG_DWORD")
                ListBox1.Items.Add("[Graficos] Se ha establecido la resolucion 1280x800")
            Case "1280x1024"
                On Error Resume Next
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\Resolution", "00000003", "REG_DWORD")
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\ResolutionA", "00000003", "REG_DWORD")
                ListBox1.Items.Add("[Graficos] Se ha establecido la resolucion 1280x1024")
            Case "1360x768"
                On Error Resume Next
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\Resolution", "00000005", "REG_DWORD")
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\ResolutionA", "00000005", "REG_DWORD")
                ListBox1.Items.Add("[Graficos] Se ha establecido la resolucion 1360x768")
            Case "1440x900"
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\Resolution", "00000008", "REG_DWORD")
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\ResolutionA", "00000008", "REG_DWORD")
                ListBox1.Items.Add("[Graficos] Se ha establecido la resolucion 1440x900")
            Case "1600x1200"
                On Error Resume Next
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\Resolution", "00000004", "REG_DWORD")
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\ResolutionA", "00000004", "REG_DWORD")
                ListBox1.Items.Add("[Graficos] Se ha establecido la resolucion 1600x1200")
            Case "1680x1050"
                On Error Resume Next
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\Resolution", "00000009", "REG_DWORD")
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\ResolutionA", "00000009", "REG_DWORD")
                ListBox1.Items.Add("[Graficos] Se ha establecido la resolucion 1680x1050")
        End Select
    End Sub

    Private Sub CheckBox2_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\MusicOnOff", "1", "REG_SZ")
            objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\MusicOn", "1", "REG_SZ")
            ListBox1.Items.Add("[Multimedia] Se ha habilitado la Musica")
        Else
            objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\MusicOn", "0", "REG_SZ")
            objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\MusicOnOff", "0", "REG_SZ")
            ListBox1.Items.Add("[Multimedia] Se ha deshabilitado la Musica")
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\SoundOnOff", "1", "REG_DWORD")
            objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\SoundOn", "1", "REG_DWORD")
            ListBox1.Items.Add("[Multimedia] Se ha habilitado el sonido")
        Else
            objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\SoundOn", "0", "REG_DWORD")
            objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\SoundOnOff", "0", "REG_DWORD")
            ListBox1.Items.Add("[Multimedia] Se ha deshabilitado el sonido")
        End If
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        Select Case ComboBox4.SelectedItem
            Case "Ventana"
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\WindowMode", "00000001", "REG_DWORD")
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\FullScreenMode", "1", "REG_DWORD")
                ListBox1.Items.Add("[Graficos] Se ha habilitado el Modo Ventana")
            Case "Completa"
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\WindowMode", "00000000", "REG_DWORD")
                objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\FullScreenMode", "0", "REG_DWORD")
                ListBox1.Items.Add("[Graficos] Se ha habilitado el Modo Pantalla Completa")
        End Select
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.Text = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\", "ID", " ")
        TextBox1.Text = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\", "UserID", " ")
        TextBox2.Text = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\", "Exe", " ")
        ListBox1.Items.Add("Aplicacion desarrollada por Azzlaer")
        ListBox1.Items.Add("para el foro de TuServerMU")
        ListBox1.Items.Add("Se ha iniciado la aplicacion MuCFG")
        ListBox1.Items.Add("Se han cargado los registros correctamente")
        Process.Start("http://www.facebook.com/freaksgamesnetworks")
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        OpenFileDialog1.ShowDialog()
        TextBox2.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\EffectSmoke", "1", "REG_DWORD")
            ListBox1.Items.Add("[Extra] Se ha habilitado el Efecto de Humo")
        Else
            objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\EffectSmoke", "0", "REG_DWORD")
            ListBox1.Items.Add("[Extra] Se ha deshabilitado el Efecto de Humo")
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\EffectScene", "1", "REG_DWORD")
            ListBox1.Items.Add("[Extra] Se ha habilitado el Efecto de Escenarios")
        Else
            objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\EffectScene", "0", "REG_DWORD")
            ListBox1.Items.Add("[Extra] Se ha deshabilitado el Efecto de Escenarios")
        End If
    End Sub

    Private Sub CheckBox5_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = True Then
            objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\EffectMist", "1", "REG_DWORD")
            ListBox1.Items.Add("[Extra] Se ha habilitado el Efecto de Niebla")
        Else
            objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\EffectMist", "0", "REG_DWORD")
            ListBox1.Items.Add("[Extra] Se ha deshabilitado el Efecto de Niebla")
        End If
    End Sub

    Private Sub RadioButton3_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\ColorDepth", "0", "REG_DWORD")
        ListBox1.Items.Add("[Extra] Se establecio la profundidad de color en 16")
    End Sub

    Private Sub RadioButton4_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged
        objShell.RegWrite("HKEY_CURRENT_USER\SOFTWARE\Webzen\Mu\Config\ColorDepth", "1", "REG_DWORD")
        ListBox1.Items.Add("[Extra] Se establecio la profundidad de color en 32")
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Using SW As New IO.StreamWriter(Application.StartupPath & "\logs.txt", True)
            For Each itm As String In Me.ListBox1.Items
                SW.WriteLine(itm)
            Next
        End Using
        Process.Start(Application.StartupPath & "\logs.txt", True)
    End Sub

    Private Sub LinkLabel2_LinkClicked_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Process.Start("http://www.facebook.com/freaksgamesnetworks")
        Process.Start("http://tuservermu.com.ve/index.php?action=profile;u=370")
    End Sub

    Private Sub LinkLabel1_LinkClicked_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Process.Start("http://tuservermu.com.ve")
    End Sub
End Class
