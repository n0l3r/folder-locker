Imports System.IO

Public Class Form2
    Dim acsn As New FolderBrowserDialog

    'settings
    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.AllowDrop = True
        Button1.Focus()
        homehigh.Visible = False
    End Sub

    Private Sub low_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles low.CheckedChanged
        If low.Checked = True Then
            high.Checked = False
        Else
            high.Checked = True
        End If
    End Sub

    Private Sub high_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles high.CheckedChanged
        If high.Checked = True Then
            low.Checked = False
        Else
            low.Checked = True
        End If
    End Sub

    Private Sub disable_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles disable.CheckedChanged
        If disable.Checked = True Then
            user.Checked = False
            computerr.Checked = False
            classic.Checked = False
            recyclebin.Checked = False
        End If
    End Sub

    Private Sub user_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles user.CheckedChanged
        If user.Checked = True Then
            disable.Checked = False
            computerr.Checked = False
            classic.Checked = False
            recyclebin.Checked = False
        End If
    End Sub

    Private Sub classic_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles classic.CheckedChanged
        If classic.Checked = True Then
            user.Checked = False
            computerr.Checked = False
            disable.Checked = False
            recyclebin.Checked = False
        End If
    End Sub

    Private Sub computerr_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles computerr.CheckedChanged
        If computerr.Checked = True Then
            user.Checked = False
            disable.Checked = False
            classic.Checked = False
            recyclebin.Checked = False
        End If
    End Sub

    Private Sub recyclebin_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles recyclebin.CheckedChanged
        If recyclebin.Checked = True Then
            user.Checked = False
            computerr.Checked = False
            classic.Checked = False
            disable.Checked = False
        End If
    End Sub

    Function ceklok() As Boolean
        If IO.File.Exists(pathlow.Text & "\Desktop.ini") Then
            ceklok = True
        Else
            ceklok = False
        End If
    End Function

    Function ceklok2() As Boolean
        If IO.File.Exists(pathhigh.Text & "\Desktop.ini") Then
            ceklok2 = True
        Else
            ceklok2 = False
        End If
    End Function

    Function cekpass() As Boolean
        If IO.File.Exists(pathlow.Text & "\kepo.lol") Then
            cekpass = True
        Else
            cekpass = False
        End If
    End Function

    Function cekpass2() As Boolean
        If IO.File.Exists(pathhigh.Text & "\kepo.lol") Then
            cekpass2 = True
        Else
            cekpass2 = False
        End If
    End Function

    'homelow

    Private Sub browselow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles browselow.Click
        If acsn.ShowDialog = Windows.Forms.DialogResult.OK Then
            pathlow.Text = acsn.SelectedPath
            Application.DoEvents()
        End If
    End Sub

    Private Sub locklow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles locklow.Click
        If Len(pathlow.Text) <= 3 Then
            If acsn.ShowDialog = Windows.Forms.DialogResult.OK Then
                pathlow.Text = acsn.SelectedPath
            End If
        Else
            If ceklok() = True Then
                MsgBox("Folder already locked", MsgBoxStyle.Critical, "ACSN Locker")
            Else
                Dim fs As FileStream = File.Create(pathlow.Text & "\Desktop.ini")
                fs.Close()
                If disable.Checked = True Then
                    Dim filee As New StreamWriter(pathlow.Text & "\Desktop.ini")
                    filee.WriteLine("[.ShellClassInfo]")
                    filee.WriteLine("CLSID={871C5380-42A0-1069-A2EA-08002B30309D}")
                    filee.Close()
                    SetAttr(pathlow.Text & "\Desktop.ini", FileAttribute.System + FileAttribute.Hidden)
                    SetAttr(pathlow.Text, FileAttribute.ReadOnly)
                    Application.DoEvents()
                End If
                If classic.Checked = True Then
                    Dim filee As New StreamWriter(pathlow.Text & "\Desktop.ini")
                    filee.WriteLine("[.ShellClassInfo]")
                    filee.WriteLine("CLSID={2559a1f2-21d7-11d4-bdaf-00c04f60b9f0}")
                    filee.Close()
                    SetAttr(pathlow.Text & "\Desktop.ini", FileAttribute.System + FileAttribute.Hidden)
                    SetAttr(pathlow.Text, FileAttribute.ReadOnly)
                    Application.DoEvents()
                End If
                If user.Checked = True Then
                    Dim filee As New StreamWriter(pathlow.Text & "\Desktop.ini")
                    filee.WriteLine("[.ShellClassInfo]")
                    filee.WriteLine("CLSID={59031A47-3F72-44A7-89C5-5595FE6B30EE}")
                    filee.Close()
                    SetAttr(pathlow.Text & "\Desktop.ini", FileAttribute.System + FileAttribute.Hidden)
                    SetAttr(pathlow.Text, FileAttribute.ReadOnly)
                    Application.DoEvents()
                End If
                If computerr.Checked = True Then
                    Dim filee As New StreamWriter(pathlow.Text & "\Desktop.ini")
                    filee.WriteLine("[.ShellClassInfo]")
                    filee.WriteLine("CLSID={20D04FE0-3AEA-1069-A2D8-08002B30309D}")
                    filee.Close()
                    SetAttr(pathlow.Text & "\Desktop.ini", FileAttribute.System + FileAttribute.Hidden)
                    SetAttr(pathlow.Text, FileAttribute.ReadOnly)
                    Application.DoEvents()
                End If
                If recyclebin.Checked = True Then
                    Dim filee As New StreamWriter(pathlow.Text & "\Desktop.ini")
                    filee.WriteLine("[.ShellClassInfo]")
                    filee.WriteLine("CLSID={645FF040-5081-101B-9F08-00AA002F954E}")
                    filee.Close()
                    SetAttr(pathlow.Text & "\Desktop.ini", FileAttribute.System + FileAttribute.Hidden)
                    SetAttr(pathlow.Text, FileAttribute.ReadOnly)
                    Application.DoEvents()
                End If
                Application.DoEvents()
                MsgBox("Success Lock your Folder", MsgBoxStyle.Information, "ACSN Folder Locker")
            End If
        End If
    End Sub

    Private Sub unlocklow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles unlocklow.Click
        If Len(pathlow.Text) <= 3 Then
            If acsn.ShowDialog = Windows.Forms.DialogResult.OK Then
                pathlow.Text = acsn.SelectedPath
            End If
        Else
            If ceklok() = False Then
                MsgBox("Folder not yet locked", MsgBoxStyle.Critical, "ACSN Low Locker")
            Else
                If cekpass() = False Then
                    SetAttr(pathlow.Text, FileAttribute.Normal)
                    SetAttr(pathlow.Text & "\Desktop.ini", FileAttribute.Normal)
                    Kill(pathlow.Text & "\Desktop.ini")
                    Application.DoEvents()
                    MsgBox("Succes UnLock your Folder", MsgBoxStyle.Information, "ACSN Folder Locker")
                Else
                    MsgBox("Your Folder Lock with Password, please change Type to UnLock", MsgBoxStyle.Critical, "ACSN Folder Locker")
                End If
            End If
        End If
    End Sub

    'homehigh

    Private Sub browsehigh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles browsehigh.Click
        If acsn.ShowDialog = Windows.Forms.DialogResult.OK Then
            pathhigh.Text = acsn.SelectedPath
        End If
    End Sub

    Private Sub lockhigh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lockhigh.Click
        If Len(pathhigh.Text) <= 3 Then
            If acsn.ShowDialog = Windows.Forms.DialogResult.OK Then
                pathhigh.Text = acsn.SelectedPath
            End If
        Else
            If ceklok2() = True Then
                MsgBox("Folder already locked", MsgBoxStyle.Critical, "ACSN Locker")
            Else
                inputpass.Visible = True
                password.Focus()
            End If
        End If
    End Sub


    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If password.Text = "" Then
            Label11.Text = "Fill your Password"
        Else
            If disable.Checked = True Then
                Dim fsg As FileStream = File.Create(pathhigh.Text & "\Desktop.ini")
                fsg.Close()
                Dim filee As New StreamWriter(pathhigh.Text & "\Desktop.ini")
                filee.WriteLine("[.ShellClassInfo]")
                filee.WriteLine("CLSID={871C5380-42A0-1069-A2EA-08002B30309D}")
                filee.Close()
                SetAttr(pathhigh.Text & "\Desktop.ini", FileAttribute.System + FileAttribute.Hidden)
                SetAttr(pathhigh.Text, FileAttribute.ReadOnly)
                Application.DoEvents()
              End If
            If classic.Checked = True Then
                Dim fs As FileStream = File.Create(pathhigh.Text & "\Desktop.ini")
                fs.Close()
                Dim filee As New StreamWriter(pathhigh.Text & "\Desktop.ini")
                filee.WriteLine("[.ShellClassInfo]")
                filee.WriteLine("CLSID={2559a1f2-21d7-11d4-bdaf-00c04f60b9f0}")
                filee.Close()
                SetAttr(pathhigh.Text & "\Desktop.ini", FileAttribute.System + FileAttribute.Hidden)
                SetAttr(pathhigh.Text, FileAttribute.ReadOnly)
                Application.DoEvents()
             End If
            If user.Checked = True Then
                Dim fs As FileStream = File.Create(pathhigh.Text & "\Desktop.ini")
                fs.Close()
                Dim filee As New StreamWriter(pathhigh.Text & "\Desktop.ini")
                filee.WriteLine("[.ShellClassInfo]")
                filee.WriteLine("CLSID={59031A47-3F72-44A7-89C5-5595FE6B30EE}")
                filee.Close()
                SetAttr(pathhigh.Text & "\Desktop.ini", FileAttribute.System + FileAttribute.Hidden)
                SetAttr(pathhigh.Text, FileAttribute.ReadOnly)
                Application.DoEvents()
             End If
            If computerr.Checked = True Then
                Dim fs As FileStream = File.Create(pathhigh.Text & "\Desktop.ini")
                fs.Close()
                Dim filee As New StreamWriter(pathhigh.Text & "\Desktop.ini")
                filee.WriteLine("[.ShellClassInfo]")
                filee.WriteLine("CLSID={20D04FE0-3AEA-1069-A2D8-08002B30309D}")
                filee.Close()
                SetAttr(pathhigh.Text & "\Desktop.ini", FileAttribute.System + FileAttribute.Hidden)
                SetAttr(pathhigh.Text, FileAttribute.ReadOnly)
                Application.DoEvents()
            End If
            If recyclebin.Checked = True Then
                Dim fs As FileStream = File.Create(pathhigh.Text & "\Desktop.ini")
                fs.Close()
                Dim filee As New StreamWriter(pathhigh.Text & "\Desktop.ini")
                filee.WriteLine("[.ShellClassInfo]")
                filee.WriteLine("CLSID={645FF040-5081-101B-9F08-00AA002F954E}")
                filee.Close()
                SetAttr(pathhigh.Text & "\Desktop.ini", FileAttribute.System + FileAttribute.Hidden)
                SetAttr(pathhigh.Text, FileAttribute.ReadOnly)
                Application.DoEvents()
            End If
            Application.DoEvents()
            Dim fsss As FileStream = File.Create(pathhigh.Text & "\kepo.lol")
            fsss.Close()
            Dim tulis As New StreamWriter(pathhigh.Text & "\kepo.lol")
            tulis.WriteLine(password.Text)
            tulis.Close()
            Application.DoEvents()
            Dim fss As FileStream = File.Create(pathhigh.Text & "\hint.lol")
            fss.Close()
            Dim tuliss As New StreamWriter(pathhigh.Text & "\hint.lol")
            tuliss.WriteLine(hpassword.Text)
            tuliss.Close()
            SetAttr(pathhigh.Text & "\kepo.lol", FileAttribute.Hidden)
            SetAttr(pathhigh.Text & "\hint.lol", FileAttribute.Hidden)
            SetAttr(pathhigh.Text, FileAttribute.ReadOnly)
            Application.DoEvents()
            MsgBox("Succes Lock your Folder", MsgBoxStyle.Information, "ACSN Folder Locker")
            password.Text = ""
            hpassword.Text = ""
            inputpass.Visible = False
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        inputpass.Visible = False
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        inputpass.Visible = False
        GroupBox5.Visible = False
    End Sub

    Private Sub unlockhigh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles unlockhigh.Click
        If Len(pathhigh.Text) <= 3 Then
            If acsn.ShowDialog = Windows.Forms.DialogResult.OK Then
                pathhigh.Text = acsn.SelectedPath
            End If
        Else
            If ceklok2() = False Then
                MsgBox("Folder not yet locked", MsgBoxStyle.Critical, "ACSN Low Locker")
            Else
                If cekpass2() = True Then
                    Application.DoEvents()
                    Dim rs As New StreamReader(pathhigh.Text & "\hint.lol")
                    Dim bs As String
                    Do Until rs.EndOfStream
                        bs = rs.ReadLine
                        Label14.Text = "Hint Password :" & bs
                    Loop
                    Application.DoEvents()
                    GroupBox5.Visible = True
                    passwordnya.Focus()
                    rs.Close()
                Else
                    MsgBox("Your Folder Lock with No Password, please change Type to UnLock", MsgBoxStyle.Critical, "ACSN Folder Locker")
                End If
            End If
        End If
    End Sub
    Private Sub passwordnya_KeyPress(ByVal Sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles passwordnya.KeyPress
        Dim sr As New StreamReader(pathhigh.Text & "\kepo.lol")
        Dim s As String
        s = sr.ReadLine
        Application.DoEvents()
        sr.Close()
        If e.KeyChar = ChrW(Keys.Enter) Then
            If passwordnya.Text = s Then
                sr.Close()
                SetAttr(pathhigh.Text, FileAttribute.Normal)
                SetAttr(pathhigh.Text & "\Desktop.ini", FileAttribute.Normal)
                SetAttr(pathhigh.Text & "\kepo.lol", FileAttribute.Normal)
                SetAttr(pathhigh.Text & "\hint.lol", FileAttribute.Normal)
                Kill(pathhigh.Text & "\Desktop.ini")
                Kill(pathhigh.Text & "\kepo.lol")
                Kill(pathhigh.Text & "\hint.lol")
                Application.DoEvents()
                s = Nothing
                Application.DoEvents()
                MsgBox("Succes UnLock your Folder", MsgBoxStyle.Information, "ACSN Folder Locker")
                GroupBox5.Visible = False
                inputpass.Visible = False
                passwordnya.Text = ""
            Else
                Label15.Text = "Wrong Password !"
                Label15.Visible = True
                passwordnya.Text = ""
                Application.DoEvents()
            End If
        End If
    End Sub

    'desain

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If low.Checked = True Then
            about.Visible = True
            settings.Visible = True
            homelow.Visible = True
            homehigh.Visible = False
            Button1.Focus()
        Else
            about.Visible = True
            settings.Visible = True
            homelow.Visible = True
            homehigh.Visible = True
            Button1.Focus()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        about.Visible = True
        settings.Visible = True
        homelow.Visible = False
        homehigh.Visible = False
        Button2.Focus()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        about.Visible = True
        settings.Visible = False
        homelow.Visible = False
        homehigh.Visible = False
        Button3.Focus()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Form2_DragDrop(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragDrop
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
        For Each Path In files
            pathhigh.Text = Path
            pathlow.Text = Path
        Next
    End Sub

    Private Sub Form2_DragEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

End Class