Imports System.Security
Imports LicenseManager


Public Class frmlogin

    Dim Cn As New SqlClient.SqlConnection

    Dim Clogin As Boolean = False
    Dim WURL As String
    Private CanOpenForm As Boolean = False




    Private Sub frmlogin_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        '///COMMENTED SIVAKUMAR ON 14-Sep-2018

        '  Me.Close()
    End Sub

    ''' <summary>
    ''' JGC MRMS Login
    ''' </summary>
    Private Sub frmlogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            '///COMMENTED SIVAKUMAR ON 14-Sep-2018


            Text = "J-MRMS V" & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." &
            My.Application.Info.Version.Build & "." & My.Application.Info.Version.Revision & " Login"
            MinimizeBox = False
            MaximizeBox = False
            '// Login person system ID //


            '=======================================================================================
            '// Review Tool //
            cb_reviewtool.Items.Clear()
            'cb_reviewtool.Items.Add("NavisWorks 5.4")
            ' cb_reviewtool.Items.Add("NavisWorks 2010")
            'cb_reviewtool.Items.Add("NavisWorks 2012")
            'cb_reviewtool.Items.Add("NavisWorks 2014")
            cb_reviewtool.Items.Add("NavisWorks 2017")
            cb_reviewtool.Items.Add("NavisWorks 2019")
            cb_reviewtool.Items.Add("NavisWorks 2022")
            cb_reviewtool.Items.Add("SmartPlant Review")
            '=======================================================================================


            '// copy execution file to local drive //
            Try
                '// Creating Temperary Directory //
                If IO.Directory.Exists("C:\Temp\") = False Then Call IO.Directory.CreateDirectory("C:\Temp\")

                '// CHECK C:\JMRMS DIRECTORY EXIST //
                If Not IO.Directory.Exists("C:\JMRMS\") Then
                    '// create new directory //
                    Call IO.Directory.CreateDirectory("C:\JMRMS\")
                    '// CREATE BATCH FILE AND COPY THE FILE 
                    Dim bfile As IO.StreamWriter
                    bfile = IO.File.CreateText("C:\Temp\jmrms_copy.bat")
                    Dim Gdic() As String = IO.Directory.GetDirectories(System.AppDomain.CurrentDomain.BaseDirectory)
                    bfile.WriteLine("COPY " & Chr(34) & System.AppDomain.CurrentDomain.BaseDirectory & Chr(34) & "  " & Chr(34) & "C:\JMRMS" & Chr(34))
                    For Each dir As String In Gdic
                        Dim dirinfo As New IO.DirectoryInfo(dir)
                        If Not IO.Directory.Exists("C:\JMRMS\" & dirinfo.Name) Then
                            Call IO.Directory.CreateDirectory("C:\JMRMS\" & dirinfo.Name)
                        End If
                        bfile.WriteLine("COPY " & Chr(34) & dir & Chr(34) & "  " & Chr(34) & "C:\JMRMS\" & dirinfo.Name & Chr(34))
                    Next
                    bfile.Close()
                    Call Shell("C:\temp\jmrms_copy.bat", AppWinStyle.Hide, True) '// RUN BATCH FILE 
                    Call IO.File.Delete("C:\temp\jmrms_copy.bat") '// DELETE BATCH FILE 
                Else
                    '// READ ORIGINaL FILE
                    Dim Oread() As String = IO.File.ReadAllLines(System.AppDomain.CurrentDomain.BaseDirectory & "\JGC Model Review Management System_WS.exe.manifest")
                    '// READ COPIED FILE 
                    Dim Cread() As String
                    If IO.File.Exists("C:\JMRMS\JGC Model Review Management System_WS.exe.manifest") Then
                        Cread = IO.File.ReadAllLines("C:\JMRMS\JGC Model Review Management System_WS.exe.manifest")
                        Try
                            If Oread(2) <> Cread(2) Then
                                Dim bfile As IO.StreamWriter
                                bfile = IO.File.CreateText("C:\temp\jmrms_copy.bat")
                                Dim Gdic() As String = IO.Directory.GetDirectories(System.AppDomain.CurrentDomain.BaseDirectory)
                                bfile.WriteLine("COPY " & Chr(34) & System.AppDomain.CurrentDomain.BaseDirectory & Chr(34) & "  " & Chr(34) & "C:\JMRMS" & Chr(34))
                                For Each dir As String In Gdic
                                    Dim dirinfo As New IO.DirectoryInfo(dir)
                                    If Not IO.Directory.Exists("C:\JMRMS\" & dirinfo.Name) Then
                                        Call IO.Directory.CreateDirectory("C:\JMRMS\" & dirinfo.Name)
                                    End If
                                    bfile.WriteLine("COPY " & Chr(34) & dir & Chr(34) & "  " & Chr(34) & "C:\JMRMS\" & dirinfo.Name & Chr(34))
                                Next
                                bfile.Close()
                                Call Shell("C:\temp\jmrms_copy.bat", AppWinStyle.Hide, True) '// RUN BATCH FILE 
                                Call IO.File.Delete("C:\temp\jmrms_copy.bat") '// DELETE BATCH FILE 
                            End If
                        Catch ex As Exception
                            '// skip //
                        End Try
                    Else
                        Dim bfile As IO.StreamWriter
                        bfile = IO.File.CreateText("C:\temp\jmrms_copy.bat")
                        Dim Gdic() As String = IO.Directory.GetDirectories(System.AppDomain.CurrentDomain.BaseDirectory)
                        bfile.WriteLine("COPY " & Chr(34) & System.AppDomain.CurrentDomain.BaseDirectory & Chr(34) & "  " & Chr(34) & "C:\JMRMS" & Chr(34))
                        For Each dir As String In Gdic
                            Dim dirinfo As New IO.DirectoryInfo(dir)
                            If Not IO.Directory.Exists("C:\JMRMS\" & dirinfo.Name) Then
                                Call IO.Directory.CreateDirectory("C:\JMRMS\" & dirinfo.Name)
                            End If
                            bfile.WriteLine("COPY " & Chr(34) & dir & Chr(34) & "  " & Chr(34) & "C:\JMRMS\" & dirinfo.Name & Chr(34))
                        Next
                        bfile.Close()
                        Call Shell("C:\temp\jmrms_copy.bat", AppWinStyle.Hide, True) '// RUN BATCH FILE 
                        Call IO.File.Delete("C:\temp\jmrms_copy.bat") '// DELETE BATCH FILE 
                    End If
                    Oread = Nothing
                    Cread = Nothing
                End If
            Catch ex As Exception
                '// skip //
            End Try

            '// READ XML FILE //
            If IO.File.Exists(My.Settings.strXMLfile) Then


                If txt_database.Text <> "" Then
                    txt_database.ReadOnly = True
                    txt_database.BackColor = Color.Wheat
                End If
                If txt_servername.Text <> "" Then
                    txt_servername.ReadOnly = True
                    txt_servername.BackColor = Color.Wheat
                End If
            End If

            '// Make connect button is Active //
            Me.but_Connect.Select()

        Catch ex As Exception
            '// skip //
        End Try
    End Sub

    ''' <summary>
    ''' Connecting to SQL Server 
    ''' </summary>
    Private Sub but_Connect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but_Connect.Click


    End Sub

    ''' <summary>
    ''' XML File Browser
    ''' </summary>
    Private Sub but_xmlbrowser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but_xmlbrowser.Click
        Dim Ofile As New OpenFileDialog
        Ofile.Title = "Specify the XML File"
        Ofile.Filter = "XML|*.xml"
        Ofile.FilterIndex = 1
        Ofile.FileName = My.Settings.strXMLfile
        Dim result = Ofile.ShowDialog
        If result = MsgBoxResult.Ok Then
            txt_xmlfile.Text = Ofile.FileName
            My.Settings.strXMLfile = Ofile.FileName



        End If
    End Sub

    Private Sub but_cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but_cancel.Click
        Me.Close()
    End Sub

    Private Sub but_Connect_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles but_Connect.MouseEnter, Button1.MouseEnter
        Cursor = Cursors.Hand
    End Sub

    Private Sub but_Connect_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles but_Connect.MouseLeave, Button1.MouseLeave
        Cursor = Cursors.Default
    End Sub

    Private Sub frmlogin_Scroll(ByVal sender As Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles Me.Scroll

    End Sub

    Private Sub frmlogin_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Dim Title As String = String.Empty
        For Each s As String In My.Application.CommandLineArgs
            If s.StartsWith("-iTitle") Then
                Title = s.Replace("-iTitle:", vbNullString)
                Exit For
            End If
        Next

        Select Case Title
            Case "REVIEW COMMENT SUMMARY"

                Dim iXML As String = String.Empty
                Dim iReviewType As String = String.Empty
                Dim iDiscipline As String = String.Empty
                Dim iStatus As String = String.Empty
                Dim iImage As String = String.Empty
                Dim iAttachment As String = String.Empty
                Dim iViewpointXML As String = String.Empty
                Dim iFilename As String = String.Empty
                Dim iNavis As String = String.Empty
                Dim iNavisfilename As String = String.Empty
                For Each s As String In My.Application.CommandLineArgs
                    ' MsgBox(s)
                    If s.StartsWith("-iXML") Then
                        iXML = s.Replace("-iXML:", vbNullString)
                    ElseIf s.StartsWith("-iReviewType") Then
                        iReviewType = s.Replace("-iReviewType:", vbNullString)
                    ElseIf s.StartsWith("-iDiscipline") Then
                        iDiscipline = s.Replace("-iDiscipline:", vbNullString)
                    ElseIf s.StartsWith("-iStatus") Then
                        iStatus = s.Replace("-iStatus:", vbNullString)
                    ElseIf s.StartsWith("-iImage") Then
                        iImage = s.Replace("-iImage:", vbNullString)
                    ElseIf s.StartsWith("-iAttachment") Then
                        iAttachment = s.Replace("-iAttachment:", vbNullString)
                    ElseIf s.StartsWith("-iViewpointXML") Then
                        iViewpointXML = s.Replace("-iViewpointXML:", vbNullString)
                    ElseIf s.StartsWith("-iNavisfilename") Then
                        iNavisfilename = s.Replace("-iNavisfilename:", vbNullString)
                    ElseIf s.StartsWith("-iFilename") Then
                        iFilename = s.Replace("-iFilename:", vbNullString)
                    ElseIf s.StartsWith("-iNavis") Then
                        iNavis = s.Replace("-iNavis:", vbNullString)
                    End If
                    Clogin = True
                Next
                If iXML.Length = 0 Or Not IO.File.Exists(iXML) Then
                    If Clogin Then
                        End
                    Else
                        Exit Sub
                    End If
                End If

        End Select
    End Sub


End Class