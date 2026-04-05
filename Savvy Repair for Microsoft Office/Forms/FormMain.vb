Option Explicit On

Imports System
Imports System.IO
Imports System.Xml
Imports System.Drawing
Imports System.Threading
Imports System.Xml.Schema
Imports System.ComponentModel
Imports System.Collections.Generic
Imports System.Text.RegularExpressions
Imports Microsoft.Win32
Imports System.Reflection

Public Class FormMain

#Region "Constants"

    Protected Const _strAppName As String = "Savvy Repair for Microsoft Office"

    Protected Const _bShowTraces As Boolean = False

#End Region

#Region "Fields"

    Dim WithEvents _worker As BackgroundWorker

#End Region

#Region "Properties"

    Public Property AppName As String

        Get
            Return _strAppName
        End Get

        Set(value As String)
        End Set

    End Property

#End Region

#Region "Ctors"

    Public Sub New()

        InitializeComponent()

        _worker = New BackgroundWorker()
        _worker.WorkerSupportsCancellation = True

    End Sub

#End Region

#Region "Events"

    Private Sub FormMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Try
            MsgBox("Depending on which kind of file you are trying to recover, " _
                   & "this program may not work if there are visible or invisible " _
                   & "Word, Excel or PowerPoint instances running in Task Manager. " _
                   & "Before hitting the OK button on this message, please be sure to hit " _
                   & "Ctl-Alt-Delete on your keyboard, start Task Manager and depending on " _
                   & "which applies to the file you are recovering, end all instances of " _
                   & """WINWORD.EXE"", ""WINWORD.EXE * 32"", ""EXCEL.EXE"", ""EXCEL.EXE * 32"", " _
                   & """POWERPNT.EXE"" or ""POWERPNT.EXE * 32"".", MsgBoxStyle.Exclamation)

        Catch

        End Try

    End Sub

    Private Sub miExit_Click(sender As System.Object, e As System.EventArgs) Handles miExit.Click

        Try

            Me.Close()

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

    End Sub

    Private Sub ToolStripMenuItem10_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem10.Click

        SelectFile()

    End Sub

    Private Sub ExitToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem1.Click

        Try

            Me.Close()

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

    End Sub

    Private Sub ToolStripMenuItem12_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem12.Click

        Impl_Recover_Auto1()

    End Sub

    Private Sub ToolStripMenuItem13_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem13.Click

        Impl_Recover_Method_1(False)

    End Sub

    Private Sub ToolStripMenuItem14_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem14.Click

        Impl_Recover_Method_3(False)

    End Sub

    Private Sub ToolStripMenuItem15_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem15.Click

        Impl_Recover_Method_4(False)

    End Sub

    Private Sub Method1BLaxValidationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Method1BLaxValidationToolStripMenuItem.Click

        Impl_Recover_Method_2(False)

    End Sub

    Private Sub Method3SalvageToolStripMenuItem_Click(sender As Object, e As EventArgs)

        Impl_Recover_Method_4(False)

    End Sub

    Private Sub ShadowExplorerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShadowExplorerToolStripMenuItem.Click

        Impl_ShadowExplorerURL()

    End Sub

    Private Sub PreviousVersionFileRecovererToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PreviousVersionFileRecovererToolStripMenuItem.Click

        Impl_PreviousVersionFileRecovererURL()

    End Sub

    Private Sub UnsavedOfficeFileRecoveryStepsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UnsavedOfficeFileRecoveryStepsToolStripMenuItem.Click

        Impl_UnsavedWordFileStepsURL()

    End Sub

    Private Sub S2ServicesCorruptFileRecoveryFreewareToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles S2ServicesCorruptFileRecoveryFreewareToolStripMenuItem.Click

        Impl_S2ServiceSourceforgeURL()

    End Sub

    Private Sub StepsToRecoverAWordFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StepsToRecoverAWordFileToolStripMenuItem.Click

        Impl_WordStepsURL()

    End Sub

    Private Sub StepsToRecoveringAnExcelFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StepsToRecoveringAnExcelFileToolStripMenuItem.Click

        Impl_ExcelStepsURL()

    End Sub

    Private Sub TryFreeOnlineServiceUseCouponS2SERVICESToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles TryFreeOnlineServiceUseCouponS2SERVICESToolStripMenuItem.Click

        Impl_OnlineFileRepairCouponVisit()

    End Sub

    Private Sub S2ServicesFreeServiceCurrentlyUnavailableToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles S2ServicesFreeServiceCurrentlyUnavailableToolStripMenuItem.Click

        Impl_SaveOfficeData()

    End Sub

    Private Sub GetDatasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GetDatasToolStripMenuItem.Click

        Impl_RepairMyWord()

    End Sub

    Private Sub StepsToRecoveringAnOpenOfficeFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StepsToRecoveringAnOpenOfficeFileToolStripMenuItem.Click

        Impl_OpenOfficeSteps()

    End Sub

    Private Sub ToolStripMenuItem17_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem17.Click

        Impl_TryWordFixURL()

    End Sub

    Private Sub TryExcelFixToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles TryExcelFixToolStripMenuItem.Click

        Impl_TryExcelFixUrl()

    End Sub

    Private Sub OfficeFixToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OfficeFixToolStripMenuItem.Click

        Impl_TryOfficeFixUrl()

    End Sub


    Private Sub TryOurManual22RepairToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles TryOurManual22RepairToolStripMenuItem.Click

        Impl_S2ServicesManualFileRepair()

    End Sub

    Private Sub SilverCodersDocToTextToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SilverCodersDocToTextToolStripMenuItem.Click

        Impl_Silvercoders_DocToText()

    End Sub

    Private Sub ToolStripMenuItem6_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem6.Click

        Impl_OpenHelpUrl()

    End Sub

    Private Sub ToolStripMenuItem7_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem7.Click

        FormAbout.Show()

    End Sub

    Private Sub ToolStripMenuItem8_Click(sender As Object, e As EventArgs)

        Impl_TryWordFixURL()

    End Sub

    Private Sub TryExcelFixToolStripMenuItem_Click(sender As Object, e As EventArgs)

        Impl_TryExcelFixUrl()

    End Sub

    Private Sub TryOurManual22RepairToolStripMenuItem_Click(sender As Object, e As EventArgs)

        Impl_S2ServicesManualFileRepair()

    End Sub

    Private Sub TryFreeOnlineServiceUseCouponS2SERVICESToolStripMenuItem_Click(sender As Object, e As EventArgs)

        Impl_OnlineFileRepairCouponVisit()

    End Sub

    Private Sub picWordFixLink_Click_1(sender As Object, e As EventArgs) Handles picWordFixLink.Click

        Impl_OpenPayPalDonateUrl()

    End Sub

    Private Sub ToolStripMenuItem16_Click_1(sender As Object, e As EventArgs) Handles ToolStripMenuItem16.Click

        Impl_OpenPayPalDonateUrl()

    End Sub

    Private Sub btnBrowse_Click(sender As System.Object, e As System.EventArgs) Handles btnBrowse.Click

        SelectFile()

    End Sub

    Private Sub btnAutoRecover_Click(sender As System.Object, e As System.EventArgs) Handles btnAutoRecover.Click

        Impl_Recover_Auto1()

    End Sub

    Private Sub btnRecoverMethods_Click(sender As System.Object, e As System.EventArgs) Handles btnRecoverMethods.Click

        Dim ptMenu As System.Drawing.Point = New System.Drawing.Point(btnRecoverMethods.Right, btnRecoverMethods.Top)
        ptMenu = PointToScreen(ptMenu)

        popupRecover.Show(ptMenu)

    End Sub

    Private Sub popupmenuitemAuto_Click(sender As System.Object, e As System.EventArgs) Handles popupmenuitemAuto.Click

        Impl_Recover_Auto1()

    End Sub

    Private Sub popupmenuitemMethod1_Click(sender As System.Object, e As System.EventArgs) Handles popupmenuitemMethod1.Click

        Impl_Recover_Method_1(False)

    End Sub

    Private Sub Method2RepairWithLaxXMLValidationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Method2RepairWithLaxXMLValidationToolStripMenuItem.Click

        Impl_Recover_Method_2(False)

    End Sub

    Private Sub popupmenuitemMethod2_Click(sender As System.Object, e As System.EventArgs) Handles popupmenuitemMethod2.Click

        Impl_Recover_Method_3(False)

    End Sub

    Private Sub popupmenuitemMethod3_Click(sender As System.Object, e As System.EventArgs) Handles popupmenuitemMethod3.Click

        Impl_Recover_Method_4(False)

    End Sub

    Private Sub ToolStripMenuItem18_Click(sender As Object, e As EventArgs)

        Impl_Recover_Method_2(False)

    End Sub

    Private Sub ToolStripMenuItem20_Click(sender As Object, e As EventArgs)

        Impl_Recover_Method_4(False)

    End Sub

    Private Sub picWordFixLink_Click(sender As System.Object, e As System.EventArgs)

        Impl_TryWordFixURL()

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)

        Impl_TryExcelFixUrl()

    End Sub

    Private Sub txtInputFile_DragEnter(sender As System.Object, e As System.Windows.Forms.DragEventArgs) Handles txtInputFile.DragEnter

        If (e.Data.GetDataPresent(DataFormats.FileDrop)) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If

    End Sub

    Private Sub txtInputFile_DragDrop(sender As System.Object, e As System.Windows.Forms.DragEventArgs) Handles txtInputFile.DragDrop

        Try

            Dim arrFiles() As String = e.Data.GetData(DataFormats.FileDrop, False)
            Dim sFile As String = Nothing

            If arrFiles IsNot Nothing Then

                If arrFiles.Length > 0 Then

                    Dim strFile As String = CType(arrFiles.GetValue(arrFiles.GetLowerBound(0)), String)

                    If File.Exists(strFile) = True Then

                        txtInputFile.Text = strFile
                        txtInputFile.Text = sFile

                    End If

                End If

            End If

        Catch

        End Try

    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs)

        Impl_S2ServicesManualFileRepair()

    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs)

        Impl_OnlineFileRepairCouponVisit()

    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs)

        Impl_OpenPayPalDonateUrl()

    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click

        Me.Close()

    End Sub

#End Region

#Region "Implementation: general"

    Private Sub Impl_TryWordFixURL()

        Try

            Dim process As New Process()
            process.StartInfo.FileName = "http://www.cimaware.com/info/info.php?lang=en&id=622&path=wordfix.html"
            process.StartInfo.Arguments = ""
            process.StartInfo.UseShellExecute = True
            process.Start()

        Catch
        End Try

    End Sub

    Private Sub Impl_TryExcelFixUrl()

        Try

            Dim process As New Process()
            process.StartInfo.FileName = "http://www.cimaware.com/info/info.php?lang=en&id=622&path=excelfix.html"
            process.StartInfo.Arguments = ""
            process.StartInfo.UseShellExecute = True
            process.Start()

        Catch
        End Try

    End Sub

    Private Sub Impl_TryOfficeFixUrl()

        Try

            Dim process As New Process()
            process.StartInfo.FileName = "http://www.cimaware.com/info/info.php?lang=en&id=622&path=officefix.html"
            process.StartInfo.Arguments = ""
            process.StartInfo.UseShellExecute = True
            process.Start()

        Catch
        End Try

    End Sub

    Private Sub Impl_OpenHelpUrl()

        Try

            Dim process As New Process()
            process.StartInfo.FileName = "http://legacy.s2services.com/savvy_repair_help.htm"
            process.StartInfo.Arguments = ""
            process.StartInfo.UseShellExecute = True
            process.Start()

        Catch
        End Try

    End Sub

    Private Sub Impl_S2ServicesManualFileRepair()

        Try
            Dim proc As New Process()
            proc.StartInfo.FileName = "http://saveofficedata.com/contact.htm"
            proc.StartInfo.Arguments = ""
            proc.StartInfo.UseShellExecute = True
            proc.Start()
        Catch

        End Try

    End Sub

    Private Sub Impl_OnlineFileRepairCouponVisit()

        Try
            MsgBox("Until Nov 1, 2013, for a free $39 value file repair attempt, first go through " _
       & "Demo recovery with your corrupt file on the Online Office Recovery site that is about to " _
       & "open. After recovery, scroll down past ""Demo Results"" and enter in the coupon code " _
       & """S2SERVICES"" in the field above the ""Submit Code"" button at the end of the ""Full " _
       & "Results"" section. Use all caps for the code but don't include the quotes.", MsgBoxStyle.Information)

            Dim process As New Process()

            process.StartInfo.FileName = "https://online.officerecovery.com/"
            process.StartInfo.Arguments = ""
            process.StartInfo.UseShellExecute = True
            process.Start()

        Catch

        End Try

    End Sub

    Private Sub Impl_OpenPayPalDonateUrl()

        Dim proc As New Process()

        Try

            proc.StartInfo.FileName = "https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=3SJY6GVWUL65S"
            proc.StartInfo.Arguments = ""
            proc.StartInfo.UseShellExecute = True
            proc.Start()

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

    End Sub

    Private Sub Impl_ShadowExplorerURL()

        Try

            Dim process As New Process()
            process.StartInfo.FileName = "http://www.shadowexplorer.com/"
            process.StartInfo.Arguments = ""
            process.StartInfo.UseShellExecute = True
            process.Start()

        Catch
        End Try

    End Sub

    Private Sub Impl_PreviousVersionFileRecovererURL()

        Try

            Dim process As New Process()
            process.StartInfo.FileName = "http://sourceforge.net/projects/vistaprevrsrcvr/"
            process.StartInfo.Arguments = ""
            process.StartInfo.UseShellExecute = True
            process.Start()

        Catch
        End Try

    End Sub

    Private Sub Impl_UnsavedWordFileStepsURL()

        Try

            Dim process As New Process()
            process.StartInfo.FileName = "http://www.makeuseof.com/tag/recover-unsaved-ms-word-2010-document-seconds/"
            process.StartInfo.Arguments = ""
            process.StartInfo.UseShellExecute = True
            process.Start()

        Catch
        End Try

    End Sub

    Private Sub Impl_S2ServiceSourceforgeURL()

        Try

            Dim process As New Process()
            process.StartInfo.FileName = "http://sourceforge.net/users/socrtwo22"
            process.StartInfo.Arguments = ""
            process.StartInfo.UseShellExecute = True
            process.Start()

        Catch
        End Try

    End Sub

    Private Sub Impl_WordStepsURL()

        Try

            Dim process As New Process()
            process.StartInfo.FileName = "http://legacy.s2services.com/word_repair.htm"
            process.StartInfo.Arguments = ""
            process.StartInfo.UseShellExecute = True
            process.Start()

        Catch
        End Try

    End Sub

    Private Sub Impl_ExcelStepsURL()

        Try

            Dim process As New Process()
            process.StartInfo.FileName = "http://legacy.s2services.com/excel.htm"
            process.StartInfo.Arguments = ""
            process.StartInfo.UseShellExecute = True
            process.Start()

        Catch
        End Try

    End Sub

    Private Sub Impl_SaveOfficeData()

        Try

            Dim process As New Process()
            process.StartInfo.FileName = "http://saveofficedata.com/"
            process.StartInfo.Arguments = ""
            process.StartInfo.UseShellExecute = True
            process.Start()

        Catch
        End Try

    End Sub

    Private Sub Impl_RepairMyWord()

        Try

            Dim process As New Process()
            process.StartInfo.FileName = "http://www.repairmyword.com/"
            process.StartInfo.Arguments = ""
            process.StartInfo.UseShellExecute = True
            process.Start()

        Catch
        End Try

    End Sub

    Private Sub Impl_OpenOfficeSteps()

        Try

            Dim process As New Process()
            process.StartInfo.FileName = "http://legacy.s2services.com/open_office.htm"
            process.StartInfo.Arguments = ""
            process.StartInfo.UseShellExecute = True
            process.Start()

        Catch
        End Try

    End Sub

    Private Sub Impl_Silvercoders_DocToText()

        Try

            Dim process As New Process()
            process.StartInfo.FileName = "http://silvercoders.com/en/products/doctotext/"
            process.StartInfo.Arguments = ""
            process.StartInfo.UseShellExecute = True
            process.Start()

        Catch
        End Try

    End Sub

    Private Sub SelectFile()

        Dim dlgOpenFile As New OpenFileDialog

        dlgOpenFile.Filter = "Microsoft 2007-2013 Format Word, Excel or PowerPoint Files " _
            & "(*.docx;*.docm;*.dotx;*.dotm;*.xlsx;*.xlsm;*.xltx;*.xltm;*.xlsb;*.xlam;*.pptx;" _
            & "*.pptm;*.potx;*.potm;*.ppam;*.ppsx;*.ppsm;*.sldx;*.sldm;*.thmx)|*.docx;*.docm;" _
            & "*.dotx;*.dotm;*.xlsx;*.xlsm;*.xltx;*.xltm;*.xlsb;*.xlam;*.pptx;*.pptm;*.potx;" _
            & "*.potm;*.ppam;*.ppsx;*.ppsm;*.sldx;*.sldm;*.thmx|Microsoft Word 2007-2013 Format " _
            & "Documents (*.docx;*.docm;*.dotx;*.dotm)|*.docx;*.docm;*.dotx;*.dotm|Microsoft " _
            & "Excel 2007-2013 Format Spreadsheets (*.xlsx;*.xlsm;*.xltx;*.xltm;*.xlsb;*.xlam)|" _
            & "*.xlsx;*.xlsm;*.xltx;*.xltm;*.xlsb;*.xlam|Microsoft PowerPoint 2007-2013 Format " _
            & "Presentations (*.pptx;*.pptm;*.potx;*.potm;*.ppam;*.ppsx;*.ppsm;*.sldx;*.sldm;" _
            & "*.thmx)|*.pptx;*.pptm;*.potx;*.potm;*.ppam;*.ppsx;*.ppsm;*.sldx;*.sldm;*.thmx|" _
            & "All files (*.*)|*.*"

        dlgOpenFile.RestoreDirectory = True

        If dlgOpenFile.ShowDialog() = DialogResult.OK Then

            txtInputFile.Text = dlgOpenFile.FileName

        End If

    End Sub

    Sub ValidationEventHandler(sender As Object, e As ValidationEventArgs)

        MessageBox.Show(e.Message)

    End Sub

    Private Sub PumpMessages()

        System.Windows.Forms.Application.DoEvents()

    End Sub

    Private Sub UpdateUI(bEnable As Boolean)

        PumpMessages()

        Try

            If bEnable = True Then

                Me.Text = Me.AppName

            Else

                Me.Text = "Processing..."

            End If

            txtInputFile.Enabled = bEnable
            btnAutoRecover.Enabled = bEnable
            btnRecoverMethods.Enabled = bEnable
            menuMain.Enabled = bEnable

        Catch
        End Try

        PumpMessages()

    End Sub

#End Region

#Region "Implementation: methods"

    Private Sub Impl_Recover_Auto1()

        MsgBox("With this automatic recovery choice, four recovery/repair algorthms will be tried: " _
            & "xml sub-file repair with strict XML validation; the same with lax validation; where " _
            & "any missing xml sub-files exist an addition to the corrupt file with analogues from " _
            & "a corresponding blank docx, xlsx or pptx file is made and then the larger file is processed " _
            & "for XML sub-file repair under strcit validation conditions; and finally a text/data " _
            & "salavaging method using SilverCoder's DocToText.", MsgBoxStyle.Information)

        Dim bResult As Boolean

        txtInputFile.Visible = False
        btnBrowse.Visible = False
        progressAuto.Visible = True
        PumpMessages()
        bResult = Impl_Recover_Method_1(True)
        PumpMessages()
        bResult = Impl_Recover_Method_2(True)
        PumpMessages()
        bResult = Impl_Recover_Method_3(True)
        PumpMessages()
        bResult = Impl_Recover_Method_4(True)

        PumpMessages()

        txtInputFile.Visible = True
        btnBrowse.Visible = True
        progressAuto.Visible = False

    End Sub

    Private Function Impl_Recover_Method_1(bSilent As Boolean) As Boolean

        Dim bResult As Boolean = False

        Try

            If String.IsNullOrEmpty(txtInputFile.Text) = True Then

                SelectFile()

            End If

            If String.IsNullOrEmpty(txtInputFile.Text) = False Then

                UpdateUI(False)

                MsgBox("Savvy Recovery for MS Office will now perform XML sub-file repair with " _
                    & "strict XML validation.", MsgBoxStyle.Information)

                Dim filename As String = Nothing
                Dim counterVariable As Integer = Nothing
                Dim previousVersionCounterVariable As Integer = Nothing
                Dim saveShadowPath As String = Nothing
                Dim sFileShadowPath As String = Nothing
                Dim sFileShadowName As String = Nothing
                Dim sFileShadowSize As String = Nothing
                Dim sFileShadowPathDate As String = Nothing
                Dim selectedsFileShadowPathDate As String = Nothing
                Dim selectedsFileShadowPathSize As String = Nothing
                Dim selectedPreviousVersion As String = Nothing
                Dim pathToComboBoxSelectedFile As String = Nothing
                Dim shadowLinkFolderName As New List(Of String)
                Dim nonErrorShadowPathList As New List(Of String)
                Dim xmlFilesCheckedForCorruption As New List(Of String)
                Dim comboBoxIndex As Integer = 0
                Dim matchCount As Integer = 0
                Dim comboBoxChoiceIndex As Integer = 0
                Dim pathToComboBoxSelectedFileSize As Integer = 0
                Dim preVersionHashTable As New Hashtable
                Dim myDocument As New XmlDocument
                Dim sFile As String = txtInputFile.Text
                Dim xmlValidateReader As StreamReader = Nothing
                Dim xmlValidateErrorReader As StreamReader = Nothing
                Dim xmlValidate2Reader As StreamReader = Nothing
                Dim xmlValidate2ErrorReader As StreamReader = Nothing
                Dim byteMatch As Match
                Dim extractedCorruptFileDirInfo As DirectoryInfo
                Dim corruptFileDirRetrievedFilesInfoArray As FileInfo()
                Dim indCorruptFileDirRetrievedFileInfo As FileInfo
                Dim xmlFilesReplacedbyDummyOnes As New List(Of String)
                Dim xmlValidateArguments As String
                Dim xmlValidateCompOut As String
                Dim xmlValidateErrorOut As String
                Dim byteMatchString As String
                Dim byteErrorLocation As String
                Dim truncatedLengthAsString As String
                Dim truncateArguments As String
                Dim truncateFullPath As String
                Dim xmllintFullPath As String
                Dim xmlValidate2CompOut As String
                Dim xmlValidate2ErrorOut As String
                Dim xmlValFullPath As String
                Dim xmlRecoverArguments As String
                Dim zipExtensionXMLRepairedFullPath As String
                Dim xmlRepairedFileName As String
                Dim xmlRepairedFullPath As String
                Dim oldXMLRepairedFileName As String
                Dim oldXMLRepairedFullPath As String
                Dim sevenZipUpArguments As String
                Dim zipRepairedBasePathAndFileNameIndexLastPeriod As Integer
                Dim byteErrorLocationInteger As Integer
                Dim truncatedLength As Integer
                Dim intTruncationAmount As Integer
                Dim x As Integer
                Dim sFileExtension As String
                Dim sFileName As String = LCase(Path.GetFileName(sFile))
                Dim sFileZip As String = sFile & ".zip"
                Dim zipRepairedsFileName As String = "zipRepaired_" & sFileName & ".zip"
                Dim sFileBasePath As String = Path.GetDirectoryName(sFile)
                Dim zipRepairedBasePathAndFileName As String = sFileBasePath & "\" & zipRepairedsFileName
                Dim repairZipArguments As String
                Dim zipFullPath As String
                Dim extractedRepairedZipOutputDirectory As String = _
                    zipRepairedBasePathAndFileName.Remove(zipRepairedBasePathAndFileName.Length - 9)
                Dim sevenZipFullPath As String
                Dim sevenZipExtractArguments As String
                Dim extractedCorruptFileDirPath As String
                Dim individualCorruptFileXMLSubFileName As String

                progressAuto.Value = x
                progressAuto.Minimum = 0
                progressAuto.Maximum = 100
                progressAuto.Visible = True
                progressAuto.Value = 10
                Dim officeRecoveryXMLRepairExecutionPath As String = _
                    Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                sFileExtension = LCase(Path.GetExtension(sFile))

                'First we repair the zip as with the previous button.
                'We start by getting the path from the textbox up top where we loaded the file
                'If the extension is .doc, .xls or ppt, it  
                'probably is not a docx, xlsx or ppts format file.

                PumpMessages()

                If sFileExtension = ".doc" Then

                    MsgBox("XML repair is only useful for files with docx, .xslx or pptx extensions " _
                           & "and format. Your file may still be in reality with a docx, xlsx or pptx " _
                           & "format. Try changing the the extension.", MsgBoxStyle.Exclamation)
                    Return bResult

                    Exit Function

                End If

                If sFileExtension = ".xls" Then

                    MsgBox("XML repair is only useful for files with docx, .xslx or pptx extensions " _
                           & "and format. Your file may still be in reality with a docx, xlsx or pptx " _
                           & "format. Try changing the the extension.", MsgBoxStyle.Exclamation)
                    Return bResult

                    Exit Function

                End If

                If sFileExtension = ".ppt" Then

                    MsgBox("XML repair is only useful for files with docx, .xslx or pptx extensions " _
                           & "and format. Your file may still be in reality with a docx, xlsx or pptx " _
                           & "format. Try changing the the extension.", MsgBoxStyle.Exclamation)
                    Return bResult

                    Exit Function

                End If

                If File.Exists(sFileZip) Then

                    File.Delete(sFileZip)

                End If

                File.Copy(sFile, sFileZip, True)

                If File.Exists(zipRepairedBasePathAndFileName) Then

                    File.Delete(zipRepairedBasePathAndFileName)

                End If

                zipFullPath = """" & Path.Combine(officeRecoveryXMLRepairExecutionPath, "zip.exe") & """"
                repairZipArguments = "-FF """ & sFileZip & """ --out """ _
                        & zipRepairedBasePathAndFileName & """"
                PumpMessages()

                Using repairZip As Process = New Process

                    repairZip.StartInfo.FileName = zipFullPath
                    repairZip.StartInfo.Arguments = repairZipArguments
                    repairZip.StartInfo.UseShellExecute = False
                    repairZip.StartInfo.CreateNoWindow = True
                    repairZip.Start()
                    repairZip.WaitForExit()
                    repairZip.Close()

                End Using

                File.Delete(sFileZip)
                progressAuto.Value = 20

                'Now we extract the repaired file.

                zipRepairedBasePathAndFileNameIndexLastPeriod = zipRepairedBasePathAndFileName.LastIndexOf(".")
                extractedRepairedZipOutputDirectory = _
                    zipRepairedBasePathAndFileName.Remove(zipRepairedBasePathAndFileNameIndexLastPeriod - 5)
                sevenZipFullPath = """" & Path.Combine(officeRecoveryXMLRepairExecutionPath, "7z.exe") & """"
                sevenZipExtractArguments = "x """ & zipRepairedBasePathAndFileName & """ -o""" & _
                                                        extractedRepairedZipOutputDirectory & """"
                PumpMessages()

                Using extractZip As Process = New Process

                    extractZip.StartInfo.FileName = sevenZipFullPath
                    extractZip.StartInfo.Arguments = sevenZipExtractArguments
                    extractZip.StartInfo.UseShellExecute = False
                    extractZip.StartInfo.CreateNoWindow = True
                    extractZip.StartInfo.RedirectStandardInput = True
                    extractZip.StartInfo.RedirectStandardOutput = True
                    extractZip.StartInfo.RedirectStandardError = True
                    extractZip.Start()
                    extractZip.StandardInput.WriteLine("u")
                    Dim extractZipReader As StreamReader = extractZip.StandardOutput
                    Dim extractZipReaderCompOut As String = extractZipReader.ReadToEnd
                    Dim extractZipErrorReader As StreamReader = extractZip.StandardError
                    Dim extractZipErrorReaderOut As String = extractZipErrorReader.ReadToEnd
                    extractZip.WaitForExit()
                    extractZip.Close()

                End Using

                progressAuto.Value = 30
                File.Delete(zipRepairedBasePathAndFileName)
                extractedCorruptFileDirPath = extractedRepairedZipOutputDirectory & "\"
                extractedCorruptFileDirInfo = New DirectoryInfo(extractedCorruptFileDirPath)
                corruptFileDirRetrievedFilesInfoArray = _
                    extractedCorruptFileDirInfo.GetFiles("*.*", SearchOption.AllDirectories)

                Dim xmlFileNamesInTheCorruptFile As New List(Of String)

                'We process the FileInfo for each of the the files on the list retrieved by GetFiles.
                'If the file ends in extensions .xml or .rels then they are XML files that can be repaired
                'by xmllint and I add them to the list called  xmlFileNamesInTheCorruptFile.

                PumpMessages()

                If corruptFileDirRetrievedFilesInfoArray IsNot Nothing Then

                    For Each indCorruptFileDirRetrievedFileInfo In corruptFileDirRetrievedFilesInfoArray

                        If InStr(indCorruptFileDirRetrievedFileInfo.Extension, ".xml") Or _
                            InStr(indCorruptFileDirRetrievedFileInfo.Extension, ".rels") Then

                            individualCorruptFileXMLSubFileName = indCorruptFileDirRetrievedFileInfo.Name
                            xmlFileNamesInTheCorruptFile.Add(individualCorruptFileXMLSubFileName)

                        End If

                    Next

                Else

                    MsgBox("Your loaded corrupt DOCX, XLSX or PPTX file is missing any " _
                        & "XML sub-files and is unrecoverable unless it is in reality an " _
                        & "old format doc, xls, or ppt format file. Try changing the extension " _
                        & "and using another recovery program suitable for those formats.")

                    Return bResult

                    Exit Function

                End If

                progressAuto.Value = 40

                'Here we process each .xml or .rels file by validating it first. If it is invalid,
                'we truncate it at the error and repair with xmllint. We try first adding no bits
                'extra to truncate, if the the file after treating with xmllint does not validate,
                'we remove 50 bits, then try 100 and running through the xmlrepai again.

                extractedCorruptFileDirInfo = New DirectoryInfo(extractedCorruptFileDirPath)
                corruptFileDirRetrievedFilesInfoArray = _
                    extractedCorruptFileDirInfo.GetFiles("*.*", SearchOption.AllDirectories)

                Dim corruptFileDirRetrievedFileInfo As FileInfo
                Dim extractedCorruptFileDirInfoArrayCount As Integer = _
                    corruptFileDirRetrievedFilesInfoArray.GetLength(0)

                Dim progressBarIncrement As Integer = 50 \ (extractedCorruptFileDirInfoArrayCount + 1)

                If progressBarIncrement = 0 Then

                    progressBarIncrement = 1

                End If

                PumpMessages()

                For Each corruptFileDirRetrievedFileInfo In corruptFileDirRetrievedFilesInfoArray

                    If InStr(corruptFileDirRetrievedFileInfo.Extension, ".xml") Or _
                        InStr(corruptFileDirRetrievedFileInfo.Extension, ".rels") Then

                        If progressAuto.Value > 90 Then

                            progressAuto.Value = 90

                        Else

                            progressAuto.Value = progressAuto.Value + progressBarIncrement

                        End If

                        PumpMessages()

                        Dim corruptFileDirRetrievedXMLFileName As String = corruptFileDirRetrievedFileInfo.Name
                        Dim corruptFileDirRetrievedXMLFileFullPath As String = corruptFileDirRetrievedFileInfo.FullName
                        Dim extractedCorruptFileDirPathCharacterCount As Integer = _
                            extractedRepairedZipOutputDirectory.Length
                        Dim corruptFileDirRetrievedXMLFileFullPathCharacterCount As Integer = _
                            corruptFileDirRetrievedXMLFileFullPath.Length
                        Dim corruptFileDirRetrievedXMLFileWithinArchivePathCharacterCount As Integer = _
                            corruptFileDirRetrievedXMLFileFullPathCharacterCount - _
                            extractedCorruptFileDirPathCharacterCount
                        Dim corruptFileDirRetrievedXMLFileWithinArchivePath As String = _
                            corruptFileDirRetrievedXMLFileFullPath.Substring(extractedCorruptFileDirPathCharacterCount, _
                            corruptFileDirRetrievedXMLFileWithinArchivePathCharacterCount)

                        xmlValFullPath = """" & Path.Combine(officeRecoveryXMLRepairExecutionPath, _
                            "xml.exe") & """"
                        xmlValidateArguments = "val -e """ & corruptFileDirRetrievedXMLFileFullPath & """"

                        Using xmlValidate As Process = New Process

                            xmlValidate.StartInfo.FileName = xmlValFullPath
                            xmlValidate.StartInfo.Arguments = xmlValidateArguments
                            xmlValidate.StartInfo.UseShellExecute = False
                            xmlValidate.StartInfo.RedirectStandardOutput = True
                            xmlValidate.StartInfo.RedirectStandardError = True
                            xmlValidate.StartInfo.CreateNoWindow = True
                            xmlValidate.Start()
                            xmlValidateReader = xmlValidate.StandardOutput
                            xmlValidateCompOut = xmlValidateReader.ReadToEnd
                            xmlValidateErrorReader = xmlValidate.StandardError
                            xmlValidateErrorOut = xmlValidateErrorReader.ReadToEnd
                            xmlValidate.WaitForExit()
                            xmlValidate.Close()

                        End Using

                        If xmlValidateCompOut.Contains("invalid") Then

                            Dim loopCounter As Integer = 0

                            Do

                                'The validator will register an error and indicate the byte location of 
                                'the error if the document.xml file has an error. We isolate this byte 
                                'location with a Regex and the DelFromleft function, changed the byte to 
                                'an integer and subtract first 0 then 50 then 100 bytes to try to steer 
                                'clear of any additional bad xml if there is some just before the error.

                                loopCounter = loopCounter + 1

                                If loopCounter = 1 Then

                                    intTruncationAmount = 0

                                ElseIf loopCounter = 2 Then

                                    intTruncationAmount = 50

                                Else

                                    intTruncationAmount = 100

                                End If

                                byteMatch = Regex.Match(xmlValidateErrorOut, _
                                 (":2.[0-9]+"))
                                byteMatchString = byteMatch.ToString
                                byteErrorLocation = DelFromLeft(":2.", byteMatchString)
                                Integer.TryParse(byteErrorLocation, byteErrorLocationInteger)

                                truncatedLength = byteErrorLocationInteger - intTruncationAmount
                                truncatedLengthAsString = String.Empty
                                truncatedLengthAsString = System.Convert.ToString(truncatedLength)
                                truncateArguments = """" & corruptFileDirRetrievedXMLFileFullPath & """ " & truncatedLengthAsString
                                truncateFullPath = """" & _
                                    Path.Combine(officeRecoveryXMLRepairExecutionPath, "trunc.exe") & """"

                                'Now we will truncate the file at bad byte minus 0, 50 or 100 bytes.

                                PumpMessages()

                                Using truncate As Process = New Process

                                    truncate.StartInfo.FileName = truncateFullPath
                                    truncate.StartInfo.Arguments = truncateArguments
                                    truncate.StartInfo.UseShellExecute = False
                                    truncate.StartInfo.CreateNoWindow = True
                                    truncate.Start()
                                    truncate.WaitForExit()
                                    truncate.Close()

                                End Using

                                xmlRecoverArguments = "--recover " & _
                                    """" & corruptFileDirRetrievedXMLFileFullPath & """" & " -o " & _
                                    """" & corruptFileDirRetrievedXMLFileFullPath & """"
                                xmllintFullPath = """" & Path.Combine(officeRecoveryXMLRepairExecutionPath, _
                                                "xmllint.exe") & """"

                                'Now we use xmllint to reconstruct the nice xml ending tags
                                'to try to slip document.xml past Word's XML validator.

                                PumpMessages()

                                Using xmlRecover As Process = New Process

                                    xmlRecover.StartInfo.FileName = xmllintFullPath
                                    xmlRecover.StartInfo.Arguments = xmlRecoverArguments
                                    xmlRecover.StartInfo.UseShellExecute = False
                                    xmlRecover.StartInfo.CreateNoWindow = True
                                    xmlRecover.Start()
                                    xmlRecover.WaitForExit()
                                    xmlRecover.Close()

                                End Using

                                'We validate again, hoping our xmlRecovery is OK by our validator.

                                PumpMessages()

                                Using xmlValidate2 As Process = New Process

                                    xmlValidate2.StartInfo.FileName = xmlValFullPath
                                    xmlValidate2.StartInfo.Arguments = xmlValidateArguments
                                    xmlValidate2.StartInfo.UseShellExecute = False
                                    xmlValidate2.StartInfo.RedirectStandardOutput = True
                                    xmlValidate2.StartInfo.RedirectStandardError = True
                                    xmlValidate2.StartInfo.CreateNoWindow = True
                                    xmlValidate2.Start()
                                    xmlValidate2Reader = xmlValidate2.StandardOutput
                                    xmlValidate2CompOut = xmlValidate2Reader.ReadToEnd
                                    xmlValidate2ErrorReader = xmlValidate2.StandardError
                                    xmlValidate2ErrorOut = xmlValidate2ErrorReader.ReadToEnd
                                    xmlValidate2.WaitForExit()
                                    xmlValidate2.Close()

                                End Using

                                'If our validator says the XML is still bad, we 
                                'go through a 2nd of truncation and XML recovery

                                If loopCounter = 3 Then

                                    Exit Do

                                End If

                            Loop Until Not xmlValidate2CompOut.Contains("invalid")

                        End If

                    End If

                Next

                'Now we will rezip our directory and open as a docx, xlsx or pptx file.

                xmlRepairedFileName = "xml_repaired_" & sFileName
                xmlRepairedFullPath = sFileBasePath & "\" & xmlRepairedFileName

                If File.Exists(xmlRepairedFullPath) Then

                    oldXMLRepairedFileName = "xml_repaired_old_" & sFileName
                    oldXMLRepairedFullPath = sFileBasePath & "\" & oldXMLRepairedFileName
                    File.Copy(xmlRepairedFullPath, oldXMLRepairedFullPath, True)
                    File.Delete(xmlRepairedFullPath)

                End If

                zipExtensionXMLRepairedFullPath = xmlRepairedFullPath & ".zip"
                sevenZipUpArguments = "a -r """ & zipExtensionXMLRepairedFullPath & """ """ & _
                       extractedRepairedZipOutputDirectory & """\*"

                Using sevenZipReZip As Process = New Process

                    sevenZipReZip.StartInfo.FileName = sevenZipFullPath
                    sevenZipReZip.StartInfo.Arguments = sevenZipUpArguments
                    sevenZipReZip.StartInfo.UseShellExecute = False
                    sevenZipReZip.StartInfo.RedirectStandardOutput = True
                    sevenZipReZip.StartInfo.RedirectStandardError = True
                    sevenZipReZip.StartInfo.CreateNoWindow = True
                    sevenZipReZip.Start()
                    Dim sevenZipReZipReader As StreamReader = sevenZipReZip.StandardOutput
                    Dim sevenZipReZipReaderCompOut As String = sevenZipReZipReader.ReadToEnd
                    Dim sevenZipReZipErrorReader As StreamReader = sevenZipReZip.StandardError
                    Dim sevenZipReZipErrorReaderOut As String = sevenZipReZipErrorReader.ReadToEnd
                    sevenZipReZip.WaitForExit()
                    sevenZipReZip.Close()

                End Using

                progressAuto.Value = 95

                File.Copy(zipExtensionXMLRepairedFullPath, xmlRepairedFullPath, True)

                Using Process = New Process

                    Process.StartInfo.UseShellExecute = False
                    Process.StartInfo.RedirectStandardOutput = True
                    Process.StartInfo.CreateNoWindow = True
                    Process.Start(xmlRepairedFullPath)
                    Process.Close()

                End Using

                File.Delete(zipExtensionXMLRepairedFullPath)
                Directory.Delete(extractedRepairedZipOutputDirectory, True)

                progressAuto.Value = 100
                progressAuto.Visible = False

            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

        UpdateUI(True)

        Return bResult

    End Function

    Private Function Impl_Recover_Method_2(bSilent As Boolean) As Boolean

        MsgBox("Savvy Recovery for MS Office will now perform XML sub-file repair with " _
            & "lax XML validation.", MsgBoxStyle.Information)
        Dim bResult As Boolean = False

        Try

            If String.IsNullOrEmpty(txtInputFile.Text) = True Then

                SelectFile()

            End If

            If String.IsNullOrEmpty(txtInputFile.Text) = False Then

                UpdateUI(False)

                Dim comboBoxIndex As Integer = 0
                Dim matchCount As Integer = 0
                Dim comboBoxChoiceIndex As Integer = 0
                Dim pathToComboBoxSelectedFileSize As Integer = 0
                Dim byteErrorLocationInteger As Integer = 0
                Dim truncatedLength As Integer = 0
                Dim intTruncationAmount As Integer = 0
                Dim x As Integer = 0
                Dim preVersionHashTable As New Hashtable
                Dim myDocument As New XmlDocument
                Dim xmlValidateReader As StreamReader = Nothing
                Dim xmlValidateErrorReader As StreamReader = Nothing
                Dim xmlValidate2Reader As StreamReader = Nothing
                Dim xmlValidate2ErrorReader As StreamReader = Nothing
                Dim byteMatch As Match = Nothing
                Dim extractedCorruptFileDirInfo As DirectoryInfo = Nothing
                Dim corruptFileDirRetrievedFilesInfoArray As FileInfo() = Nothing
                Dim indCorruptFileDirRetrievedFileInfo As FileInfo = Nothing
                Dim xmlFilesReplacedbyDummyOnes As New List(Of String)
                Dim shadowLinkFolderName As New List(Of String)
                Dim nonErrorShadowPathList As New List(Of String)
                Dim xmlFilesCheckedForCorruption As New List(Of String)
                Dim sFile As String = txtInputFile.Text
                Dim officeRecoveryXMLRepairExecutionPath As String = Nothing
                Dim xmlValidateArguments As String = Nothing
                Dim xmlValidateCompOut As String = Nothing
                Dim xmlValidateErrorOut As String = Nothing
                Dim byteMatchString As String = Nothing
                Dim byteErrorLocation As String = Nothing
                Dim truncatedLengthAsString As String = Nothing
                Dim truncateArguments As String = Nothing
                Dim truncateFullPath As String = Nothing
                Dim xmllintFullPath As String = Nothing
                Dim xmlValidate2CompOut As String = Nothing
                Dim xmlValidate2ErrorOut As String = Nothing
                Dim xmlValFullPath As String = Nothing
                Dim xmlRecoverArguments As String = Nothing
                Dim zipExtensionXMLRepairedFullPath As String = Nothing
                Dim xmlRepairedFileName As String = Nothing
                Dim xmlRepairedFullPath As String = Nothing
                Dim oldXMLRepairedFileName As String = Nothing
                Dim oldXMLRepairedFullPath As String = Nothing
                Dim sevenZipUpArguments As String = Nothing
                Dim zipRepairedBasePathAndFileNameIndexLastPeriod As Integer = Nothing
                Dim sFileExtension As String = Nothing
                Dim sFileName As String = LCase(Path.GetFileName(sFile))
                Dim sFileZip As String = sFile & "_1b.zip"
                Dim zipRepairedsFileName As String = "zipRepaired_1b_" & sFileName & ".zip"
                Dim sFileBasePath As String = Path.GetDirectoryName(sFile)
                Dim zipRepairedBasePathAndFileName As String = sFileBasePath & "\" & zipRepairedsFileName
                Dim repairZipArguments As String = Nothing
                Dim zipFullPath As String = Nothing
                Dim extractedRepairedZipOutputDirectory As String = _
                    zipRepairedBasePathAndFileName.Remove(zipRepairedBasePathAndFileName.Length - 9)
                Dim sevenZipFullPath As String = Nothing
                Dim sevenZipExtractArguments As String = Nothing
                Dim extractedCorruptFileDirPath As String = Nothing
                Dim individualCorruptFileXMLSubFileName As String = Nothing

                progressAuto.Value = x
                progressAuto.Minimum = 0
                progressAuto.Maximum = 100
                progressAuto.Visible = True
                progressAuto.Value = 10
                officeRecoveryXMLRepairExecutionPath = _
                    Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                sFileExtension = LCase(Path.GetExtension(sFile))

                'First we repair the zip as with the previous button.
                'We start by getting the path from the textbox up top where we loaded the file
                'If the extension is .doc, .xls or ppt, it  
                'probably is not a docx, xlsx or ppts format file.

                If sFileExtension = ".doc" Then

                    MsgBox("XML repair is only useful for files with docx, .xslx or pptx extensions " _
                           & "and format. Your file may still be in reality with a docx, xlsx or pptx " _
                           & "format. Try changing the the extension.", MsgBoxStyle.Exclamation)

                    Return bResult

                    Exit Function

                End If

                If sFileExtension = ".xls" Then

                    MsgBox("XML repair is only useful for files with docx, .xslx or pptx extensions " _
                           & "and format. Your file may still be in reality with a docx, xlsx or pptx " _
                           & "format. Try changing the the extension.", MsgBoxStyle.Exclamation)

                    Return bResult

                    Exit Function

                End If

                If sFileExtension = ".ppt" Then

                    MsgBox("XML repair is only useful for files with docx, .xslx or pptx extensions " _
                           & "and format. Your file may still be in reality with a docx, xlsx or pptx " _
                           & "format. Try changing the the extension.", MsgBoxStyle.Exclamation)

                    Return bResult

                    Exit Function

                End If

                If File.Exists(sFileZip) Then

                    File.Delete(sFileZip)

                End If

                File.Copy(sFile, sFileZip, True)

                If File.Exists(zipRepairedBasePathAndFileName) Then

                    File.Delete(zipRepairedBasePathAndFileName)

                End If

                zipFullPath = """" & Path.Combine(officeRecoveryXMLRepairExecutionPath, "zip.exe") & """"
                repairZipArguments = "-FF """ & sFileZip & """ --out """ _
                        & zipRepairedBasePathAndFileName & """"

                Using repairZip As Process = New Process

                    repairZip.StartInfo.FileName = zipFullPath
                    repairZip.StartInfo.Arguments = repairZipArguments
                    repairZip.StartInfo.UseShellExecute = False
                    repairZip.StartInfo.CreateNoWindow = True
                    repairZip.Start()
                    repairZip.WaitForExit()
                    repairZip.Close()

                End Using

                File.Delete(sFileZip)
                progressAuto.Value = 20

                'Now we extract the repaired file.

                zipRepairedBasePathAndFileNameIndexLastPeriod = zipRepairedBasePathAndFileName.LastIndexOf(".")
                extractedRepairedZipOutputDirectory = _
                    zipRepairedBasePathAndFileName.Remove(zipRepairedBasePathAndFileNameIndexLastPeriod - 5)
                sevenZipFullPath = """" & Path.Combine(officeRecoveryXMLRepairExecutionPath, "7z.exe") & """"
                sevenZipExtractArguments = "x """ & zipRepairedBasePathAndFileName & """ -o""" & _
                                                        extractedRepairedZipOutputDirectory & """"

                Using extractZip As Process = New Process

                    extractZip.StartInfo.FileName = sevenZipFullPath
                    extractZip.StartInfo.Arguments = sevenZipExtractArguments
                    extractZip.StartInfo.UseShellExecute = False
                    extractZip.StartInfo.CreateNoWindow = True
                    extractZip.StartInfo.RedirectStandardInput = True
                    extractZip.StartInfo.RedirectStandardOutput = True
                    extractZip.StartInfo.RedirectStandardError = True
                    extractZip.Start()
                    extractZip.StandardInput.WriteLine("u")
                    Dim extractZipReader As StreamReader = extractZip.StandardOutput
                    Dim extractZipReaderCompOut As String = extractZipReader.ReadToEnd
                    Dim extractZipErrorReader As StreamReader = extractZip.StandardError
                    Dim extractZipErrorReaderOut As String = extractZipErrorReader.ReadToEnd
                    extractZip.WaitForExit()
                    extractZip.Close()

                End Using

                progressAuto.Value = 30
                File.Delete(zipRepairedBasePathAndFileName)
                extractedCorruptFileDirPath = extractedRepairedZipOutputDirectory & "\"
                extractedCorruptFileDirInfo = New DirectoryInfo(extractedCorruptFileDirPath)
                corruptFileDirRetrievedFilesInfoArray = _
                    extractedCorruptFileDirInfo.GetFiles("*.*", SearchOption.AllDirectories)

                Dim xmlFileNamesInTheCorruptFile As New List(Of String)

                'We process the FileInfo for each of the the files on the list retrieved by GetFiles.
                'If the file ends in extensions .xml or .rels then they are XML files that can be repaired
                'by xmllint and I add them to the list called  xmlFileNamesInTheCorruptFile.

                If corruptFileDirRetrievedFilesInfoArray IsNot Nothing Then

                    For Each indCorruptFileDirRetrievedFileInfo In corruptFileDirRetrievedFilesInfoArray

                        If InStr(indCorruptFileDirRetrievedFileInfo.Extension, ".xml") Or _
                            InStr(indCorruptFileDirRetrievedFileInfo.Extension, ".rels") Then

                            individualCorruptFileXMLSubFileName = indCorruptFileDirRetrievedFileInfo.Name
                            xmlFileNamesInTheCorruptFile.Add(individualCorruptFileXMLSubFileName)

                        End If

                    Next

                Else

                    MsgBox("Your loaded corrupt DOCX, XLSX or PPTX file is missing " _
                        & "any XML sub-files and is unrecoverable unless it is in reality " _
                        & "an old format doc, xls, or ppt format file. Try changing the extension " _
                        & "and using another recovery program suitable for those formats.")

                    Return bResult

                    Exit Function

                End If

                progressAuto.Value = 40

                'Here we process each .xml or .rels file by validating it first. If it is invalid,
                'we truncate it at the error and repair with xmllint. We try first adding no bits
                'extra to truncate, if the the file after treating with xmllint does not validate,
                'we remove 50 bits, then try 100 and running through the xmlrepai again.

                extractedCorruptFileDirInfo = New DirectoryInfo(extractedCorruptFileDirPath)
                corruptFileDirRetrievedFilesInfoArray = _
                    extractedCorruptFileDirInfo.GetFiles("*.*", SearchOption.AllDirectories)

                Dim corruptFileDirRetrievedFileInfo As FileInfo
                Dim extractedCorruptFileDirInfoArrayCount As Integer = _
                    corruptFileDirRetrievedFilesInfoArray.GetLength(0)

                Dim progressBarIncrement As Integer = 50 \ (extractedCorruptFileDirInfoArrayCount + 1)

                If progressBarIncrement = 0 Then

                    progressBarIncrement = 1

                End If

                For Each corruptFileDirRetrievedFileInfo In corruptFileDirRetrievedFilesInfoArray

                    If InStr(corruptFileDirRetrievedFileInfo.Extension, ".xml") Or _
                        InStr(corruptFileDirRetrievedFileInfo.Extension, ".rels") Then

                        If progressAuto.Value > 90 Then

                            progressAuto.Value = 90

                        Else

                            progressAuto.Value = progressAuto.Value + progressBarIncrement

                        End If

                        Dim corruptFileDirRetrievedXMLFileName As String = corruptFileDirRetrievedFileInfo.Name
                        Dim corruptFileDirRetrievedXMLFileFullPath As String = corruptFileDirRetrievedFileInfo.FullName
                        Dim extractedCorruptFileDirPathCharacterCount As Integer = _
                            extractedRepairedZipOutputDirectory.Length
                        Dim corruptFileDirRetrievedXMLFileFullPathCharacterCount As Integer = _
                            corruptFileDirRetrievedXMLFileFullPath.Length
                        Dim corruptFileDirRetrievedXMLFileWithinArchivePathCharacterCount As Integer = _
                            corruptFileDirRetrievedXMLFileFullPathCharacterCount - _
                            extractedCorruptFileDirPathCharacterCount
                        Dim corruptFileDirRetrievedXMLFileWithinArchivePath As String = _
                            corruptFileDirRetrievedXMLFileFullPath.Substring(extractedCorruptFileDirPathCharacterCount, _
                            corruptFileDirRetrievedXMLFileWithinArchivePathCharacterCount)

                        xmlValFullPath = """" & Path.Combine(officeRecoveryXMLRepairExecutionPath, _
                            "xmlval.exe") & """"
                        xmlValidateArguments = """" & corruptFileDirRetrievedXMLFileFullPath & """"

                        Using xmlValidate As Process = New Process

                            xmlValidate.StartInfo.FileName = xmlValFullPath
                            xmlValidate.StartInfo.Arguments = xmlValidateArguments
                            xmlValidate.StartInfo.UseShellExecute = False
                            xmlValidate.StartInfo.RedirectStandardOutput = True
                            xmlValidate.StartInfo.RedirectStandardError = True
                            xmlValidate.StartInfo.CreateNoWindow = True
                            xmlValidate.Start()
                            xmlValidateReader = xmlValidate.StandardOutput
                            xmlValidateCompOut = xmlValidateReader.ReadToEnd
                            xmlValidateErrorReader = xmlValidate.StandardError
                            xmlValidateErrorOut = xmlValidateErrorReader.ReadToEnd
                            xmlValidate.WaitForExit()
                            xmlValidate.Close()

                        End Using

                        If xmlValidateErrorOut.Contains("byte") Then

                            Dim loopCounter As Integer = 0

                            Do

                                'The validator will register an error and indicate the byte location of 
                                'the error if the document.xml file has an error. We isolate this byte 
                                'location with a Regex and the DelFromleft function, changed the byte to 
                                'an integer and subtract first 0 then 50 then 100 bytes to try to steer 
                                'clear of any additional bad xml if there is some just before the error.

                                loopCounter = loopCounter + 1

                                If loopCounter = 1 Then

                                    intTruncationAmount = 22

                                ElseIf loopCounter = 2 Then

                                    intTruncationAmount = 50

                                Else

                                    intTruncationAmount = 100

                                End If

                                byteMatch = Regex.Match(xmlValidateErrorOut, _
                                 "byte [0-9]+")
                                byteMatchString = byteMatch.ToString
                                byteErrorLocation = DelFromLeft("byte ", byteMatchString)
                                Integer.TryParse(byteErrorLocation, byteErrorLocationInteger)

                                truncatedLength = byteErrorLocationInteger - intTruncationAmount
                                truncatedLengthAsString = String.Empty
                                truncatedLengthAsString = System.Convert.ToString(truncatedLength)
                                truncateArguments = """" & corruptFileDirRetrievedXMLFileFullPath & """ " & truncatedLengthAsString
                                truncateFullPath = """" & _
                                    Path.Combine(officeRecoveryXMLRepairExecutionPath, "trunc.exe") & """"

                                'Now we will truncate the file at bad byte minus 0, 50 or 100 bytes.

                                Using truncate As Process = New Process

                                    truncate.StartInfo.FileName = truncateFullPath
                                    truncate.StartInfo.Arguments = truncateArguments
                                    truncate.StartInfo.UseShellExecute = False
                                    truncate.StartInfo.CreateNoWindow = True
                                    truncate.Start()
                                    truncate.WaitForExit()
                                    truncate.Close()

                                End Using

                                xmlRecoverArguments = "--recover " & _
                                    """" & corruptFileDirRetrievedXMLFileFullPath & """" & " -o " & _
                                    """" & corruptFileDirRetrievedXMLFileFullPath & """"
                                xmllintFullPath = """" & Path.Combine(officeRecoveryXMLRepairExecutionPath, _
                                                "xmllint.exe") & """"

                                'Now we use xmllint to reconstruct the nice xml ending tags
                                'to try to slip document.xml past Word's XML validator.

                                Using xmlRecover As Process = New Process

                                    xmlRecover.StartInfo.FileName = xmllintFullPath
                                    xmlRecover.StartInfo.Arguments = xmlRecoverArguments
                                    xmlRecover.StartInfo.UseShellExecute = False
                                    xmlRecover.StartInfo.CreateNoWindow = True
                                    xmlRecover.Start()
                                    xmlRecover.WaitForExit()
                                    xmlRecover.Close()

                                End Using

                                'We validate again, hoping our xmlRecovery is OK by our validator.

                                Using xmlValidate2 As Process = New Process

                                    xmlValidate2.StartInfo.FileName = xmlValFullPath
                                    xmlValidate2.StartInfo.Arguments = xmlValidateArguments
                                    xmlValidate2.StartInfo.UseShellExecute = False
                                    xmlValidate2.StartInfo.RedirectStandardOutput = True
                                    xmlValidate2.StartInfo.RedirectStandardError = True
                                    xmlValidate2.StartInfo.CreateNoWindow = True
                                    xmlValidate2.Start()
                                    xmlValidate2Reader = xmlValidate2.StandardOutput
                                    xmlValidate2CompOut = xmlValidate2Reader.ReadToEnd
                                    xmlValidate2ErrorReader = xmlValidate2.StandardError
                                    xmlValidate2ErrorOut = xmlValidate2ErrorReader.ReadToEnd
                                    xmlValidate2.WaitForExit()
                                    xmlValidate2.Close()

                                End Using

                                'If our validator says the XML is still bad, we 
                                'go through a 2nd of truncation and XML recovery

                                If loopCounter = 3 Then

                                    Exit Do

                                End If

                            Loop Until Not xmlValidateErrorOut.Contains("byte")

                        End If

                    End If

                Next

                'Now we will rezip our directory and open as a docx, xlsx or pptx file.

                xmlRepairedFileName = "lax_xml_repaired_" & sFileName
                xmlRepairedFullPath = sFileBasePath & "\" & xmlRepairedFileName

                If File.Exists(xmlRepairedFullPath) Then

                    oldXMLRepairedFileName = "lax_xml_repaired_old_" & sFileName
                    oldXMLRepairedFullPath = sFileBasePath & "\" & oldXMLRepairedFileName
                    File.Copy(xmlRepairedFullPath, oldXMLRepairedFullPath, True)
                    File.Delete(xmlRepairedFullPath)

                End If

                zipExtensionXMLRepairedFullPath = xmlRepairedFullPath & ".zip"
                sevenZipUpArguments = "a -r """ & zipExtensionXMLRepairedFullPath & """ """ & _
                       extractedRepairedZipOutputDirectory & """\*"

                Using sevenZipReZip As Process = New Process

                    sevenZipReZip.StartInfo.FileName = sevenZipFullPath
                    sevenZipReZip.StartInfo.Arguments = sevenZipUpArguments
                    sevenZipReZip.StartInfo.UseShellExecute = False
                    sevenZipReZip.StartInfo.RedirectStandardOutput = True
                    sevenZipReZip.StartInfo.RedirectStandardError = True
                    sevenZipReZip.StartInfo.CreateNoWindow = True
                    sevenZipReZip.Start()
                    Dim sevenZipReZipReader As StreamReader = sevenZipReZip.StandardOutput
                    Dim sevenZipReZipReaderCompOut As String = sevenZipReZipReader.ReadToEnd
                    Dim sevenZipReZipErrorReader As StreamReader = sevenZipReZip.StandardError
                    Dim sevenZipReZipErrorReaderOut As String = sevenZipReZipErrorReader.ReadToEnd
                    sevenZipReZip.WaitForExit()
                    sevenZipReZip.Close()

                End Using

                progressAuto.Value = 95

                File.Copy(zipExtensionXMLRepairedFullPath, xmlRepairedFullPath, True)

                Using officeFileHandler As Process = New Process

                    Process.Start(xmlRepairedFullPath)

                End Using

                File.Delete(zipExtensionXMLRepairedFullPath)
                Directory.Delete(extractedRepairedZipOutputDirectory, True)

                progressAuto.Value = 100
                progressAuto.Visible = False

            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

        UpdateUI(True)

        Return bResult

    End Function

    Private Function Impl_Recover_Method_3(bSilent As Boolean) As Boolean

        MsgBox("Savvy Recovery for MS Office will now perform XML sub-file repair with " _
            & "strict XML validation and the addition of missing XML sub-files from " _
            & "a corresponding blank one.", MsgBoxStyle.Information)

        Dim bResult As Boolean = False

        Try

            If String.IsNullOrEmpty(txtInputFile.Text) = True Then

                SelectFile()

            End If

            If String.IsNullOrEmpty(txtInputFile.Text) = False Then

                UpdateUI(False)

                Dim shadowLinkFolderName As New List(Of String)
                Dim nonErrorShadowPathList As New List(Of String)
                Dim xmlFilesCheckedForCorruption As New List(Of String)
                Dim comboBoxIndex As Integer = 0
                Dim matchCount As Integer = 0
                Dim comboBoxChoiceIndex As Integer = 0
                Dim pathToComboBoxSelectedFileSize As Integer = 0
                Dim preVersionHashTable As New Hashtable
                Dim myDocument As New XmlDocument
                Dim xmlValidateReader As StreamReader = Nothing
                Dim xmlValidateErrorReader As StreamReader = Nothing
                Dim xmlValidate2Reader As StreamReader = Nothing
                Dim xmlValidate2ErrorReader As StreamReader = Nothing
                Dim byteMatch As Match
                Dim extractedCorruptFileDirInfo As DirectoryInfo
                Dim corruptFileDirRetrievedFilesInfoArray As FileInfo()
                Dim indCorruptFileDirRetrievedFileInfo As FileInfo
                Dim xmlFilesReplacedbyDummyOnes As New List(Of String)
                Dim sFile As String = txtInputFile.Text
                Dim officeRecoveryXMLRepairExecutionPath As String = Nothing
                Dim sevenZipFullPath As String = Nothing
                Dim sevenZipExtractArguments As String = Nothing
                Dim extractedCorruptFileDirPath As String = Nothing
                Dim individualCorruptFileXMLSubFileName As String = Nothing
                Dim xmlValidateArguments As String = Nothing
                Dim xmlValidateCompOut As String = Nothing
                Dim xmlValidateErrorOut As String = Nothing
                Dim byteMatchString As String = Nothing
                Dim byteErrorLocation As String = Nothing
                Dim truncatedLengthAsString As String = Nothing
                Dim truncateArguments As String = Nothing
                Dim truncateFullPath As String = Nothing
                Dim xmllintFullPath As String = Nothing
                Dim xmlValidate2CompOut As String = Nothing
                Dim xmlValidate2ErrorOut As String = Nothing
                Dim xmlValFullPath As String = Nothing
                Dim xmlRecoverArguments As String = Nothing
                Dim zipExtensionXMLRepairedFullPath As String = Nothing
                Dim xmlRepairedFileName As String = Nothing
                Dim xmlRepairedFullPath As String = Nothing
                Dim oldXMLRepairedFileName As String = Nothing
                Dim oldXMLRepairedFullPath As String = Nothing
                Dim sevenZipUpArguments As String = Nothing
                Dim sFileExtension As String = Nothing
                Dim sFileName As String = LCase(Path.GetFileName(sFile))
                Dim sFileZip As String = sFile & "_2.zip"
                Dim zipRepairedsFileName As String = "zipRepaired_2_" & sFileName & ".zip"
                Dim sFileBasePath As String = Path.GetDirectoryName(sFile)
                Dim zipRepairedBasePathAndFileName As String = sFileBasePath & "\" & zipRepairedsFileName
                Dim repairZipArguments As String
                Dim zipFullPath As String
                Dim extractedRepairedZipOutputDirectory As String = _
                    zipRepairedBasePathAndFileName.Remove(zipRepairedBasePathAndFileName.Length - 9)
                Dim zipRepairedBasePathAndFileNameIndexLastPeriod As Integer = 0
                Dim byteErrorLocationInteger As Integer = 0
                Dim truncatedLength As Integer = 0
                Dim intTruncationAmount As Integer = 0
                Dim x As Integer = 0

                progressAuto.Value = x
                progressAuto.Minimum = 0
                progressAuto.Maximum = 100
                progressAuto.Visible = True
                progressAuto.Value = 10
                officeRecoveryXMLRepairExecutionPath = _
                    Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                sFileExtension = LCase(Path.GetExtension(sFile))

                'First we repair the zip as with the previous button.
                'We start by getting the path from the textbox up top where we loaded the file
                'If the extension is .doc, .xls or ppt, it  
                'probably is not a docx, xlsx or ppts format file.

                If sFileExtension = ".doc" Then

                    MsgBox("XML repair is only useful for files with docx, .xslx or pptx extensions " _
                           & "and format. Your file may still be in reality with a docx, xlsx or pptx " _
                           & "format. Try changing the the extension.", MsgBoxStyle.Exclamation)

                    Return bResult

                    Exit Function

                End If

                If sFileExtension = ".xls" Then

                    MsgBox("XML repair is only useful for files with docx, .xslx or pptx extensions " _
                           & "and format. Your file may still be in reality with a docx, xlsx or pptx " _
                           & "format. Try changing the the extension.", MsgBoxStyle.Exclamation)

                    Return bResult

                    Exit Function

                End If

                If sFileExtension = ".ppt" Then

                    MsgBox("XML repair is only useful for files with docx, .xslx or pptx extensions " _
                           & "and format. Your file may still be in reality with a docx, xlsx or pptx " _
                           & "format. Try changing the the extension.", MsgBoxStyle.Exclamation)

                    Return bResult

                    Exit Function

                End If

                If File.Exists(sFileZip) Then

                    File.Delete(sFileZip)

                End If

                File.Copy(sFile, sFileZip, True)

                If File.Exists(zipRepairedBasePathAndFileName) Then

                    File.Delete(zipRepairedBasePathAndFileName)

                End If

                zipFullPath = """" & Path.Combine(officeRecoveryXMLRepairExecutionPath, "zip.exe") & """"
                repairZipArguments = "-FF """ & sFileZip & """ --out """ _
                        & zipRepairedBasePathAndFileName & """"

                Using repairZip As Process = New Process

                    repairZip.StartInfo.FileName = zipFullPath
                    repairZip.StartInfo.Arguments = repairZipArguments
                    repairZip.StartInfo.UseShellExecute = False
                    repairZip.StartInfo.CreateNoWindow = True
                    repairZip.Start()
                    repairZip.WaitForExit()
                    repairZip.Close()

                End Using

                File.Delete(sFileZip)
                progressAuto.Value = 20

                'Now we extract the repaired file.

                zipRepairedBasePathAndFileNameIndexLastPeriod = zipRepairedBasePathAndFileName.LastIndexOf(".")
                extractedRepairedZipOutputDirectory = _
                    zipRepairedBasePathAndFileName.Remove(zipRepairedBasePathAndFileNameIndexLastPeriod - 5)
                sevenZipFullPath = """" & Path.Combine(officeRecoveryXMLRepairExecutionPath, "7z.exe") & """"
                sevenZipExtractArguments = "x """ & zipRepairedBasePathAndFileName & """ -o""" & _
                                                        extractedRepairedZipOutputDirectory & """"

                Using extractZip As Process = New Process

                    extractZip.StartInfo.FileName = sevenZipFullPath
                    extractZip.StartInfo.Arguments = sevenZipExtractArguments
                    extractZip.StartInfo.UseShellExecute = False
                    extractZip.StartInfo.CreateNoWindow = True
                    extractZip.StartInfo.RedirectStandardInput = True
                    extractZip.StartInfo.RedirectStandardOutput = True
                    extractZip.StartInfo.RedirectStandardError = True
                    extractZip.Start()
                    extractZip.StandardInput.WriteLine("u")
                    Dim extractZipReader As StreamReader = extractZip.StandardOutput
                    Dim extractZipReaderCompOut As String = extractZipReader.ReadToEnd
                    Dim extractZipErrorReader As StreamReader = extractZip.StandardError
                    Dim extractZipErrorReaderOut As String = extractZipErrorReader.ReadToEnd
                    extractZip.WaitForExit()
                    extractZip.Close()

                End Using

                progressAuto.Value = 30
                File.Delete(zipRepairedBasePathAndFileName)
                extractedCorruptFileDirPath = extractedRepairedZipOutputDirectory & "\"
                extractedCorruptFileDirInfo = New DirectoryInfo(extractedCorruptFileDirPath)
                corruptFileDirRetrievedFilesInfoArray = _
                    extractedCorruptFileDirInfo.GetFiles("*.*", SearchOption.AllDirectories)

                Dim xmlFileNamesInTheCorruptFile As New List(Of String)

                If corruptFileDirRetrievedFilesInfoArray IsNot Nothing Then

                    For Each indCorruptFileDirRetrievedFileInfo In corruptFileDirRetrievedFilesInfoArray

                        If InStr(indCorruptFileDirRetrievedFileInfo.Extension, ".xml") Or _
                            InStr(indCorruptFileDirRetrievedFileInfo.Extension, ".rels") Then

                            individualCorruptFileXMLSubFileName = indCorruptFileDirRetrievedFileInfo.Name
                            xmlFileNamesInTheCorruptFile.Add(individualCorruptFileXMLSubFileName)

                        End If

                    Next

                Else

                    MsgBox("Your loaded corrupt DOCX, XLSX or PPTX file is missing any XML " _
                        & "sub-files and is unrecoverable unless it is in reality an old " _
                        & "format doc, xls, or ppt format file. Try changing the extension " _
                        & "and using another recovery program suitable for those formats.")

                    Return bResult

                    Exit Function

                End If

                progressAuto.Value = 40

                Dim dummyFileDirPath As String

                If sFileExtension = ".docx" Then

                    dummyFileDirPath = officeRecoveryXMLRepairExecutionPath & "\do_not_delete\word"

                ElseIf sFileExtension = ".xlsx" Then

                    dummyFileDirPath = officeRecoveryXMLRepairExecutionPath & "\do_not_delete\excel"

                Else

                    dummyFileDirPath = officeRecoveryXMLRepairExecutionPath & "\do_not_delete\powerpoint"

                End If

                Dim dummyFileDirInfo As DirectoryInfo = New DirectoryInfo(dummyFileDirPath)
                Dim retrievedSubFilesFromDummyInfoArray As FileInfo() = _
                    dummyFileDirInfo.GetFiles("*.*", SearchOption.AllDirectories)
                Dim indRetrievedSubFileFromDummyInfo As FileInfo
                Dim xmlFileNamesInTheDummyFile As New List(Of String)

                For Each indRetrievedSubFileFromDummyInfo In retrievedSubFilesFromDummyInfoArray

                    If InStr(indRetrievedSubFileFromDummyInfo.Extension, ".xml") _
                        Or InStr(indRetrievedSubFileFromDummyInfo.Extension, ".rels") Then

                        Dim xmlSubFileFromDummyName As String = _
                            indRetrievedSubFileFromDummyInfo.Name
                        Dim xmlSubFileFromDummyNameCharacterCount As Integer = _
                            xmlSubFileFromDummyName.Length

                        xmlFileNamesInTheDummyFile.Add(xmlSubFileFromDummyName)

                        Dim xmlSubFileFromDummyFullPath As String = _
                            indRetrievedSubFileFromDummyInfo.FullName
                        Dim dummyFolderCharacterCount As Integer = dummyFileDirPath.Length
                        Dim xmlSubFileFromDummyFullPathCharacterCount As Integer = _
                            xmlSubFileFromDummyFullPath.Length
                        Dim xmlSubFileFromDummyWithinArchivePathCharacterCount As Integer = _
                            xmlSubFileFromDummyFullPathCharacterCount - _
                            dummyFolderCharacterCount
                        Dim xmlSubFileFromDummyWithinArchivePath As String = _
                            xmlSubFileFromDummyFullPath.Substring(dummyFolderCharacterCount, _
                            xmlSubFileFromDummyWithinArchivePathCharacterCount)
                        Dim missingXMLReplacementNewFullPath As String = _
                            extractedRepairedZipOutputDirectory & _
                            xmlSubFileFromDummyWithinArchivePath
                        Dim missingXMLReplacementNewFullPathCharacterCount As Integer = _
                            missingXMLReplacementNewFullPath.Length
                        Dim missingXMLReplacementDirectory = missingXMLReplacementNewFullPath.Remove _
                            (missingXMLReplacementNewFullPath.Length - _
                            xmlSubFileFromDummyNameCharacterCount)

                        If Not xmlFileNamesInTheCorruptFile.Contains(xmlSubFileFromDummyName) Then

                            If (Not Directory.Exists(missingXMLReplacementDirectory)) Then
                                Directory.CreateDirectory(missingXMLReplacementDirectory)
                            End If

                            File.Copy(xmlSubFileFromDummyFullPath, missingXMLReplacementNewFullPath, True)

                            Continue For

                        Else

                            Continue For

                        End If

                    End If

                Next

                progressAuto.Value = 50

                Dim DummyDirInfoArrayFileCount As Integer = xmlFileNamesInTheDummyFile.Count
                Dim xmlAlteredDirInfo As DirectoryInfo = New DirectoryInfo(extractedCorruptFileDirPath)
                Dim xmlAlteredDirRetrievedFilesInfoArray As FileInfo() = _
                    xmlAlteredDirInfo.GetFiles("*.*", SearchOption.AllDirectories)
                Dim indAlteredDirRetrievedFileInfo As FileInfo
                Dim xmlAlteredDirRetrievedFilesInfoArrayCount As Integer = _
                    xmlAlteredDirRetrievedFilesInfoArray.GetLength(0)

                Dim progressBarIncrement As Integer = 40 \ (xmlAlteredDirRetrievedFilesInfoArrayCount + 1)

                If progressBarIncrement = 0 Then

                    progressBarIncrement = 1

                End If

                For Each indAlteredDirRetrievedFileInfo In xmlAlteredDirRetrievedFilesInfoArray

                    If InStr(indAlteredDirRetrievedFileInfo.Extension, ".xml") Or _
                        InStr(indAlteredDirRetrievedFileInfo.Extension, ".rels") Then

                        If progressAuto.Value > 90 Then

                            progressAuto.Value = 90

                        Else

                            progressAuto.Value = progressAuto.Value + progressBarIncrement

                        End If

                        Dim alteredDirRetrievedXMLFileName As String = indAlteredDirRetrievedFileInfo.Name
                        Dim alteredDirRetrievedXMLFileFullPath As String = indAlteredDirRetrievedFileInfo.FullName
                        Dim extractedCorruptFileDirPathCharacterCount As Integer = _
                            extractedRepairedZipOutputDirectory.Length
                        Dim alteredDirRetrievedXMLFileFullPathCharacterCount As Integer = _
                            alteredDirRetrievedXMLFileFullPath.Length
                        Dim alteredDirRetrievedXMLFileWithinArchivePathCharacterCount As Integer = _
                            alteredDirRetrievedXMLFileFullPathCharacterCount - _
                            extractedCorruptFileDirPathCharacterCount
                        Dim alteredDirRetrievedXMLFileWithinArchivePath As String = _
                            alteredDirRetrievedXMLFileFullPath.Substring(extractedCorruptFileDirPathCharacterCount, _
                            alteredDirRetrievedXMLFileWithinArchivePathCharacterCount)
                        Dim replacementDummyXMLOriginalFullPath As String = dummyFileDirPath _
                                & alteredDirRetrievedXMLFileWithinArchivePath

                        xmlValFullPath = """" & Path.Combine(officeRecoveryXMLRepairExecutionPath, _
                            "xml.exe") & """"
                        xmlValidateArguments = "val -e """ & alteredDirRetrievedXMLFileFullPath & """"

                        Using xmlValidate As Process = New Process

                            xmlValidate.StartInfo.FileName = xmlValFullPath
                            xmlValidate.StartInfo.Arguments = xmlValidateArguments
                            xmlValidate.StartInfo.UseShellExecute = False
                            xmlValidate.StartInfo.RedirectStandardOutput = True
                            xmlValidate.StartInfo.RedirectStandardError = True
                            xmlValidate.StartInfo.CreateNoWindow = True
                            xmlValidate.Start()
                            xmlValidateReader = xmlValidate.StandardOutput
                            xmlValidateCompOut = xmlValidateReader.ReadToEnd
                            xmlValidateErrorReader = xmlValidate.StandardError
                            xmlValidateErrorOut = xmlValidateErrorReader.ReadToEnd
                            xmlValidate.WaitForExit()
                            xmlValidate.Close()

                        End Using

                        If xmlValidateCompOut.Contains("invalid") Then

                            Dim loopCounter As Integer = 0

                            Do

                                'The validator will register an error and indicate the byte location 
                                'of the error if the document.xml file has an error. We isolate
                                ' this byte location with a Regex and the DelFromleft function, 
                                'changed the byte to an integer and subtract 50 more bytes to try
                                ' to steer clear of any additional bad xml just before the error.

                                loopCounter = loopCounter + 1

                                If loopCounter = 1 Then

                                    intTruncationAmount = 0

                                ElseIf loopCounter = 2 Then

                                    intTruncationAmount = 50

                                Else

                                    intTruncationAmount = 100

                                End If

                                byteMatch = Regex.Match(xmlValidateErrorOut, _
                                 (":2.[0-9]+"))
                                byteMatchString = byteMatch.ToString
                                byteErrorLocation = DelFromLeft(":2.", byteMatchString)
                                Integer.TryParse(byteErrorLocation, byteErrorLocationInteger)

                                If byteErrorLocationInteger < 50 Then

                                    'If the corrupt XML file is less than 50 bytes long and can't be recovered. 
                                    'It will deleted if its not found in the dummy or by the dummy file if it is
                                    'found in the dummy file.

                                    If Not xmlFileNamesInTheDummyFile.Contains(alteredDirRetrievedXMLFileName) Then

                                        File.Delete(alteredDirRetrievedXMLFileFullPath)

                                        Exit Do

                                        Continue For

                                    Else

                                        File.Copy(replacementDummyXMLOriginalFullPath, _
                                                  alteredDirRetrievedXMLFileFullPath, True)
                                        xmlFilesReplacedbyDummyOnes.Add(alteredDirRetrievedXMLFileFullPath)
                                        Exit Do

                                        Continue For

                                    End If

                                End If

                                truncatedLength = byteErrorLocationInteger - intTruncationAmount
                                truncatedLengthAsString = String.Empty
                                truncatedLengthAsString = System.Convert.ToString(truncatedLength)
                                truncateArguments = """" & alteredDirRetrievedXMLFileFullPath & """ " & truncatedLengthAsString
                                truncateFullPath = """" & _
                                    Path.Combine(officeRecoveryXMLRepairExecutionPath, "trunc.exe") & """"

                                Using truncate As Process = New Process

                                    truncate.StartInfo.FileName = truncateFullPath
                                    truncate.StartInfo.Arguments = truncateArguments
                                    truncate.StartInfo.UseShellExecute = False
                                    truncate.StartInfo.CreateNoWindow = True
                                    truncate.Start()
                                    truncate.WaitForExit()
                                    truncate.Close()

                                End Using

                                'Now we use xmllint to reconstruct the nice xml ending tags
                                'to try to slip document.xml past Word's XML validator.

                                xmlRecoverArguments = "--recover """ & _
                                    alteredDirRetrievedXMLFileFullPath & """ -o """ & _
                                    alteredDirRetrievedXMLFileFullPath & """"
                                xmllintFullPath = """" & Path.Combine(officeRecoveryXMLRepairExecutionPath, _
                                                "xmllint.exe") & """"

                                Using xmlRecover As Process = New Process

                                    xmlRecover.StartInfo.FileName = xmllintFullPath
                                    xmlRecover.StartInfo.Arguments = xmlRecoverArguments
                                    xmlRecover.StartInfo.UseShellExecute = False
                                    xmlRecover.StartInfo.CreateNoWindow = True
                                    xmlRecover.Start()
                                    xmlRecover.WaitForExit()
                                    xmlRecover.Close()

                                End Using

                                'We validate again, hoping our xmlRecovery is OK by our validator.

                                Using xmlValidate2 As Process = New Process

                                    xmlValidate2.StartInfo.FileName = xmlValFullPath
                                    xmlValidate2.StartInfo.Arguments = xmlValidateArguments
                                    xmlValidate2.StartInfo.UseShellExecute = False
                                    xmlValidate2.StartInfo.RedirectStandardOutput = True
                                    xmlValidate2.StartInfo.RedirectStandardError = True
                                    xmlValidate2.StartInfo.CreateNoWindow = True
                                    xmlValidate2.Start()
                                    xmlValidate2Reader = xmlValidate2.StandardOutput
                                    xmlValidate2CompOut = xmlValidate2Reader.ReadToEnd
                                    xmlValidate2ErrorReader = xmlValidate2.StandardError
                                    xmlValidate2ErrorOut = xmlValidate2ErrorReader.ReadToEnd
                                    xmlValidate2.WaitForExit()
                                    xmlValidate2.Close()

                                End Using

                                'If our validator says the XML is still bad, we 
                                'go through a 2nd of truncation and XML recovery

                                If loopCounter = 3 Then

                                    Exit Do

                                End If

                            Loop Until Not xmlValidate2CompOut.Contains("invalid")

                        End If

                    End If

                Next

                'Now we will rezip our directory and open as a docx, xlsx or pptx file.

                xmlRepairedFileName = "add_missing_subfiles_repaired_" & sFileName
                xmlRepairedFullPath = sFileBasePath & "\" & xmlRepairedFileName

                If File.Exists(xmlRepairedFullPath) Then

                    oldXMLRepairedFileName = "add_missing_subfiles_repaired_old_" & sFileName
                    oldXMLRepairedFullPath = sFileBasePath & "\" & oldXMLRepairedFileName
                    File.Copy(xmlRepairedFullPath, oldXMLRepairedFullPath, True)
                    File.Delete(xmlRepairedFullPath)

                End If

                zipExtensionXMLRepairedFullPath = xmlRepairedFullPath & ".zip"
                sevenZipUpArguments = "a -r """ & zipExtensionXMLRepairedFullPath & """ """ & _
                       extractedRepairedZipOutputDirectory & """\*"

                Using sevenZipReZip As Process = New Process

                    sevenZipReZip.StartInfo.FileName = sevenZipFullPath
                    sevenZipReZip.StartInfo.Arguments = sevenZipUpArguments
                    sevenZipReZip.StartInfo.UseShellExecute = False
                    sevenZipReZip.StartInfo.RedirectStandardOutput = True
                    sevenZipReZip.StartInfo.RedirectStandardError = True
                    sevenZipReZip.StartInfo.CreateNoWindow = True
                    sevenZipReZip.Start()
                    Dim sevenZipReZipReader As StreamReader = sevenZipReZip.StandardOutput
                    Dim sevenZipReZipReaderCompOut As String = sevenZipReZipReader.ReadToEnd
                    Dim sevenZipReZipErrorReader As StreamReader = sevenZipReZip.StandardError
                    Dim sevenZipReZipErrorReaderOut As String = sevenZipReZipErrorReader.ReadToEnd
                    sevenZipReZip.WaitForExit()
                    sevenZipReZip.Close()

                End Using

                progressAuto.Value = 95

                File.Copy(zipExtensionXMLRepairedFullPath, xmlRepairedFullPath, True)

                Using officeFileHandler As Process = New Process

                    Process.Start(xmlRepairedFullPath)

                End Using

                File.Delete(zipExtensionXMLRepairedFullPath)
                Directory.Delete(extractedRepairedZipOutputDirectory, True)

                progressAuto.Value = 100
                progressAuto.Visible = False

            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

        UpdateUI(True)

        Return bResult

    End Function

    Private Function Impl_Recover_Method_4(bSilent As Boolean) As Boolean

        MsgBox("Savvy Recovery for MS Office will now perform salvage recovery using " _
            & "SilverCoder's DocToText.", MsgBoxStyle.Information)

        Dim bResult As Boolean = False

        Try

            If String.IsNullOrEmpty(txtInputFile.Text) = True Then

                SelectFile()

            End If

            If String.IsNullOrEmpty(txtInputFile.Text) = False Then

                UpdateUI(False)

                Dim sFile As String = txtInputFile.Text
                Dim shadowLinkFolderName As New List(Of String)
                Dim nonErrorShadowPathList As New List(Of String)
                Dim xmlFilesCheckedForCorruption As New List(Of String)
                Dim preVersionHashTable As New Hashtable
                Dim myDocument As New XmlDocument
                Dim xmlValidateReader As StreamReader = Nothing
                Dim xmlValidateErrorReader As StreamReader = Nothing
                Dim xmlValidate2Reader As StreamReader = Nothing
                Dim xmlValidate2ErrorReader As StreamReader = Nothing
                Dim xmlFilesReplacedbyDummyOnes As New List(Of String)
                Dim officeRecoveryXMLRepairExecutionPath As String = Nothing
                Dim sFileExtension As String = Nothing
                Dim sFileName As String = LCase(Path.GetFileName(sFile))
                Dim sFileZip As String = sFile & "_3.zip"
                Dim zipRepairedsFileName As String = "zipRepaired_3_" & sFileName & ".zip"
                Dim zipRepairedsFileNameNonZipExt As String = "zipRepaired_3_" & sFileName
                Dim zipRepairedsFileNameNonZipExtNoSpace As String = zipRepairedsFileNameNonZipExt.Replace(" ", "_")
                Dim sFileBasePath As String = Path.GetDirectoryName(sFile)
                Dim salvagedsFileName As String = "salvaged_" & sFileName
                Dim salvagedsFileNameNoSpace As String = salvagedsFileName.Replace(" ", "_")
                Dim salvagedsFileNameAndBasePathNoSpace As String = sFileBasePath & "\" & salvagedsFileNameNoSpace
                Dim zipRepairedBasePathAndFileName As String = sFileBasePath & "\" & zipRepairedsFileName
                Dim zipRepairedFullPathNoSpacesInFileName = sFileBasePath & "\" & zipRepairedsFileNameNonZipExtNoSpace
                Dim repairZipArguments As String = Nothing
                Dim zipFullPath As String = Nothing
                Dim extractedRepairedZipOutputDirectory As String = _
                    zipRepairedBasePathAndFileName.Remove(zipRepairedBasePathAndFileName.Length - 9)
                Dim comboBoxIndex As Integer = 0
                Dim matchCount As Integer = 0
                Dim comboBoxChoiceIndex As Integer = 0
                Dim pathToComboBoxSelectedFileSize As Integer = 0
                Dim x As Integer = 0

                PumpMessages()
                progressAuto.Value = x
                progressAuto.Minimum = 0
                progressAuto.Maximum = 100
                progressAuto.Visible = True
                progressAuto.Value = 20
                officeRecoveryXMLRepairExecutionPath = _
                    Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                sFileExtension = LCase(Path.GetExtension(sFile))

                'First we repair the zip as with the previous button.
                'We start by getting the path from the textbox up top where we loaded the file
                'If the extension is .doc, .xls or ppt, it  
                'probably is not a docx, xlsx or ppts format file.

                If sFileExtension = ".doc" Then

                    MsgBox("XML repair is only useful for files with docx, .xslx or pptx extensions " _
                           & "and format. Your file may still be in reality with a docx, xlsx or pptx " _
                           & "format. Try changing the the extension.", MsgBoxStyle.Exclamation)

                    Return bResult

                    Exit Function

                End If

                If sFileExtension = ".xls" Then

                    MsgBox("XML repair is only useful for files with docx, .xslx or pptx extensions " _
                           & "and format. Your file may still be in reality with a docx, xlsx or pptx " _
                           & "format. Try changing the the extension.", MsgBoxStyle.Exclamation)

                    Return bResult

                    Exit Function

                End If

                If sFileExtension = ".ppt" Then

                    MsgBox("XML repair is only useful for files with docx, .xslx or pptx extensions " _
                           & "and format. Your file may still be in reality with a docx, xlsx or pptx " _
                           & "format. Try changing the the extension.", MsgBoxStyle.Exclamation)

                    Return bResult

                    Exit Function

                End If

                If File.Exists(sFileZip) Then

                    File.Delete(sFileZip)

                End If

                File.Copy(sFile, sFileZip, True)

                If File.Exists(zipRepairedBasePathAndFileName) Then

                    File.Delete(zipRepairedBasePathAndFileName)

                End If

                zipFullPath = """" & Path.Combine(officeRecoveryXMLRepairExecutionPath, "zip.exe") & """"
                repairZipArguments = "-FF """ & sFileZip & """ --out """ _
                        & zipRepairedBasePathAndFileName & """"

                Using repairZip As Process = New Process

                    repairZip.StartInfo.FileName = zipFullPath
                    repairZip.StartInfo.Arguments = repairZipArguments
                    repairZip.StartInfo.UseShellExecute = False
                    repairZip.StartInfo.CreateNoWindow = True
                    repairZip.Start()
                    repairZip.WaitForExit()
                    repairZip.Close()

                End Using

                File.Delete(sFileZip)
                progressAuto.Value = 40
                File.Copy(zipRepairedBasePathAndFileName, zipRepairedsFileNameNonZipExtNoSpace, True)
                File.Delete(zipRepairedBasePathAndFileName)

                Dim doctotextArguments As String = "--fix-xml --unzip-cmd=""7z.exe x %a %f -o%d"" " & _
                    zipRepairedsFileNameNonZipExtNoSpace
                Dim readerdoctotext As StreamReader
                Dim doctotextOutput As String
                Dim readerdoctotextErrors As StreamReader
                Dim doctotextErrorsOutput As String

                Using doctotextProcess As Process = New Process

                    doctotextProcess.StartInfo.FileName = "doctotext.exe"
                    doctotextProcess.StartInfo.Arguments = doctotextArguments
                    doctotextProcess.StartInfo.UseShellExecute = False
                    doctotextProcess.StartInfo.RedirectStandardError = True
                    doctotextProcess.StartInfo.RedirectStandardInput = True
                    doctotextProcess.StartInfo.RedirectStandardOutput = True
                    doctotextProcess.StartInfo.CreateNoWindow = True
                    doctotextProcess.Start()

                    readerdoctotext = doctotextProcess.StandardOutput
                    doctotextOutput = readerdoctotext.ReadToEnd
                    readerdoctotextErrors = doctotextProcess.StandardError
                    doctotextErrorsOutput = readerdoctotextErrors.ReadToEnd

                    doctotextProcess.WaitForExit()
                    doctotextProcess.Close()

                End Using
                progressAuto.Value = 60

                Dim oldSalvagedFileName As String

                If File.Exists(salvagedsFileNameNoSpace) Then

                    oldSalvagedFileName = "salvaged_old_" & salvagedsFileNameNoSpace
                    File.Copy(salvagedsFileNameNoSpace, oldSalvagedFileName, True)
                    File.Delete(salvagedsFileNameNoSpace)

                End If

                File.Create(salvagedsFileNameNoSpace).Dispose()

                Using objWriter As System.IO.StreamWriter = New System.IO.StreamWriter(salvagedsFileName)

                    objWriter.Write(doctotextOutput)
                    objWriter.Close()

                End Using

                File.Delete(zipRepairedsFileNameNonZipExtNoSpace)

                Dim salvagedsFileNameNoSpacesWithLastXRemoved As String = salvagedsFileNameNoSpace.TrimEnd("x")
                Dim originalDirectorySalvagedsFullPathNoSpacesWithLastXRemoved As String = _
                    sFileBasePath & "\" & salvagedsFileNameNoSpacesWithLastXRemoved

                Rename(salvagedsFileNameNoSpace, salvagedsFileNameNoSpacesWithLastXRemoved)
                File.Copy(salvagedsFileNameNoSpacesWithLastXRemoved, _
                          originalDirectorySalvagedsFullPathNoSpacesWithLastXRemoved, True)
                File.Delete(salvagedsFileNameNoSpacesWithLastXRemoved)

                progressAuto.Value = 80

                MsgBox("Savvy Repair will now attempt to open a salvaged version of your file. " _
                       & "The extension has been changed to the old format doc, xls or ppt, to allow " _
                       & "opening of the essentially just text in the originating programs. Please " _
                       & "take the opportunity to do a Save As to resave the file in the in the original " _
                       & """Office Open"" docx, xlsx or pptx format.", MsgBoxStyle.Exclamation)

                Using officeFileHandler As Process = New Process

                    Process.Start(originalDirectorySalvagedsFullPathNoSpacesWithLastXRemoved)

                End Using

                progressAuto.Value = 100
                progressAuto.Visible = False

            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

        UpdateUI(True)

        Return bResult

    End Function


#End Region

#Region "Helpers"

    Public Function DelFromRight(ByVal sChars As String, ByVal sLine As String) As String
        'Removes unwanted characters from right of given string
        ' EXAMPLE
        '  MsgBox DelFromRight(" TEST", "THIS IS A TEST")
        'displays "THIS IS A"

        sLine = ReverseString(sLine)
        sChars = ReverseString(sChars)
        sLine = DelFromLeft(sChars, sLine)
        DelFromRight = ReverseString(sLine)
        Exit Function

    End Function

    Public Function DelFromLeft(ByVal sChars As String, _
      ByVal sLine As String) As String

        ' Removes unwanted characters from left of given string
        '  EXAMPLE
        '      MsgBox DelFromLeft("THIS", "THIS IS A TEST")
        '        displays  "IS A TEST"

        Dim iCount As Integer
        Dim sChar As String

        DelFromLeft = ""
        ' Remove unwanted characters to left of folder name
        If InStr(sLine, sChars) > 0 Then
            For iCount = 1 To Len(sChars)
                ' Retrieve character from start string to 
                'look for in folder string (sLine)
                sChar = Mid$(sChars, iCount, 1)
                ' Remove all characters to left of found string
                sLine = Mid$(sLine, InStr(sLine, sChar) + 1)

            Next iCount
        End If
        DelFromLeft = sLine
        Exit Function

    End Function

    Public Function ReverseString(ByVal InputString As String) _
      As String

        'If you have vb6, you can use
        'StrReverse instead of this function

        Dim lLen As Long, lCtr As Long
        Dim sChar As String
        Dim sAns As String = ""

        lLen = Len(InputString)
        For lCtr = lLen To 1 Step -1
            sChar = Mid(InputString, lCtr, 1)
            sAns = sAns & sChar
        Next

        ReverseString = sAns

    End Function

    Public Sub TraceMsgBox(strMessage As String)

        If _bShowTraces = True Then
            MsgBox(strMessage)
        End If

    End Sub

    Private Sub worker_DoWork(sender As Object, e As DoWorkEventArgs) Handles _worker.DoWork

        Try

        Catch
        End Try

    End Sub

#End Region

    

End Class
