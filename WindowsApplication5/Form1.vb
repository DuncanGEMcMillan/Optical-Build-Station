Imports System.Windows.Forms.DataVisualization.Charting
Imports System
Imports System.IO
Imports System.Text
Imports NationalInstruments.DAQmx
Imports NationalInstruments
Imports System.Data
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports Steema.TeeChart.Functions.PolyFitting
Imports Steema.TeeChart.Functions.TrendFunction
Imports Steema.TeeChart.Styles
Imports Steema.TeeChart.Texts
Imports Steema.TeeChart.Legend
Imports Steema.TeeChart.Functions.Poly

Public Class Form1
    Inherits System.Windows.Forms.Form

    Dim arrayList2 As New System.Collections.ArrayList()
    Dim arrayListMirr1 As New System.Collections.ArrayList()
    Dim arrayListMirr2 As New System.Collections.ArrayList()
    Dim arrayListLaPha As New System.Collections.ArrayList()
    Dim arrayListSMSR As New System.Collections.ArrayList()
    Dim arrayListPeakPower As New System.Collections.ArrayList()
    Dim arrayListChannel As New System.Collections.ArrayList()

    Dim BurnInFolder = "G:\10G TOSA Product\PIC Development\PIC_Measurement_Data\Burn In"
    ' MES
    Dim EazyWorksService As New EazyWorks.EazyWebServiceSoapClient
    Dim mes_username As String = "dmcmillan"
    Dim mes_pass_key As String = "k18jlV4kaVEoLJfNqOPm5mTrXBESZ9VCAd7i8WEa23bEQioaBGTH0g=="

    'Public Const Rse As NationalInstruments.DAQmx.AITerminalConfiguration = 10083

    'Private myTask As Task  'Main Task which is Assigned when the Get Voltage Button is Clicked
    'Private reader As AnalogMultiChannelReader

    'Private totalSamples As Int32        'Global Container for the number samples to be acquired
    'Private acquiredSamplesCount As Int32 = 0   'Iteration variable which hold the current sample being acquired

    'Private dataColumn As DataColumn()             'Channels of Data
    'Private dataTable As DataTable = New DataTable 'Table to Display Data

    'Private data As AnalogWaveform(Of Double)()
    'Private runningTask As Task
    'Private analogInReader As AnalogMultiChannelReader
    ''Private myAsyncCallback As AsyncCallback = New AsyncCallback(AddressOf AnalogInCallback)
    ''Private dataColumn() As DataColumn = Nothing
    ''Private dataTable As DataTable = New DataTable

    Friend WithEvents loopTimer As System.Windows.Forms.Timer
    Friend WithEvents acquisitionResultGroupBox_ As System.Windows.Forms.GroupBox
    ' Dim loopTimer As System.Windows.Forms.Timer
    Dim CurrentRecipe As New ArrayList

    Private Sub TChart1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If FRNumber_TB1.Text <> "" And UID_TB1.Text = "" Then
            GetUID_from_TravelerID() ' Call function using FR ID to get MES Part details UID, Part Number and PIC Type
        End If
        If UID_TB1.Text <> "" And FRNumber_TB1.Text = "" Then
            GetFRNumber_from_UID() ' Call function using FR ID to get MES Part details UID, Part Number and PIC Type
        End If
        If FRNumber_TB1.Text = "" And UID_TB1.Text = "" Then
            MsgBox("No Job/Traveler or PIC UID scanned in!!")
        End If

        GetLaserBurnInRecipe(PIC_PD_TB1.Text, IDensityS1_CB.Text)

        GainLimitS1P1_TB.Text = CurrentRecipe(0)
        LasPhaLimitS1P2_TB.Text = CurrentRecipe(1)
        Mirr1LimitS1P3_TB.Text = CurrentRecipe(2)
        Mirr2LimitS1P4_TB.Text = CurrentRecipe(3)
        Phase1LimitS1P5_TB.Text = CurrentRecipe(4)
        SOA1LimitS1P6_TB.Text = CurrentRecipe(5)
        SOA2LimitS1P7_TB.Text = CurrentRecipe(6)
        GainLimit2S1P8_TB.Text = CurrentRecipe(7)
        Phase2LimitS3P1_TB.Text = CurrentRecipe(8)
        Phase3Limit2S3P2_TB.Text = CurrentRecipe(9)
        Phase4Limit2S3P3_TB.Text = CurrentRecipe(10)
        VBiasNLimit2S3P4_TB.Text = CurrentRecipe(11)
        VBiasPLimit2S3P5_TB.Text = CurrentRecipe(12)
        'PowMonLimit2S3P6_TB.Text = CurrentRecipe(13)

        ILD_S1P1_TB.Text = Convert.ToDouble(CurrentRecipe(0)) + Convert.ToDouble(CurrentRecipe(7)) ' Gain
        ILD_S1P1_VScrollBar.Maximum = Convert.ToDouble(CurrentRecipe(0)) + Convert.ToDouble(CurrentRecipe(7)) + 9
        ILD_S1P1_VScrollBar.Minimum = 0

        ILD_S1P2_TB.Text = CurrentRecipe(1) ' Laser Phase
        ILD_S1P2_VScrollBar.Maximum = Convert.ToDouble(CurrentRecipe(1)) + 9
        ILD_S1P2_VScrollBar.Minimum = 0

        ILD_S1P3_TB.Text = CurrentRecipe(2) ' Mirror1
        ILD_S1P3_VScrollBar.Maximum = Convert.ToDouble(CurrentRecipe(2)) + 9
        ILD_S1P3_VScrollBar.Minimum = 0

        ILD_S1P4_TB.Text = CurrentRecipe(3) 'Mirror2
        ILD_S1P4_VScrollBar.Maximum = Convert.ToDouble(CurrentRecipe(3)) + 9
        ILD_S1P4_VScrollBar.Minimum = 0

        ILD_S1P5_TB.Text = CurrentRecipe(4) ' Phase1
        ILD_S1P5_VScrollBar.Maximum = Convert.ToDouble(CurrentRecipe(4)) + 9
        ILD_S1P5_VScrollBar.Minimum = 0

        ILD_S1P6_TB.Text = CurrentRecipe(5) ' SOA1
        ILD_S1P6_VScrollBar.Maximum = Convert.ToDouble(CurrentRecipe(5)) + 9
        ILD_S1P6_VScrollBar.Minimum = 0

        ILD_S1P7_TB.Text = CurrentRecipe(6) ' SOA2
        ILD_S1P7_VScrollBar.Maximum = Convert.ToDouble(CurrentRecipe(6)) + 9
        ILD_S1P7_VScrollBar.Minimum = 0

        'ILD_S1P8_TB.Text = CurrentRecipe(7) Add numbers together for tied Gain sections

        ILD_S2P1_TB.Text = CurrentRecipe(8) ' Phase2
        ILD_S2P1_VScrollBar.Maximum = Convert.ToDouble(CurrentRecipe(8)) + 9
        ILD_S2P1_VScrollBar.Minimum = 0

        ILD_S2P2_TB.Text = CurrentRecipe(9) ' Phase3
        ILD_S2P2_VScrollBar.Maximum = Convert.ToDouble(CurrentRecipe(9)) + 9
        ILD_S2P2_VScrollBar.Minimum = 0

        ILD_S2P3_TB.Text = CurrentRecipe(10) ' Phase4
        ILD_S2P3_VScrollBar.Maximum = Convert.ToDouble(CurrentRecipe(10)) + 9
        ILD_S2P3_VScrollBar.Minimum = 0

        ILD_S2P4_TB.Text = CurrentRecipe(11) ' VBiasN
        ILD_S2P4_VScrollBar.Maximum = Convert.ToDouble(CurrentRecipe(11)) + 9
        ILD_S2P4_VScrollBar.Minimum = 0

        ILD_S2P5_TB.Text = CurrentRecipe(12) ' VBiasP
        ILD_S2P5_VScrollBar.Maximum = Convert.ToDouble(CurrentRecipe(12)) + 9
        ILD_S2P5_VScrollBar.Minimum = 0

        'ILD_S2P4_TB.Text = CurrentRecipe(13)

        SynchTestSettings() ' Function to synch the burn in settings to test panel display

        ' Adjust PIC picture to reflect the DUT B1 to B16
        Dim NetworkLocation As String = "G:\10G TOSA Product\PIC Development\PIC_Design\PIC B-Type Images\"

        If PIC_PD_TB1.Text.Contains("B1") Then
            PIC_Picture.Load(NetworkLocation + "B01.JPG")
        End If
        If PIC_PD_TB1.Text.Contains("B2") Then
            PIC_Picture.Load(NetworkLocation + "B02.JPG")
        End If
        If PIC_PD_TB1.Text.Contains("B3") Then
            PIC_Picture.Load(NetworkLocation + "B03.JPG")
        End If

        If PIC_PD_TB1.Text.Contains("B4") Then
            PIC_Picture.Load(NetworkLocation + "B04.JPG")
        End If

        If PIC_PD_TB1.Text.Contains("B5") Then
            PIC_Picture.Load(NetworkLocation + "B05.JPG")
        End If

        If PIC_PD_TB1.Text.Contains("B6") Then
            PIC_Picture.Load(NetworkLocation + "B06.JPG")
        End If

        If PIC_PD_TB1.Text.Contains("B7") Then
            PIC_Picture.Load(NetworkLocation + "B07.JPG")
        End If

        If PIC_PD_TB1.Text.Contains("B8") Then
            PIC_Picture.Load(NetworkLocation + "B08.JPG")
        End If

        If PIC_PD_TB1.Text.Contains("B9") Then
            PIC_Picture.Load(NetworkLocation + "B09.JPG")
        End If

        If PIC_PD_TB1.Text.Contains("B10") Then
            PIC_Picture.Load(NetworkLocation + "B10.JPG")
        End If

        If PIC_PD_TB1.Text.Contains("B11") Then
            PIC_Picture.Load(NetworkLocation + "B11.JPG")
        End If

        If PIC_PD_TB1.Text.Contains("B12") Then
            PIC_Picture.Load(NetworkLocation + "B12.JPG")
        End If

        If PIC_PD_TB1.Text.Contains("B13") Then
            PIC_Picture.Load(NetworkLocation + "B13.JPG")
        End If

        If PIC_PD_TB1.Text.Contains("B14") Then
            PIC_Picture.Load(NetworkLocation + "B14.JPG")
        End If
        If PIC_PD_TB1.Text.Contains("B15") Then
            PIC_Picture.Load(NetworkLocation + "B15.JPG")
        End If
        If PIC_PD_TB1.Text.Contains("B16") Then
            PIC_Picture.Load(NetworkLocation + "B16.JPG")
        End If

    End Sub

    Private Sub GetUID_from_TravelerID()
        ' Return uid for Slot1 test PIC
        Dim s_o As String = "{""sFlowName"":""FR_Flow""}"
        Dim s_arCheckVals As String = "[{""_EZQueryItem"":true,""sNm"":""sFRID"",""sTyp"":""string"",""sOperation"":""=="",""sVal"":""" + FRNumber_TB1.Text + """ ,""bAnd"":true,""bOr"":false,""bCase"": false,""arJbConv"": []}]"
        Dim s_arJSColm As String = "[{""sTyp"":""string"",""sNm"":""sFRarPRUniqueID"",""sTitle"":""UID""}]"
        Dim s_oSort As String = "null"
        Dim UID As String
        Dim PD As String
        Dim PDName As String
        Dim FRState As String
        Dim FRStep As String
        Dim FRNotes As String

        If FRNumber_TB1.Text <> "" Then
            UID = EazyWorksService.mGenerateJSON(mes_username, mes_pass_key, s_arJSColm, s_o, s_arCheckVals, s_oSort)
            If UID = "[[""UID""]]" Then
                ' Resubmit FR UID truncating the initial 5 characters
                Dim Modified_FRUID = FRNumber_TB1.Text
                Modified_FRUID = Modified_FRUID.Remove(0, 4)
                s_arCheckVals = "[{""_EZQueryItem"":true,""sNm"":""sFRID"",""sTyp"":""string"",""sOperation"":""contains"",""sVal"":""" + Modified_FRUID + """ ,""bAnd"":true,""bOr"":false,""bCase"": false,""arJbConv"": []}]"
                UID = EazyWorksService.mGenerateJSON(mes_username, mes_pass_key, s_arJSColm, s_o, s_arCheckVals, s_oSort)
            End If
            If UID = "[[""UID""]]" Then
                MsgBox("Job/Traveler Failed!! Try scanned UID instead")
            End If
            Dim numbers = (From s In UID
                   Where Char.IsDigit(s)
                   Select Int32.Parse(s)).ToArray() ' Retrieve only numbers from string
            ' Build uid string from array
            UID = String.Join("", numbers)
            UID_TB1.Text = UID
            UID_TB1.Refresh()

            ' Traveler Part Status
            s_arJSColm = "[{""sTyp"":""string"",""sNm"":""sFRProcessState"",""sTitle"":""""}]"
            FRState = EazyWorksService.mGenerateJSON(mes_username, mes_pass_key, s_arJSColm, s_o, s_arCheckVals, s_oSort)
            ' Modify string
            FRState = FRState.Remove(0, 9)
            FRState = FRState.Remove(FRState.Length - 3, 3)
            FRState_TB.Text = FRState
            FRState_TB.Refresh()

            ' Traveler's Current Step
            s_arJSColm = "[{""sTyp"":""string"",""sNm"":""sFRSD"",""sTitle"":""""}]"
            FRStep = EazyWorksService.mGenerateJSON(mes_username, mes_pass_key, s_arJSColm, s_o, s_arCheckVals, s_oSort)
            FRStep = FRStep.Remove(0, 9)
            FRStep = FRStep.Remove(FRStep.Length - 3, 3)
            CurrentStep_TB.Text = FRStep
            CurrentStep_TB.Refresh()

            ' Get FR Notes
            ' sFRNotes
            s_arJSColm = "[{""sTyp"":""string"",""sNm"":""sFRNotes"",""sTitle"":""""}]"
            FRNotes = EazyWorksService.mGenerateJSON(mes_username, mes_pass_key, s_arJSColm, s_o, s_arCheckVals, s_oSort)
            FRNotes = FRNotes.Remove(0, 9)
            FRNotes = FRNotes.Remove(FRNotes.Length - 3, 3)
            FRNotes = FRNotes.Replace("\n", " ") ' Remove the carriage returns from the string
            FRNotes_TB.Text = FRNotes
            FRNotes_TB.Refresh()

            ' Get Traveler Part Definition Goal
            s_arJSColm = "[{""sTyp"":""string"",""sNm"":""sFRPDGoal"",""sTitle"":""""}]"
            PD = EazyWorksService.mGenerateJSON(mes_username, mes_pass_key, s_arJSColm, s_o, s_arCheckVals, s_oSort)
            If PD.Length > 13 Then
                PD = PD.Remove(0, 9)
                PD = PD.Remove(PD.Length - 3, 3)
                PIC_PD_TB1.Text = PD
                PIC_PD_TB1.Refresh()

                'Get Part Definition Laser Type B1-B16
                s_o = "{""sFlowName"":""PD_Flow""}"
                s_arCheckVals = "[{""_EZQueryItem"":true,""sNm"":""sPDPN"",""sTyp"":""string"",""sOperation"":""=="",""sVal"":""" + PIC_PD_TB1.Text + """ ,""bAnd"":true,""bOr"":false,""bCase"": false,""arJbConv"": []}]"
                s_arJSColm = "[{""sTyp"":""string"",""sNm"":""sPDName"",""sTitle"":""""}]"
                PDName = EazyWorksService.mGenerateJSON(mes_username, mes_pass_key, s_arJSColm, s_o, s_arCheckVals, s_oSort)
                If PDName.Length > 6 Then
                    PDName = PDName.Remove(0, 7)
                    PDName = PDName.Replace("[", "")
                    PDName = PDName.Replace("]", "")
                End If
                PIC_PD_TB1.Text = PDName
                PIC_PD_TB1.Refresh()
            End If
        End If
    End Sub

    Private Sub SynchTestSettings()
        ' Update text box values to mirror Current Controller test settings derived from MES query and current density selection
        'DUT 1
        Dim s_o As String = ""

        ' Move to PIC Burn In step on the DUT's Traveler (FD14-00061 Rev B)
        If FRNumber_TB1.Text <> "" Then
            s_o = "{""bNoChecks"": true,""sFR"": """ + FRNumber_TB1.Text + """,""sSF"": ""SF14-000000330""}"""
            EazyWorksService.mFR_MoveToStep(mes_username, mes_pass_key, s_o)
        End If
    End Sub

    Private Function GetLaserBurnInRecipe(PICType As String, CurrentDensity As String)
        ' Pass in laser type details and return recipe for currents
        ' Default all sections to zero
        Dim GainLimit As String = "0"
        Dim LasPhaLimit As String = "0"
        Dim Mirr1Limit As String = "0"
        Dim Mirr2Limit As String = "0"
        Dim Phase1Limit As String = "0"
        Dim SOA1Limit As String = "0"
        Dim SOA2Limit As String = "0"
        Dim GainLimit2 As String = "0"
        Dim Phase2Limit As String = "0"
        Dim Phase3Limit As String = "0"
        Dim Phase4Limit As String = "0"
        Dim VBiasNLimit As String = "0"
        Dim VBiasPLimit As String = "0"
        Dim PowerMonitorLimit As String = "0"

        CurrentRecipe.Clear()

        ' B1 Section
        Dim Gain2 As Double = 0
        Dim LaserPhase As Double = 0
        Dim Mirr2 As Double = 0
        Dim Phase2 As Double = 0
        Dim SOA2 As Double = 0
        Dim Phase4 As Double = 0
        Dim Mod2 As Double = 0
        '0.00060749	Term 2
        '0.00060749	Term 1
        Dim PowMon As Double = 0
        Dim Mod1 As Double = 0
        Dim Phase3 As Double = 0
        Dim SOA1 As Double = 0
        Dim Phase1 As Double = 0
        Dim Mirr1 As Double = 0
        Dim Gain1 As Double = 0

        If PICType.Contains("B1") Then
            Gain2 = 0.00122815
            LaserPhase = 0.0002625
            Mirr2 = 0.00111948
            Phase2 = 0.00034913
            SOA2 = 0.00092189
            Phase4 = 0.00037815
            Mod2 = 0.00063138
            '0.00060749	Term 2
            '0.00060749	Term 1
            PowMon = 0.00070086
            Mod1 = 0.00063138
            Phase3 = 0.00037815
            SOA1 = 0.00092189
            Phase1 = 0.00034913
            Mirr1 = 0.00193676
            Gain1 = 0.0007084
            GainLimit = Convert.ToString((Gain1 * 10000 * Convert.ToDouble(CurrentDensity)))
            LasPhaLimit = Convert.ToString(LaserPhase * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr1Limit = Convert.ToString(Mirr1 * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr2Limit = Convert.ToString(Mirr2 * 10000 * Convert.ToDouble(CurrentDensity))
            Phase1Limit = Convert.ToString(Phase1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA1Limit = Convert.ToString(SOA1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA2Limit = Convert.ToString(SOA2 * 10000 * Convert.ToDouble(CurrentDensity))
            GainLimit2 = Convert.ToString((Gain2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase2Limit = Convert.ToString((Phase2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase3Limit = Convert.ToString((Phase3 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase4Limit = Convert.ToString((Phase4 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasNLimit = Convert.ToString((Mod1 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasPLimit = Convert.ToString((Mod2 * 10000 * Convert.ToDouble(CurrentDensity)))
            PowerMonitorLimit = Convert.ToString((PowMon * 10000 * Convert.ToDouble(CurrentDensity)))
        End If
        If PICType.Contains("B2") Then
            ' B2 Section
            Gain2 = 0.00122815
            LaserPhase = 0.0002625
            Mirr2 = 0.00111948
            Phase2 = 0.00034913
            SOA2 = 0.00092189
            Phase4 = 0.0002768
            Mod2 = 0.0008632
            '0.00060749	Term 2
            '0.00060749	Term 1
            PowMon = 0.00070086
            Mod1 = 0.00087004
            Phase3 = 0.0
            SOA1 = 0.00092189
            Phase1 = 0.00034913
            Mirr1 = 0.00193673
            Gain1 = 0.0007084

            GainLimit = Convert.ToString((Gain1 * 10000 * Convert.ToDouble(CurrentDensity)))
            LasPhaLimit = Convert.ToString(LaserPhase * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr1Limit = Convert.ToString(Mirr1 * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr2Limit = Convert.ToString(Mirr2 * 10000 * Convert.ToDouble(CurrentDensity))
            Phase1Limit = Convert.ToString(Phase1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA1Limit = Convert.ToString(SOA1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA2Limit = Convert.ToString(SOA2 * 10000 * Convert.ToDouble(CurrentDensity))
            GainLimit2 = Convert.ToString((Gain2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase2Limit = Convert.ToString((Phase2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase3Limit = Convert.ToString((Phase3 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase4Limit = Convert.ToString((Phase4 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasNLimit = Convert.ToString((Mod1 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasPLimit = Convert.ToString((Mod2 * 10000 * Convert.ToDouble(CurrentDensity)))
            PowerMonitorLimit = Convert.ToString((PowMon * 10000 * Convert.ToDouble(CurrentDensity)))
        End If
        If PICType.Contains("B3") Then
            Gain2 = 0.0012299
            LaserPhase = 0.0002625
            Mirr2 = 0.00111948
            Phase2 = 0.00034913
            SOA2 = 0.00092189
            Phase4 = 0.00034815
            Mod2 = 0.00063138
            'Term 2	0.00060749
            'Term 1	0.00060749
            PowMon = 0.00070086
            Mod1 = 0.00063138
            Phase3 = 0.00037815
            SOA1 = 0.00092189
            Phase1 = 0.00034913
            Mirr1 = 0.00193673
            Gain1 = 0.00071015

            GainLimit = Convert.ToString((Gain1 * 10000 * Convert.ToDouble(CurrentDensity)))
            LasPhaLimit = Convert.ToString(LaserPhase * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr1Limit = Convert.ToString(Mirr1 * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr2Limit = Convert.ToString(Mirr2 * 10000 * Convert.ToDouble(CurrentDensity))
            Phase1Limit = Convert.ToString(Phase1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA1Limit = Convert.ToString(SOA1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA2Limit = Convert.ToString(SOA2 * 10000 * Convert.ToDouble(CurrentDensity))
            GainLimit2 = Convert.ToString((Gain2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase2Limit = Convert.ToString((Phase2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase3Limit = Convert.ToString((Phase3 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase4Limit = Convert.ToString((Phase4 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasNLimit = Convert.ToString((Mod1 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasPLimit = Convert.ToString((Mod2 * 10000 * Convert.ToDouble(CurrentDensity)))
            PowerMonitorLimit = Convert.ToString((PowMon * 10000 * Convert.ToDouble(CurrentDensity)))
        End If
        If PICType.Contains("B4") Then
            Gain2 = 0.0012299
            LaserPhase = 0.0002625
            Mirr2 = 0.00111948
            Phase2 = 0.00034913
            SOA2 = 0.00092189
            Phase4 = 0.0002768
            Mod2 = 0.0008632
            'Term 2	0.00060749
            'Term 1	0.00060749
            PowMon = 0.00070086
            Mod1 = 0.00087004
            SOA1 = 0.00092189
            Phase1 = 0.00034913
            Mirr1 = 0.00193673
            Gain1 = 0.00071015

            GainLimit = Convert.ToString((Gain1 * 10000 * Convert.ToDouble(CurrentDensity)))
            LasPhaLimit = Convert.ToString(LaserPhase * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr1Limit = Convert.ToString(Mirr1 * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr2Limit = Convert.ToString(Mirr2 * 10000 * Convert.ToDouble(CurrentDensity))
            Phase1Limit = Convert.ToString(Phase1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA1Limit = Convert.ToString(SOA1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA2Limit = Convert.ToString(SOA2 * 10000 * Convert.ToDouble(CurrentDensity))
            GainLimit2 = Convert.ToString((Gain2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase2Limit = Convert.ToString((Phase2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase3Limit = Convert.ToString((Phase3 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase4Limit = Convert.ToString((Phase4 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasNLimit = Convert.ToString((Mod1 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasPLimit = Convert.ToString((Mod2 * 10000 * Convert.ToDouble(CurrentDensity)))
            PowerMonitorLimit = Convert.ToString((PowMon * 10000 * Convert.ToDouble(CurrentDensity)))
        End If

        If PICType.Contains("B5") Then
            Gain2 = 0.0015799
            LaserPhase = 0.0002625
            Mirr2 = 0.00111948
            Phase2 = 0.00034913
            SOA2 = 0.00092189
            Phase4 = 0.00037815
            Mod2 = 0.00063138
            'Term 2	0.00060749
            'Term 1	0.00060749
            PowMon = 0.00070086
            Mod1 = 0.00063138
            Phase3 = 0.00037815
            SOA1 = 0.00092189
            Phase1 = 0.00034913
            Mirr1 = 0.00193673
            Gain1 = 0.00106015

            GainLimit = Convert.ToString((Gain1 * 10000 * Convert.ToDouble(CurrentDensity)))
            LasPhaLimit = Convert.ToString(LaserPhase * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr1Limit = Convert.ToString(Mirr1 * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr2Limit = Convert.ToString(Mirr2 * 10000 * Convert.ToDouble(CurrentDensity))
            Phase1Limit = Convert.ToString(Phase1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA1Limit = Convert.ToString(SOA1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA2Limit = Convert.ToString(SOA2 * 10000 * Convert.ToDouble(CurrentDensity))
            GainLimit2 = Convert.ToString((Gain2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase2Limit = Convert.ToString((Phase2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase3Limit = Convert.ToString((Phase3 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase4Limit = Convert.ToString((Phase4 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasNLimit = Convert.ToString((Mod1 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasPLimit = Convert.ToString((Mod2 * 10000 * Convert.ToDouble(CurrentDensity)))
            PowerMonitorLimit = Convert.ToString((PowMon * 10000 * Convert.ToDouble(CurrentDensity)))
        End If
        If PICType.Contains("B6") Then
            Gain2 = 0.0012299
            LaserPhase = 0.0002625
            Mirr2 = 0.00111948
            Phase2 = 0.00034913
            SOA2 = 0.00092189
            Phase4 = 0.00037815
            Mod2 = 0.00063138
            PowMon = 0.00070086
            Mod1 = 0.00063138
            Phase3 = 0.00037815
            SOA1 = 0.00092189
            Phase1 = 0.00034913
            Mirr1 = 0.00193673
            Gain1 = 0.00071015
            GainLimit = Convert.ToString((Gain1 * 10000 * Convert.ToDouble(CurrentDensity)))
            LasPhaLimit = Convert.ToString(LaserPhase * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr1Limit = Convert.ToString(Mirr1 * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr2Limit = Convert.ToString(Mirr2 * 10000 * Convert.ToDouble(CurrentDensity))
            Phase1Limit = Convert.ToString(Phase1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA1Limit = Convert.ToString(SOA1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA2Limit = Convert.ToString(SOA2 * 10000 * Convert.ToDouble(CurrentDensity))
            GainLimit2 = Convert.ToString((Gain2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase2Limit = Convert.ToString((Phase2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase3Limit = Convert.ToString((Phase3 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase4Limit = Convert.ToString((Phase4 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasNLimit = Convert.ToString((Mod1 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasPLimit = Convert.ToString((Mod2 * 10000 * Convert.ToDouble(CurrentDensity)))
            PowerMonitorLimit = Convert.ToString((PowMon * 10000 * Convert.ToDouble(CurrentDensity)))
        End If
        If PICType.Contains("B7") Then
            Gain2 = 0.0015799
            LaserPhase = 0.0002625
            Mirr2 = 0.00111948
            Phase2 = 0.00034913
            SOA2 = 0.00092189
            Phase4 = 0.00037815
            Mod2 = 0.00063138
            PowMon = 0.00070086
            Mod1 = 0.00063138
            Phase3 = 0.00037815
            SOA1 = 0.00092189
            Phase1 = 0.00034913
            Mirr1 = 0.00193673
            Gain1 = 0.00106015
            GainLimit = Convert.ToString((Gain1 * 10000 * Convert.ToDouble(CurrentDensity)))
            LasPhaLimit = Convert.ToString(LaserPhase * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr1Limit = Convert.ToString(Mirr1 * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr2Limit = Convert.ToString(Mirr2 * 10000 * Convert.ToDouble(CurrentDensity))
            Phase1Limit = Convert.ToString(Phase1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA1Limit = Convert.ToString(SOA1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA2Limit = Convert.ToString(SOA2 * 10000 * Convert.ToDouble(CurrentDensity))
            GainLimit2 = Convert.ToString((Gain2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase2Limit = Convert.ToString((Phase2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase3Limit = Convert.ToString((Phase3 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase4Limit = Convert.ToString((Phase4 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasNLimit = Convert.ToString((Mod1 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasPLimit = Convert.ToString((Mod2 * 10000 * Convert.ToDouble(CurrentDensity)))
            PowerMonitorLimit = Convert.ToString((PowMon * 10000 * Convert.ToDouble(CurrentDensity)))
        End If
        If PICType.Contains("B8") Then
            Gain2 = 0.0015799
            LaserPhase = 0.0002625
            Mirr2 = 0.00111948
            Phase2 = 0.00034913
            SOA2 = 0.00092189
            Phase4 = 0.0002768
            Mod2 = 0.0008632
            PowMon = 0.00070086
            Mod1 = 0.00087004
            SOA1 = 0.00092189
            Phase1 = 0.00034913
            Mirr1 = 0.00193673
            Gain1 = 0.00106015
            GainLimit = Convert.ToString((Gain1 * 10000 * Convert.ToDouble(CurrentDensity)))
            LasPhaLimit = Convert.ToString(LaserPhase * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr1Limit = Convert.ToString(Mirr1 * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr2Limit = Convert.ToString(Mirr2 * 10000 * Convert.ToDouble(CurrentDensity))
            Phase1Limit = Convert.ToString(Phase1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA1Limit = Convert.ToString(SOA1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA2Limit = Convert.ToString(SOA2 * 10000 * Convert.ToDouble(CurrentDensity))
            GainLimit2 = Convert.ToString((Gain2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase2Limit = Convert.ToString((Phase2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase3Limit = Convert.ToString((Phase3 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase4Limit = Convert.ToString((Phase4 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasNLimit = Convert.ToString((Mod1 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasPLimit = Convert.ToString((Mod2 * 10000 * Convert.ToDouble(CurrentDensity)))
            PowerMonitorLimit = Convert.ToString((PowMon * 10000 * Convert.ToDouble(CurrentDensity)))
        End If
        If PICType.Contains("B9") Then
            Gain2 = 0.00137428
            LaserPhase = 0.0002625
            Mirr2 = 0.00154998
            Phase2 = 0.00034913
            SOA2 = 0.00092565
            Phase4 = 0.00037815
            Mod2 = 0.00063138
            PowMon = 0.00070086
            Mod1 = 0.00063138
            Phase3 = 0.00037815
            SOA1 = 0.00091813
            Phase1 = 0.00034913
            Mirr1 = 0.00265598
            Gain1 = 0.00056578
            GainLimit = Convert.ToString((Gain1 * 10000 * Convert.ToDouble(CurrentDensity)))
            LasPhaLimit = Convert.ToString(LaserPhase * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr1Limit = Convert.ToString(Mirr1 * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr2Limit = Convert.ToString(Mirr2 * 10000 * Convert.ToDouble(CurrentDensity))
            Phase1Limit = Convert.ToString(Phase1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA1Limit = Convert.ToString(SOA1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA2Limit = Convert.ToString(SOA2 * 10000 * Convert.ToDouble(CurrentDensity))
            GainLimit2 = Convert.ToString((Gain2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase2Limit = Convert.ToString((Phase2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase3Limit = Convert.ToString((Phase3 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase4Limit = Convert.ToString((Phase4 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasNLimit = Convert.ToString((Mod1 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasPLimit = Convert.ToString((Mod2 * 10000 * Convert.ToDouble(CurrentDensity)))
            PowerMonitorLimit = Convert.ToString((PowMon * 10000 * Convert.ToDouble(CurrentDensity)))
        End If
        If PICType.Contains("B10") Then
            Gain2 = 0.00137428
            LaserPhase = 0.0002625
            Mirr2 = 0.00154998
            Phase2 = 0.00034913
            SOA2 = 0.00092565
            Phase4 = 0.0002768
            Mod2 = 0.0008632
            PowMon = 0.00070086
            Mod1 = 0.00087004
            SOA1 = 0.00091813
            Phase1 = 0.00034913
            Mirr1 = 0.00265598
            Gain1 = 0.00056578
            GainLimit = Convert.ToString((Gain1 * 10000 * Convert.ToDouble(CurrentDensity)))
            LasPhaLimit = Convert.ToString(LaserPhase * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr1Limit = Convert.ToString(Mirr1 * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr2Limit = Convert.ToString(Mirr2 * 10000 * Convert.ToDouble(CurrentDensity))
            Phase1Limit = Convert.ToString(Phase1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA1Limit = Convert.ToString(SOA1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA2Limit = Convert.ToString(SOA2 * 10000 * Convert.ToDouble(CurrentDensity))
            GainLimit2 = Convert.ToString((Gain2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase2Limit = Convert.ToString((Phase2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase3Limit = Convert.ToString((Phase3 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase4Limit = Convert.ToString((Phase4 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasNLimit = Convert.ToString((Mod1 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasPLimit = Convert.ToString((Mod2 * 10000 * Convert.ToDouble(CurrentDensity)))
            PowerMonitorLimit = Convert.ToString((PowMon * 10000 * Convert.ToDouble(CurrentDensity)))
        End If
        If PICType.Contains("B11") Then
            Gain2 = 0.00137428
            LaserPhase = 0.0002625
            Mirr2 = 0.00154998
            Phase2 = 0.00034913
            SOA2 = 0.00092565
            Phase4 = 0.00037815
            Mod2 = 0.00063138
            PowMon = 0.00070086
            Mod1 = 0.00063138
            Phase3 = 0.00037815
            SOA1 = 0.00091813
            Phase1 = 0.00034913
            Mirr1 = 0.00265598
            Gain1 = 0.00056578
            GainLimit = Convert.ToString((Gain1 * 10000 * Convert.ToDouble(CurrentDensity)))
            LasPhaLimit = Convert.ToString(LaserPhase * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr1Limit = Convert.ToString(Mirr1 * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr2Limit = Convert.ToString(Mirr2 * 10000 * Convert.ToDouble(CurrentDensity))
            Phase1Limit = Convert.ToString(Phase1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA1Limit = Convert.ToString(SOA1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA2Limit = Convert.ToString(SOA2 * 10000 * Convert.ToDouble(CurrentDensity))
            GainLimit2 = Convert.ToString((Gain2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase2Limit = Convert.ToString((Phase2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase3Limit = Convert.ToString((Phase3 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase4Limit = Convert.ToString((Phase4 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasNLimit = Convert.ToString((Mod1 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasPLimit = Convert.ToString((Mod2 * 10000 * Convert.ToDouble(CurrentDensity)))
            PowerMonitorLimit = Convert.ToString((PowMon * 10000 * Convert.ToDouble(CurrentDensity)))
        End If
        If PICType.Contains("B12") Then
            Gain2 = 0.0012299
            LaserPhase = 0.0002625
            Mirr2 = 0.00111948
            Phase2 = 0.00034913
            SOA2 = 0.00092565
            Phase4 = 0.00037815
            Mod2 = 0.00063138
            PowMon = 0.00070086
            Mod1 = 0.00063138
            Phase3 = 0.00037815
            SOA1 = 0.00091813
            Phase1 = 0.00034913
            Mirr1 = 0.00193673
            Gain1 = 0.00071015
            GainLimit = Convert.ToString((Gain1 * 10000 * Convert.ToDouble(CurrentDensity)))
            LasPhaLimit = Convert.ToString(LaserPhase * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr1Limit = Convert.ToString(Mirr1 * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr2Limit = Convert.ToString(Mirr2 * 10000 * Convert.ToDouble(CurrentDensity))
            Phase1Limit = Convert.ToString(Phase1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA1Limit = Convert.ToString(SOA1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA2Limit = Convert.ToString(SOA2 * 10000 * Convert.ToDouble(CurrentDensity))
            GainLimit2 = Convert.ToString((Gain2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase2Limit = Convert.ToString((Phase2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase3Limit = Convert.ToString((Phase3 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase4Limit = Convert.ToString((Phase4 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasNLimit = Convert.ToString((Mod1 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasPLimit = Convert.ToString((Mod2 * 10000 * Convert.ToDouble(CurrentDensity)))
            PowerMonitorLimit = Convert.ToString((PowMon * 10000 * Convert.ToDouble(CurrentDensity)))
        End If
        If PICType.Contains("B13") Then
            Gain2 = 0.00172428
            LaserPhase = 0.0002625
            Mirr2 = 0.00154998
            Phase2 = 0.00034913
            SOA2 = 0.00092565
            Phase4 = 0.00037815
            Mod2 = 0.00063138
            PowMon = 0.00070086
            Mod1 = 0.00063138
            Phase3 = 0.00037815
            SOA1 = 0.00092025
            Phase1 = 0.00034913
            Mirr1 = 0.00265598
            Gain1 = 0.00091578
            GainLimit = Convert.ToString((Gain1 * 10000 * Convert.ToDouble(CurrentDensity)))
            LasPhaLimit = Convert.ToString(LaserPhase * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr1Limit = Convert.ToString(Mirr1 * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr2Limit = Convert.ToString(Mirr2 * 10000 * Convert.ToDouble(CurrentDensity))
            Phase1Limit = Convert.ToString(Phase1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA1Limit = Convert.ToString(SOA1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA2Limit = Convert.ToString(SOA2 * 10000 * Convert.ToDouble(CurrentDensity))
            GainLimit2 = Convert.ToString((Gain2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase2Limit = Convert.ToString((Phase2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase3Limit = Convert.ToString((Phase3 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase4Limit = Convert.ToString((Phase4 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasNLimit = Convert.ToString((Mod1 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasPLimit = Convert.ToString((Mod2 * 10000 * Convert.ToDouble(CurrentDensity)))
            PowerMonitorLimit = Convert.ToString((PowMon * 10000 * Convert.ToDouble(CurrentDensity)))
        End If
        If PICType.Contains("B14") Then
            Gain2 = 0.0015799
            LaserPhase = 0.0002625
            Mirr2 = 0.00111948
            Phase2 = 0.00034913
            Phase4 = 0.00037815
            Mod2 = 0.00063138
            PowMon = 0.00070086
            Mod1 = 0.00063138
            Phase3 = 0.00037815
            Phase1 = 0.00034913
            Mirr1 = 0.00193673
            Gain1 = 0.00106015
            GainLimit = Convert.ToString((Gain1 * 10000 * Convert.ToDouble(CurrentDensity)))
            LasPhaLimit = Convert.ToString(LaserPhase * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr1Limit = Convert.ToString(Mirr1 * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr2Limit = Convert.ToString(Mirr2 * 10000 * Convert.ToDouble(CurrentDensity))
            Phase1Limit = Convert.ToString(Phase1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA1Limit = Convert.ToString(SOA1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA2Limit = Convert.ToString(SOA2 * 10000 * Convert.ToDouble(CurrentDensity))
            GainLimit2 = Convert.ToString((Gain2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase2Limit = Convert.ToString((Phase2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase3Limit = Convert.ToString((Phase3 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase4Limit = Convert.ToString((Phase4 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasNLimit = Convert.ToString((Mod1 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasPLimit = Convert.ToString((Mod2 * 10000 * Convert.ToDouble(CurrentDensity)))
            PowerMonitorLimit = Convert.ToString((PowMon * 10000 * Convert.ToDouble(CurrentDensity)))
        End If
        If PICType.Contains("B15") Then
            Gain2 = 0.0015799
            LaserPhase = 0.0002625
            Mirr2 = 0.00068898
            Phase2 = 0.00034913
            SOA2 = 0.00092401
            Phase4 = 0.00037815
            Mod2 = 0.00063138
            PowMon = 0.00070086
            Mod1 = 0.00063138
            Phase3 = 0.00037815
            SOA1 = 0.00092025
            Phase1 = 0.00034913
            Mirr1 = 0.00121748
            Gain1 = 0.0013489
            GainLimit = Convert.ToString((Gain1 * 10000 * Convert.ToDouble(CurrentDensity)))
            LasPhaLimit = Convert.ToString(LaserPhase * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr1Limit = Convert.ToString(Mirr1 * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr2Limit = Convert.ToString(Mirr2 * 10000 * Convert.ToDouble(CurrentDensity))
            Phase1Limit = Convert.ToString(Phase1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA1Limit = Convert.ToString(SOA1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA2Limit = Convert.ToString(SOA2 * 10000 * Convert.ToDouble(CurrentDensity))
            GainLimit2 = Convert.ToString((Gain2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase2Limit = Convert.ToString((Phase2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase3Limit = Convert.ToString((Phase3 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase4Limit = Convert.ToString((Phase4 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasNLimit = Convert.ToString((Mod1 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasPLimit = Convert.ToString((Mod2 * 10000 * Convert.ToDouble(CurrentDensity)))
            PowerMonitorLimit = Convert.ToString((PowMon * 10000 * Convert.ToDouble(CurrentDensity)))
        End If
        If PICType.Contains("B16") Then
            Gain2 = 0.0015799
            LaserPhase = 0.0002625
            Mirr2 = 0.00068898
            Phase2 = 0.00034913
            SOA2 = 0.00092189
            Phase4 = 0.0002768
            Mod2 = 0.0008632
            PowMon = 0.00070086
            Mod1 = 0.00087004
            SOA1 = 0.00092189
            Phase1 = 0.00034913
            Mirr1 = 0.00121748
            Gain1 = 0.0013489
            GainLimit = Convert.ToString((Gain1 * 10000 * Convert.ToDouble(CurrentDensity)))
            LasPhaLimit = Convert.ToString(LaserPhase * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr1Limit = Convert.ToString(Mirr1 * 10000 * Convert.ToDouble(CurrentDensity))
            Mirr2Limit = Convert.ToString(Mirr2 * 10000 * Convert.ToDouble(CurrentDensity))
            Phase1Limit = Convert.ToString(Phase1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA1Limit = Convert.ToString(SOA1 * 10000 * Convert.ToDouble(CurrentDensity))
            SOA2Limit = Convert.ToString(SOA2 * 10000 * Convert.ToDouble(CurrentDensity))
            GainLimit2 = Convert.ToString((Gain2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase2Limit = Convert.ToString((Phase2 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase3Limit = Convert.ToString((Phase3 * 10000 * Convert.ToDouble(CurrentDensity)))
            Phase4Limit = Convert.ToString((Phase4 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasNLimit = Convert.ToString((Mod1 * 10000 * Convert.ToDouble(CurrentDensity)))
            VBiasPLimit = Convert.ToString((Mod2 * 10000 * Convert.ToDouble(CurrentDensity)))
            PowerMonitorLimit = Convert.ToString((PowMon * 10000 * Convert.ToDouble(CurrentDensity)))
        End If
        ' Return Current Recipe
        CurrentRecipe.Add(GainLimit)
        CurrentRecipe.Add(LasPhaLimit)
        CurrentRecipe.Add(Mirr1Limit)
        CurrentRecipe.Add(Mirr2Limit)
        CurrentRecipe.Add(Phase1Limit)
        CurrentRecipe.Add(SOA1Limit)
        CurrentRecipe.Add(SOA2Limit)
        CurrentRecipe.Add(GainLimit2)
        CurrentRecipe.Add(Phase2Limit)
        CurrentRecipe.Add(Phase3Limit)
        CurrentRecipe.Add(Phase4Limit)
        CurrentRecipe.Add(VBiasNLimit)
        CurrentRecipe.Add(VBiasPLimit)
        CurrentRecipe.Add(PowerMonitorLimit)

        Return CurrentRecipe
        Return True
    End Function

    Private Sub ILD_S1P1_VScrollBar_Scroll(sender As Object, e As ScrollEventArgs) Handles ILD_S1P1_VScrollBar.Scroll
        ILD_S1P1_TB.Text = Format(ILD_S1P1_VScrollBar.Value) ' Gain 1 and 2
    End Sub

    Private Sub ILD_S1P2_VScrollBar_Scroll(sender As Object, e As ScrollEventArgs) Handles ILD_S1P2_VScrollBar.Scroll
        ILD_S1P2_TB.Text = Format(ILD_S1P2_VScrollBar.Value) ' Laser Phase
    End Sub

    Private Sub ILD_S1P4_VScrollBar_Scroll(sender As Object, e As ScrollEventArgs) Handles ILD_S1P4_VScrollBar.Scroll
        ILD_S1P4_TB.Text = Format(ILD_S1P4_VScrollBar.Value) ' Mirror 2
    End Sub

    Private Sub ILD_S1P5_VScrollBar_Scroll(sender As Object, e As ScrollEventArgs) Handles ILD_S1P5_VScrollBar.Scroll
        ILD_S1P5_TB.Text = Format(ILD_S1P5_VScrollBar.Value) ' Phase 1
    End Sub

    Private Sub ILD_S1P6_VScrollBar_Scroll(sender As Object, e As ScrollEventArgs) Handles ILD_S1P6_VScrollBar.Scroll
        ILD_S1P6_TB.Text = Format(ILD_S1P6_VScrollBar.Value) ' SOA 1
    End Sub

    Private Sub ILD_S2P3_VScrollBar_Scroll(sender As Object, e As ScrollEventArgs) Handles ILD_S2P3_VScrollBar.Scroll
        ILD_S2P3_TB.Text = Format(ILD_S2P3_VScrollBar.Value) ' Phase 4
    End Sub

    Private Sub ILD_S2P5_VScrollBar_Scroll(sender As Object, e As ScrollEventArgs) Handles ILD_S2P5_VScrollBar.Scroll
        ILD_S2P5_TB.Text = Format(ILD_S2P5_VScrollBar.Value) ' Modulator 1
    End Sub

    Private Sub ILD_S1P3_VScrollBar_Scroll(sender As Object, e As ScrollEventArgs) Handles ILD_S1P3_VScrollBar.Scroll
        ILD_S1P3_TB.Text = Format(ILD_S1P3_VScrollBar.Value) ' Mirror 1
    End Sub

    Private Sub ILD_S2P1_VScrollBar_Scroll(sender As Object, e As ScrollEventArgs) Handles ILD_S2P1_VScrollBar.Scroll
        ILD_S2P1_TB.Text = Format(ILD_S2P1_VScrollBar.Value) ' Phase 2
    End Sub

    Private Sub ILD_S1P7_VScrollBar_Scroll(sender As Object, e As ScrollEventArgs) Handles ILD_S1P7_VScrollBar.Scroll
        ILD_S1P7_TB.Text = Format(ILD_S1P7_VScrollBar.Value) ' SOA 2
    End Sub

    Private Sub ILD_S2P2_VScrollBar_Scroll(sender As Object, e As ScrollEventArgs) Handles ILD_S2P2_VScrollBar.Scroll
        ILD_S2P2_TB.Text = Format(ILD_S2P2_VScrollBar.Value) ' Phase 3
    End Sub

    Private Sub ILD_S2P4_VScrollBar_Scroll(sender As Object, e As ScrollEventArgs) Handles ILD_S2P4_VScrollBar.Scroll
        ILD_S2P4_TB.Text = Format(ILD_S2P4_VScrollBar.Value) ' Modulator 1
    End Sub

    Private Sub ILD_S1P1_TB_TextChanged(sender As Object, e As EventArgs) Handles ILD_S1P1_TB.TextChanged
        If ILD_S1P1_TB.Text <> "" And GainLimit2S1P8_TB.Text <> "" Then
            If ILD_S1P1_TB.Text > Convert.ToDouble(GainLimitS1P1_TB.Text) + Convert.ToDouble(GainLimit2S1P8_TB.Text) Then
                ILD_S1P1_TB.Text = Convert.ToDouble(GainLimitS1P1_TB.Text) + Convert.ToDouble(GainLimit2S1P8_TB.Text)
            End If
        End If
    End Sub
    Private Sub ILD_S1P1_TB_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ILD_S1P1_TB.KeyPress
        If (ILD_S1P1_TB.Text.IndexOf(".") >= 0 And e.KeyChar = ".") Then e.Handled = True
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) And Not e.KeyChar = "." Then
            e.Handled = True
        End If
    End Sub

    Private Sub ILD_S1P4_TB_TextChanged(sender As Object, e As EventArgs) Handles ILD_S1P4_TB.TextChanged
        If ILD_S1P4_TB.Text <> "" And Mirr2LimitS1P4_TB.Text <> "" Then
            If ILD_S1P4_TB.Text > Convert.ToDouble(Mirr2LimitS1P4_TB.Text) Then
                ILD_S1P4_TB.Text = Mirr2LimitS1P4_TB.Text
            End If
        End If
    End Sub
    Private Sub ILD_S1P4_TB_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ILD_S1P4_TB.KeyPress
        If (ILD_S1P4_TB.Text.IndexOf(".") >= 0 And e.KeyChar = ".") Then e.Handled = True
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) And Not e.KeyChar = "." Then
            e.Handled = True
        End If
    End Sub

    Private Sub ILD_S1P5_TB_TextChanged(sender As Object, e As EventArgs) Handles ILD_S1P5_TB.TextChanged
        If ILD_S1P5_TB.Text <> "" And Phase1LimitS1P5_TB.Text <> "" Then
            If ILD_S1P5_TB.Text > Convert.ToDouble(Phase1LimitS1P5_TB.Text) Then
                ILD_S1P5_TB.Text = Phase1LimitS1P5_TB.Text
            End If
        End If
    End Sub
    Private Sub ILD_S1P5_TB_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ILD_S1P5_TB.KeyPress
        If (ILD_S1P5_TB.Text.IndexOf(".") >= 0 And e.KeyChar = ".") Then e.Handled = True
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) And Not e.KeyChar = "." Then
            e.Handled = True
        End If
    End Sub

    Private Sub ILD_S1P6_TB_TextChanged(sender As Object, e As EventArgs) Handles ILD_S1P6_TB.TextChanged
        If ILD_S1P6_TB.Text <> "" And SOA1LimitS1P6_TB.Text <> "" Then
            If ILD_S1P6_TB.Text > Convert.ToDouble(SOA1LimitS1P6_TB.Text) Then
                ILD_S1P6_TB.Text = SOA1LimitS1P6_TB.Text
            End If
        End If
    End Sub
    Private Sub ILD_S1P6_TB_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ILD_S1P6_TB.KeyPress
        If (ILD_S1P6_TB.Text.IndexOf(".") >= 0 And e.KeyChar = ".") Then e.Handled = True
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) And Not e.KeyChar = "." Then
            e.Handled = True
        End If
    End Sub
    Private Sub ILD_S2P3_TB_TextChanged(sender As Object, e As EventArgs) Handles ILD_S2P3_TB.TextChanged
        If ILD_S2P3_TB.Text <> "" And Phase4Limit2S3P3_TB.Text <> "" Then
            If ILD_S2P3_TB.Text > Convert.ToDouble(Phase4Limit2S3P3_TB.Text) Then
                ILD_S2P3_TB.Text = Phase4Limit2S3P3_TB.Text
            End If
        End If
    End Sub
    Private Sub ILD_S2P3_TB_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ILD_S2P3_TB.KeyPress
        If (ILD_S2P3_TB.Text.IndexOf(".") >= 0 And e.KeyChar = ".") Then e.Handled = True
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) And Not e.KeyChar = "." Then
            e.Handled = True
        End If
    End Sub

    Private Sub ILD_S2P5_TB_TextChanged(sender As Object, e As EventArgs) Handles ILD_S2P5_TB.TextChanged
        If ILD_S2P5_TB.Text <> "" And VBiasPLimit2S3P5_TB.Text <> "" Then
            If ILD_S2P5_TB.Text > Convert.ToDouble(VBiasPLimit2S3P5_TB.Text) Then
                ILD_S2P5_TB.Text = VBiasPLimit2S3P5_TB.Text
            End If
        End If
    End Sub
    Private Sub ILD_S2P5_TB_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ILD_S2P5_TB.KeyPress
        If (ILD_S2P5_TB.Text.IndexOf(".") >= 0 And e.KeyChar = ".") Then e.Handled = True
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) And Not e.KeyChar = "." Then
            e.Handled = True
        End If
    End Sub

    Private Sub ILD_S2P1_TB_TextChanged(sender As Object, e As EventArgs) Handles ILD_S2P1_TB.TextChanged
        If ILD_S2P1_TB.Text <> "" And Phase2LimitS3P1_TB.Text <> "" Then
            If ILD_S2P1_TB.Text > Convert.ToDouble(Phase2LimitS3P1_TB.Text) Then
                ILD_S2P1_TB.Text = Phase2LimitS3P1_TB.Text
            End If
        End If
    End Sub
    Private Sub ILD_S2P1_TB_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ILD_S2P1_TB.KeyPress
        If (ILD_S2P1_TB.Text.IndexOf(".") >= 0 And e.KeyChar = ".") Then e.Handled = True
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) And Not e.KeyChar = "." Then
            e.Handled = True
        End If
    End Sub

    Private Sub ILD_S1P3_TB_TextChanged(sender As Object, e As EventArgs) Handles ILD_S1P3_TB.TextChanged
        If ILD_S1P3_TB.Text <> "" And Mirr1LimitS1P3_TB.Text <> "" Then
            If ILD_S1P3_TB.Text > Convert.ToDouble(Mirr1LimitS1P3_TB.Text) Then
                ILD_S1P3_TB.Text = Mirr1LimitS1P3_TB.Text
            End If
        End If
    End Sub
    Private Sub ILD_S1P3_TB_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ILD_S1P3_TB.KeyPress
        If (ILD_S1P3_TB.Text.IndexOf(".") >= 0 And e.KeyChar = ".") Then e.Handled = True
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) And Not e.KeyChar = "." Then
            e.Handled = True
        End If
    End Sub

    Private Sub ILD_S1P2_TB_TextChanged(sender As Object, e As EventArgs) Handles ILD_S1P2_TB.TextChanged
        If ILD_S1P2_TB.Text <> "" And LasPhaLimitS1P2_TB.Text <> "" Then
            If ILD_S1P2_TB.Text > Convert.ToDouble(LasPhaLimitS1P2_TB.Text) Then
                ILD_S1P2_TB.Text = LasPhaLimitS1P2_TB.Text
            End If
        End If
    End Sub
    Private Sub ILD_S1P2_TB_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ILD_S1P2_TB.KeyPress
        If (ILD_S1P2_TB.Text.IndexOf(".") >= 0 And e.KeyChar = ".") Then e.Handled = True
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) And Not e.KeyChar = "." Then
            e.Handled = True
        End If
    End Sub

    Private Sub ILD_S1P7_TB_TextChanged(sender As Object, e As EventArgs) Handles ILD_S1P7_TB.TextChanged
        If ILD_S1P7_TB.Text <> "" And SOA2LimitS1P7_TB.Text <> "" Then
            If ILD_S1P7_TB.Text > Convert.ToDouble(SOA2LimitS1P7_TB.Text) Then
                ILD_S1P7_TB.Text = SOA2LimitS1P7_TB.Text
            End If
        End If
    End Sub
    Private Sub ILD_S1P7_TB_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ILD_S1P7_TB.KeyPress
        If (ILD_S1P7_TB.Text.IndexOf(".") >= 0 And e.KeyChar = ".") Then e.Handled = True
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) And Not e.KeyChar = "." Then
            e.Handled = True
        End If
    End Sub

    Private Sub ILD_S2P2_TB_TextChanged(sender As Object, e As EventArgs) Handles ILD_S2P2_TB.TextChanged
        If ILD_S2P2_TB.Text <> "" And Phase3Limit2S3P2_TB.Text <> "" Then
            If ILD_S2P2_TB.Text > Convert.ToDouble(Phase3Limit2S3P2_TB.Text) Then
                ILD_S2P2_TB.Text = Phase3Limit2S3P2_TB.Text
            End If
        End If
    End Sub
    Private Sub ILD_S2P2_TB_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ILD_S2P2_TB.KeyPress
        If (ILD_S2P2_TB.Text.IndexOf(".") >= 0 And e.KeyChar = ".") Then e.Handled = True
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) And Not e.KeyChar = "." Then
            e.Handled = True
        End If
    End Sub

    Private Sub ILD_S2P4_TB_TextChanged(sender As Object, e As EventArgs) Handles ILD_S2P4_TB.TextChanged
        If ILD_S2P4_TB.Text <> "" And VBiasNLimit2S3P4_TB.Text <> "" Then
            If ILD_S2P4_TB.Text > Convert.ToDouble(VBiasNLimit2S3P4_TB.Text) Then
                ILD_S2P4_TB.Text = VBiasNLimit2S3P4_TB.Text
            End If
        End If
    End Sub
    Private Sub ILD_S2P4_TB_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ILD_S2P4_TB.KeyPress
        If (ILD_S2P4_TB.Text.IndexOf(".") >= 0 And e.KeyChar = ".") Then e.Handled = True
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) And Not e.KeyChar = "." Then
            e.Handled = True
        End If
    End Sub

    Private Sub ILD_S1P8_VScrollBar_Scroll(sender As Object, e As ScrollEventArgs) Handles ILD_S1P8_VScrollBar.Scroll
        ' Not Used as it's tied to Gain 1 through wirebonds
    End Sub

    Private Sub SetPIC_Button_Click(sender As Object, e As EventArgs) Handles SetPIC_Button.Click
        Dim ChannelID As Integer = Channels_CB.SelectedIndex
        ILD_S1P3_TB.Text = arrayListMirr1.Item(ChannelID)
        ILD_S1P4_TB.Text = arrayListMirr2.Item(ChannelID)
        ILD_S1P2_TB.Text = arrayListLaPha.Item(ChannelID)
        ILD_S1P7_TB.Text = "50"
        ILD_S1P6_TB.Text = "50"
        ILD_S1P1_TB.Text = "50"
        ILD_S1P8_TB.Text = "100"
        ' Set Slot 1 Currents for all ports 1-8
        SetIAllDUT1()
    End Sub

    Private Sub SetIAllDUT1()
        ' Set Port 1 Gain 1
        Dim SLOT As Integer = 1
        Dim Conv_Text_double As Double
        'Application.DoEvents()
        Try
            Using com7 As IO.Ports.SerialPort =
                My.Computer.Ports.OpenSerialPort(DUT1COM_CB.Text)
                Dim ILDP1_TB_Double As Double = Convert.ToDouble(ILD_S1P1_TB.Text) / 1000 ' Check for spec limit compliance
                If ILDP1_TB_Double > Convert.ToDouble(GainLimitS1P1_TB.Text) / 1000 Then
                    ILD_S1P1_TB.Text = Convert.ToString(GainLimitS1P1_TB.Text / 1000)
                Else : ILD_S1P1_TB.Text = Convert.ToString(ILDP1_TB_Double)
                End If
                com7.WriteLine(":SLOT " + Convert.ToString(SLOT) + ";:PORT 1;:ILD:SET " + ILD_S1P1_TB.Text + ";:ILD:START 0;:ILD:STOP 0")

                ' Read back set point and measured current
                com7.WriteLine(":ILD:SET?") ' Read PORT ILD Setting
                Dim ILDSet As String = com7.ReadLine()
                Dim ILDSet_Text As String = ILDSet.Replace(":ILD:SET ", "")
                Conv_Text_double = Convert.ToDouble(ILDSet_Text.Replace(Chr(0), "")) * 1000 ' Convert to milliamps
                ILDSet_Text = Convert.ToString(Format(Conv_Text_double, "#,##0.000"))

                com7.WriteLine(":ILD:ACT?") ' Read PORT ILD Measured
                Dim ILDMeas As String = com7.ReadLine()
                Dim ILDMeas_Text As String = ILDMeas.Replace(":ILD:ACT ", "")
                Conv_Text_double = Convert.ToDouble(ILDMeas_Text.Replace(Chr(0), "")) * 1000 ' Convert to milliamps
                ILDMeas_Text = Convert.ToString(Format(Conv_Text_double, "#,##0.000"))

                ILD_S1P1_TB.Text = ILDSet_Text
                ILD_S1P1_TB.Refresh()
                ILD_MeasS1P1_TB.Text = ILDMeas_Text
                ILD_MeasS1P1_TB.Refresh()

                ' Set Port 2 Laser Phase
                Dim ILDP2_TB_Double As Double = Convert.ToDouble(ILD_S1P2_TB.Text) / 1000 ' Check for spec limit compliance
                If ILDP2_TB_Double > Convert.ToDouble(LasPhaLimitS1P2_TB.Text) / 1000 Then
                    ILD_S1P2_TB.Text = Convert.ToString(Convert.ToDouble(LasPhaLimitS1P2_TB.Text) / 1000)
                Else : ILD_S1P2_TB.Text = Convert.ToString(ILDP2_TB_Double)
                End If
                com7.WriteLine(":SLOT " + Convert.ToString(SLOT) + ";:PORT 2;:ILD:SET " + ILD_S1P2_TB.Text + ";:ILD:START 0;:ILD:STOP 0")

                ' Read back set point and measured current
                com7.WriteLine(":ILD:SET?") ' Read PORT ILD Setting
                ILDSet = com7.ReadLine()
                ILDSet_Text = ILDSet.Replace(":ILD:SET ", "")
                Conv_Text_double = Convert.ToDouble(ILDSet_Text.Replace(Chr(0), "")) * 1000 ' Convert to milliamps
                ILDSet_Text = Convert.ToString(Format(Conv_Text_double, "#,##0.000"))

                com7.WriteLine(":ILD:ACT?") ' Read PORT ILD Measured
                ILDMeas = com7.ReadLine()
                ILDMeas_Text = ILDMeas.Replace(":ILD:ACT ", "")
                Conv_Text_double = Convert.ToDouble(ILDMeas_Text.Replace(Chr(0), "")) * 1000 ' Convert to milliamps
                ILDMeas_Text = Convert.ToString(Format(Conv_Text_double, "#,##0.000"))

                ILD_S1P2_TB.Text = ILDSet_Text
                ILD_S1P2_TB.Refresh()
                ILD_MeasS1P2_TB.Text = ILDMeas_Text
                ILD_MeasS1P2_TB.Refresh()

                ' Set Port 3 Mirror 1
                Dim ILDP3_TB_Double As Double = Convert.ToDouble(ILD_S1P3_TB.Text) / 1000 ' Check for spec limit compliance
                If ILDP3_TB_Double > Convert.ToDouble(Mirr1LimitS1P3_TB) / 1000 Then
                    ILD_S1P3_TB.Text = Convert.ToString(Convert.ToDouble(Mirr1LimitS1P3_TB) / 1000)
                Else : ILD_S1P3_TB.Text = Convert.ToString(ILDP3_TB_Double)
                End If
                com7.WriteLine(":SLOT " + Convert.ToString(SLOT) + ";:PORT 3;:ILD:SET " + ILD_S1P3_TB.Text + ";:ILD:START 0;:ILD:STOP 0")

                ' Read back set point and measured current
                com7.WriteLine(":ILD:SET?") ' Read PORT ILD Setting
                ILDSet = com7.ReadLine()
                ILDSet_Text = ILDSet.Replace(":ILD:SET ", "")
                Conv_Text_double = Convert.ToDouble(ILDSet_Text.Replace(Chr(0), "")) * 1000 ' Convert to milliamps
                ILDSet_Text = Convert.ToString(Format(Conv_Text_double, "#,##0.000"))

                com7.WriteLine(":ILD:ACT?") ' Read PORT ILD Measured
                ILDMeas = com7.ReadLine()
                ILDMeas_Text = ILDMeas.Replace(":ILD:ACT ", "")
                Conv_Text_double = Convert.ToDouble(ILDMeas_Text.Replace(Chr(0), "")) * 1000 ' Convert to milliamps
                ILDMeas_Text = Convert.ToString(Format(Conv_Text_double, "#,##0.000"))

                ILD_S1P3_TB.Text = ILDSet_Text
                ILD_S1P3_TB.Refresh()
                ILD_MeasS1P3_TB.Text = ILDMeas_Text
                ILD_MeasS1P3_TB.Refresh()

                ' Set Port 4 Mirror 2
                Dim ILDP4_TB_Double As Double = Convert.ToDouble(ILD_S1P4_TB.Text) / 1000 ' Check for spec limit compliance
                If ILDP4_TB_Double > Convert.ToDouble(Mirr2LimitS1P4_TB) / 1000 Then
                    ILD_S1P4_TB.Text = Convert.ToString(Convert.ToDouble(Mirr2LimitS1P4_TB) / 1000)
                Else : ILD_S1P4_TB.Text = Convert.ToString(ILDP4_TB_Double)
                End If
                com7.WriteLine(":SLOT " + Convert.ToString(SLOT) + ";:PORT 4;:ILD:SET " + ILD_S1P4_TB.Text + ";:ILD:START 0;:ILD:STOP 0")

                ' Read back set point and measured current
                com7.WriteLine(":ILD:SET?") ' Read PORT ILD Setting
                ILDSet = com7.ReadLine()
                ILDSet_Text = ILDSet.Replace(":ILD:SET ", "")
                Conv_Text_double = Convert.ToDouble(ILDSet_Text.Replace(Chr(0), "")) * 1000 ' Convert to milliamps
                ILDSet_Text = Convert.ToString(Format(Conv_Text_double, "#,##0.000000"))

                com7.WriteLine(":ILD:ACT?") ' Read PORT ILD Measured
                ILDMeas = com7.ReadLine()
                ILDMeas_Text = ILDMeas.Replace(":ILD:ACT ", "")
                Conv_Text_double = Convert.ToDouble(ILDMeas_Text.Replace(Chr(0), "")) * 1000 ' Convert to milliamps
                ILDMeas_Text = Convert.ToString(Format(Conv_Text_double, "#,##0.000000"))

                ILD_S1P4_TB.Text = ILDSet_Text
                ILD_S1P4_TB.Refresh()
                ILD_MeasS1P4_TB.Text = ILDMeas_Text
                ILD_MeasS1P4_TB.Refresh()

                ' Set Port 5 Phase 1
                Dim ILDP5_TB_Double As Double = Convert.ToDouble(ILD_S1P5_TB.Text) / 1000 ' Check for spec limit compliance
                If ILDP5_TB_Double > Convert.ToDouble(Phase1LimitS1P5_TB.Text) / 1000 Then
                    ILD_S1P5_TB.Text = Convert.ToString(Convert.ToDouble(Phase1LimitS1P5_TB.Text) / 1000)
                Else : ILD_S1P5_TB.Text = Convert.ToString(ILDP5_TB_Double)
                End If
                com7.WriteLine(":SLOT " + Convert.ToString(SLOT) + ";:PORT 5;:ILD:SET " + ILD_S1P5_TB.Text + ";:ILD:START 0;:ILD:STOP 0")

                ' Read back set point and measured current
                com7.WriteLine(":ILD:SET?") ' Read PORT ILD Setting
                ILDSet = com7.ReadLine()
                ILDSet_Text = ILDSet.Replace(":ILD:SET ", "")
                Conv_Text_double = Convert.ToDouble(ILDSet_Text.Replace(Chr(0), "")) * 1000 ' Convert to milliamps
                ILDSet_Text = Convert.ToString(Format(Conv_Text_double, "#,##0.000"))

                com7.WriteLine(":ILD:ACT?") ' Read PORT ILD Measured
                ILDMeas = com7.ReadLine()
                ILDMeas_Text = ILDMeas.Replace(":ILD:ACT ", "")
                Conv_Text_double = Convert.ToDouble(ILDMeas_Text.Replace(Chr(0), "")) * 1000 ' Convert to milliamps
                ILDMeas_Text = Convert.ToString(Format(Conv_Text_double, "#,##0.000"))

                ILD_S1P5_TB.Text = ILDSet_Text
                ILD_S1P5_TB.Refresh()
                ILD_MeasS1P5_TB.Text = ILDMeas_Text
                ILD_MeasS1P5_TB.Refresh()

                ' Set Port 6 SOA1
                Dim ILDP6_TB_Double As Double = Convert.ToDouble(ILD_S1P6_TB.Text) / 1000 ' Check for spec limit compliance
                If ILDP6_TB_Double > Convert.ToDouble(SOA1LimitS1P6_TB.Text) / 1000 Then
                    ILD_S1P6_TB.Text = Convert.ToString(Convert.ToDouble(SOA1LimitS1P6_TB.Text) / 1000)
                Else : ILD_S1P6_TB.Text = Convert.ToString(ILDP6_TB_Double)
                End If
                com7.WriteLine(":SLOT " + Convert.ToString(SLOT) + ";:PORT 6;:ILD:SET " + ILD_S1P6_TB.Text + ";:ILD:START 0;:ILD:STOP 0")

                ' Read back set point and measured current
                com7.WriteLine(":ILD:SET?") ' Read PORT ILD Setting
                ILDSet = com7.ReadLine()
                ILDSet_Text = ILDSet.Replace(":ILD:SET ", "")
                Conv_Text_double = Convert.ToDouble(ILDSet_Text.Replace(Chr(0), "")) * 1000 ' Convert to milliamps
                ILDSet_Text = Convert.ToString(Format(Conv_Text_double, "#,##0.000000"))

                com7.WriteLine(":ILD:ACT?") ' Read PORT ILD Measured
                ILDMeas = com7.ReadLine()
                ILDMeas_Text = ILDMeas.Replace(":ILD:ACT ", "")
                Conv_Text_double = Convert.ToDouble(ILDMeas_Text.Replace(Chr(0), "")) * 1000 ' Convert to milliamps
                ILDMeas_Text = Convert.ToString(Format(Conv_Text_double, "#,##0.000000"))

                ILD_S1P6_TB.Text = ILDSet_Text
                ILD_S1P6_TB.Refresh()
                ILD_MeasS1P6_TB.Text = ILDMeas_Text
                ILD_MeasS1P6_TB.Refresh()


                ' Set Port 7 SOA2
                Dim ILDP7_TB_Double As Double = Convert.ToDouble(ILD_S1P7_TB.Text) / 1000 ' Check for spec limit compliance
                If ILDP7_TB_Double > Convert.ToDouble(SOA2LimitS1P7_TB.Text) / 1000 Then
                    ILD_S1P7_TB.Text = Convert.ToString(Convert.ToDouble(SOA2LimitS1P7_TB.Text) / 1000)
                Else : ILD_S1P7_TB.Text = Convert.ToString(ILDP7_TB_Double)
                End If
                com7.WriteLine(":SLOT " + Convert.ToString(SLOT) + ";:PORT 7;:ILD:SET " + ILD_S1P7_TB.Text + ";:ILD:START 0;:ILD:STOP 0")

                ' Read back set point and measured current
                com7.WriteLine(":ILD:SET?") ' Read PORT ILD Setting
                ILDSet = com7.ReadLine()
                ILDSet_Text = ILDSet.Replace(":ILD:SET ", "")
                Conv_Text_double = Convert.ToDouble(ILDSet_Text.Replace(Chr(0), "")) * 1000 ' Convert to milliamps
                ILDSet_Text = Convert.ToString(Format(Conv_Text_double, "#,##0.000000"))

                com7.WriteLine(":ILD:ACT?") ' Read PORT ILD Measured
                ILDMeas = com7.ReadLine()
                ILDMeas_Text = ILDMeas.Replace(":ILD:ACT ", "")
                Conv_Text_double = Convert.ToDouble(ILDMeas_Text.Replace(Chr(0), "")) * 1000 ' Convert to milliamps
                ILDMeas_Text = Convert.ToString(Format(Conv_Text_double, "#,##0.000000"))

                ILD_S1P7_TB.Text = ILDSet_Text
                ILD_S1P7_TB.Refresh()
                ILDMeas_S1P7_TB.Text = ILDMeas_Text
                ILDMeas_S1P7_TB.Refresh()
            End Using
        Catch ex As Exception
            MsgBox("Failed to set the device!")
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles GetλMap_Button.Click
        Dim strExcelFilePath As String = "" ' File location if wavelength map is created
        Dim strFilename As String = ""

        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkBooks As Excel.Workbooks = Nothing
        Dim xlWorkBook As Excel.Workbook = Nothing
        Dim xlWorkSheet As Excel.Worksheet = Nothing
        Dim xlWorkSheets As Excel.Sheets = Nothing
        Dim xlCells As Excel.Range = Nothing

        xlApp = New Excel.Application
        xlApp.DisplayAlerts = False
        xlWorkBooks = xlApp.Workbooks

        strExcelFilePath = "G:\10G TOSA Product\PIC Development\PIC_Measurement_Data\UID" + UID_TB1.Text + "\DC_Raw_Data\TOSA\"
        'strExcelFilePath = "C:\Users\dmcmillan\Documents\"
        ' Get filename with wavelength map in this folder
        Dim arrayList As New System.Collections.ArrayList()
        strFilename = ""
        Dim txtFiles = Directory.GetFiles(strExcelFilePath, "*_wavelength_search_data*.CSV", SearchOption.TopDirectoryOnly).[Select](Function(nm) Path.GetFileName(nm))
        For Each filenm As String In txtFiles
            ArrayList.Add(filenm)
        Next
        If arrayList.Count > 0 Then
            strFilename = arrayList.Item(0)
            Using MyReader As New Microsoft.VisualBasic.
                                          FileIO.TextFieldParser(
                                            strExcelFilePath + strFilename)
                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters(",")
                Dim currentRow As String()

                While Not MyReader.EndOfData
                    Try
                        currentRow = MyReader.ReadFields()
                        Dim currentField As String
                        If currentRow(5) = "1.0000000E+00" Then
                            arrayListMirr1.Add(currentRow(0))
                            arrayListMirr2.Add(currentRow(1))
                            arrayListLaPha.Add(currentRow(6))
                            arrayListChannel.Add(currentRow(2))
                            arrayListSMSR.Add(currentRow(7))
                            arrayListPeakPower.Add(currentRow(8))
                        End If
                        'For Each currentField In currentRow
                        '    MsgBox(currentField)
                        'Next
                    Catch ex As Microsoft.VisualBasic.
                                FileIO.MalformedLineException
                        MsgBox("Line " & ex.Message &
                        "is not valid and will be skipped.")
                    End Try
                End While
            End Using

            Dim ChannelCount As Integer = 0
            Do Until ChannelCount >= arrayListChannel.Count
                Channels_CB.Items.Add(arrayListChannel.Item(ChannelCount))
                ChannelCount += 1
            Loop
            Channels_CB.Text = Channels_CB.Items(0)
        Else : MsgBox("No Wavelength Data currently available!!")
            Channels_CB.Text = "No Data!"
        End If
        Channels_CB.Refresh()
    End Sub

    Private Sub GetFRNumber_from_UID()
        ' Return uid for Slot1 test PIC
        Dim s_o As String = "{""sFlowName"":""PR_Flow""}"
        Dim s_arCheckVals As String = "[{""_EZQueryItem"":true,""sNm"":""sPRUniqueID"",""sTyp"":""string"",""sOperation"":""=="",""sVal"":""" + UID_TB1.Text + """ ,""bAnd"":true,""bOr"":false,""bCase"": false,""arJbConv"": []}]"
        Dim s_arJSColm As String = "[{""sTyp"":""string"",""sNm"":""sPRFR"",""sTitle"":""FR UID""}]"
        Dim s_oSort As String = "null"
        Dim FR_UID As String
        Dim PD As String
        Dim PDName As String

        If UID_TB1.Text <> "" Then
            FR_UID = EazyWorksService.mGenerateJSON(mes_username, mes_pass_key, s_arJSColm, s_o, s_arCheckVals, s_oSort)
            ' FR_UID = String.Join("", numbers)
            '14 characters to be trimmed from start and 13 trimmed from end of Query return
            FR_UID = FR_UID.Remove(0, 14)
            Dim LastChar As Integer = FR_UID.Length
            FR_UID = FR_UID.Remove(LastChar - 13, 13)
            FRNumber_TB1.Text = FR_UID
            FRNumber_TB1.Refresh()
            ' Get Traveler Part Definition Goal from FRID using FR Flow
            s_o = "{""sFlowName"":""FR_Flow""}"
            s_arCheckVals = "[{""_EZQueryItem"":true,""sNm"":""sJob"",""sTyp"":""string"",""sOperation"":""=="",""sVal"":""" + FR_UID + """ ,""bAnd"":true,""bOr"":false,""bCase"": false,""arJbConv"": []}]"
            s_arJSColm = "[{""sTyp"":""string"",""sNm"":""sFRPDGoal"",""sTitle"":""""}]"
            PD = EazyWorksService.mGenerateJSON(mes_username, mes_pass_key, s_arJSColm, s_o, s_arCheckVals, s_oSort)
            If PD.Length > 13 Then
                PD = PD.Remove(0, 9)
                PD = PD.Remove(14, 3)
                PIC_PD_TB1.Text = PD
                PIC_PD_TB1.Refresh()
                'Get Part Definition Laser Type B1-B16
                s_o = "{""sFlowName"":""PD_Flow""}"
                s_arCheckVals = "[{""_EZQueryItem"":true,""sNm"":""sPDPN"",""sTyp"":""string"",""sOperation"":""=="",""sVal"":""" + PIC_PD_TB1.Text + """ ,""bAnd"":true,""bOr"":false,""bCase"": false,""arJbConv"": []}]"
                s_arJSColm = "[{""sTyp"":""string"",""sNm"":""sPDName"",""sTitle"":""""}]"
                PDName = EazyWorksService.mGenerateJSON(mes_username, mes_pass_key, s_arJSColm, s_o, s_arCheckVals, s_oSort)
                If PDName.Length > 6 Then
                    PDName = PDName.Remove(0, 7)
                    PDName = PDName.Replace("[", "")
                    PDName = PDName.Replace("]", "")
                End If
                PIC_PD_TB1.Text = PDName
                PIC_PD_TB1.Refresh()
            End If
        End If
        'Throw New NotImplementedException
    End Sub

End Class
