Imports System.IO.Ports
Imports System.Threading
Imports MySql.Data.MySqlClient
Imports Microsoft.Office.Interop

Public Class ReadHoldingRegistersExampleForm
    Dim MysqlConn As MySqlConnection

    Dim COMMAND As MySqlCommand


    'Declare variablses & constants 
    Private serialPort As SerialPort = Nothing
    Private Sub ReadHoldingRegistersExampleForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            serialPort = New SerialPort("COM7", 9600, Parity.None, 8, StopBits.One)
            serialPort.Open()
        Catch ex As Exception
            ' MessageBox.Show(Me, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ReadHoldingRegistersExampleForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Try
            serialPort.Close()
        Catch ex As Exception
            ' MessageBox.Show(Me, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnReadHoldingRegisters_Click(sender As Object, e As EventArgs)
        Try
            Dim slaveAddress As Byte = 1
            Dim functionCode As Byte = 3
            Dim startAddress As UShort = 32784
            Dim numberOfPoints = 17

            Dim frame As Byte() = Me.ReadHolingRegisters(slaveAddress, functionCode, startAddress, numberOfPoints)
            txtSendMsg.Text = Me.DisplayValue(frame)
            serialPort.Write(frame, 0, frame.Length)
            Thread.Sleep(100)


            If serialPort.BytesToRead > 5 Then
                Dim buffRecei As Byte() = New Byte(serialPort.BytesToRead) {}
                serialPort.Read(buffRecei, 0, buffRecei.Length)
                txtReceiMsg.Text = Me.DisplayValue(buffRecei)

                Dim data As Byte() = New Byte(buffRecei.Length - 5) {}
                Array.Copy(buffRecei, 3, data, 0, data.Length)
                Dim result As UInt16() = DataType.Word.ToArray(data) '

                'Display Result
                txtResult.Text = String.Empty
                For Each item As UInt16 In result
                    txtResult.Text += String.Format("{0}", item)
                Next


            End If

        Catch ex As Exception
            'MessageBox.Show(Me, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Read Holding registers
    ''' </summary>
    ''' <param name="slaveAddress"> Salve Address</param>
    ''' <param name="functionCode"> Functon</param>
    ''' <param name="startAddress"> Starting Address</param>
    ''' <param name="numberOfPoints"> Quantity of Registers</param>
    ''' <returns>Byte()</returns>
    ''' <remarks> </remarks>
    Private Function ReadHolingRegisters(slaveAddress As Byte, functionCode As Byte, startAddress As UShort, numberOfPoints As UShort) As Byte()
        Dim frame As Byte() = New Byte(7) {}
        frame(0) = slaveAddress
        frame(1) = functionCode
        frame(2) = CByte(startAddress / 256)
        frame(3) = CByte(startAddress Mod 256)
        frame(4) = CByte(numberOfPoints >> 8)
        frame(5) = CByte(numberOfPoints)
        Dim crc As Byte() = Me.CRC(frame)
        frame(6) = crc(0)
        frame(7) = crc(1)
        Return frame

    End Function

    Private Function CRC(data As Byte()) As Byte()
        Dim CRCFull As UShort = &HFFFF
        Dim CRCHigh As Byte = &HF, CRCLow As Byte = &HFF
        Dim CRCLSB As Char
        Dim result As Byte() = New Byte(1) {}
        For i As Integer = 0 To (data.Length) - 3
            CRCFull = CUShort(CRCFull Xor data(i))

            For j As Integer = 0 To 7
                CRCLSB = ChrW(CRCFull And &H1)
                CRCFull = CUShort((CRCFull >> 1) And &H7FFF)
                If Convert.ToInt32(CRCLSB) = 1 Then
                    CRCFull = CUShort(CRCFull Xor &HA001)
                End If
            Next
        Next
        CRCHigh = CByte((CRCFull >> 8) And &HFF)
        CRCLow = CByte((CRCFull And &HFF))
        CRCLow = CByte(CRCFull And &HFF)
        Return New Byte(1) {CRCLow, CRCHigh}
    End Function
    Private Function DisplayValue(values As Byte()) As String
        Dim result As String = String.Empty
        For Each item As Byte In values
            result += String.Format("{0:X2}", item)
        Next
        Return result
    End Function

    Private Sub RegisterTime_Tick(sender As Object, e As EventArgs)
        Try
            Dim slaveAddress As Byte = 1
            Dim functionCode As Byte = 3
            Dim startAddress As UShort = 32784
            Dim numberOfPoints = 1

            Dim frame As Byte() = Me.ReadHolingRegisters(slaveAddress, functionCode, startAddress, numberOfPoints)
            txtSendMsg.Text = Me.DisplayValue(frame)
            serialPort.Write(frame, 0, frame.Length)
            Thread.Sleep(100)


            If serialPort.BytesToRead > 5 Then
                Dim buffRecei As Byte() = New Byte(serialPort.BytesToRead) {}
                serialPort.Read(buffRecei, 0, buffRecei.Length)
                txtReceiMsg.Text = Me.DisplayValue(buffRecei)

                Dim data As Byte() = New Byte(buffRecei.Length - 5) {}
                Array.Copy(buffRecei, 3, data, 0, data.Length)
                Dim result As UInt16() = DataType.Word.ToArray(data) '

                'Display Result
                txtResult.Text = String.Empty
                For Each item As UInt16 In result
                    txtResult.Text += String.Format("{0}", item)
                Next

            End If

        Catch ex As Exception
            ' MessageBox.Show(Me, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)

        MysqlConn = New MySqlConnection
        MysqlConn.ConnectionString = "server ='localhost';userid='root';database='database';password=''"
        Dim READER As MySqlDataReader

        Try
            '  Dim subst As String = txtReceiMsg.Text.Substring(0, 6)
            '  If subst = 10311 Then
            MysqlConn.Open()
            Dim Query As String
            Query = "insert into mydata(treceimsg,tresult)
     values ('" & txtReceiMsg.Text & "','" & txtResult.Text & "')"
            COMMAND = New MySqlCommand(Query, MysqlConn)
            READER = COMMAND.ExecuteReader

            ' MessageBox.Show("Data successfully insert to Database")
            MysqlConn.Close()
            '   End If
        Catch ex As Exception
            ' MessageBox.Show(Me, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Try
            Dim slaveAddress As Byte = 1
            Dim functionCode As Byte = 3
            Dim startAddress As UShort = 32784
            Dim numberOfPoints = 17

            Dim frame As Byte() = Me.ReadHolingRegisters(slaveAddress, functionCode, startAddress, numberOfPoints)
            txtSendMsg.Text = Me.DisplayValue(frame)
            serialPort.Write(frame, 0, frame.Length)
            Thread.Sleep(100)


            If serialPort.BytesToRead > 5 Then
                Dim buffRecei As Byte() = New Byte(serialPort.BytesToRead) {}
                serialPort.Read(buffRecei, 0, buffRecei.Length)
                txtReceiMsg.Text = Me.DisplayValue(buffRecei)

                Dim data As Byte() = New Byte(buffRecei.Length - 5) {}
                Array.Copy(buffRecei, 3, data, 0, data.Length)
                Dim result As UInt16() = DataType.Word.ToArray(data) '

                'Display Result
                txtResult.Text = String.Empty
                For Each item As UInt16 In result
                    txtResult.Text += String.Format("{0}", item)
                Next

                '  Process.Start("http://localhost/apitest?msg=" & txtSendMsg.Text & "&txtReceiMsg=" &
                ' txtReceiMsg.Text &
                ' "&txtResult=" & txtResult.Text)

            End If

        Catch ex As Exception
            ' MessageBox.Show(Me, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        MysqlConn = New MySqlConnection
        MysqlConn.ConnectionString ="server ='localhost';userid='root';database='database';password=''"
        ' "server ='nhimex-prod.c0hfebe8dpuc.ap-southeast-1.rds.amazonaws.com';userid='nhimexAtiqDev';database='nhimex';password='nhimexdev2020'"
        Dim READER As MySqlDataReader

        Try


            If ((txtReceiMsg.Text.Length >= 32) And (txtReceiMsg.Text.Length <= 40) And (txtResult.Text.Length <= 12)) Then
                Dim subst As String = txtReceiMsg.Text.Substring(0, 6)
                '(txtReceiMsg.Text.Length <= 10) And
                If subst = 10311 Then

                    Dim intCount As String = 9999999

                    If (txtResult.Text <= intCount) Then
                        Dim intCounts As String = txtResult.Text / 1000000

                        Dim Filtering_number As String = 65.536

                        If (intCounts <= Filtering_number) Then
                            '   Dim numberOfRecords As String
                            '   numberOfRecords = "Select tresult from mydata where tresult =' " & txtResult.Text & " " '"
                            '   If numberOfRecords <> 0 Then
                            '     MessageBox.Show("Data ")

                            ' Else
                            MysqlConn.Open()
                            Dim Query As String
                            Label4.Text = Date.Now.ToString("dd MMM yyyy hh:mm:ss")
                            Query = "insert into mydata(treceimsg,tresult,DateTime)
                        values ('" & txtReceiMsg.Text & "','" & intCounts & "' ,'" & Label4.Text & "')"
                            COMMAND = New MySqlCommand(Query, MysqlConn)
                            READER = COMMAND.ExecuteReader

                            MysqlConn.Close()
                        End If
                    End If
                End If
            End If
            ' End If
        Catch ex As Exception
            ' MessageBox.Show(Me, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub



    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        Label4.Text = Date.Now.ToString("dd MMM yyyy hh:mm:ss")
    End Sub
End Class
