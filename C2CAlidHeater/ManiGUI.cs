
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using System.IO.Ports;
using System.Diagnostics;
using System.Reflection;
using System.Globalization;
using System.Configuration;

namespace C2CAlidHeater
{
    public partial class ManiGUI : Form
    {
        private SerialPort serialPort1;
        private string ComPort = string.Empty;
        private string rxString = string.Empty;
        private bool rxFinished = false;
        private uint seq = 0;
        private uint seqNo = 100;
        private List<string> comPort_list = new List<string>();
        private bool readingFromAtMega = false;

        private bool PgainCh0_editing = false;
        private bool IgainCh0_editing = false;
        private bool DgainCh0_editing = false;
        private bool TempSensor0_editing = false;
        private bool TempWindow0_editing = false;
        private bool SettleTime0_editing = false;

        private bool PgainCh1_editing = false;
        private bool IgainCh1_editing = false;
        private bool DgainCh1_editing = false;
        private bool TempSensor1_editing = false;
        private bool TempWindow1_editing = false;
        private bool SettleTime1_editing = false;

        private bool PgainCh2_editing = false;
        private bool IgainCh2_editing = false;
        private bool DgainCh2_editing = false;
        private bool TempSensor2_editing = false;
        private bool TempWindow2_editing = false;
        private bool SettleTime2_editing = false;

        private bool PgainCh3_editing = false;
        private bool IgainCh3_editing = false;
        private bool DgainCh3_editing = false;
        private bool TempSensor3_editing = false;
        private bool TempWindow3_editing = false;
        private bool SettleTime3_editing = false;

        private const int writeTimeOut = 200;     // [ms]

        private const int TempSensor0       = 100;
        private const int P_ch0             = 101;
        private const int I_ch0             = 102;
        private const int D_ch0             = 103;
        private const int TempSetPoint0     = 104;
        private const int PgainCh0          = 105;
        private const int IgainCh0          = 106;
        private const int DgainCh0          = 107;
        private const int TempWindowCh0     = 108;
        private const int SettleTimeTempCh0 = 109;
        private const int TempStableCh0     = 110;
        private const int SetTempSetPoint0  = 150;
        private const int SetPgainCh0       = 151;
        private const int SetIgainCh0       = 152;
        private const int SetDgainCh0       = 153;
        private const int SetHeaterOnOffCh0 = 154;
        private const int SetTempWindowCh0  = 155;
        private const int SetSettleTimeT0   = 156;

        
        private const int TempSensor1       = 200;
        private const int P_ch1             = 201;
        private const int I_ch1             = 202;
        private const int D_ch1             = 203;
        private const int TempSetPoint1     = 204;
        private const int PgainCh1          = 205;
        private const int IgainCh1          = 206;
        private const int DgainCh1          = 207;
        private const int TempWindowCh1     = 208;
        private const int SettleTimeTempCh1 = 209;
        private const int TempStableCh1     = 210;
        private const int SetTempSetPoint1  = 250;
        private const int SetPgainCh1       = 251;
        private const int SetIgainCh1       = 252;
        private const int SetDgainCh1       = 253;
        private const int SetHeaterOnOffCh1 = 254;
        private const int SetTempWindowCh1  = 255;
        private const int SetSettleTimeT1   = 256;

        private const int TempSensor2       = 300;
        private const int P_ch2             = 301;
        private const int I_ch2             = 302;
        private const int D_ch2             = 303;
        private const int TempSetPoint2     = 304;
        private const int PgainCh2          = 305;
        private const int IgainCh2          = 306;
        private const int DgainCh2          = 307;
        private const int TempWindowCh2     = 308;
        private const int SettleTimeTempCh2 = 309;
        private const int TempStableCh2     = 310;
        private const int SetTempSetPoint2  = 350;
        private const int SetPgainCh2       = 351;
        private const int SetIgainCh2       = 352;
        private const int SetDgainCh2       = 353;
        private const int SetHeaterOnOffCh2 = 354;
        private const int SetTempWindowCh2  = 355;
        private const int SetSettleTimeT2   = 356;

        private const int TempSensor3       = 400;
        private const int P_ch3             = 401;
        private const int I_ch3             = 402;
        private const int D_ch3             = 403;
        private const int TempSetPoint3     = 404;
        private const int PgainCh3          = 405;
        private const int IgainCh3          = 406;
        private const int DgainCh3          = 407;
        private const int TempWindowCh3     = 408;
        private const int SettleTimeTempCh3 = 409;
        private const int TempStableCh3     = 410;
        private const int SetTempSetPoint3  = 450;
        private const int SetPgainCh3       = 451;
        private const int SetIgainCh3       = 452;
        private const int SetDgainCh3       = 453;
        private const int SetHeaterOnOffCh3 = 454;
        private const int SetTempWindowCh3  = 455;
        private const int SetSettleTimeT3   = 456;

        private const int WrtParamToEEPROM  = 500;

        Configuration configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

        public ManiGUI()
        {
            InitializeComponent();
            ShowComPort();
            GetComPorts();
            //ComPortSearch();    // Remove before release
        }

        private void GetComPorts()
        {
            foreach (string s in SerialPort.GetPortNames())
            {
                comPort_list.Add(s);
            }
        }

        private void ShowComPort()
        {
            //show list of valid com ports
            foreach (string s in SerialPort.GetPortNames())
            {
                listBox_comPorts.Items.Add(s);
            }
        }

        private void ReadParams()
        {
            string dataRx = string.Empty;

            // *** Channel 0 ***
            dataRx = ReadFloat(TempSetPoint0);        
            if (dataRx != null)
            {
                textBox_SetPointCh0.Text = dataRx;
            }

            dataRx = ReadFloat(PgainCh0);        
            if (dataRx != null)
            {
                textBox_PgainCh0.Text = dataRx;
            }

            dataRx = ReadFloat(IgainCh0);
            if (dataRx != null)
            {
                textBox_IgainCh0.Text = dataRx;
            }

            dataRx = ReadFloat(DgainCh0);
            if (dataRx != null)
            {
                textBox_DgainCh0.Text = dataRx;
            }

            dataRx = ReadFloat(TempWindowCh0);
            if (dataRx != null)
            {
                textBox_TempWindowCh0.Text = dataRx;
            }

            dataRx = ReadFloat(SettleTimeTempCh0);
            if (dataRx != null)
            {
                textBox_SettleTimeCh0.Text = dataRx;
            }

            dataRx = ReadFloat(TempStableCh0);
            if (dataRx != null)
            {
                if(dataRx == "0")
                {
                    radioButton_tempStableCh0.Checked = false;
                }
                else if(dataRx == "1")
                {
                    radioButton_tempStableCh0.Checked = true;
                }
            }

            // *** Channel 1 ***
            dataRx = ReadFloat(TempSetPoint1);
            if (dataRx != null)
            {
                textBox_SetPointCh1.Text = dataRx;
            }

            dataRx = ReadFloat(PgainCh1);
            if (dataRx != null)
            {
                textBox_PgainCh1.Text = dataRx;
            }

            dataRx = ReadFloat(IgainCh1);
            if (dataRx != null)
            {
                textBox_IgainCh1.Text = dataRx;
            }

            dataRx = ReadFloat(DgainCh1);
            if (dataRx != null)
            {
                textBox_DgainCh1.Text = dataRx;
            }

            dataRx = ReadFloat(TempWindowCh1);
            if (dataRx != null)
            {
                textBox_TempWindowCh1.Text = dataRx;
            }

            dataRx = ReadFloat(SettleTimeTempCh1);
            if (dataRx != null)
            {
                textBox_SettleTimeCh1.Text = dataRx;
            }

            dataRx = ReadFloat(TempStableCh1);
            if (dataRx != null)
            {
                if (dataRx == "0")
                {
                    radioButton_tempStableCh1.Checked = false;
                }
                else if (dataRx == "1")
                {
                    radioButton_tempStableCh1.Checked = true;
                }
            }

            // *** Channel 2 ***
            dataRx = ReadFloat(TempSetPoint2);
            if (dataRx != null)
            {
                textBox_SetPointCh2.Text = dataRx;
            }

            dataRx = ReadFloat(PgainCh2);
            if (dataRx != null)
            {
                textBox_PgainCh2.Text = dataRx;
            }

            dataRx = ReadFloat(IgainCh2);
            if (dataRx != null)
            {
                textBox_IgainCh2.Text = dataRx;
            }

            dataRx = ReadFloat(DgainCh2);
            if (dataRx != null)
            {
                textBox_DgainCh2.Text = dataRx;
            }

            dataRx = ReadFloat(TempWindowCh2);
            if (dataRx != null)
            {
                textBox_TempWindowCh2.Text = dataRx;
            }

            dataRx = ReadFloat(SettleTimeTempCh2);
            if (dataRx != null)
            {
                textBox_SettleTimeCh2.Text = dataRx;
            }

            dataRx = ReadFloat(TempStableCh2);
            if (dataRx != null)
            {
                if (dataRx == "0")
                {
                    radioButton_tempStableCh2.Checked = false;
                }
                else if (dataRx == "1")
                {
                    radioButton_tempStableCh2.Checked = true;
                }
            }

            // *** Channel 3 ***
            dataRx = ReadFloat(TempSetPoint3);
            if (dataRx != null)
            {
                textBox_SetPointCh3.Text = dataRx;
            }

            dataRx = ReadFloat(PgainCh3);
            if (dataRx != null)
            {
                textBox_PgainCh3.Text = dataRx;
            }

            dataRx = ReadFloat(IgainCh3);
            if (dataRx != null)
            {
                textBox_IgainCh3.Text = dataRx;
            }

            dataRx = ReadFloat(DgainCh3);
            if (dataRx != null)
            {
                textBox_DgainCh3.Text = dataRx;
            }

            dataRx = ReadFloat(TempWindowCh3);
            if (dataRx != null)
            {
                textBox_TempWindowCh3.Text = dataRx;
            }

            dataRx = ReadFloat(SettleTimeTempCh3);
            if (dataRx != null)
            {
                textBox_SettleTimeCh3.Text = dataRx;
            }

            dataRx = ReadFloat(TempStableCh3);
            if (dataRx != null)
            {
                if (dataRx == "0")
                {
                    radioButton_tempStableCh3.Checked = false;
                }
                else if (dataRx == "1")
                {
                    radioButton_tempStableCh3.Checked = true;
                }
            }
        }

        private void MainGUI_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (serialPort1.IsOpen == true)
            {
                serialPort1.Close();
            }
        }

        private bool ConnectToSerialPort(string portName)
        {
            serialPort1 = new SerialPort(portName);
            bool openSuccess = false;

            if (!serialPort1.IsOpen)
            {
                try
                {
                    serialPort1.BaudRate = 57600;
                    serialPort1.DataBits = 8;
                    serialPort1.DiscardNull = false;
                    serialPort1.DtrEnable = false;
                    serialPort1.Handshake = Handshake.None;
                    serialPort1.Parity = Parity.None;
                    serialPort1.ParityReplace = 63;
                    serialPort1.ReadBufferSize = 4096;
                    serialPort1.ReadTimeout = 5;
                    serialPort1.ReceivedBytesThreshold = 1;
                    serialPort1.RtsEnable = false;
                    serialPort1.StopBits = StopBits.One;
                    serialPort1.WriteBufferSize = 2048;
                    serialPort1.WriteTimeout = 10;
                    serialPort1.Open();

                    //serialPort1.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(_serialPort1_DataRx);
                    openSuccess = true;
                }
                catch (Exception e)
                {
                    openSuccess = false;
                    //MessageBox.Show(e.Message + "\n\r" + "Could not open ", portName);
                }
            }
            return openSuccess;
        }

        private void _serialPort1_DataRx(object sender, SerialDataReceivedEventArgs e)
        {
            //Initialize a buffer to hold the received data 
            byte[] buffer = new byte[serialPort1.ReadBufferSize];

            //There is no accurate method for checking how many bytes are read 
            //unless you check the return from the Read method 
            int bytesRead = serialPort1.Read(buffer, 0, buffer.Length);

            //For the example assume the data we are received is ASCII data. 
            rxString += Encoding.ASCII.GetString(buffer, 0, bytesRead);

            if (rxString.Contains("\r"))
            {
                rxFinished = true;
            }
            //Debug.WriteLine(rxString);

        }

        private void listBox_comPorts_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComPort = listBox_comPorts.SelectedItem.ToString();
        }

        private string CrcAdd(string dataTX)
        {
            string tx = string.Empty;
            int stringLength;
            byte[] chArr;
            ushort crc;

            stringLength = dataTX.Length;
            chArr = System.Text.ASCIIEncoding.Default.GetBytes(dataTX);
            crc = GenCrc16(chArr, stringLength);
            dataTX += crc.ToString("X4");
            dataTX += "\r";

            return dataTX;
        }

        private bool CrcCompare(string data_string)
        {
            int crc_rx = 0;

            if(data_string.Length > 4)
            {
                string crc_string = data_string.Substring(data_string.Length - 5, 4);   // Isolate crc from string
                string rx_msg = data_string.Substring(0, data_string.Length - 5);       // Exclude crc from string
                ushort crc_gen = GenCrc16(System.Text.ASCIIEncoding.Default.GetBytes(rx_msg), rx_msg.Length); // calculate received crc
                if(int.TryParse(crc_string, NumberStyles.HexNumber, CultureInfo.CurrentCulture, out crc_rx))
                { 
                    if(crc_gen == crc_rx)
                    {
                        //Debug.WriteLine(" CRC true ");
                        return true;
                    }
                    else
                    {
                        Debug.WriteLine(" CRC false ");
                        return false;
                    }
                }
                else
                {
                    Debug.WriteLine(" CRC unparseable ");
                    return false;
                }
            }
            else
            {
                Debug.WriteLine(" String to short ");
                return false;
            }
            
        }

        private ushort GenCrc16(byte[] c, int nByte)
        {
            ushort Polynominal = 0x1021;
            ushort InitValue = 0x0;

            ushort i, j, index = 0;
            ushort CRC = InitValue;
            ushort Remainder, tmp, short_c;
            for (i = 0; i < nByte; i++)
            {
                short_c = (ushort)(0x00ff & (ushort)c[index]);
                tmp = (ushort)((CRC >> 8) ^ short_c);
                Remainder = (ushort)(tmp << 8);
                for (j = 0; j < 8; j++)
                {


                    if ((Remainder & 0x8000) != 0)
                    {
                        Remainder = (ushort)((Remainder << 1) ^ Polynominal);
                    }
                    else
                    {
                        Remainder = (ushort)(Remainder << 1);
                    }
                }
                CRC = (ushort)((CRC << 8) ^ Remainder);
                index++;
            }
            return CRC;
        }

        private uint SeqNoGen()
        {
            seqNo++;
            if (seqNo >= 1000)
            {
                seqNo = 100;
            }
            return seqNo;
        }

        private string ReadFloat(uint id)
        {
            readingFromAtMega = true;
            string output_string = null;
            uint seqNo_return = 0;

            uint seqNo = SeqNoGen();

            serialPort1.Write(CrcAdd("#" + seqNo.ToString() + "?VR" + id.ToString()));
            RxWait();

            if (rxString != "")
            {
                if (rxString[0] == '!')
                {
                    if (uint.TryParse(rxString.Substring(1, 3), out seqNo_return))
                    {
                        if (seqNo == seqNo_return)
                        {
                            if (CrcCompare(rxString) == true)
                            {
                                output_string = rxString.Substring(4,rxString.Length - 9);
                            }
                        }
                    }
                    else
                    {
                        output_string = null;
                    }
                }
            }
            else
            {
                output_string = null;
            }
            rxString = string.Empty;
            readingFromAtMega = false;
            return output_string;
        }

        private bool WriteFloat(string d_string, int id)
        {
            bool return_val = false;
            uint seqNo_return = 0;
            bool timerScanStatus = timer_Scan.Enabled;
            var timeSpan = Stopwatch.StartNew();

            while (readingFromAtMega == true)
            {
                if (timeSpan.ElapsedMilliseconds > writeTimeOut)
                {
                    timeSpan.Stop();
                    return false;
                }
            }
            timeSpan.Stop();
            if (timer_Scan.Enabled)
            {
                timer_Scan.Enabled = false;
            }

            uint seqNo = SeqNoGen();
            serialPort1.Write(CrcAdd("#" +  seqNo.ToString() + "&VS" + id.ToString() + d_string));
            RxWait();

            if (rxString != "")
            {
                if (rxString[0] == '!')
                {
                    if (uint.TryParse(rxString.Substring(1, 3), out seqNo_return))
                    {
                        if (seqNo == seqNo_return)
                        {
                            return_val = true;
                        }
                    }
                    else
                    {
                        return_val = false;
                    }
                }
                else
                {
                    return_val = false;
                }
            }
            else
            {
               return_val = false;
            }

            rxString = string.Empty;
            timer_Scan.Enabled = timerScanStatus;
            return return_val;
        }

        private string WriteFloatRetString(string d_string, int id)
        {
            string return_val = "false";
            uint seqNo_return = 0;
            bool timerScanStatus = timer_Scan.Enabled;

            if (timer_Scan.Enabled)
            {
                timer_Scan.Enabled = false;
            }

            var timeSpan = Stopwatch.StartNew();
            while (readingFromAtMega == true)
            {
                if (timeSpan.ElapsedMilliseconds > writeTimeOut)
                {
                    timeSpan.Stop();
                    return "false";                    
                }
            }
            timeSpan.Stop();

            uint seqNo = SeqNoGen();
            serialPort1.Write(CrcAdd("#" + seqNo.ToString() + "&VS" + id.ToString() + d_string));
            RxWait();

            if (rxString != "")
            {
                if (rxString[0] == '!')
                {
                    if (uint.TryParse(rxString.Substring(1, 3), out seqNo_return))
                    {
                        if (seqNo == seqNo_return)
                        {
                            return_val = rxString;
                        }
                    }
                    else
                    {
                        return_val = "false";
                    }
                }
                else
                {
                    return_val = "false";
                }
            }
            else
            {
                return_val = "false";
            }

            rxString = string.Empty;
            timer_Scan.Enabled = timerScanStatus;
            return return_val;
        }

        private bool WriteParamToEEPROM(int id)
        {
            bool return_val = false;
            uint seqNo_return = 0;
            bool timerScanStatus = timer_Scan.Enabled;
            var timeSpan = Stopwatch.StartNew();

            while(readingFromAtMega == true)
            {
                if (timeSpan.ElapsedMilliseconds > writeTimeOut)
                {
                    return false;
                }
            }
            if(timer_Scan.Enabled)
            {
                timer_Scan.Enabled = false;
            }

            uint seqNo = SeqNoGen();
            serialPort1.Write(CrcAdd("#" +  seqNo.ToString() + "&VS" + id.ToString()));
            RxWait();

            if (rxString != "")
            {
                if (rxString[0] == '!')
                {
                    if (uint.TryParse(rxString.Substring(1, 3), out seqNo_return))
                    {
                        if (seqNo == seqNo_return)
                        {
                            return_val = true;
                        }
                    }
                    else
                    {
                        return_val = false;
                    }
                }
                else
                {
                    return_val = false;
                }
            }
            else
            {
               return_val = false;
            }

            rxString = string.Empty;
            timer_Scan.Enabled = timerScanStatus;
            return return_val;
        }
        
        private void RxWait()
        {
            rxFinished = false;
            var timeSpan = Stopwatch.StartNew();

            while (rxFinished == false && timeSpan.ElapsedMilliseconds < 150)     // Waiting for all received data or timeout
            {
            }
            timeSpan.Stop();
            if(timeSpan.ElapsedMilliseconds >= 150)
            {
                Debug.WriteLine("Reception time out");
            }
            //Debug.Write(rxString);
            //Debug.Write(rxFinished + " ");
            //Debug.Write(timeSpan.ElapsedMilliseconds + " ");
        }

        private void RxWait(uint time)
        {
            rxFinished = false;
            var timeSpan = Stopwatch.StartNew();

            while (rxFinished == false && timeSpan.ElapsedMilliseconds < time)     // Waiting for all received data or timeout
            {
            }
            timeSpan.Stop();
        }

        private void ComPortSearch()
        {
            if (configuration.AppSettings.Settings["comPort"].Value.Contains("COM"))
            {
                ConnectToSerialPort(configuration.AppSettings.Settings["comPort"].Value);
                try
                {
                    textBox_ComPort.Text = "Try connect to " + configuration.AppSettings.Settings["comPort"].Value;
                    serialPort1.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(_serialPort1_DataRx);
                    serialPort1.Write(CrcAdd("#123?IF"));
                }
                catch (Exception e)
                {
                    //MessageBox.Show(e.Message);
                }
                RxWait(1000);
                if (rxString.Contains("!123C2CA"))
                {
                    textBox_ComPort.Text = "Connected to " + configuration.AppSettings.Settings["comPort"].Value;
                    rxString = string.Empty;

                    ReadParams();
                }
            }
            else
            {
                foreach (string comPort in comPort_list)
                {
                    if (ConnectToSerialPort(comPort))
                    {
                        textBox_ComPort.Text = "Try connect to " + comPort;
                        try
                        {
                            serialPort1.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(_serialPort1_DataRx);
                            serialPort1.Write(CrcAdd("#123?IF"));
                        }
                        catch (Exception e)
                        {
                            //MessageBox.Show(e.Message);
                        }
                        RxWait(1000);
                        if (rxString.Contains("!123C2CA"))
                        {
                            textBox_ComPort.Text = "Connected to " + comPort;
                            rxString = string.Empty;

                            configuration.AppSettings.Settings["comPort"].Value = comPort;
                            configuration.Save(ConfigurationSaveMode.Full, true);

                            ReadParams();
                            break;
                        }
                        else
                        {
                            serialPort1.DataReceived -= new System.IO.Ports.SerialDataReceivedEventHandler(_serialPort1_DataRx);
                            try
                            {
                                serialPort1.Close();
                            }
                            catch (Exception e)
                            {
                               // MessageBox.Show(e.Message);
                            }
                        }
                    }
                }
            }
        }

        private void timer_Scan_Tick(object sender, EventArgs e)
        {
            string dataRx = string.Empty;
            float f;

            switch(seq)
            {
                // *** Channel 0 ***
                case 0:
                    dataRx = ReadFloat(TempSensor0);
                    if (dataRx != null)
                    {
                        textBox_TempCh0.Text = dataRx;
                        if(float.TryParse(dataRx, NumberStyles.AllowDecimalPoint, CultureInfo.GetCultureInfo("en-US"), out f))
                        {
                            chart1.Series["Ch 0"].Points.AddY(f);
                        }                        
                    }
                    seq = 101;
                    break;

                case 101:
                    dataRx = ReadFloat(P_ch0);        
                    if (dataRx != null)
                    {
                        textBox_Pch0.Text = dataRx;
                    }
                    seq = 102;
                    break;

                case 102:
                    dataRx = ReadFloat(I_ch0);
                    if (dataRx != null)
                    {
                        textBox_Ich0.Text = dataRx;
                    }
                    seq = 103;
                    break;

                case 103:
                    dataRx = ReadFloat(D_ch0);
                    if (dataRx != null)
                    {
                        textBox_Dch0.Text = dataRx;
                    }
                    seq = 104;
                    break;

                case 104:
                    dataRx = ReadFloat(TempStableCh0);
                    if (dataRx != null)
                    {
                        if(dataRx == "0")
                        {
                            radioButton_tempStableCh0.Checked = false;
                        }
                        else if(dataRx == "1")
                        {
                            radioButton_tempStableCh0.Checked = true;
                        }
                    }
                    seq = 200;
                    break;

                // *** Channel 1 ***
                case 200:
                    dataRx = ReadFloat(TempSensor1);
                    if (dataRx != null)
                    {
                        textBox_TempCh1.Text = dataRx;
                        if (float.TryParse(dataRx, NumberStyles.AllowDecimalPoint, CultureInfo.GetCultureInfo("en-US"), out f))
                        {
                            chart1.Series["Ch 1"].Points.AddY(f);
                        }
                    }
                    seq = 201;
                    break;

                case 201:
                    dataRx = ReadFloat(P_ch1);
                    if (dataRx != null)
                    {
                        textBox_Pch1.Text = dataRx;
                    }
                    seq = 202;
                    break;

                case 202:
                    dataRx = ReadFloat(I_ch1);
                    if (dataRx != null)
                    {
                        textBox_Ich1.Text = dataRx;
                    }
                    seq = 203;
                    break;

                case 203:
                    dataRx = ReadFloat(D_ch1);
                    if (dataRx != null)
                    {
                        textBox_Dch1.Text = dataRx;
                    }
                    seq = 204;
                    break;

                case 204:
                    dataRx = ReadFloat(TempStableCh1);
                    if (dataRx != null)
                    {
                        if (dataRx == "0")
                        {
                            radioButton_tempStableCh1.Checked = false;
                        }
                        else if (dataRx == "1")
                        {
                            radioButton_tempStableCh1.Checked = true;
                        }
                    }
                    seq = 300;
                    break;

                // *** Channel 2 ***
                case 300:
                    dataRx = ReadFloat(TempSensor2);
                    if (dataRx != null)
                    {
                        textBox_TempCh2.Text = dataRx;
                        if (float.TryParse(dataRx, NumberStyles.AllowDecimalPoint, CultureInfo.GetCultureInfo("en-US"), out f))
                        {
                            chart1.Series["Ch 2"].Points.AddY(f);
                        }
                    }
                    seq = 301;
                    break;

                case 301:
                    dataRx = ReadFloat(P_ch2);
                    if (dataRx != null)
                    {
                        textBox_Pch2.Text = dataRx;
                    }
                    seq = 302;
                    break;

                case 302:
                    dataRx = ReadFloat(I_ch2);
                    if (dataRx != null)
                    {
                        textBox_Ich2.Text = dataRx;
                    }
                    seq = 303;
                    break;

                case 303:
                    dataRx = ReadFloat(D_ch2);
                    if (dataRx != null)
                    {
                        textBox_Dch2.Text = dataRx;            }
                    seq = 304;
                    break;

                case 304:
                    dataRx = ReadFloat(TempStableCh2);
                    if (dataRx != null)
                    {
                        if (dataRx == "0")
                        {
                            radioButton_tempStableCh2.Checked = false;
                        }
                        else if (dataRx == "1")
                        {
                            radioButton_tempStableCh2.Checked = true;
                        }
                    }
                    seq = 400;
                    break;

                // *** Channel 3 ***
                case 400:
                    dataRx = ReadFloat(TempSensor3);
                    if (dataRx != null)
                    {
                        textBox_TempCh3.Text = dataRx;
                        if (float.TryParse(dataRx, NumberStyles.AllowDecimalPoint, CultureInfo.GetCultureInfo("en-US"), out f))
                        {
                            chart1.Series["Ch 3"].Points.AddY(f);
                        }
                    }
                    seq = 401;
                    break;

                case 401:
                    dataRx = ReadFloat(P_ch3);
                    if (dataRx != null)
                    {
                        textBox_Pch3.Text = dataRx;
                    }
                    seq = 402;
                    break;

                case 402:
                    dataRx = ReadFloat(I_ch3);
                    if (dataRx != null)
                    {
                        textBox_Ich3.Text = dataRx;
                    }
                    seq = 403;
                    break;

                case 403:
                    dataRx = ReadFloat(D_ch3);
                    if (dataRx != null)
                    {
                        textBox_Dch3.Text = dataRx;
                    }
                    seq = 404;
                    break;

                case 404:
                    dataRx = ReadFloat(TempStableCh3);
                    if (dataRx != null)
                    {
                        if (dataRx == "0")
                        {
                            radioButton_tempStableCh3.Checked = false;
                        }
                        else if (dataRx == "1")
                        {
                            radioButton_tempStableCh3.Checked = true;
                        }
                    }
                    seq = 0;
                    break;
            }
        }

        private string CheckFloatAndRetString(string param)
        {
            float f;
            string s;
            if (float.TryParse(param, NumberStyles.AllowDecimalPoint, CultureInfo.GetCultureInfo("en-US"), out f) ||
                float.TryParse(param, NumberStyles.AllowDecimalPoint, CultureInfo.CurrentCulture, out f))
            {
                s = f.ToString("0.000");
                if(s.Contains(','))
                {
                    s = s.Replace(",", ".");
                    return s;
                }
                else
                {
                    return s;
                }
            }
            else
            {
                return null;
            }
        }

        private string CheckIntAndRetString(string param)
        {
            int i;
            string s;
            if (int.TryParse(param, out i))
            {
                i = Math.Abs(i);
                //if(i > 10)
                //{
                //    i = 10;
                //}
                s = i.ToString("0");
                return s;
            }
            else
            {
                return null;
            }
        }

        private void button_Scan_Click(object sender, EventArgs e)
        {
            timer_Scan.Enabled = !timer_Scan.Enabled;
        }

        private void button_Connect_Click(object sender, EventArgs e)
        {
            ComPortSearch();
        }

        private void button_SetParam_Click(object sender, EventArgs e)
        {
            WriteParamToEEPROM(WrtParamToEEPROM);
        }

        private void button_Refresh_Click(object sender, EventArgs e)
        {
            ReadParams();
        }

        private void button_HeatOnOffCh0_Click(object sender, EventArgs e)
        {
            if (!radioButton_HeatOnOffCh0.Checked)
            {
                if (WriteFloatRetString("1", SetHeaterOnOffCh0).Contains("ON") == true)
                {
                    radioButton_HeatOnOffCh0.Checked = true;
                }
            }
            else
            {
                if (WriteFloatRetString("0", SetHeaterOnOffCh0).Contains("OFF") == true)
                {
                    radioButton_HeatOnOffCh0.Checked = false;
                }
            }
        }

        private void button_ClrChart_Click(object sender, EventArgs e)
        {
            foreach (var series in chart1.Series)
            {
                series.Points.Clear();
            }
        }

        private void textBox_SetPointCh0_KeyDown(object sender, KeyEventArgs e)
        {
            string dataRx;

            if (e.KeyCode == Keys.Enter)    // Enter key writes data to controller
            {
                string param = CheckFloatAndRetString(textBox_SetPointCh0.Text);
                if (param != null)
                {
                    WriteFloat(param, SetTempSetPoint0);
                    textBox_SetPointCh0.Text = param;
                    TempSensor0_editing = false;
                }
            }
            else
            {
                if (e.KeyCode == Keys.Escape)   // Escape key reads back value from controller
                {
                    dataRx = ReadFloat(TempSetPoint0);
                    if (dataRx != null)
                    {
                        textBox_SetPointCh0.Text = dataRx;
                    }
                    TempSensor0_editing = false;
                }
                else if (TempSensor0_editing == false) // Any other key purge the input field
                {
                    textBox_SetPointCh0.Clear();
                    TempSensor0_editing = true;
                }
            }
        }

        private void textBox_PgainCh0_KeyDown(object sender, KeyEventArgs e)
        {
            string dataRx;

            if (e.KeyCode == Keys.Enter)    // Enter key writes data to controller
            {
                string param = CheckFloatAndRetString(textBox_PgainCh0.Text);
                if(param != null)
                {
                    WriteFloat(param, SetPgainCh0);
                    textBox_PgainCh0.Text = param;
                    PgainCh0_editing = false;
                }
            }
            else
            {
                if (e.KeyCode == Keys.Escape)   // Escape key reads back value from controller
                {
                    dataRx = ReadFloat(PgainCh0);
                    if (dataRx != null)
                    {
                        textBox_PgainCh0.Text = dataRx;
                    }
                    PgainCh0_editing = false;
                }
                else if (PgainCh0_editing == false) // Any other key purge the input field
                {
                    textBox_PgainCh0.Clear();
                    PgainCh0_editing = true;
                }
            }
        }

        private void textBox_IgainCh0_KeyDown(object sender, KeyEventArgs e)
        {
            string dataRx;

            if (e.KeyCode == Keys.Enter)    // Enter key writes data to controller
            {
                string param = CheckFloatAndRetString(textBox_IgainCh0.Text);
                if (param != null)
                {
                    WriteFloat(param, SetIgainCh0);
                    textBox_IgainCh0.Text = param;
                    IgainCh0_editing = false;
                }
            }
            else
            {
                if (e.KeyCode == Keys.Escape)   // Escape key reads back value from controller
                {
                    dataRx = ReadFloat(IgainCh0);
                    if (dataRx != null)
                    {
                        textBox_IgainCh0.Text = dataRx;
                    }
                    IgainCh0_editing = false;
                }
                else if (IgainCh0_editing == false) // Any other key purge the input field
                {
                    textBox_IgainCh0.Clear();
                    IgainCh0_editing = true;
                }
            }
        }

        private void textBox_DgainCh0_KeyDown(object sender, KeyEventArgs e)
        {
            string dataRx;

            if (e.KeyCode == Keys.Enter)    // Enter key writes data to controller
            {
                string param = CheckFloatAndRetString(textBox_DgainCh0.Text);
                if (param != null)
                {
                    WriteFloat(param, SetDgainCh0);
                    textBox_DgainCh0.Text = param;
                    DgainCh0_editing = false;
                }
            }
            else
            {
                if (e.KeyCode == Keys.Escape)   // Escape key reads back value from controller
                {
                    dataRx = ReadFloat(DgainCh0);
                    if (dataRx != null)
                    {
                        textBox_DgainCh0.Text = dataRx;
                    }
                    DgainCh0_editing = false;
                }
                else if (DgainCh0_editing == false) // Any other key purge the input field
                {
                    textBox_DgainCh0.Clear();
                    DgainCh0_editing = true;
                }
            }
        }

        private void textBox_SetPointCh1_KeyDown(object sender, KeyEventArgs e)
        {
            string dataRx;

            if (e.KeyCode == Keys.Enter)    // Enter key writes data to controller
            {
                string param = CheckFloatAndRetString(textBox_SetPointCh1.Text);
                if (param != null)
                {
                    WriteFloat(param, SetTempSetPoint1);
                    textBox_SetPointCh1.Text = param;
                    TempSensor1_editing = false;
                }
            }
            else
            {
                if (e.KeyCode == Keys.Escape)   // Escape key reads back value from controller
                {
                    dataRx = ReadFloat(TempSetPoint1);
                    if (dataRx != null)
                    {
                        textBox_SetPointCh1.Text = dataRx;
                    }
                    TempSensor1_editing = false;
                }
                else if (TempSensor1_editing == false) // Any other key purge the input field
                {
                    textBox_SetPointCh1.Clear();
                    TempSensor1_editing = true;
                }
            }
        }

        private void textBox_PgainCh1_KeyDown(object sender, KeyEventArgs e)
        {
            string dataRx;

            if (e.KeyCode == Keys.Enter)    // Enter key writes data to controller
            {
                string param = CheckFloatAndRetString(textBox_PgainCh1.Text);
                if (param != null)
                {
                    WriteFloat(param, SetPgainCh1);
                    textBox_PgainCh1.Text = param;
                    PgainCh1_editing = false;
                }
            }
            else
            {
                if (e.KeyCode == Keys.Escape)   // Escape key reads back value from controller
                {
                    dataRx = ReadFloat(PgainCh1);
                    if (dataRx != null)
                    {
                        textBox_PgainCh1.Text = dataRx;
                    }
                    PgainCh1_editing = false;
                }
                else if (PgainCh1_editing == false) // Any other key purge the input field
                {
                    textBox_PgainCh1.Clear();
                    PgainCh1_editing = true;
                }
            }
        }

        private void textBox_IgainCh1_KeyDown(object sender, KeyEventArgs e)
        {
            string dataRx;

            if (e.KeyCode == Keys.Enter)    // Enter key writes data to controller
            {
                string param = CheckFloatAndRetString(textBox_IgainCh1.Text);
                if (param != null)
                {
                    WriteFloat(param, SetIgainCh1);
                    textBox_IgainCh1.Text = param;
                    IgainCh1_editing = false;
                }
            }
            else
            {
                if (e.KeyCode == Keys.Escape)   // Escape key reads back value from controller
                {
                    dataRx = ReadFloat(IgainCh1);
                    if (dataRx != null)
                    {
                        textBox_IgainCh1.Text = dataRx;
                    }
                    IgainCh1_editing = false;
                }
                else if (IgainCh1_editing == false) // Any other key purge the input field
                {
                    textBox_IgainCh1.Clear();
                    IgainCh1_editing = true;
                }
            }
        }

        private void textBox_DgainCh1_KeyDown(object sender, KeyEventArgs e)
        {
            string dataRx;

            if (e.KeyCode == Keys.Enter)    // Enter key writes data to controller
            {
                string param = CheckFloatAndRetString(textBox_DgainCh1.Text);
                if (param != null)
                {
                    WriteFloat(param, SetDgainCh1);
                    textBox_DgainCh1.Text = param;
                    DgainCh1_editing = false;
                }
            }
            else
            {
                if (e.KeyCode == Keys.Escape)   // Escape key reads back value from controller
                {
                    dataRx = ReadFloat(DgainCh1);
                    if (dataRx != null)
                    {
                        textBox_DgainCh1.Text = dataRx;
                    }
                    DgainCh1_editing = false;
                }
                else if (DgainCh1_editing == false) // Any other key purge the input field
                {
                    textBox_DgainCh1.Clear();
                    DgainCh1_editing = true;
                }
            }
        }

        private void textBox_SetPointCh2_KeyDown(object sender, KeyEventArgs e)
        {
            string dataRx;

            if (e.KeyCode == Keys.Enter)    // Enter key writes data to controller
            {
                string param = CheckFloatAndRetString(textBox_SetPointCh2.Text);
                if (param != null)
                {
                    WriteFloat(param, SetTempSetPoint2);
                    textBox_SetPointCh2.Text = param;
                    TempSensor2_editing = false;
                }
            }
            else
            {
                if (e.KeyCode == Keys.Escape)   // Escape key reads back value from controller
                {
                    dataRx = ReadFloat(TempSetPoint2);
                    if (dataRx != null)
                    {
                        textBox_SetPointCh2.Text = dataRx;
                    }
                    TempSensor2_editing = false;
                }
                else if (TempSensor2_editing == false) // Any other key purge the input field
                {
                    textBox_SetPointCh2.Clear();
                    TempSensor2_editing = true;
                }
            }
        }

        private void textBox_PgainCh2_KeyDown(object sender, KeyEventArgs e)
        {
            string dataRx;

            if (e.KeyCode == Keys.Enter)    // Enter key writes data to controller
            {
                string param = CheckFloatAndRetString(textBox_PgainCh2.Text);
                if (param != null)
                {
                    WriteFloat(param, SetPgainCh2);
                    textBox_PgainCh2.Text = param;
                    PgainCh2_editing = false;
                }
            }
            else
            {
                if (e.KeyCode == Keys.Escape)   // Escape key reads back value from controller
                {
                    dataRx = ReadFloat(PgainCh2);
                    if (dataRx != null)
                    {
                        textBox_PgainCh2.Text = dataRx;
                    }
                    PgainCh2_editing = false;
                }
                else if (PgainCh2_editing == false) // Any other key purge the input field
                {
                    textBox_PgainCh2.Clear();
                    PgainCh2_editing = true;
                }
            }

        }

        private void textBox_IgainCh2_KeyDown(object sender, KeyEventArgs e)
        {
            string dataRx;

            if (e.KeyCode == Keys.Enter)    // Enter key writes data to controller
            {
                string param = CheckFloatAndRetString(textBox_IgainCh2.Text);
                if (param != null)
                {
                    WriteFloat(param, SetIgainCh2);
                    textBox_IgainCh2.Text = param;
                    IgainCh2_editing = false;
                }
            }
            else
            {
                if (e.KeyCode == Keys.Escape)   // Escape key reads back value from controller
                {
                    dataRx = ReadFloat(IgainCh2);
                    if (dataRx != null)
                    {
                        textBox_IgainCh2.Text = dataRx;
                    }
                    IgainCh2_editing = false;
                }
                else if (IgainCh2_editing == false) // Any other key purge the input field
                {
                    textBox_IgainCh2.Clear();
                    IgainCh2_editing = true;
                }
            }
        }

        private void textBox_DgainCh2_KeyDown(object sender, KeyEventArgs e)
        {
            string dataRx;

            if (e.KeyCode == Keys.Enter)    // Enter key writes data to controller
            {
                string param = CheckFloatAndRetString(textBox_DgainCh2.Text);
                if (param != null)
                {
                    WriteFloat(param, SetDgainCh2);
                    textBox_DgainCh2.Text = param;
                    DgainCh2_editing = false;
                }
            }
            else
            {
                if (e.KeyCode == Keys.Escape)   // Escape key reads back value from controller
                {
                    dataRx = ReadFloat(DgainCh2);
                    if (dataRx != null)
                    {
                        textBox_DgainCh2.Text = dataRx;
                    }
                    DgainCh2_editing = false;
                }
                else if (DgainCh2_editing == false) // Any other key purge the input field
                {
                    textBox_DgainCh2.Clear();
                    DgainCh2_editing = true;
                }
            }
        }

        private void textBox_SetPointCh3_KeyDown(object sender, KeyEventArgs e)
        {
            string dataRx;

            if (e.KeyCode == Keys.Enter)    // Enter key writes data to controller
            {
                string param = CheckFloatAndRetString(textBox_SetPointCh3.Text);
                if (param != null)
                {
                    WriteFloat(param, SetTempSetPoint3);
                    textBox_SetPointCh3.Text = param;
                    TempSensor3_editing = false;
                }
            }
            else
            {
                if (e.KeyCode == Keys.Escape)   // Escape key reads back value from controller
                {
                    dataRx = ReadFloat(TempSetPoint3);
                    if (dataRx != null)
                    {
                        textBox_SetPointCh3.Text = dataRx;
                    }
                    TempSensor3_editing = false;
                }
                else if (TempSensor3_editing == false) // Any other key purge the input field
                {
                    textBox_SetPointCh3.Clear();
                    TempSensor3_editing = true;
                }
            }
        }

        private void textBox_PgainCh3_KeyDown(object sender, KeyEventArgs e)
        {
            string dataRx;

            if (e.KeyCode == Keys.Enter)    // Enter key writes data to controller
            {
                string param = CheckFloatAndRetString(textBox_PgainCh3.Text);
                if (param != null)
                {
                    WriteFloat(param, SetPgainCh3);
                    textBox_PgainCh3.Text = param;
                    PgainCh3_editing = false;
                }
            }
            else
            {
                if (e.KeyCode == Keys.Escape)   // Escape key reads back value from controller
                {
                    dataRx = ReadFloat(PgainCh3);
                    if (dataRx != null)
                    {
                        textBox_PgainCh3.Text = dataRx;
                    }
                    PgainCh3_editing = false;
                }
                else if (PgainCh3_editing == false) // Any other key purge the input field
                {
                    textBox_PgainCh3.Clear();
                    PgainCh3_editing = true;
                }
            }
        }

        private void textBox_IgainCh3_KeyDown(object sender, KeyEventArgs e)
        {
            string dataRx;

            if (e.KeyCode == Keys.Enter)    // Enter key writes data to controller
            {
                string param = CheckFloatAndRetString(textBox_IgainCh3.Text);
                if (param != null)
                {
                    WriteFloat(param, SetIgainCh3);
                    textBox_IgainCh3.Text = param;
                    IgainCh3_editing = false;
                }
            }
            else
            {
                if (e.KeyCode == Keys.Escape)   // Escape key reads back value from controller
                {
                    dataRx = ReadFloat(IgainCh3);
                    if (dataRx != null)
                    {
                        textBox_IgainCh3.Text = dataRx;
                    }
                    IgainCh3_editing = false;
                }
                else if (IgainCh3_editing == false) // Any other key purge the input field
                {
                    textBox_IgainCh3.Clear();
                    IgainCh3_editing = true;
                }
            }
        }

        private void textBox_DgainCh3_KeyDown(object sender, KeyEventArgs e)
        {
            string dataRx;

            if (e.KeyCode == Keys.Enter)    // Enter key writes data to controller
            {
                string param = CheckFloatAndRetString(textBox_DgainCh3.Text);
                if (param != null)
                {
                    WriteFloat(param, SetDgainCh3);
                    textBox_DgainCh3.Text = param;
                    DgainCh3_editing = false;
                }
            }
            else
            {
                if (e.KeyCode == Keys.Escape)   // Escape key reads back value from controller
                {
                    dataRx = ReadFloat(DgainCh3);
                    if (dataRx != null)
                    {
                        textBox_DgainCh3.Text = dataRx;
                    }
                    DgainCh3_editing = false;
                }
                else if (DgainCh3_editing == false) // Any other key purge the input field
                {
                    textBox_DgainCh3.Clear();
                    DgainCh3_editing = true;
                }
            }
        }

        private void button_HeatOnOffCh1_Click(object sender, EventArgs e)
        {
            if (!radioButton_HeatOnOffCh1.Checked)
            {
                if (WriteFloatRetString("1", SetHeaterOnOffCh1).Contains("ON") == true)
                {
                    radioButton_HeatOnOffCh1.Checked = true;
                }
            }
            else
            {
                if (WriteFloatRetString("0", SetHeaterOnOffCh1).Contains("OFF") == true)
                {
                    radioButton_HeatOnOffCh1.Checked = false;
                }
            }

        }

        private void button_HeatOnOffCh2_Click(object sender, EventArgs e)
        {
            if (!radioButton_HeatOnOffCh2.Checked)
            {
                if (WriteFloatRetString("1", SetHeaterOnOffCh2).Contains("ON") == true)
                {
                    radioButton_HeatOnOffCh2.Checked = true;
                }
            }
            else
            {
                if (WriteFloatRetString("0", SetHeaterOnOffCh2).Contains("OFF") == true)
                {
                    radioButton_HeatOnOffCh2.Checked = false;
                }
            }

        }

        private void button_HeatOnOffCh3_Click(object sender, EventArgs e)
        {
            if (!radioButton_HeatOnOffCh3.Checked)
            {
                if (WriteFloatRetString("1", SetHeaterOnOffCh3).Contains("ON") == true)
                {
                    radioButton_HeatOnOffCh3.Checked = true;
                }
            }
            else
            {
                if (WriteFloatRetString("0", SetHeaterOnOffCh3).Contains("OFF") == true)
                {
                    radioButton_HeatOnOffCh3.Checked = false;
                }
            }

        }

        private void textBox_SettleTimeCh0_KeyDown(object sender, KeyEventArgs e)
        {
            string dataRx;

            if (e.KeyCode == Keys.Enter)    // Enter key writes data to controller
            {
                string param = CheckIntAndRetString(textBox_SettleTimeCh0.Text);
                if (param != null)
                {
                    WriteFloat(param, SetSettleTimeT0);
                    textBox_SettleTimeCh0.Text = param;
                    SettleTime0_editing = false;
                }
            }
            else
            {
                if (e.KeyCode == Keys.Escape)   // Escape key reads back value from controller
                {
                    dataRx = ReadFloat(SettleTimeTempCh0);
                    if (dataRx != null)
                    {
                        textBox_SettleTimeCh0.Text = dataRx;
                    }
                    SettleTime0_editing = false;
                }
                else if (SettleTime0_editing == false) // Any other key purge the input field
                {
                    textBox_SettleTimeCh0.Clear();
                    SettleTime0_editing = true;
                }
            }
        }

        private void textBox_SettleTimeCh1_KeyDown(object sender, KeyEventArgs e)
        {
            string dataRx;

            if (e.KeyCode == Keys.Enter)    // Enter key writes data to controller
            {
                string param = CheckIntAndRetString(textBox_SettleTimeCh1.Text);
                if (param != null)
                {
                    WriteFloat(param, SetSettleTimeT1);
                    textBox_SettleTimeCh1.Text = param;
                    SettleTime1_editing = false;
                }
            }
            else
            {
                if (e.KeyCode == Keys.Escape)   // Escape key reads back value from controller
                {
                    dataRx = ReadFloat(SettleTimeTempCh1);
                    if (dataRx != null)
                    {
                        textBox_SettleTimeCh1.Text = dataRx;
                    }
                    SettleTime1_editing = false;
                }
                else if (SettleTime1_editing == false) // Any other key purge the input field
                {
                    textBox_SettleTimeCh1.Clear();
                    SettleTime1_editing = true;
                }
            }
        }

        private void textBox_SettleTimeCh2_KeyDown(object sender, KeyEventArgs e)
        {
            string dataRx;

            if (e.KeyCode == Keys.Enter)    // Enter key writes data to controller
            {
                string param = CheckIntAndRetString(textBox_SettleTimeCh2.Text);
                if (param != null)
                {
                    WriteFloat(param, SetSettleTimeT2);
                    textBox_SettleTimeCh2.Text = param;
                    SettleTime2_editing = false;
                }
            }
            else
            {
                if (e.KeyCode == Keys.Escape)   // Escape key reads back value from controller
                {
                    dataRx = ReadFloat(SettleTimeTempCh2);
                    if (dataRx != null)
                    {
                        textBox_SettleTimeCh2.Text = dataRx;
                    }
                    SettleTime2_editing = false;
                }
                else if (SettleTime2_editing == false) // Any other key purge the input field
                {
                    textBox_SettleTimeCh2.Clear();
                    SettleTime2_editing = true;
                }
            }
        }

        private void textBox_SettleTimeCh3_KeyDown(object sender, KeyEventArgs e)
        {
            string dataRx;

            if (e.KeyCode == Keys.Enter)    // Enter key writes data to controller
            {
                string param = CheckIntAndRetString(textBox_SettleTimeCh3.Text);
                if (param != null)
                {
                    WriteFloat(param, SetSettleTimeT3);
                    textBox_SettleTimeCh3.Text = param;
                    SettleTime3_editing = false;
                }
            }
            else
            {
                if (e.KeyCode == Keys.Escape)   // Escape key reads back value from controller
                {
                    dataRx = ReadFloat(SettleTimeTempCh3);
                    if (dataRx != null)
                    {
                        textBox_SettleTimeCh3.Text = dataRx;
                    }
                    SettleTime3_editing = false;
                }
                else if (SettleTime3_editing == false) // Any other key purge the input field
                {
                    textBox_SettleTimeCh3.Clear();
                    SettleTime3_editing = true;
                }
            }
        }

        private void textBox_TempWindowCh0_KeyDown(object sender, KeyEventArgs e)
        {
            string dataRx;

            if (e.KeyCode == Keys.Enter)    // Enter key writes data to controller
            {
                string param = CheckFloatAndRetString(textBox_TempWindowCh0.Text);
                if (param != null)
                {
                    WriteFloat(param, SetTempWindowCh0);
                    textBox_TempWindowCh0.Text = param;
                    TempWindow0_editing = false;
                }
            }
            else
            {
                if (e.KeyCode == Keys.Escape)   // Escape key reads back value from controller
                {
                    dataRx = ReadFloat(TempWindowCh0);
                    if (dataRx != null)
                    {
                        textBox_TempWindowCh0.Text = dataRx;
                    }
                    TempWindow0_editing = false;
                }
                else if (TempWindow0_editing == false) // Any other key purge the input field
                {
                    textBox_TempWindowCh0.Clear();
                    TempWindow0_editing = true;
                }
            }
        }

        private void textBox_TempWindowCh1_KeyDown(object sender, KeyEventArgs e)
        {
            string dataRx;

            if (e.KeyCode == Keys.Enter)    // Enter key writes data to controller
            {
                string param = CheckFloatAndRetString(textBox_TempWindowCh1.Text);
                if (param != null)
                {
                    WriteFloat(param, SetTempWindowCh1);
                    textBox_TempWindowCh1.Text = param;
                    TempWindow1_editing = false;
                }
            }
            else
            {
                if (e.KeyCode == Keys.Escape)   // Escape key reads back value from controller
                {
                    dataRx = ReadFloat(TempWindowCh1);
                    if (dataRx != null)
                    {
                        textBox_TempWindowCh1.Text = dataRx;
                    }
                    TempWindow1_editing = false;
                }
                else if (TempWindow1_editing == false) // Any other key purge the input field
                {
                    textBox_TempWindowCh1.Clear();
                    TempWindow1_editing = true;
                }
            }
        }

        private void textBox_TempWindowCh2_KeyDown(object sender, KeyEventArgs e)
        {
            string dataRx;

            if (e.KeyCode == Keys.Enter)    // Enter key writes data to controller
            {
                string param = CheckFloatAndRetString(textBox_TempWindowCh2.Text);
                if (param != null)
                {
                    WriteFloat(param, SetTempWindowCh2);
                    textBox_TempWindowCh2.Text = param;
                    TempWindow2_editing = false;
                }
            }
            else
            {
                if (e.KeyCode == Keys.Escape)   // Escape key reads back value from controller
                {
                    dataRx = ReadFloat(TempWindowCh2);
                    if (dataRx != null)
                    {
                        textBox_TempWindowCh2.Text = dataRx;
                    }
                    TempWindow2_editing = false;
                }
                else if (TempWindow2_editing == false) // Any other key purge the input field
                {
                    textBox_TempWindowCh2.Clear();
                    TempWindow2_editing = true;
                }
            }
        }

        private void textBox_TempWindowCh3_KeyDown(object sender, KeyEventArgs e)
        {
            string dataRx;

            if (e.KeyCode == Keys.Enter)    // Enter key writes data to controller
            {
                string param = CheckFloatAndRetString(textBox_TempWindowCh3.Text);
                if (param != null)
                {
                    WriteFloat(param, SetTempWindowCh3);
                    textBox_TempWindowCh3.Text = param;
                    TempWindow3_editing = false;
                }
            }
            else
            {
                if (e.KeyCode == Keys.Escape)   // Escape key reads back value from controller
                {
                    dataRx = ReadFloat(TempWindowCh3);
                    if (dataRx != null)
                    {
                        textBox_TempWindowCh3.Text = dataRx;
                    }
                    TempWindow3_editing = false;
                }
                else if (TempWindow3_editing == false) // Any other key purge the input field
                {
                    textBox_TempWindowCh3.Clear();
                    TempWindow3_editing = true;
                }
            }
        }

        private void button_test_Click(object sender, EventArgs e)
        {
            configuration.AppSettings.Settings["comPort"].Value = "Test";
            configuration.Save(ConfigurationSaveMode.Full, true);
 
            textBox_test.Text = configuration.AppSettings.Settings["comPort"].Value;
        }
    }
}
