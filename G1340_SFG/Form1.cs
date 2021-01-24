using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace G1340_SFG
{
    public partial class Form1 : Form
    {
        List<string> OffsetValue = null;
        List<string> AfterOffset = null;
        List<string> valueSensor1340 = null;
        List<string> valueSensor1340Linear = null;
        List<string> plotIt = null;
        TcpClient goClient = null;
        NetworkStream goStream = null;
        StreamReader goReader = null;
        TcpClient plcClient = null;
        NetworkStream plcStream = null;
        StreamReader plcReader = null;
        ConnectDB connection = null;

        bool loopState = false;

        public Form1()
        {
            InitializeComponent();
            OffsetValue = new List<string>(30);
            valueSensor1340 = new List<string>(30);
            AfterOffset = new List<string>(30);
            valueSensor1340Linear = new List<string>(30);
            plotIt = new List<string>(30);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            btnConDB.Enabled = false;
            btnConnectGocator.Enabled = false;
            btnStart.Enabled = false;

        }
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OffsetValue.Clear();
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string sFileName = openFileDialog1.FileName;
                tbOffset.Text = openFileDialog1.SafeFileName;
                StreamReader reader = new StreamReader(sFileName);
                while (!reader.EndOfStream)
                {
                    string line = reader.ReadLine();
                    //MessageBox.Show(line);
                    string[] data1 = line.Split(',');
                    try
                    {
                        if (!string.IsNullOrEmpty(data1[4]))
                        {
                            OffsetValue.Add(data1[3]);
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Error Offset");
                    }

                }
                MessageBox.Show(string.Join(",", OffsetValue));
            }
            if (!string.IsNullOrEmpty(tbOffset.Text))
            {
                btnConnectGocator.Enabled = true;
                btnConDB.Enabled = true;
            }
        }

        private string ReadData(TcpClient client, NetworkStream stream, StreamReader reader) //Read Data from PLC
        {
            stream = client.GetStream();
            reader = new StreamReader(stream);
            char[] readerText = new char[64];
            reader.Read(readerText, 0, readerText.Length);
            string dataSensor = new string(readerText);
            Array.Clear(readerText, 0, readerText.Length);
            return dataSensor;
        }

        private void WriteData(TcpClient client, NetworkStream stream, string mgs)  //Write Data to send to Client
        {
            stream = client.GetStream();
            byte[] buff = Encoding.ASCII.GetBytes(mgs + "\r\n");
            stream.Write(buff, 0, buff.Length);
            Array.Clear(buff, 0, buff.Length);
        }

        private void btnConDB_Click(object sender, EventArgs e)
        {
            try
            {
                connection = new ConnectDB();
                connection.connect(@"Data Source=172.28.103.50;Initial Catalog=SiamFiberGlass;User ID=sa;Password=server");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Database Error  {ex.Message}", "Error connection", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            btnStart.Enabled = true;
        }

        private void btnConnectGocator_Click(object sender, EventArgs e)
        {
            try
            {
                try
                {
                    goClient = new TcpClient();
                    goClient.Connect("192.168.0.10", 8190);
                    //goClient.Connect("127.0.0.1", 21);
                }
                catch (Exception)
                {
                    throw new Exception("Gocator");
                }

                try
                {
                    plcClient = new TcpClient();
                    //plcClient.Connect("192.168.0.6", 4000);
                    plcClient.Connect("127.0.0.1", 21);
                }
                catch (Exception)
                {
                    throw new Exception("PLC");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"System connection error :{ex.Message}", "Error connection", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(tbModel.Text))
            {
                try
                {
                    Thread mainSys = new Thread(new ThreadStart(mainRun));
                    loopState = true;
                    mainSys.Start();
                }
                catch
                {
                }
            }
            else
            {
                MessageBox.Show("Please type a Model Name");
                tbModel.Focus();
            }
            
        }

        private void mainRun()
        {
            
            while (loopState) // while (true)
            {
                //WriteData(goClient, goStream, "Stop");
                string rawMsg = ReadData(plcClient, plcStream, plcReader);
                string[] Msg = rawMsg.Split(',');
                switch (Msg[0])
                {
                    case "ONP":
                        WriteData(goClient, goStream, "Stop");
                        WriteData(goClient, goStream, "Start");
                        ReadData(goClient, goStream, goReader);
                        Thread.Sleep(3000);
                        WriteData(goClient, goStream, "Stop");
                        ReadData(goClient, goStream, goReader);
                        getPointP();

                        //sent finish to PLC
                        WriteData(plcClient, plcStream, "Finish");
                        break;

                    case "ONLS": //Start scan with Linear
                        WriteData(goClient, goStream, "Start");
                        WriteData(plcClient, plcStream, "Readyyyyyyyy"); //Send Ready to PLC then PLC run a Linear

                        break;

                    case "ONLE": //End scan with Linear
                        getPointL();
                        WriteData(goClient, goStream, "Stop"); //PLC Stop a Linear and Send Stop to Gocator
                        //Thread.Sleep(5000);
                        //getPointL(); //ไม่มา

                        WriteData(plcClient, plcStream, "Finish");
                        break;

                    case "L":
                        plotDataGridP();
                        WriteData(goClient, goStream, "LoadJob,Line.job");
                        WriteData(plcClient, plcStream, "Ready");
                        break;
                    case "P":
                        plotDataGridL();
                        WriteData(goClient, goStream, "LoadJob,Point.job");
                        WriteData(plcClient, plcStream, "Ready");
                        break;
                    default:
                        break;
                }
            }
        }

        private void getPointP()
        {
            ReadData(goClient, goStream, goReader);
            WriteData(goClient, goStream, "Result");
            
            string rawMsg = ReadData(goClient, goStream, goReader);
            string[] dataMsg = rawMsg.Split(',');

            if ((dataMsg.Length > 0) && dataMsg[0] == "OK")
            {
                try
                {
                        for (int i = 1; i < dataMsg.Length; i++)
                        {
                        valueSensor1340.Add((Convert.ToDouble(dataMsg[i]) / 1000.0).ToString());
                        }
                }
                catch (Exception ex)
                {
                    
                   MessageBox.Show($"mainRun : read data error !! : {ex.Message}");
                    
                }
            }
        }
        private void visualizationData(string table, List<string> dataSensor, DataGridView dgv , string product)
        {
            string D = DateTime.Now.ToString("yyyy-MM-dd");
            string T = DateTime.Now.ToString("HH:mm:ss");
            string DT = D + " " + T;
            try
            {
                switch (dataSensor.Count)
                {
                    case 1:
                        dgv.Invoke(new Action(() => dgv.Rows.Add((dgv.Rows.Count + 1).ToString(), D, dataSensor[0])));
                        connection.insert("insert into " + table + "(dTime,Models,P1) values('" + DT + "','" + product + "'," + dataSensor[0] + ")");
                        break;
                    case 2:
                        dgv.Invoke(new Action(() => dgv.Rows.Add((dgv.Rows.Count + 1).ToString(), DT, dataSensor[0], dataSensor[1])));
                        connection.insert("insert into " + table + "(dTime,Models,P1,P2) values('" + DT + "','" + product + "'," + dataSensor[0] + "," + dataSensor[1] + ")");
                        break;
                    case 3:
                        dgv.Invoke(new Action(() => dgv.Rows.Add((dgv.Rows.Count + 1).ToString(), DT, dataSensor[0], dataSensor[1], dataSensor[2])));
                        connection.insert("insert into " + table + "(dTime,Models,P1,P2,P3) values('" + DT + "','" + product + "'," + dataSensor[0] + "," + dataSensor[1] + "," + dataSensor[2] + ")");
                        break;
                    case 4:
                        dgv.Invoke(new Action(() => dgv.Rows.Add((dgv.Rows.Count + 1).ToString(), DT, dataSensor[0], dataSensor[1], dataSensor[2], dataSensor[3])));
                        connection.insert("insert into " + table + "(dTime,Models,P1,P2,P3,P4) values('" + DT + "','" + product + "'," + dataSensor[0] + "," + dataSensor[1] + "," + dataSensor[2] + "," + dataSensor[3] + ")");
                        break;
                    case 5:
                        dgv.Invoke(new Action(() => dgv.Rows.Add((dgv.Rows.Count + 1).ToString(), DT, dataSensor[0], dataSensor[1], dataSensor[2], dataSensor[3], dataSensor[4])));
                        connection.insert("insert into " + table + "(dTime,Models,P1,P2,P3,P4,P5) values('" + DT + "','" + product + "'," + dataSensor[0] + "," + dataSensor[1] + "," + dataSensor[2] + "," + dataSensor[3] + "," + dataSensor[4] + ")");
                        break;
                    case 6:
                        dgv.Invoke(new Action(() => dgv.Rows.Add((dgv.Rows.Count + 1).ToString(), DT, dataSensor[0], dataSensor[1], dataSensor[2], dataSensor[3], dataSensor[4], dataSensor[5])));
                        connection.insert("insert into " + table + "(dTime,Models,P1,P2,P3,P4,P5,P6) values('" + DT + "','" + product + "'," + dataSensor[0] + "," + dataSensor[1] + "," + dataSensor[2] + "," + dataSensor[3] + "," + dataSensor[4] + "," + dataSensor[5] + ")");
                        break;
                    case 7:
                        dgv.Invoke(new Action(() => dgv.Rows.Add((dgv.Rows.Count + 1).ToString(), DT, dataSensor[0], dataSensor[1], dataSensor[2], dataSensor[3], dataSensor[4], dataSensor[5], dataSensor[6])));
                        connection.insert("insert into " + table + "(dTime,Models,P1,P2,P3,P4,P5,P6,P7) values('" + DT + "','" + product + "'," + dataSensor[0] + "," + dataSensor[1] + "," + dataSensor[2] + "," + dataSensor[3] + "," + dataSensor[4] + "," + dataSensor[5] + "," + dataSensor[6] + ")");
                        break;
                    case 8:
                        dgv.Invoke(new Action(() => dgv.Rows.Add((dgv.Rows.Count + 1).ToString(), DT, dataSensor[0], dataSensor[1], dataSensor[2], dataSensor[3], dataSensor[4], dataSensor[5], dataSensor[6], dataSensor[7])));
                        connection.insert("insert into " + table + "(dTime,Models,P1,P2,P3,P4,P5,P6,P7,P8) values('" + DT + "','" + product + "'," + dataSensor[0] + "," + dataSensor[1] + "," + dataSensor[2] + "," + dataSensor[3] + "," + dataSensor[4] + "," + dataSensor[5] + "," + dataSensor[6] + "," + dataSensor[7] + ")");
                        break;
                    case 9:
                        dgv.Invoke(new Action(() => dgv.Rows.Add((dgv.Rows.Count + 1).ToString(), DT, dataSensor[0], dataSensor[1], dataSensor[2], dataSensor[3], dataSensor[4], dataSensor[5], dataSensor[6], dataSensor[7], dataSensor[8])));
                        connection.insert("insert into " + table + "(dTime,Models,P1,P2,P3,P4,P5,P6,P7,P8,P9) values('" + DT + "','" + product + "'," + dataSensor[0] + "," + dataSensor[1] + "," + dataSensor[2] + "," + dataSensor[3] + "," + dataSensor[4] + "," + dataSensor[5] + "," + dataSensor[6] + "," + dataSensor[7] + "," + dataSensor[8] + ")");
                        break;
                    case 10:
                        dgv.Invoke(new Action(() => dgv.Rows.Add((dgv.Rows.Count + 1).ToString(), DT, dataSensor[0], dataSensor[1], dataSensor[2], dataSensor[3], dataSensor[4], dataSensor[5], dataSensor[6], dataSensor[7], dataSensor[8], dataSensor[9])));
                        connection.insert("insert into " + table + "(dTime,Models,P1,P2,P3,P4,P5,P6,P7,P8,P9,P10) values('" + DT + "','" + product + "'," + dataSensor[0] + "," + dataSensor[1] + "," + dataSensor[2] + "," + dataSensor[3] + "," + dataSensor[4] + "," + dataSensor[5] + "," + dataSensor[6] + "," + dataSensor[7] + "," + dataSensor[8] + "," + dataSensor[9] + ")");
                        break;
                    case 11:
                        dgv.Invoke(new Action(() => dgv.Rows.Add((dgv.Rows.Count + 1).ToString(), DT, dataSensor[0], dataSensor[1], dataSensor[2], dataSensor[3], dataSensor[4], dataSensor[5], dataSensor[6], dataSensor[7], dataSensor[8], dataSensor[9], dataSensor[10])));
                        connection.insert("insert into " + table + "(dTime,Models,P1,P2,P3,P4,P5,P6,P7,P8,P9,P10,P11) values('" + DT + "','" + product + "'," + dataSensor[0] + "," + dataSensor[1] + "," + dataSensor[2] + "," + dataSensor[3] + "," + dataSensor[4] + "," + dataSensor[5] + "," + dataSensor[6] + "," + dataSensor[7] + "," + dataSensor[8] + "," + dataSensor[9] + "," + dataSensor[10] + ")");
                        break;
                    case 12:
                        dgv.Invoke(new Action(() => dgv.Rows.Add((dgv.Rows.Count + 1).ToString(), DT, dataSensor[0], dataSensor[1], dataSensor[2], dataSensor[3], dataSensor[4], dataSensor[5], dataSensor[6], dataSensor[7], dataSensor[8], dataSensor[9], dataSensor[10], dataSensor[11])));
                        connection.insert("insert into " + table + "(dTime,Models,P1,P2,P3,P4,P5,P6,P7,P8,P9,P10,P11,P12) values('" + DT + "','" + product + "'," + dataSensor[0] + "," + dataSensor[1] + "," + dataSensor[2] + "," + dataSensor[3] + "," + dataSensor[4] + "," + dataSensor[5] + "," + dataSensor[6] + "," + dataSensor[7] + "," + dataSensor[8] + "," + dataSensor[9] + "," + dataSensor[10] + "," + dataSensor[11] + ")");
                        break;
                    case 13:
                        dgv.Invoke(new Action(() => dgv.Rows.Add((dgv.Rows.Count + 1).ToString(), DT, dataSensor[0], dataSensor[1], dataSensor[2], dataSensor[3], dataSensor[4], dataSensor[5], dataSensor[6], dataSensor[7], dataSensor[8], dataSensor[9], dataSensor[10], dataSensor[11], dataSensor[12])));
                        connection.insert("insert into " + table + "(dTime,Models,P1,P2,P3,P4,P5,P6,P7,P8,P9,P10,P11,P12,P13) values('" + DT + "','" + product + "'," + dataSensor[0] + "," + dataSensor[1] + "," + dataSensor[2] + "," + dataSensor[3] + "," + dataSensor[4] + "," + dataSensor[5] + "," + dataSensor[6] + "," + dataSensor[7] + "," + dataSensor[8] + "," + dataSensor[9] + "," + dataSensor[10] + "," + dataSensor[11] + "," + dataSensor[12] + ")");
                        break;
                    case 14:
                        dgv.Invoke(new Action(() => dgv.Rows.Add((dgv.Rows.Count + 1).ToString(), DT, dataSensor[0], dataSensor[1], dataSensor[2], dataSensor[3], dataSensor[4], dataSensor[5], dataSensor[6], dataSensor[7], dataSensor[8], dataSensor[9], dataSensor[10], dataSensor[11], dataSensor[12], dataSensor[13])));
                        connection.insert("insert into " + table + "(dTime,Models,P1,P2,P3,P4,P5,P6,P7,P8,P9,P10,P11,P12,P13,P14) values('" + DT + "','" + product + "'," + dataSensor[0] + "," + dataSensor[1] + "," + dataSensor[2] + "," + dataSensor[3] + "," + dataSensor[4] + "," + dataSensor[5] + "," + dataSensor[6] + "," + dataSensor[7] + "," + dataSensor[8] + "," + dataSensor[9] + "," + dataSensor[10] + "," + dataSensor[11] + "," + dataSensor[12] + "," + dataSensor[13] + ")");
                        break;
                    case 15:
                        dgv.Invoke(new Action(() => dgv.Rows.Add((dgv.Rows.Count + 1).ToString(), DT, dataSensor[0], dataSensor[1], dataSensor[2], dataSensor[3], dataSensor[4], dataSensor[5], dataSensor[6], dataSensor[7], dataSensor[8], dataSensor[9], dataSensor[10], dataSensor[11], dataSensor[12], dataSensor[13], dataSensor[14])));
                        connection.insert("insert into " + table + "(dTime,Models,P1,P2,P3,P4,P5,P6,P7,P8,P9,P10,P11,P12,P13,P14,P15) values('" + DT + "','" + product + "'," + dataSensor[0] + "," + dataSensor[1] + "," + dataSensor[2] + "," + dataSensor[3] + "," + dataSensor[4] + "," + dataSensor[5] + "," + dataSensor[6] + "," + dataSensor[7] + "," + dataSensor[8] + "," + dataSensor[9] + "," + dataSensor[10] + "," + dataSensor[11] + "," + dataSensor[12] + "," + dataSensor[13] + "," + dataSensor[14] + ")");
                        break;
                    case 16:
                        dgv.Invoke(new Action(() => dgv.Rows.Add((dgv.Rows.Count + 1).ToString(), DT, dataSensor[0], dataSensor[1], dataSensor[2], dataSensor[3], dataSensor[4], dataSensor[5], dataSensor[6], dataSensor[7], dataSensor[8], dataSensor[9], dataSensor[10], dataSensor[11], dataSensor[12], dataSensor[13], dataSensor[14], dataSensor[15])));
                        connection.insert("insert into " + table + "(dTime,Models,P1,P2,P3,P4,P5,P6,P7,P8,P9,P10,P11,P12,P13,P14,P15,P16) values('" + DT + "','" + product + "'," + dataSensor[0] + "," + dataSensor[1] + "," + dataSensor[2] + "," + dataSensor[3] + "," + dataSensor[4] + "," + dataSensor[5] + "," + dataSensor[6] + "," + dataSensor[7] + "," + dataSensor[8] + "," + dataSensor[9] + "," + dataSensor[10] + "," + dataSensor[11] + "," + dataSensor[12] + "," + dataSensor[13] + "," + dataSensor[14] + "," + dataSensor[15] + ")");
                        break;
                    case 17:
                        dgv.Invoke(new Action(() => dgv.Rows.Add((dgv.Rows.Count + 1).ToString(), DT, dataSensor[0], dataSensor[1], dataSensor[2], dataSensor[3], dataSensor[4], dataSensor[5], dataSensor[6], dataSensor[7], dataSensor[8], dataSensor[9], dataSensor[10], dataSensor[11], dataSensor[12], dataSensor[13], dataSensor[14], dataSensor[15], dataSensor[16])));
                        connection.insert("insert into " + table + "(dTime,Models,P1,P2,P3,P4,P5,P6,P7,P8,P9,P10,P11,P12,P13,P14,P15,P16,P17) values('" + DT + "','" + product + "'," + dataSensor[0] + "," + dataSensor[1] + "," + dataSensor[2] + "," + dataSensor[3] + "," + dataSensor[4] + "," + dataSensor[5] + "," + dataSensor[6] + "," + dataSensor[7] + "," + dataSensor[8] + "," + dataSensor[9] + "," + dataSensor[10] + "," + dataSensor[11] + "," + dataSensor[12] + "," + dataSensor[13] + "," + dataSensor[14] + "," + dataSensor[15] + "," + dataSensor[16] + ")");
                        break;
                    case 18:
                        dgv.Invoke(new Action(() => dgv.Rows.Add((dgv.Rows.Count + 1).ToString(), DT, dataSensor[0], dataSensor[1], dataSensor[2], dataSensor[3], dataSensor[4], dataSensor[5], dataSensor[6], dataSensor[7], dataSensor[8], dataSensor[9], dataSensor[10], dataSensor[11], dataSensor[12], dataSensor[13], dataSensor[14], dataSensor[15], dataSensor[16], dataSensor[17])));
                        connection.insert("insert into " + table + "(dTime,Models,P1,P2,P3,P4,P5,P6,P7,P8,P9,P10,P11,P12,P13,P14,P15,P16,P17,P18) values('" + DT + "','" + product + "'," + dataSensor[0] + "," + dataSensor[1] + "," + dataSensor[2] + "," + dataSensor[3] + "," + dataSensor[4] + "," + dataSensor[5] + "," + dataSensor[6] + "," + dataSensor[7] + "," + dataSensor[8] + "," + dataSensor[9] + "," + dataSensor[10] + "," + dataSensor[11] + "," + dataSensor[12] + "," + dataSensor[13] + "," + dataSensor[14] + "," + dataSensor[15] + "," + dataSensor[16] + "," + dataSensor[17] + ")");
                        break;
                    case 19:
                        dgv.Invoke(new Action(() => dgv.Rows.Add((dgv.Rows.Count + 1).ToString(), DT, dataSensor[0], dataSensor[1], dataSensor[2], dataSensor[3], dataSensor[4], dataSensor[5], dataSensor[6], dataSensor[7], dataSensor[8], dataSensor[9], dataSensor[10], dataSensor[11], dataSensor[12], dataSensor[13], dataSensor[14], dataSensor[15], dataSensor[16], dataSensor[17], dataSensor[18])));
                        connection.insert("insert into " + table + "(dTime,Models,P1,P2,P3,P4,P5,P6,P7,P8,P9,P10,P11,P12,P13,P14,P15,P16,P17,P18,P19) values('" + DT + "','" + product + "'," + dataSensor[0] + "," + dataSensor[1] + "," + dataSensor[2] + "," + dataSensor[3] + "," + dataSensor[4] + "," + dataSensor[5] + "," + dataSensor[6] + "," + dataSensor[7] + "," + dataSensor[8] + "," + dataSensor[9] + "," + dataSensor[10] + "," + dataSensor[11] + "," + dataSensor[12] + "," + dataSensor[13] + "," + dataSensor[14] + "," + dataSensor[15] + "," + dataSensor[16] + "," + dataSensor[17] + "," + dataSensor[18] + ")");
                        break;
                    case 20:
                        dgv.Invoke(new Action(() => dgv.Rows.Add((dgv.Rows.Count + 1).ToString(), DT, dataSensor[0], dataSensor[1], dataSensor[2], dataSensor[3], dataSensor[4], dataSensor[5], dataSensor[6], dataSensor[7], dataSensor[8], dataSensor[9], dataSensor[10], dataSensor[11], dataSensor[12], dataSensor[13], dataSensor[14], dataSensor[15], dataSensor[16], dataSensor[17], dataSensor[18], dataSensor[19])));
                        connection.insert("insert into " + table + "(dTime,Models,P1,P2,P3,P4,P5,P6,P7,P8,P9,P10,P11,P12,P13,P14,P15,P16,P17,P18,P19,P20) values('" + DT + "','" + product + "'," + dataSensor[0] + "," + dataSensor[1] + "," + dataSensor[2] + "," + dataSensor[3] + "," + dataSensor[4] + "," + dataSensor[5] + "," + dataSensor[6] + "," + dataSensor[7] + "," + dataSensor[8] + "," + dataSensor[9] + "," + dataSensor[10] + "," + dataSensor[11] + "," + dataSensor[12] + "," + dataSensor[13] + "," + dataSensor[14] + "," + dataSensor[15] + "," + dataSensor[16] + "," + dataSensor[17] + "," + dataSensor[18] + "," + dataSensor[19] + ")");
                        break;
                    case 21:
                        dgv.Invoke(new Action(() => dgv.Rows.Add((dgv.Rows.Count + 1).ToString(), DT, dataSensor[0], dataSensor[1], dataSensor[2], dataSensor[3], dataSensor[4], dataSensor[5], dataSensor[6], dataSensor[7], dataSensor[8], dataSensor[9], dataSensor[10], dataSensor[11], dataSensor[12], dataSensor[13], dataSensor[14], dataSensor[15], dataSensor[16], dataSensor[17], dataSensor[18], dataSensor[19], dataSensor[20])));
                        connection.insert("insert into " + table + "(dTime,Models,P1,P2,P3,P4,P5,P6,P7,P8,P9,P10,P11,P12,P13,P14,P15,P16,P17,P18,P19,P20,P21) values('" + DT + "','" + product + "'," + dataSensor[0] + "," + dataSensor[1] + "," + dataSensor[2] + "," + dataSensor[3] + "," + dataSensor[4] + "," + dataSensor[5] + "," + dataSensor[6] + "," + dataSensor[7] + "," + dataSensor[8] + "," + dataSensor[9] + "," + dataSensor[10] + "," + dataSensor[11] + "," + dataSensor[12] + "," + dataSensor[13] + "," + dataSensor[14] + "," + dataSensor[15] + "," + dataSensor[16] + "," + dataSensor[17] + "," + dataSensor[18] + "," + dataSensor[19] + "," + dataSensor[20] + ")");
                        break;
                    case 22:
                        dgv.Invoke(new Action(() => dgv.Rows.Add((dgv.Rows.Count + 1).ToString(), DT, dataSensor[0], dataSensor[1], dataSensor[2], dataSensor[3], dataSensor[4], dataSensor[5], dataSensor[6], dataSensor[7], dataSensor[8], dataSensor[9], dataSensor[10], dataSensor[11], dataSensor[12], dataSensor[13], dataSensor[14], dataSensor[15], dataSensor[16], dataSensor[17], dataSensor[18], dataSensor[19], dataSensor[20], dataSensor[21])));
                        connection.insert("insert into " + table + "(dTime,Models,P1,P2,P3,P4,P5,P6,P7,P8,P9,P10,P11,P12,P13,P14,P15,P16,P17,P18,P19,P20,P21,P22) values('" + DT + "','" + product + "'," + dataSensor[0] + "," + dataSensor[1] + "," + dataSensor[2] + "," + dataSensor[3] + "," + dataSensor[4] + "," + dataSensor[5] + "," + dataSensor[6] + "," + dataSensor[7] + "," + dataSensor[8] + "," + dataSensor[9] + "," + dataSensor[10] + "," + dataSensor[11] + "," + dataSensor[12] + "," + dataSensor[13] + "," + dataSensor[14] + "," + dataSensor[15] + "," + dataSensor[16] + "," + dataSensor[17] + "," + dataSensor[18] + "," + dataSensor[19] + "," + dataSensor[20] + "," + dataSensor[21] + ")");
                        break;
                    case 23:
                        dgv.Invoke(new Action(() => dgv.Rows.Add((dgv.Rows.Count + 1).ToString(), DT, dataSensor[0], dataSensor[1], dataSensor[2], dataSensor[3], dataSensor[4], dataSensor[5], dataSensor[6], dataSensor[7], dataSensor[8], dataSensor[9], dataSensor[10], dataSensor[11], dataSensor[12], dataSensor[13], dataSensor[14], dataSensor[15], dataSensor[16], dataSensor[17], dataSensor[18], dataSensor[19], dataSensor[20], dataSensor[21], dataSensor[22])));
                        connection.insert("insert into " + table + "(dTime,Models,P1,P2,P3,P4,P5,P6,P7,P8,P9,P10,P11,P12,P13,P14,P15,P16,P17,P18,P19,P20,P21,P22,P23) values('" + DT + "','" + product + "'," + dataSensor[0] + "," + dataSensor[1] + "," + dataSensor[2] + "," + dataSensor[3] + "," + dataSensor[4] + "," + dataSensor[5] + "," + dataSensor[6] + "," + dataSensor[7] + "," + dataSensor[8] + "," + dataSensor[9] + "," + dataSensor[10] + "," + dataSensor[11] + "," + dataSensor[12] + "," + dataSensor[13] + "," + dataSensor[14] + "," + dataSensor[15] + "," + dataSensor[16] + "," + dataSensor[17] + "," + dataSensor[18] + "," + dataSensor[19] + "," + dataSensor[20] + "," + dataSensor[21] + "," + dataSensor[22] + ")");
                        break;
                    case 24:
                        dgv.Invoke(new Action(() => dgv.Rows.Add((dgv.Rows.Count + 1).ToString(), DT, dataSensor[0], dataSensor[1], dataSensor[2], dataSensor[3], dataSensor[4], dataSensor[5], dataSensor[6], dataSensor[7], dataSensor[8], dataSensor[9], dataSensor[10], dataSensor[11], dataSensor[12], dataSensor[13], dataSensor[14], dataSensor[15], dataSensor[16], dataSensor[17], dataSensor[18], dataSensor[19], dataSensor[20], dataSensor[21], dataSensor[22], dataSensor[23], dataSensor[24])));
                        connection.insert("insert into " + table + "(dTime,Models,P1,P2,P3,P4,P5,P6,P7,P8,P9,P10,P11,P12,P13,P14,P15,P16,P17,P18,P19,P20,P21,P22,P23,P24) values('" + DT + "','" + product + "'," + dataSensor[0] + "," + dataSensor[1] + "," + dataSensor[2] + "," + dataSensor[3] + "," + dataSensor[4] + "," + dataSensor[5] + "," + dataSensor[6] + "," + dataSensor[7] + "," + dataSensor[8] + "," + dataSensor[9] + "," + dataSensor[10] + "," + dataSensor[11] + "," + dataSensor[12] + "," + dataSensor[13] + "," + dataSensor[14] + "," + dataSensor[15] + "," + dataSensor[16] + "," + dataSensor[17] + "," + dataSensor[18] + "," + dataSensor[19] + "," + dataSensor[20] + "," + dataSensor[21] + "," + dataSensor[22] + "," + dataSensor[23] + ")");
                        break;
                    case 25:
                        dgv.Invoke(new Action(() => dgv.Rows.Add((dgv.Rows.Count + 1).ToString(), DT, dataSensor[0], dataSensor[1], dataSensor[2], dataSensor[3], dataSensor[4], dataSensor[5], dataSensor[6], dataSensor[7], dataSensor[8], dataSensor[9], dataSensor[10], dataSensor[11], dataSensor[12], dataSensor[13], dataSensor[14], dataSensor[15], dataSensor[16], dataSensor[17], dataSensor[18], dataSensor[19], dataSensor[20], dataSensor[21], dataSensor[22], dataSensor[23], dataSensor[24])));
                        connection.insert("insert into " + table + "(dTime,Models,P1,P2,P3,P4,P5,P6,P7,P8,P9,P10,P11,P12,P13,P14,P15,P16,P17,P18,P19,P20,P21,P22,P23,P24,P25) values('" + DT + "','" + product + "'," + dataSensor[0] + "," + dataSensor[1] + "," + dataSensor[2] + "," + dataSensor[3] + "," + dataSensor[4] + "," + dataSensor[5] + "," + dataSensor[6] + "," + dataSensor[7] + "," + dataSensor[8] + "," + dataSensor[9] + "," + dataSensor[10] + "," + dataSensor[11] + "," + dataSensor[12] + "," + dataSensor[13] + "," + dataSensor[14] + "," + dataSensor[15] + "," + dataSensor[16] + "," + dataSensor[17] + "," + dataSensor[18] + "," + dataSensor[19] + "," + dataSensor[20] + "," + dataSensor[21] + "," + dataSensor[22] + "," + dataSensor[23] + "," + dataSensor[24] + ")");
                        break;
                    case 26:
                        dgv.Invoke(new Action(() => dgv.Rows.Add((dgv.Rows.Count + 1).ToString(), DT, dataSensor[0], dataSensor[1], dataSensor[2], dataSensor[3], dataSensor[4], dataSensor[5], dataSensor[6], dataSensor[7], dataSensor[8], dataSensor[9], dataSensor[10], dataSensor[11], dataSensor[12], dataSensor[13], dataSensor[14], dataSensor[15], dataSensor[16], dataSensor[17], dataSensor[18], dataSensor[19], dataSensor[20], dataSensor[21], dataSensor[22], dataSensor[23], dataSensor[24], dataSensor[25])));
                        connection.insert("insert into " + table + "(dTime,Models,P1,P2,P3,P4,P5,P6,P7,P8,P9,P10,P11,P12,P13,P14,P15,P16,P17,P18,P19,P20,P21,P22,P23,P24,P25,P26) values('" + DT + "','" + product + "'," + dataSensor[0] + "," + dataSensor[1] + "," + dataSensor[2] + "," + dataSensor[3] + "," + dataSensor[4] + "," + dataSensor[5] + "," + dataSensor[6] + "," + dataSensor[7] + "," + dataSensor[8] + "," + dataSensor[9] + "," + dataSensor[10] + "," + dataSensor[11] + "," + dataSensor[12] + "," + dataSensor[13] + "," + dataSensor[14] + "," + dataSensor[15] + "," + dataSensor[16] + "," + dataSensor[17] + "," + dataSensor[18] + "," + dataSensor[19] + "," + dataSensor[20] + "," + dataSensor[21] + "," + dataSensor[22] + "," + dataSensor[23] + "," + dataSensor[24] + "," + dataSensor[25] + ")");
                        break;
                    case 27:
                        dgv.Invoke(new Action(() => dgv.Rows.Add((dgv.Rows.Count + 1).ToString(), DT, dataSensor[0], dataSensor[1], dataSensor[2], dataSensor[3], dataSensor[4], dataSensor[5], dataSensor[6], dataSensor[7], dataSensor[8], dataSensor[9], dataSensor[10], dataSensor[11], dataSensor[12], dataSensor[13], dataSensor[14], dataSensor[15], dataSensor[16], dataSensor[17], dataSensor[18], dataSensor[19], dataSensor[20], dataSensor[21], dataSensor[22], dataSensor[23], dataSensor[24], dataSensor[25], dataSensor[26])));
                        connection.insert("insert into " + table + "(dTime,Models,P1,P2,P3,P4,P5,P6,P7,P8,P9,P10,P11,P12,P13,P14,P15,P16,P17,P18,P19,P20,P21,P22,P23,P24,P25,P26,P27) values('" + DT + "','" + product + "'," + dataSensor[0] + "," + dataSensor[1] + "," + dataSensor[2] + "," + dataSensor[3] + "," + dataSensor[4] + "," + dataSensor[5] + "," + dataSensor[6] + "," + dataSensor[7] + "," + dataSensor[8] + "," + dataSensor[9] + "," + dataSensor[10] + "," + dataSensor[11] + "," + dataSensor[12] + "," + dataSensor[13] + "," + dataSensor[14] + "," + dataSensor[15] + "," + dataSensor[16] + "," + dataSensor[17] + "," + dataSensor[18] + "," + dataSensor[19] + "," + dataSensor[20] + "," + dataSensor[21] + "," + dataSensor[22] + "," + dataSensor[23] + "," + dataSensor[24] + "," + dataSensor[25] + "," + dataSensor[26] + ")");
                        break;
                    case 28:
                        dgv.Invoke(new Action(() => dgv.Rows.Add((dgv.Rows.Count + 1).ToString(), DT, dataSensor[0], dataSensor[1], dataSensor[2], dataSensor[3], dataSensor[4], dataSensor[5], dataSensor[6], dataSensor[7], dataSensor[8], dataSensor[9], dataSensor[10], dataSensor[11], dataSensor[12], dataSensor[13], dataSensor[14], dataSensor[15], dataSensor[16], dataSensor[17], dataSensor[18], dataSensor[19], dataSensor[20], dataSensor[21], dataSensor[22], dataSensor[23], dataSensor[24], dataSensor[25], dataSensor[26], dataSensor[27])));
                        connection.insert("insert into " + table + "(dTime,Models,P1,P2,P3,P4,P5,P6,P7,P8,P9,P10,P11,P12,P13,P14,P15,P16,P17,P18,P19,P20,P21,P22,P23,P24,P25,P26,P27,P28) values('" + DT + "','" + product + "'," + dataSensor[0] + "," + dataSensor[1] + "," + dataSensor[2] + "," + dataSensor[3] + "," + dataSensor[4] + "," + dataSensor[5] + "," + dataSensor[6] + "," + dataSensor[7] + "," + dataSensor[8] + "," + dataSensor[9] + "," + dataSensor[10] + "," + dataSensor[11] + "," + dataSensor[12] + "," + dataSensor[13] + "," + dataSensor[14] + "," + dataSensor[15] + "," + dataSensor[16] + "," + dataSensor[17] + "," + dataSensor[18] + "," + dataSensor[19] + "," + dataSensor[20] + "," + dataSensor[21] + "," + dataSensor[22] + "," + dataSensor[23] + "," + dataSensor[24] + "," + dataSensor[25] + "," + dataSensor[26] + "," + dataSensor[27] + ")");
                        break;
                    case 29:
                        dgv.Invoke(new Action(() => dgv.Rows.Add((dgv.Rows.Count + 1).ToString(), DT, dataSensor[0], dataSensor[1], dataSensor[2], dataSensor[3], dataSensor[4], dataSensor[5], dataSensor[6], dataSensor[7], dataSensor[8], dataSensor[9], dataSensor[10], dataSensor[11], dataSensor[12], dataSensor[13], dataSensor[14], dataSensor[15], dataSensor[16], dataSensor[17], dataSensor[18], dataSensor[19], dataSensor[20], dataSensor[21], dataSensor[22], dataSensor[23], dataSensor[24], dataSensor[25], dataSensor[26], dataSensor[27], dataSensor[28])));
                        connection.insert("insert into " + table + "(dTime,Models,P1,P2,P3,P4,P5,P6,P7,P8,P9,P10,P11,P12,P13,P14,P15,P16,P17,P18,P19,P20,P21,P22,P23,P24,P25,P26,P27,P28,P29) values('" + DT + "','" + product + "'," + dataSensor[0] + "," + dataSensor[1] + "," + dataSensor[2] + "," + dataSensor[3] + "," + dataSensor[4] + "," + dataSensor[5] + "," + dataSensor[6] + "," + dataSensor[7] + "," + dataSensor[8] + "," + dataSensor[9] + "," + dataSensor[10] + "," + dataSensor[11] + "," + dataSensor[12] + "," + dataSensor[13] + "," + dataSensor[14] + "," + dataSensor[15] + "," + dataSensor[16] + "," + dataSensor[17] + "," + dataSensor[18] + "," + dataSensor[19] + "," + dataSensor[20] + "," + dataSensor[21] + "," + dataSensor[22] + "," + dataSensor[23] + "," + dataSensor[24] + "," + dataSensor[25] + "," + dataSensor[26] + "," + dataSensor[27] + "," + dataSensor[28] + ")");
                        break;
                    case 30:
                        dgv.Invoke(new Action(() => dgv.Rows.Add((dgv.Rows.Count + 1).ToString(), DT, dataSensor[0], dataSensor[1], dataSensor[2], dataSensor[3], dataSensor[4], dataSensor[5], dataSensor[6], dataSensor[7], dataSensor[8], dataSensor[9], dataSensor[10], dataSensor[11], dataSensor[12], dataSensor[13], dataSensor[14], dataSensor[15], dataSensor[16], dataSensor[17], dataSensor[18], dataSensor[19], dataSensor[20], dataSensor[21], dataSensor[22], dataSensor[23], dataSensor[24], dataSensor[25], dataSensor[26], dataSensor[27], dataSensor[28], dataSensor[29])));
                        connection.insert("insert into " + table + "(dTime,Models,P1,P2,P3,P4,P5,P6,P7,P8,P9,P10,P11,P12,P13,P14,P15,P16,P17,P18,P19,P20,P21,P22,P23,P24,P25,P26,P27,P28,P29,P30) values('" + DT + "','" + product + "'," + dataSensor[0] + "," + dataSensor[1] + "," + dataSensor[2] + "," + dataSensor[3] + "," + dataSensor[4] + "," + dataSensor[5] + "," + dataSensor[6] + "," + dataSensor[7] + "," + dataSensor[8] + "," + dataSensor[9] + "," + dataSensor[10] + "," + dataSensor[11] + "," + dataSensor[12] + "," + dataSensor[13] + "," + dataSensor[14] + "," + dataSensor[15] + "," + dataSensor[16] + "," + dataSensor[17] + "," + dataSensor[18] + "," + dataSensor[19] + "," + dataSensor[20] + "," + dataSensor[21] + "," + dataSensor[22] + "," + dataSensor[23] + "," + dataSensor[24] + "," + dataSensor[25] + "," + dataSensor[26] + "," + dataSensor[27] + "," + dataSensor[28] + "," + dataSensor[29] + ")");
                        break;
                }
                dataSensor.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error visualizationData  \n" + ex.ToString(), "Data Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void plotDataGridP() 
        {
            if (valueSensor1340.Count == OffsetValue.Count)
            {
                for (int i = 0; i < valueSensor1340.Count; i++)
                {
                    double testvalue = Convert.ToDouble(valueSensor1340[i]);
                    double testoffset = Convert.ToDouble(OffsetValue[i]);
                    double sumtest = (testvalue - testoffset);
                    AfterOffset.Add(Convert.ToString(sumtest));
                }

                plotIt.Add((dataGridView1.Rows.Count + 1).ToString());
                plotIt.Add(DateTime.Now.ToString());
                plotIt.Add(tbModel.Text);

                foreach (var value in AfterOffset)
                {
                    plotIt.Add(value);
                }
                
                dataGridView1.Invoke(new Action(() => dataGridView1.Rows.Add((plotIt.ToArray()))));
                plotIt.Clear();
                valueSensor1340.Clear();
                AfterOffset.Clear();

            }
            else
            {
                MessageBox.Show("valueSensor and OffsetValue NOT MATCH");
                MessageBox.Show("Please check an OffsetValue and press Start");
                plotIt.Clear();
                valueSensor1340.Clear();
                AfterOffset.Clear();
                loopState = false;

            }
            visualizationData("DataDimensionMethod1", AfterOffset, dataGridView1, tbModel.Text);
            plotIt.Clear();
            valueSensor1340.Clear();
            AfterOffset.Clear();
        }

        TcpClient gg;
       
        NetworkStream sgg;
        StreamReader rg;
        private void getPointL()
        {
            gg = new TcpClient();
            gg.Connect("192.168.0.10", 8190);

            //WriteData(gg, sgg, "Result");

            string rawMsg = ReadData(gg, sgg, rg);
            //MessageBox.Show("getPoint : rawMsg = " + rawMsg);
            string[] dataMsg = rawMsg.Split(',');
            //MessageBox.Show("rawMsg.Length : " + dataMsg.Length);

            if ((dataMsg.Length > 0) && dataMsg[0] == "OK")
            {
                try
                {

                    for (int i = 1; i < dataMsg.Length; i++)
                    {
                        valueSensor1340Linear.Add((Convert.ToDouble(dataMsg[i]) / 1000.0).ToString());
                    }

                }
                catch (Exception ex)
                {

                    MessageBox.Show($"mainRun : read data error !! : {ex.Message}");

                }
            }
            gg.Dispose();
        }

        private void plotDataGridL()
        {
            //MessageBox.Show("valueSensor.Count : " + valueSensor1340Linear.Count + "\r\n" + "OffsetValue.Count : " + OffsetValue.Count);
            if (valueSensor1340Linear.Count >= 1)
            {

                plotIt.Add((dataGridView2.Rows.Count + 1).ToString()); //plotIt is List
                plotIt.Add(DateTime.Now.ToString());
                plotIt.Add(tbModel.Text);

                foreach (string value in valueSensor1340Linear)
                {
                    plotIt.Add(value);
                }

                dataGridView2.Invoke(new Action(() => dataGridView2.Rows.Add((plotIt.ToArray()))));
                plotIt.Clear();
                valueSensor1340Linear.Clear();
                AfterOffset.Clear();
            }
            visualizationData("DataDimensionMethod2", plotIt, dataGridView2, tbModel.Text);
            //MessageBox.Show("ValueSensor and OffsetValue NOT EQUAL");
            plotIt.Clear();
            valueSensor1340Linear.Clear();
            AfterOffset.Clear();
        }

        private void btnCSV_Click(object sender, EventArgs e)
        {
            ExtoCSV(dataGridView1, "ExportDataG1340");
        }

        public void ExtoCSV(DataGridView gv, string str)
        {
            String Date = DateTime.Now.ToString("ddMMyyyy");
            if (gv.Rows.Count > 0)
            {
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "CSV (*.csv)|*.csv";
                sfd.FileName = str + Date + ".csv";
                bool fileError = false;
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    if (File.Exists(sfd.FileName))
                    {
                        try
                        {
                            File.Delete(sfd.FileName);
                        }
                        catch (IOException ex)
                        {
                            fileError = true;
                            MessageBox.Show("It wasn't possible to write the data to the disk." + ex.Message);
                        }
                    }
                    if (!fileError)
                    {
                        int columnCount = gv.Columns.Count;
                        string columnNames = "";
                        string[] outputCsv = new string[gv.Rows.Count + 2];
                        for (int i = 0; i < columnCount; i++)
                        {
                            columnNames += gv.Columns[i].HeaderText.ToString() + ",";
                        }
                        outputCsv[0] += "Expoert Data";
                        outputCsv[1] += columnNames;

                        for (int i = 2; (i - 2) < gv.Rows.Count; i++)
                        {
                            for (int j = 0; j < columnCount; j++)
                            {
                                try
                                {
                                    outputCsv[i] += gv.Rows[i - 2].Cells[j].Value.ToString() + ",";
                                }
                                catch
                                {

                                }

                            }
                        }

                        File.WriteAllLines(sfd.FileName, outputCsv, Encoding.UTF8);
                        MessageBox.Show("Data Exported Successfully !!!", "Info");
                    }
                }
            }
            else
            {
                MessageBox.Show("No Record To Export !!!", "Info");
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            loopState = false;
            Environment.Exit(Environment.ExitCode);
        }


    }
}
