using System;
using System.Data;
using System.Windows.Forms;
using System.IO.Ports;
using System.IO;
using System.Threading;
using System.Text;
using System.Drawing;
using System.Data.SqlClient;

namespace GETDATA
{
　public partial class GETDATA : Form
　{
　　public class Global //全域變數
　　{
      public static byte[] buffer;
      public static byte[] buffer1;
      public static string Baudrate = "3000000";
　　　public static string selectCOM;
　　　public static string row2; //port
　　　public static string row3; //mode
　　　public static string row4; //compare
　　　public static string row5; //hw
　　　public static string row6; //fw
　　　public static string row7; //SPI Flash
　　　public static byte[] a1; //select Mode byte[] for cmopare
　　　public static byte[] a2; //select BT_MAC byte[] for cmopare
      public static byte[] a3; //select HW byte[] for cmopare
      public static byte[] a4; //select FW byte[] for cmopare
      public static byte[] a5; //select SPI_Flash byte[] for cmopare
      public static byte[] b1;
　　　public static byte[] b2 = { 0x05, 0x5b, 0x07, 0x00, 0x00, 0x0a, 0x03, 0x00, 0x42, 0x32, 0x30, 0x00, 0x00, 0x00 }; //Compare Mode
　　　public static byte[] b3;
　　　public static byte[] b4 = { 0x2a, 0xd1, 0xa0 }; //Compare BT_MAC
      public static byte[] b5;
      public static byte[] b6 = { 0x05, 0x01, 0x01, 0x02, 0x01, 0x3c, 0xf2, 0xa5}; //Compare HW & FW & SPI Flash
    }

　　public GETDATA()
　　{
　　　InitializeComponent();
　　}

　　private void ScanCOM_Click(object sender, EventArgs e)
　　{
　　　ReCOM();
　　}

　　private void Form1_Load(object sender, EventArgs e)
　　{
　　　ReCOM();
    }

　　private void Start_Click(object sender, EventArgs e)
　　{
      Thread.Sleep(5);
      DataTable dataTable = new DataTable();
　　　dataTable.Columns.Add("Port", typeof(int)); //row[0]
　　　dataTable.Columns.Add("COM", typeof(string)); //row[1]
      dataTable.Columns.Add("Mode", typeof(string)); //row[2]
      dataTable.Columns.Add("BT_MAC", typeof(string)); //row[3]
      dataTable.Columns.Add("HW", typeof(string)); //row[4]
      dataTable.Columns.Add("FW", typeof(string)); //row[5]
      dataTable.Columns.Add("SPI_Flash", typeof(string)); //row[6]
      dataGridView.DataSource = dataTable;

　　　if (comlist1.SelectedIndex != -1)
　　　{
　　　　Global.selectCOM = comlist1.Text;
　　　　SendByteArray();　　　　　　　

　　　　DataRow row = dataTable.NewRow();
　　　　row[0] = 1;
　　　　row[1] = Global.selectCOM;
　　　　row[2] = Global.row2;
　　　　row[3] = Global.row3;
　　　　row[4] = Global.row4;
　　　　row[5] = Global.row5;
　　　　row[6] = Global.row6;
　　　　dataTable.Rows.Add(row);
　　　}
　　　if (comlist2.SelectedIndex != -1)
　　　{
　　　　Global.selectCOM = comlist2.Text;
　　　　SendByteArray();

　　　　DataRow row = dataTable.NewRow();
　　　　row[0] = 2;
　　　　row[1] = Global.selectCOM;
        row[2] = Global.row2;
        row[3] = Global.row3;
        row[4] = Global.row4;
        row[5] = Global.row5;
        row[6] = Global.row6;
        dataTable.Rows.Add(row);
　　　}
　　　if (comlist3.SelectedIndex != -1)
　　　{
　　　　Global.selectCOM = comlist3.Text;
　　　　SendByteArray();

　　　　DataRow row = dataTable.NewRow();
　　　　row[0] = 3;
　　　　row[1] = Global.selectCOM;
        row[2] = Global.row2;
        row[3] = Global.row3;
        row[4] = Global.row4;
        row[5] = Global.row5;
        row[6] = Global.row6;
        dataTable.Rows.Add(row);
　　　}
　　　if (comlist4.SelectedIndex != -1)
　　　{
　　　　Global.selectCOM = comlist4.Text;
　　　　SendByteArray();

　　　　DataRow row = dataTable.NewRow();
　　　　row[0] = 4;
　　　　row[1] = Global.selectCOM;
　　　　row[2] = Global.row2;
　　　　row[3] = Global.row3;
　　　　row[4] = Global.row4;
　　　　row[5] = Global.row5;
　　　　row[6] = Global.row6;
　　　　dataTable.Rows.Add(row);
　　　}
　　}

　　public void ReCOM()
　　{
　　　//rescan combobox
　　　comlist1.Items.Clear();
　　　comlist2.Items.Clear();
　　　comlist3.Items.Clear();
　　　comlist4.Items.Clear();

　　　string[] ports = SerialPort.GetPortNames();
　　　foreach (string port in ports)
　　　{
　　　　comlist1.Items.Add(port);
　　　　comlist2.Items.Add(port);
　　　　comlist3.Items.Add(port);
　　　　comlist4.Items.Add(port);
　　　}
　　}
　　
　　public void SendByteArray()
　　{            
      SerialPort sp = new SerialPort();
      try
　　　{
　　　　sp.BaudRate = 3000000; //波特率　   
　　　　sp.DataBits = 8; //資料位
　　　　sp.PortName = Global.selectCOM; ; //兩個停止位
　　　　sp.StopBits = System.IO.Ports.StopBits.One; //無奇偶校驗位
　　　　sp.Parity = System.IO.Ports.Parity.None;
　　　　sp.ReadTimeout = 100;
　　　　sp.Dispose();
　　　　Thread.Sleep(5);
　　　　sp.Close();
　　　　Thread.Sleep(5);
　　　　sp.Open();
        sp.DiscardInBuffer(); //RX
        sp.DiscardOutBuffer(); //TX
        if (!sp.IsOpen)
　　　　{
　　　　　richTextBox1.SelectionColor = Color.Red;
　　　　　richTextBox1.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + Global.selectCOM + " Connect Error\r\n");
          richTextBox1.ScrollToCaret();
          return;
　　　　}
　　　　else
　　　　{
　　　　　richTextBox1.SelectionColor = Color.Lime;
　　　　　richTextBox1.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + Global.selectCOM + " Connect Done\r\n");
          richTextBox1.ScrollToCaret();
        }
　　　　sp.DataReceived += sp_DataReceived;
　　　}
　　　catch (Exception ex)
　　　{
　　　　sp.Dispose();
　　　　richTextBox1.SelectionColor = Color.Red;
　　　　richTextBox1.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + Global.selectCOM + " " + ex.Message + "\r\n");
        richTextBox1.ScrollToCaret();
      }
      Thread.Sleep(100);

      //
      //傳送Mode - 05 5a 04 00 00 0a 02 f2
      // 
　　　int i = 1;
　　　while (i <= 5)
　　　{
　　　　Global.buffer = new byte[14]; //清空buffer
        Global.buffer1 = new byte[14]; //清空buffer1
        sp.Write(new byte[] { 0x05, 0x5a, 0x03, 0x00, 0x0b, 0x02, 0x00 }, 0, 7); //喚醒設備
        Thread.Sleep(5);
        sp.Read(Global.buffer1, 0, Global.buffer1.Length);
        Thread.Sleep(5);
        sp.Write(new byte[] { 0x05, 0x5a, 0x06, 0x00, 0x00, 0x05, 0x00, 0x00, 0x00, 0x00 }, 0, 10); //喚醒設備
        Thread.Sleep(5);
        sp.Read(Global.buffer1, 0, Global.buffer1.Length);
        Thread.Sleep(5);
        sp.Write(new byte[] { 0x05, 0x5a, 0x02, 0x00, 0x01, 0x0e }, 0, 6); //喚醒設備
        Thread.Sleep(5);
        sp.Read(Global.buffer1, 0, Global.buffer1.Length);
        Thread.Sleep(5);
　　　　sp.DiscardInBuffer(); //RX
        sp.DiscardOutBuffer(); //TX

        sp.Write(new byte[] { 0x05, 0x5a, 0x04, 0x00, 0x00, 0x0a, 0x02, 0xf2 }, 0, 8);
        Thread.Sleep(2);
        sp.Read(Global.buffer, 0, Global.buffer.Length);
        Encoding.Default.GetString(Global.buffer);
        Global.b1 = Global.buffer;

        if (BitConverter.ToString(Global.b1) == BitConverter.ToString(Global.b2))
　　　　{
　　　　　Global.a1 = new byte[3];
　　　　　Global.a1[0] = Global.b1[8];
　　　　　Global.a1[1] = Global.b1[9];
　　　　　Global.a1[2] = Global.b1[10];
          Global.row2 = Encoding.Default.GetString(Global.a1);
          richTextBox1.SelectionColor = Color.Lime;
　　　　　richTextBox1.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + Global.selectCOM + " Get Mode { " + Global.row2 + " } Done" + "\r\n");
　　　　　Thread.Sleep(5);
　　　　　richTextBox1.ScrollToCaret();　　　　　
　　　　　i = i + 15;
　　　　}
　　　　else
　　　　{
　　　　 i++;
　　　　}
　　　}
　　　if (i == 6)
　　　{
          Global.row2 = String.Format("{0:X2}{1:X2}{2:X2}{3:X2}{4:X2}{5:X2}", Global.buffer[8], Global.buffer[9], Global.buffer[10], Global.buffer[11], Global.buffer[12], Global.buffer[13]);
　　　　　richTextBox1.SelectionColor = Color.Red;
　　　　　richTextBox1.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + Global.selectCOM + " Get Mode { " + Global.row2 + " } Retry*5 Error" + "\r\n");
　　　　　Thread.Sleep(5);
　　　　　richTextBox1.ScrollToCaret();
　　　}
      Thread.Sleep(50);
      sp.DiscardInBuffer(); //RX
      sp.DiscardOutBuffer(); //TX

      //
      //傳送BT_MAC CMD - 05 5a 04 00 00 0a 00 36
      //
　　　i = 1;
　　　while (i <= 5)
　　　{
　　　　Global.buffer = new byte[14]; //清空buffer
        Global.buffer1 = new byte[14]; //清空buffer1
        sp.Write(new byte[] { 0x05, 0x5a, 0x03, 0x00, 0x0b, 0x02, 0x00 }, 0, 7); //喚醒設備
        Thread.Sleep(5);
        sp.Read(Global.buffer1, 0, Global.buffer1.Length);
        Thread.Sleep(5);
        sp.Write(new byte[] { 0x05, 0x5a, 0x06, 0x00, 0x00, 0x05, 0x00, 0x00, 0x00, 0x00 }, 0, 10); //喚醒設備
        Thread.Sleep(5);
        sp.Read(Global.buffer1, 0, Global.buffer1.Length);
        Thread.Sleep(5);
        sp.Write(new byte[] { 0x05, 0x5a, 0x02, 0x00, 0x01, 0x0e }, 0, 6); //喚醒設備
        Thread.Sleep(5);
        sp.Read(Global.buffer1, 0, Global.buffer1.Length);
        Thread.Sleep(5);
        sp.DiscardInBuffer(); //RX
        sp.DiscardOutBuffer(); //TX

        sp.Write(new byte[] { 0x05, 0x5a, 0x04, 0x00, 0x00, 0x0a, 0x00, 0x36 }, 0, 8);
        Thread.Sleep(2);
        sp.Read(Global.buffer, 0, Global.buffer.Length);
        Encoding.Default.GetString(Global.buffer);
        Global.b3 = Global.buffer;

        Global.a2 = new byte[3];
        Global.a2[0] = Global.b3[11];
        Global.a2[1] = Global.b3[12];
        Global.a2[2] = Global.b3[13];
        Global.b3 = Global.a2;
        if (BitConverter.ToString(Global.b3) == BitConverter.ToString(Global.b4))
　　　　{
　　　　　Global.row3 = (String.Format("{0:X2}{1:X2}{2:X2}{3:X2}{4:X2}{5:X2}", Global.buffer[13], Global.buffer[12], Global.buffer[11], Global.buffer[10], Global.buffer[9], Global.buffer[8])); //成功秀出byte並反轉　　　　　
　　　　　richTextBox1.SelectionColor = Color.Lime;
　　　　　richTextBox1.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + Global.selectCOM + " Get BT_MAC { " + Global.row3 + " } Done" + "\r\n");
　　　　　richTextBox1.ScrollToCaret();
　　　　　Array.Clear(Global.a2, 0, Global.a2.Length);
　　　　　i = i + 15;
　　　　}
　　　　else
　　　　{
　　　　　i++;
　　　　}
　　　}
　　　if (i == 6)
　　　{
　　　　Global.row3 = (String.Format("{0:X2}{1:X2}{2:X2}{3:X2}{4:X2}{5:X2}", Global.buffer[13], Global.buffer[12], Global.buffer[11], Global.buffer[10], Global.buffer[9], Global.buffer[8])); //失敗秀出byte並反轉
　　　　Thread.Sleep(5);
　　　　richTextBox1.SelectionColor = Color.Red;
　　　　richTextBox1.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + Global.selectCOM + " Get BT_MAC { " + Global.row3 + " } Retry*5 Error" + "\r\n");
　　　　richTextBox1.ScrollToCaret();
　　　}
      Thread.Sleep(50);
      sp.DiscardInBuffer(); //RX
      sp.DiscardOutBuffer(); //TX

      //
      //傳送HW & FW & SPI_Flash - 05 5a 04 00 00 0a 01 f4
      //
      i = 1;
　　　while (i <= 5)
　　　{
        Global.buffer = new byte[18]; //清空buffer
        Global.buffer1 = new byte[18]; //清空buffer1
        sp.Write(new byte[] { 0x05, 0x5a, 0x03, 0x00, 0x0b, 0x02, 0x00 }, 0, 7); //喚醒設備
        Thread.Sleep(5);
        sp.Read(Global.buffer1, 0, Global.buffer1.Length);
        Thread.Sleep(5);
        sp.Write(new byte[] { 0x05, 0x5a, 0x06, 0x00, 0x00, 0x05, 0x00, 0x00, 0x00, 0x00 }, 0, 10); //喚醒設備
        Thread.Sleep(5);
        sp.Read(Global.buffer1, 0, Global.buffer1.Length);
        Thread.Sleep(5);
        sp.Write(new byte[] { 0x05, 0x5a, 0x02, 0x00, 0x01, 0x0e }, 0, 6); //喚醒設備
        Thread.Sleep(5);
        sp.Read(Global.buffer1, 0, Global.buffer1.Length);
        Thread.Sleep(5);
        sp.DiscardInBuffer(); //RX
        sp.DiscardOutBuffer(); //TX

        sp.Write(new byte[] { 0x05, 0x5a, 0x04, 0x00, 0x00, 0x0a, 0x01, 0xf4 }, 0, 8);
        Thread.Sleep(200);
        sp.Read(Global.buffer, 0, Global.buffer.Length);
        Encoding.Default.GetString(Global.buffer);
        Global.b5 = Global.buffer;

        //Compare HW
        Global.a3 = new byte[8];
        Global.a3[0] = Global.b5[10];
        Global.a3[1] = Global.b5[11];
        Global.a3[2] = Global.b5[12];
        Global.a3[3] = Global.b5[13];
        Global.a3[4] = Global.b5[14];
        Global.a3[5] = Global.b5[15];
        Global.a3[6] = Global.b5[16];
        Global.a3[7] = Global.b5[17];
        Global.b5 = Global.a3;
        if (BitConverter.ToString(Global.b5) == BitConverter.ToString(Global.b6))
　　　　{
        　//Output HW
　　　　　Global.row4 = (String.Format("{0:X2}", Global.buffer[10]));　　　　
　　　　　richTextBox1.SelectionColor = Color.Lime;
　　　　　richTextBox1.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + Global.selectCOM + " Get HW { " + Global.row4 + " } Done" + "\r\n");
　　　　　richTextBox1.ScrollToCaret();
          //Output FW
          Global.row5 = (String.Format("{0:X}.{1:X}.{2:X}.{3:X}", Global.buffer[11], Global.buffer[12], Global.buffer[13], Global.buffer[14]));
          richTextBox1.SelectionColor = Color.Lime;
          richTextBox1.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + Global.selectCOM + " Get FW { " + Global.row5 + " } Done" + "\r\n");
          richTextBox1.ScrollToCaret();
          //Output SPI_Flash
          Global.row6 = (String.Format("{0:X2}{1:X2}{2:X2}", Global.buffer[15], Global.buffer[16], Global.buffer[17]));
          richTextBox1.SelectionColor = Color.Lime;
          richTextBox1.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + Global.selectCOM + " Get SPI_Flash { " + Global.row6 + " } Done" + "\r\n");
          richTextBox1.ScrollToCaret();
          Array.Clear(Global.a3, 0, Global.a3.Length);
　　　　　i = i + 15;
　　　　}
　　　　else
　　　　{
　　　　　i++;
　　　  }
　　　}
　　　if (i == 6)
　　　{　　　　
　　　　Thread.Sleep(5);
        if (Global.buffer[10] != 0x05)
        {
          Global.row4 = (String.Format("{0:X2}", Global.buffer[10]));
          richTextBox1.SelectionColor = Color.Red;
          richTextBox1.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + Global.selectCOM + " Get HW { " + Global.row4 + " } Retry*5 Error" + "\r\n");
          richTextBox1.ScrollToCaret();
        }
        else
        {
          richTextBox1.SelectionColor = Color.Lime;
          richTextBox1.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + Global.selectCOM + " Get HW { 05 } Done" + "\r\n");
          richTextBox1.ScrollToCaret();
        }

        Global.row5 = (String.Format("{0:X2}{1:X2}{2:X2}{3:X2}", Global.buffer[11], Global.buffer[12], Global.buffer[13], Global.buffer[14]));
        if (Global.row5 != "01010201")
        {
          richTextBox1.SelectionColor = Color.Red;
          richTextBox1.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + Global.selectCOM + " Get FW { " + Global.row5 + " } Retry*5 Error" + "\r\n");
          richTextBox1.ScrollToCaret();
        }
        else
        {
          richTextBox1.SelectionColor = Color.Lime;
          richTextBox1.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + Global.selectCOM + " Get HW { 1.1.2.1 } Done" + "\r\n");
          richTextBox1.ScrollToCaret();
        }
        
        Global.row6 = (String.Format("{0:X2}{1:X2}", Global.buffer[15], Global.buffer[16]));
        if (Global.row6 != "3cf2")
        {
          richTextBox1.SelectionColor = Color.Red;
          richTextBox1.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + Global.selectCOM + " Get SPI_Flash { " + String.Format("{0:X2}{1:X2}{2:X2}", Global.buffer[15], Global.buffer[16], Global.buffer[17]) + " } Retry*5 Error" + "\r\n");
          richTextBox1.ScrollToCaret();
        }
        else
        {
          richTextBox1.SelectionColor = Color.Red;
          richTextBox1.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + Global.selectCOM + " Get SPI_Flash { " + String.Format("{0:X2}{1:X2}{2:X2}", Global.buffer[15], Global.buffer[16], Global.buffer[17]) + " } Retry*5 Error" + "\r\n");
          richTextBox1.ScrollToCaret();
        }  
      }
　　　sp.Dispose();
　　　sp.Close();
      richTextBox1.SelectionColor = Color.Lime;
      richTextBox1.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + Global.selectCOM + " DisConnect Done" + "\r\n");
      richTextBox1.ScrollToCaret();
    }

　　private void Clear_Click(object sender, EventArgs e)
　　{
　　　//clear dataGridView cut
　　　dataGridView.CancelEdit();
　　　dataGridView.Columns.Clear();
　　　dataGridView.DataSource = null;

　　　//clear combobox cut
　　　comlist1.SelectedIndex = -1;
　　　comlist2.SelectedIndex = -1;
　　　comlist3.SelectedIndex = -1;
　　　comlist4.SelectedIndex = -1;

　　　SerialPort sp = new SerialPort();
　　　sp.Dispose();
　　　sp.Close();
      richTextBox1.SelectionColor = Color.Lime;
      richTextBox1.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + Global.selectCOM + " DisConnect Done" + "\r\n");
      richTextBox1.ScrollToCaret();

      //clear richTextBox1
      richTextBox1.Clear();
　　}

　　private void sp_DataReceived(object sender, SerialDataReceivedEventArgs e)
　　{

　　}

　　private void FrmExport_Load(object sender, EventArgs e)
　　{
　　　SqlConnection sqlCon;
　　　string conString = null;
　　　string sqlQuery = null;

　　　conString = "Data Source=.;Initial Catalog=DemoTest;Integrated Security=SSPI;";
　　　sqlCon = new SqlConnection(conString);
　　　sqlCon.Open();
　　　sqlQuery = "SELECT * FROM tblEmployee";
　　　SqlDataAdapter dscmd = new SqlDataAdapter(sqlQuery, sqlCon);
　　　DataTable dtData = new DataTable();
　　　dscmd.Fill(dtData);
　　　dataGridView.DataSource = dtData;
　　}

　　private void Save_to_Excel_Click(object sender, EventArgs e)
　　{
　　　if (dataGridView.Rows.Count > 0)
　　　{
　　　　SaveFileDialog sfd = new SaveFileDialog();
　　　　sfd.Filter = "CSV (*.csv)|*.csv";
　　　　sfd.FileName = DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".csv";
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
　　　　　　　richTextBox1.SelectionColor = Color.Red;
　　　　　　　richTextBox1.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " It wasn't possible to write the data to the disk." + ex.Message + "\r\n");
　　　　　　　richTextBox1.ScrollToCaret();
　　　　　　}
　　　　　}
　　　　　if (!fileError)
　　　　　{
　　　　　　try
　　　　　　{
　　　　　　　int columnCount = dataGridView.Columns.Count;
　　　　　　　string columnNames = "";
　　　　　　　string[] outputCsv = new string[dataGridView.Rows.Count + 1];
　　　　　　　for (int i = 0; i < columnCount; i++)
　　　　　　　{
　　　　　　　　columnNames += dataGridView.Columns[i].HeaderText.ToString() + ",";
　　　　　　　}
　　　　　　　outputCsv[0] += columnNames;

　　　　　　　for (int i = 1; (i - 1) < dataGridView.Rows.Count; i++)
　　　　　　　{
　　　　　　　　for (int j = 0; j < columnCount; j++)
　　　　　　　　{
　　　　　　　　　outputCsv[i] += dataGridView.Rows[i - 1].Cells[j].Value.ToString() + ",";
　　　　　　　　}
　　　　　　　}

　　　　　　　File.WriteAllLines(sfd.FileName, outputCsv, Encoding.UTF8);
　　　　　　　richTextBox1.SelectionColor = Color.Lime;
　　　　　　　richTextBox1.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " Data Exported Successfully to " + sfd.FileName + "\r\n");
　　　　　　　richTextBox1.ScrollToCaret();
　　　　　　}
　　　　　　catch (Exception ex)
　　　　　　{

　　　　　　　richTextBox1.SelectionColor = Color.Red;
　　　　　　　richTextBox1.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " Error :" + ex.Message + "\r\n");
　　　　　　　richTextBox1.ScrollToCaret();
　　　　　　}
　　　　　}
　　　　}
　　　}
　　　else
　　　{
　　　　richTextBox1.SelectionColor = Color.Red;
　　　　richTextBox1.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " No Record To Export !!!" + "\r\n");
　　　　richTextBox1.ScrollToCaret();
　　　}
　　}
　}
}
