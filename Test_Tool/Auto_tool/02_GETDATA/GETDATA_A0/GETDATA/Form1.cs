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
　　　public static string Baudrate = "3000000";
　　　public static string selectCOM;
　　　public static string row2;
　　　public static string row3;
　　　public static string row4;
　　　public static string row5;
　　　public static string row6;
　　　public static string row7;
　　　public static byte[] a1;
　　　public static byte[] a2;
　　　public static byte[] b1;
　　　public static byte[] b2 = { 0x05, 0x5b, 0x07, 0x00, 0x00, 0x0a, 0x03, 0x00, 0x42, 0x32, 0x30, 0x00, 0x00, 0x00 }; //Compare Mode
　　　public static byte[] b3;
　　　public static byte[] b4 = { 0x2a, 0xd1, 0xa0 }; //Compare BT_MAC　
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
　　　DataTable dataTable = new DataTable();
　　　dataTable.Columns.Add("Port", typeof(int));
　　　dataTable.Columns.Add("COM", typeof(string));
　　　dataTable.Columns.Add("Mode", typeof(string));
　　　//dataTable.Columns.Add("Flash_ID", typeof(string));
　　　//dataTable.Columns.Add("FW", typeof(string));
　　　//dataTable.Columns.Add("HW", typeof(string));
　　　//dataTable.Columns.Add("SPI", typeof(string));
　　　dataTable.Columns.Add("BT_MAC", typeof(string));
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
　　　　//row[4] = " ";
　　　　//row[5] = " ";
　　　　//row[6] = " ";
　　　　//row[7] = " ";
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
　　　　//row[4] = " ";
　　　　//row[5] = " ";
　　　　//row[6] = " ";
　　　　//row[7] = " ";
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
　　　　//row[4] = " ";
　　　　//row[5] = " ";
　　　　//row[6] = " ";
　　　　//row[7] = " ";
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
　　　　//row[4] = " ";
　　　　//row[5] = " ";
　　　　//row[6] = " ";
　　　　//row[7] = " ";
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
　　　　sp.BaudRate = 3000000;　//波特率　   
　　　　sp.DataBits = 8;　//資料位
　　　　sp.PortName = Global.selectCOM; ;　//兩個停止位
　　　　sp.StopBits = System.IO.Ports.StopBits.One; //無奇偶校驗位
　　　　sp.Parity = System.IO.Ports.Parity.None;
　　　　sp.ReadTimeout = 100;
　　　　sp.Dispose();
　　　　Thread.Sleep(5);
　　　　sp.Close();
　　　　Thread.Sleep(5);
　　　　sp.Open();
　　　　if (!sp.IsOpen)
　　　　{
　　　　　richTextBox1.SelectionColor = Color.Red;
　　　　　richTextBox1.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + Global.selectCOM + " Open Error\r\n");
　　　　　return;
　　　　}
　　　　else
　　　　{
　　　　　richTextBox1.SelectionColor = Color.Lime;
　　　　　richTextBox1.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + Global.selectCOM + " Open Done\r\n");　　　　   
　　　　}
　　　　sp.DataReceived += sp_DataReceived;
　　　}
　　　catch (Exception ex)
　　　{
　　　　sp.Dispose();
　　　　richTextBox1.SelectionColor = Color.Red;
　　　　richTextBox1.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + Global.selectCOM + " " + ex.Message + "\r\n");
　　　}

　　　//
　　　//傳送Mode - 05 5a 04 00 00 0a 02 f2
　　　// 
　　　byte[] buffer = new byte[14]; //清空buffer
　　　byte[] buffer1 = new byte[14]; //清空buffer1
　　　sp.Write(new byte[] { 0x05, 0x5a, 0x03, 0x00, 0x0b, 0x02, 0x00 }, 0, 7); //喚醒設備
　　　Thread.Sleep(5);
　　　sp.Read(buffer1, 0, buffer1.Length);
　　　Thread.Sleep(5);
　　　sp.Write(new byte[] { 0x05, 0x5a, 0x06, 0x00, 0x00, 0x05, 0x00, 0x00, 0x00, 0x00 }, 0, 10); //喚醒設備
　　　Thread.Sleep(5);
　　　sp.Read(buffer1, 0, buffer1.Length);
　　　Thread.Sleep(5);
　　　sp.Write(new byte[] { 0x05, 0x5a, 0x02, 0x00, 0x01, 0x0e }, 0, 6); //喚醒設備
　　　Thread.Sleep(5);
　　　sp.Read(buffer1, 0, buffer1.Length);
　　　Thread.Sleep(5);

　　　sp.Write(new byte[] { 0x05, 0x5a, 0x04, 0x00, 0x00, 0x0a, 0x02, 0xf2 }, 0, 8);
　　　Thread.Sleep(5);
　　　sp.Read(buffer, 0, buffer.Length);
　　　Encoding.Default.GetString(buffer);
　　　Global.b1 = buffer;

　　　int i = 1;
　　　while (i <= 5)
　　　{
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
　　　　　Array.Clear(Global.a1, 0, Global.a1.Length);
　　　　　i = i + 15;
　　　　}
　　　　else
　　　　{
　　　　　sp.Write(new byte[] { 0x05, 0x5a, 0x03, 0x00, 0x0b, 0x02, 0x00 }, 0, 7); //喚醒設備
　　　　　Thread.Sleep(5);
　　　　　sp.Read(buffer1, 0, buffer1.Length);
　　　　　Thread.Sleep(5);
　　　　　sp.Write(new byte[] { 0x05, 0x5a, 0x06, 0x00, 0x00, 0x05, 0x00, 0x00, 0x00, 0x00 }, 0, 10); //喚醒設備
　　　　　Thread.Sleep(5);
　　　　　sp.Read(buffer1, 0, buffer1.Length);
　　　　　Thread.Sleep(5);
　　　　　sp.Write(new byte[] { 0x05, 0x5a, 0x02, 0x00, 0x01, 0x0e }, 0, 6); //喚醒設備
　　　　　Thread.Sleep(5);
　　　　　sp.Read(buffer1, 0, buffer1.Length);
　　　　　Thread.Sleep(5);

　　　　　sp.Write(new byte[] { 0x05, 0x5a, 0x04, 0x00, 0x00, 0x0a, 0x02, 0xf2 }, 0, 8);
　　　　　Thread.Sleep(5);
　　　　　sp.Read(buffer, 0, buffer.Length);
　　　　　Encoding.Default.GetString(buffer);
　　　　　Global.b1 = buffer;

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
　　　　　　Array.Clear(Global.a1, 0, Global.a1.Length);
　　　　　　i = i + 15;
　　　　　}
　　　　　else
　　　　　{
　　　　　　i++;
　　　　　}
　　　　　Array.Clear(Global.b1, 0, Global.b1.Length);
　　　　}
　　　　Array.Clear(Global.b1, 0, Global.b1.Length);
　　　}
　　　if (i == 6)
　　　{
　　　　richTextBox1.SelectionColor = Color.Red;
　　　　richTextBox1.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + Global.selectCOM + " Get Mode { " + Global.b1[8].ToString("X2") + Global.b1[9].ToString("X2") + Global.b1[10].ToString("X2") + " } Retry*5 Error" + "\r\n");
　　　　Thread.Sleep(5);
　　　　richTextBox1.ScrollToCaret();
　　　}
　　　Array.Clear(Global.b1, 0, Global.b1.Length);

　　　//
　　　//傳送BT_MAC CMD - 05 5a 04 00 00 0a 00 36
　　　//
　　　Array.Clear(buffer, 0, buffer.Length);
　　　Array.Clear(buffer1, 0, buffer1.Length);
　　　sp.Write(new byte[] { 0x05, 0x5a, 0x03, 0x00, 0x0b, 0x02, 0x00 }, 0, 7); //喚醒設備
　　　Thread.Sleep(5);
　　　sp.Read(buffer1, 0, buffer1.Length);
　　　Thread.Sleep(5);
　　　sp.Write(new byte[] { 0x05, 0x5a, 0x06, 0x00, 0x00, 0x05, 0x00, 0x00, 0x00, 0x00 }, 0, 10); //喚醒設備
　　　Thread.Sleep(5);
　　　sp.Read(buffer1, 0, buffer1.Length);
　　　Thread.Sleep(5);
　　　sp.Write(new byte[] { 0x05, 0x5a, 0x02, 0x00, 0x01, 0x0e }, 0, 6); //喚醒設備
　　　Thread.Sleep(5);
　　　sp.Read(buffer1, 0, buffer1.Length);
　　　Thread.Sleep(5);

　　　sp.Write(new byte[] { 0x05, 0x5a, 0x04, 0x00, 0x00, 0x0a, 0x00, 0x36 }, 0, 8);
　　　Thread.Sleep(5);
　　　sp.Read(buffer, 0, buffer.Length);
　　　Encoding.Default.GetString(buffer);

　　　Global.a2 = new byte[3];
　　　Global.a2[0] = Global.b1[11];
　　　Global.a2[1] = Global.b1[12];
　　　Global.a2[2] = Global.b1[13];
　　　Global.b3 = Global.a2;

　　　i = 1;
　　　while (i <= 5)
　　　{
　　　　if (BitConverter.ToString(Global.b3) == BitConverter.ToString(Global.b4))
　　　　{
　　　　　Global.row3 = (String.Format("{0:X2}{1:X2}{2:X2}{3:X2}{4:X2}{5:X2}", buffer[13], buffer[12], buffer[11], buffer[10], buffer[9], buffer[8])); //成功秀出byte並反轉　　　　　
　　　　　richTextBox1.SelectionColor = Color.Lime;
　　　　　richTextBox1.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + Global.selectCOM + " Get BT_MAC { " + Global.row3 + " } Done" + "\r\n");
　　　　　richTextBox1.ScrollToCaret();
　　　　　Array.Clear(Global.a2, 0, Global.a2.Length);
　　　　　i = i + 15;
　　　　}
　　　　else
　　　　{
　　　　　Array.Clear(buffer, 0, buffer.Length);
　　　　　Array.Clear(buffer1, 0, buffer1.Length);
　　　　　sp.Write(new byte[] { 0x05, 0x5a, 0x03, 0x00, 0x0b, 0x02, 0x00 }, 0, 7); //喚醒設備
　　　　　Thread.Sleep(5);
　　　　　sp.Read(buffer1, 0, buffer1.Length);
　　　　　Thread.Sleep(5);
　　　　　sp.Write(new byte[] { 0x05, 0x5a, 0x06, 0x00, 0x00, 0x05, 0x00, 0x00, 0x00, 0x00 }, 0, 10); //喚醒設備
　　　　　Thread.Sleep(5);
　　　　　sp.Read(buffer1, 0, buffer1.Length);
　　　　　Thread.Sleep(5);
　　　　　sp.Write(new byte[] { 0x05, 0x5a, 0x02, 0x00, 0x01, 0x0e }, 0, 6); //喚醒設備
　　　　　Thread.Sleep(5);
　　　　　sp.Read(buffer1, 0, buffer1.Length);
　　　　　Thread.Sleep(5);

　　　　　//
　　　　　//傳送BT_MAC CMD - 05 5a 04 00 00 0a 00 36
　　　　　//
　　　　　sp.Write(new byte[] { 0x05, 0x5a, 0x04, 0x00, 0x00, 0x0a, 0x00, 0x36 }, 0, 8);
　　　　　Thread.Sleep(5);
　　　　　sp.Read(buffer, 0, buffer.Length);
　　　　　Encoding.Default.GetString(buffer);

　　　　　Global.a2 = new byte[3];
　　　　　Global.a2[0] = Global.b1[11];
　　　　　Global.a2[1] = Global.b1[12];
　　　　　Global.a2[2] = Global.b1[13];
　　　　　Global.b3 = Global.a2;

　　　　　if (BitConverter.ToString(Global.b3) == BitConverter.ToString(Global.b4))
　　　　　{
　　　　　　Global.row3 = (String.Format("{0:X2}{1:X2}{2:X2}{3:X2}{4:X2}{5:X2}", buffer[13], buffer[12], buffer[11], buffer[10], buffer[9], buffer[8])); //成功秀出byte並反轉
　　　　　　i = i + 15; //跳出迴圈
　　　　　　richTextBox1.SelectionColor = Color.Lime;
　　　　　　richTextBox1.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + Global.selectCOM + " Get BT_MAC { " + Global.row3 + " } Done" + "\r\n");
　　　　　　richTextBox1.ScrollToCaret();
　　　　　　Array.Clear(Global.a2, 0, Global.a2.Length);
　　　　　}
　　　　　else
　　　　　{
　　　　　　i++;
　　　　　}
　　　　　Array.Clear(Global.b3, 0, Global.b3.Length);
　　　　}
　　　　Array.Clear(Global.b3, 0, Global.b3.Length);
　　　}
　　　if (i == 6)
　　　{
　　　　Global.row3 = (String.Format("{0:X2}{1:X2}{2:X2}{3:X2}{4:X2}{5:X2}", buffer[13], buffer[12], buffer[11], buffer[10], buffer[9], buffer[8])); //失敗秀出byte並反轉
　　　　Thread.Sleep(5);
　　　　richTextBox1.SelectionColor = Color.Red;
　　　　richTextBox1.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + Global.selectCOM + " Get BT_MAC { " + Global.row3 + " } Retry*5 Error" + "\r\n");
　　　　richTextBox1.ScrollToCaret();
　　　}
　　　Array.Clear(Global.b1, 0, Global.b1.Length);
　　　Array.Clear(Global.b3, 0, Global.b3.Length);
　　　sp.Dispose();
　　　sp.Close();　　  
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
