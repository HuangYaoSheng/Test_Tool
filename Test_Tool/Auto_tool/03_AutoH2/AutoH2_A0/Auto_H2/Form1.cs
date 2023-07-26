// AutoH2_A0
// 使用第三方程式 h2testw.exe www.ctmagazin.de
// 使用Autoit.dll免費開源研編寫自動化程式
// 程式開發者YS.huang

using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Management;
using System.Diagnostics;
using System.Runtime.InteropServices;
using AutoIt;
using System.Drawing;
using System.Collections.Generic;

namespace AutoH2
{
    public partial class AutoH2 : Form
    {
        private ComboBox[] comboBoxes;
        private DriveInfo[] drives;
        private string[] selectedDrives;
        private ManagementEventWatcher driveInsertedWatcher;
        private ManagementEventWatcher driveRemovedWatcher;

        [DllImport("user32.dll")]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        private const int SWP_NOSIZE = 0x0001;
        private const int SWP_NOZORDER = 0x0004;
        private const int BM_CLICK = 0x00F5;

        public AutoH2()
        {
            InitializeComponent();

            comboBoxes = new ComboBox[]
            {
                comboBox1, comboBox2, comboBox3, comboBox4, comboBox5, comboBox6, comboBox7, comboBox8,
                comboBox9, comboBox10, comboBox11, comboBox12, comboBox13, comboBox14, comboBox15, comboBox16
            };

            foreach (ComboBox comboBox in comboBoxes)
            {
                comboBox.DropDownStyle = ComboBoxStyle.DropDownList;
                comboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;
            }

            selectedDrives = new string[comboBoxes.Length];

            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;

            this.Load += AutoH2_Load;
            this.FormClosing += AutoH2_FormClosing;
        }

        private void AutoH2_Load(object sender, EventArgs e)
        {
            RegisterDriveNotifications();
            UpdateDriveList();

            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
            int screenHeight = Screen.PrimaryScreen.Bounds.Height;

            int windowWidth = this.Width;
            int windowHeight = this.Height;

            int left = (screenWidth - windowWidth) / 2;
            int top = (screenHeight - windowHeight) / 2;

            this.StartPosition = FormStartPosition.Manual;
            this.Location = new Point(left, top);
        }

        private void AutoH2_FormClosing(object sender, FormClosingEventArgs e)
        {
            UnregisterDriveNotifications();
        }

        private void UpdateDriveList()
        {
            drives = DriveInfo.GetDrives()
                .Where(drive => drive.DriveType == DriveType.Removable)
                .OrderBy(drive => drive.Name)
                .ToArray();

            UpdateDriveComboBoxes();
        }

        private void UpdateDriveComboBoxes()
        {
            if (this.InvokeRequired)
            {
                this.Invoke((MethodInvoker)UpdateDriveComboBoxes);
                return;
            }

            string[] existingDriveNames = new string[comboBoxes.Length];

            for (int i = 0; i < comboBoxes.Length; i++)
            {
                existingDriveNames[i] = comboBoxes[i].Text;
            }

            // Create a new list to hold the drive names in the order of insertion
            List<string> driveNames = new List<string>();

            foreach (DriveInfo drive in drives)
            {
                driveNames.Add(drive.Name);
            }

            // Remove the existing drive names from the driveNames list
            foreach (string existingDriveName in existingDriveNames)
            {
                if (driveNames.Contains(existingDriveName))
                {
                    driveNames.Remove(existingDriveName);
                }
            }

            // Add the existing drive names at the end of the driveNames list
            driveNames.AddRange(existingDriveNames.Where(name => !string.IsNullOrEmpty(name)));

            // Update each comboBox with the drive names in the order of insertion
            for (int i = 0; i < comboBoxes.Length; i++)
            {
                ComboBox comboBox = comboBoxes[i];
                comboBox.SelectedIndexChanged -= ComboBox_SelectedIndexChanged;

                comboBox.Items.Clear();
                comboBox.Items.Add("");

                // Add the drive names to the comboBox in the order of insertion
                foreach (string driveName in driveNames)
                {
                    comboBox.Items.Add(driveName);
                }

                comboBox.Text = existingDriveNames[i] ?? "";

                comboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;
            }
        }



        private void ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox selectedComboBox = (ComboBox)sender;

            bool isDuplicateSelected = false;

            for (int i = 0; i < comboBoxes.Length; i++)
            {
                ComboBox comboBox = comboBoxes[i];

                if (comboBox != selectedComboBox && comboBox.Text == selectedComboBox.Text)
                {
                    isDuplicateSelected = true;
                    break;
                }
            }

            if (isDuplicateSelected)
            {
                selectedComboBox.SelectedIndex = -1;
            }
        }

        private void RegisterDriveNotifications()
        {
            driveInsertedWatcher = new ManagementEventWatcher("SELECT * FROM Win32_VolumeChangeEvent WHERE EventType = 2");
            driveInsertedWatcher.EventArrived += (s, args) => UpdateDriveList();
            driveInsertedWatcher.Start();

            driveRemovedWatcher = new ManagementEventWatcher("SELECT * FROM Win32_VolumeChangeEvent WHERE EventType = 3");
            driveRemovedWatcher.EventArrived += (s, args) => UpdateDriveList();
            driveRemovedWatcher.Start();
        }

        private void UnregisterDriveNotifications()
        {
            if (driveInsertedWatcher != null)
            {
                driveInsertedWatcher.Stop();
                driveInsertedWatcher.Dispose();
            }

            if (driveRemovedWatcher != null)
            {
                driveRemovedWatcher.Stop();
                driveRemovedWatcher.Dispose();
            }
        }

        private void scan_Click(object sender, EventArgs e)
        {
            UpdateDriveList();
        }

        private void clear_Click(object sender, EventArgs e)
        {
            foreach (ComboBox comboBox in comboBoxes)
            {
                comboBox.Items.Clear();
                comboBox.Text = "";
            }
        }

        private void run_h2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < comboBoxes.Length; i++)
            {
                string rootDirectory = AppDomain.CurrentDomain.BaseDirectory;
                string selectedDrivePath = comboBoxes[i].SelectedItem?.ToString();
                if (!string.IsNullOrEmpty(selectedDrivePath))
                {
                    switch (i)
                    {
                        case 0:
                            string exePath1 = Path.Combine(rootDirectory, "h2testw.exe");
                            Process process1 = Process.Start(exePath1);
                            if (process1 != null)
                            {
                                process1.WaitForInputIdle();
                                IntPtr handle1 = process1.MainWindowHandle;
                                if (handle1 != IntPtr.Zero)
                                {
                                    SetWindowPos(handle1, IntPtr.Zero, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
                                    SelectEnglishOption(handle1, selectedDrivePath);
                                }
                            }
                            break;

                        case 1:
                            string exePath2 = Path.Combine(rootDirectory, "h2testw.exe");
                            Process process2 = Process.Start(exePath2);
                            if (process2 != null)
                            {
                                process2.WaitForInputIdle();
                                IntPtr handle2 = process2.MainWindowHandle;
                                if (handle2 != IntPtr.Zero)
                                {
                                    SetWindowPos(handle2, IntPtr.Zero, 0, 255, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
                                    SelectEnglishOption(handle2, selectedDrivePath);
                                }
                            }
                            break;

                        case 2:
                            string exePath3 = Path.Combine(rootDirectory, "h2testw.exe");
                            Process process3 = Process.Start(exePath3);
                            if (process3 != null)
                            {
                                process3.WaitForInputIdle();
                                IntPtr handle3 = process3.MainWindowHandle;
                                if (handle3 != IntPtr.Zero)
                                {
                                    SetWindowPos(handle3, IntPtr.Zero, 0, 514, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
                                    SelectEnglishOption(handle3, selectedDrivePath);
                                }
                            }
                            break;

                        case 3:
                            string exePath4 = Path.Combine(rootDirectory, "h2testw.exe");
                            Process process4 = Process.Start(exePath4);
                            if (process4 != null)
                            {
                                process4.WaitForInputIdle();
                                IntPtr handle4 = process4.MainWindowHandle;
                                if (handle4 != IntPtr.Zero)
                                {
                                    SetWindowPos(handle4, IntPtr.Zero, 0, 771, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
                                    SelectEnglishOption(handle4, selectedDrivePath);
                                }
                            }
                            break;

                        case 4:
                            string exePath5 = Path.Combine(rootDirectory, "h2testw.exe");
                            Process process5 = Process.Start(exePath5);
                            if (process5 != null)
                            {
                                process5.WaitForInputIdle();
                                IntPtr handle5 = process5.MainWindowHandle;
                                if (handle5 != IntPtr.Zero)
                                {
                                    SetWindowPos(handle5, IntPtr.Zero, 418, 0, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
                                    SelectEnglishOption(handle5, selectedDrivePath);
                                }
                            }
                            break;

                        case 5:
                            string exePath6 = Path.Combine(rootDirectory, "h2testw.exe");
                            Process process6 = Process.Start(exePath6);
                            if (process6 != null)
                            {
                                process6.WaitForInputIdle();
                                IntPtr handle6 = process6.MainWindowHandle;
                                if (handle6 != IntPtr.Zero)
                                {
                                    SetWindowPos(handle6, IntPtr.Zero, 418, 255, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
                                    SelectEnglishOption(handle6, selectedDrivePath);
                                }
                            }
                            break;

                        case 6:
                            string exePath7 = Path.Combine(rootDirectory, "h2testw.exe");
                            Process process7 = Process.Start(exePath7);
                            if (process7 != null)
                            {
                                process7.WaitForInputIdle();
                                IntPtr handle7 = process7.MainWindowHandle;
                                if (handle7 != IntPtr.Zero)
                                {
                                    SetWindowPos(handle7, IntPtr.Zero, 418, 514, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
                                    SelectEnglishOption(handle7, selectedDrivePath);
                                }
                            }
                            break;

                        case 7:
                            string exePath8 = Path.Combine(rootDirectory, "h2testw.exe");
                            Process process8 = Process.Start(exePath8);
                            if (process8 != null)
                            {
                                process8.WaitForInputIdle();
                                IntPtr handle8 = process8.MainWindowHandle;
                                if (handle8 != IntPtr.Zero)
                                {
                                    SetWindowPos(handle8, IntPtr.Zero, 418, 771, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
                                    SelectEnglishOption(handle8, selectedDrivePath);
                                }
                            }
                            break;

                        case 8:
                            string exePath9 = Path.Combine(rootDirectory, "h2testw.exe");
                            Process process9 = Process.Start(exePath9);
                            if (process9 != null)
                            {
                                process9.WaitForInputIdle();
                                IntPtr handle9 = process9.MainWindowHandle;
                                if (handle9 != IntPtr.Zero)
                                {
                                    SetWindowPos(handle9, IntPtr.Zero, 836, 0, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
                                    SelectEnglishOption(handle9, selectedDrivePath);
                                }
                            }
                            break;

                        case 9:
                            string exePath10 = Path.Combine(rootDirectory, "h2testw.exe");
                            Process process10 = Process.Start(exePath10);
                            if (process10 != null)
                            {
                                process10.WaitForInputIdle();
                                IntPtr handle10 = process10.MainWindowHandle;
                                if (handle10 != IntPtr.Zero)
                                {
                                    SetWindowPos(handle10, IntPtr.Zero, 836, 255, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
                                    SelectEnglishOption(handle10, selectedDrivePath);
                                }
                            }
                            break;

                        case 10:
                            string exePath11 = Path.Combine(rootDirectory, "h2testw.exe");
                            Process process11 = Process.Start(exePath11);
                            if (process11 != null)
                            {
                                process11.WaitForInputIdle();
                                IntPtr handle11 = process11.MainWindowHandle;
                                if (handle11 != IntPtr.Zero)
                                {
                                    SetWindowPos(handle11, IntPtr.Zero, 836, 514, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
                                    SelectEnglishOption(handle11, selectedDrivePath);
                                }
                            }
                            break;

                        case 11:
                            string exePath12 = Path.Combine(rootDirectory, "h2testw.exe");
                            Process process12 = Process.Start(exePath12);
                            if (process12 != null)
                            {
                                process12.WaitForInputIdle();
                                IntPtr handle12 = process12.MainWindowHandle;
                                if (handle12 != IntPtr.Zero)
                                {
                                    SetWindowPos(handle12, IntPtr.Zero, 836, 771, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
                                    SelectEnglishOption(handle12, selectedDrivePath);
                                }
                            }
                            break;

                        case 12:
                            string exePath13 = Path.Combine(rootDirectory, "h2testw.exe");
                            Process process13 = Process.Start(exePath13);
                            if (process13 != null)
                            {
                                process13.WaitForInputIdle();
                                IntPtr handle13 = process13.MainWindowHandle;
                                if (handle13 != IntPtr.Zero)
                                {
                                    SetWindowPos(handle13, IntPtr.Zero, 1254, 0, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
                                    SelectEnglishOption(handle13, selectedDrivePath);
                                }
                            }
                            break;

                        case 13:
                            string exePath14 = Path.Combine(rootDirectory, "h2testw.exe");
                            Process process14 = Process.Start(exePath14);
                            if (process14 != null)
                            {
                                process14.WaitForInputIdle();
                                IntPtr handle14 = process14.MainWindowHandle;
                                if (handle14 != IntPtr.Zero)
                                {
                                    SetWindowPos(handle14, IntPtr.Zero, 1254, 255, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
                                    SelectEnglishOption(handle14, selectedDrivePath);
                                }
                            }
                            break;

                        case 14:
                            string exePath15 = Path.Combine(rootDirectory, "h2testw.exe");
                            Process process15 = Process.Start(exePath15);
                            if (process15 != null)
                            {
                                process15.WaitForInputIdle();
                                IntPtr handle15 = process15.MainWindowHandle;
                                if (handle15 != IntPtr.Zero)
                                {
                                    SetWindowPos(handle15, IntPtr.Zero, 1254, 514, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
                                    SelectEnglishOption(handle15, selectedDrivePath);
                                }
                            }
                            break;

                        case 15:
                            string exePath16 = Path.Combine(rootDirectory, "h2testw.exe");
                            Process process16 = Process.Start(exePath16);
                            if (process16 != null)
                            {
                                process16.WaitForInputIdle();
                                IntPtr handle16 = process16.MainWindowHandle;
                                if (handle16 != IntPtr.Zero)
                                {
                                    SetWindowPos(handle16, IntPtr.Zero, 1254, 771, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
                                    SelectEnglishOption(handle16, selectedDrivePath);
                                }
                            }
                            break;
                    }
                }
            }
        }

        private void SelectEnglishOption(IntPtr mainWindowHandle, string selectedDrivePath)
        {
            AutoItX.WinActivate(mainWindowHandle);
            AutoItX.ControlClick("H2testw", "", "[CLASS:Button; INSTANCE:11]");
            AutoItX.Send("{TAB}");
            AutoItX.Sleep(250);
            AutoItX.Send("{ENTER}");
            AutoItX.Sleep(250);
            AutoItX.ControlSetText("[CLASS:#32770]", "", "[CLASS:Edit; INSTANCE:1]", selectedDrivePath);
            AutoItX.ControlClick("[CLASS:#32770]", "", "[CLASS:Button; INSTANCE:2]");

            int H2count = 0;
            bool H2success = false;

            int h2count = 0;
            bool h2success = false;

            while (H2count < 2 && !H2success)
            {
                AutoItX.Sleep(500);
                if (AutoItX.ControlCommand("H2testw", "", "[CLASS:Button; INSTANCE:5]", "IsEnabled", "") == "0")
                {
                    H2count++;
                }
                else
                {
                    H2success = true;
                }
            }

            if (H2success)
            {
                AutoItX.Sleep(250);
                AutoItX.ControlClick("[CLASS:#32770]", "", "[CLASS:Button; INSTANCE:5]");
            }

            while (h2count < 2 && !h2success)
            {
                AutoItX.Sleep(500);
                if (AutoItX.ControlCommand("h2testw", "", "[CLASS:Button; INSTANCE:1]", "IsEnabled", "") == "0")
                {
                    h2count++;
                }
                else
                {
                    h2success = true;
                }
            }

            if (h2success)
            {
                AutoItX.Sleep(250);
                AutoItX.ControlClick("[CLASS:#32770]", "", "[CLASS:Button; INSTANCE:1]");
            }
        }
    }
}