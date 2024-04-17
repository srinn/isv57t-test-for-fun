using Modbus.Device;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Management;
using System.Data;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Security.Cryptography.X509Certificates;
using System.Diagnostics;
using System.Net;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public SerialPort serialPort = null;
        public ModbusSerialMaster modbusSerialMaster = null;
        public string oldPortName = "";
        public Random random = new Random();
        public const byte slaveAddress = 63;
        public const int maxPosition = 1000000000;
        public const int maxRpm = 3000;
        public bool runQueue = false;
        public bool changeRpm = false;
        public int targetRpm = 0;
        public const int maxConnectSerialRetry = 50;

        public class ComboboxItem
        {
            public string Text { get; set; }
            public string Value { get; set; }
            public override string ToString()
            {
                return Text;
            }
        }

        public double expRand()
        {
            return (1.0 - 0.0) * (random.NextDouble() - 0.0) / (1.5 - 0.0) + 0.0;
        }

        public void SetPortNames()
        {
            using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE Caption like '%(COM%'"))
            {
                var portnames = SerialPort.GetPortNames();
                var ports = searcher.Get().Cast<ManagementBaseObject>().ToList().Select(p => p["Caption"].ToString());

                comboBox1.Items.Clear();
                foreach (var port in portnames)
                {
                    ComboboxItem item = new ComboboxItem();
                    item.Value = port;
                    item.Text = port + " - " + ports.FirstOrDefault(s => s.Contains(port));
                    comboBox1.Items.Add(item);
                }
            }
        }
        private void modbusAndDelay(ushort address, ushort value)
        {
            if (modbusSerialMaster != null)
            {
                modbusSerialMaster.WriteSingleRegisterAsync(slaveAddress, address, value).Wait();
            }
        }

        private void freezeControls(bool isFreezing)
        {
            comboBox1.Enabled = !isFreezing;
            trackBar1.Enabled = !isFreezing;
            checkBox1.Enabled = !isFreezing;
        }

        public Form1()
        {
            InitializeComponent();
            SetPortNames();
            label1.Text = trackBar1.Value + " RPM";
            trackBar1.Maximum = maxRpm;
            trackBar1.Minimum = -maxRpm;
        }

        private void closeSerialPortAndModbus()
        {
            if (modbusSerialMaster != null)
            {
                stopCommandQueue();
                textBox1.Paste("기존 Modbus 연결 해제 중..\r\n");
                modbusSerialMaster.WriteSingleRegisterAsync(slaveAddress, 407, 0).Wait();
                modbusAndDelay(407, 0);
                modbusSerialMaster.Dispose();
                Task.Delay(50).Wait();
                modbusSerialMaster = null;
                textBox1.Paste("Modbus 연결 해제 완료\r\n");
            }
            if (serialPort != null)
            {
                textBox1.Paste(oldPortName + " 연결 해제 중..\r\n");
                serialPort.Close();
                Task.Delay(50).Wait();
                serialPort = null;
                textBox1.Paste(oldPortName + " 연결 해제 완료\r\n");
            }
        }

        private async void commandQueue()
        {
            if (runQueue) return;

            Stopwatch stopwatch = new Stopwatch();
            runQueue = true;
            int maxDiff = -32767;
            int minDiff = 32767;
            int diff;
            try
            {
                while (runQueue)
                {
                    if (!runQueue) break;
                    stopwatch.Start();
                    if (serialPort != null && modbusSerialMaster != null)
                    {
                        //await Task.Delay(3000);
                        ushort[] readPositions = await modbusSerialMaster.ReadHoldingRegistersAsync(slaveAddress, 499, 4);

                        //if (changeRpm)
                        //{
                        //    changeRpm = false;
                        //    textBox1.Paste(trackBar1.Value + " RPM 으로 변경 중..\r\n");
                        //    await modbusSerialMaster.WriteSingleRegisterAsync(slaveAddress, 219, (ushort)trackBar1.Value);
                        //    textBox1.Paste(trackBar1.Value + " RPM 으로 변경 완료\r\n");
                        //}
                        //await modbusSerialMaster.WriteMultipleRegistersAsync(slaveAddress, 406, new ushort[] { 16, 0 });
                        //await modbusSerialMaster.WriteMultipleRegistersAsync(slaveAddress, 406, new ushort[]{16, 0});
                        //await modbusSerialMaster.WriteSingleRegisterAsync(slaveAddress, 406, 16);
                        //await modbusSerialMaster.WriteSingleRegisterAsync(slaveAddress, 406, 17);
                        //await modbusSerialMaster.WriteSingleRegisterAsync(slaveAddress, 407, 4);
                        //ushort readRepeatTimes = (await modbusSerialMaster.ReadHoldingRegistersAsync(slaveAddress, 237, 1))[0];
                        stopwatch.Stop();
                        TimeSpan ts = stopwatch.Elapsed;
                        stopwatch.Restart();

                        diff = readPositions[0] - readPositions[1];
                        if (diff > short.MaxValue) diff -= ushort.MaxValue;
                        if (diff < short.MinValue) diff += ushort.MaxValue;
                        if (diff > maxDiff) maxDiff = diff;
                        if (diff < minDiff) minDiff = diff;

                        textBox1.Paste(String.Format("{0:00}.{1:000} {2} ~ {3} Diff: {4} Analog1: {5} Analog3: {6} VelocityFeedback: {7}\r\n", ts.Seconds, ts.Milliseconds, minDiff, maxDiff, diff, readPositions[0], readPositions[1], (short)readPositions[3]));
                    }
                }
            }
            catch
            {
                runQueue = false;
            }
        }

        private void stopCommandQueue()
        {
            runQueue = false;
        }

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                label1.Text = string.Format("{0:F2} %", (double)trackBar1.Value / (double)trackBar1.Maximum * 100.0);
                if (serialPort != null && modbusSerialMaster != null)
                {
                    textBox1.Paste(string.Format("{0:F2} % 로 변경 중..\r\n", (double)trackBar1.Value / (double)trackBar1.Maximum * 100.0));
                    modbusAndDelay(219, (ushort)trackBar1.Value);
                    textBox1.Paste(trackBar1.Value + " RPM 으로 변경 완료\r\n");
                }
            }
            else
            {
                label1.Text = trackBar1.Value + " RPM";
                if (serialPort != null && modbusSerialMaster != null)
                {
                    textBox1.Paste(trackBar1.Value + " RPM 으로 변경 중..\r\n");
                    modbusAndDelay(219, (ushort)trackBar1.Value);
                    textBox1.Paste(trackBar1.Value + " RPM 으로 변경 완료\r\n");
                }
                //if (serialPort != null && modbusSerialMaster != null)
                //{
                //    changeRpm = true;
                //}
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!(comboBox1.SelectedItem as ComboboxItem).Value.Equals(oldPortName))
            {
                freezeControls(true);
                closeSerialPortAndModbus();

                oldPortName = (comboBox1.SelectedItem as ComboboxItem).Value;

                serialPort = new SerialPort(oldPortName, 38400);
                serialPort.Parity = Parity.None;
                serialPort.DataBits = 8;
                serialPort.StopBits = StopBits.One;
                serialPort.ReadTimeout = 100;
                serialPort.WriteTimeout = 100;

                bool isSerialOpened = false;
                
                textBox1.Paste(oldPortName + " 연결 시도 중..\r\n");
                for (int i = 0; i < maxConnectSerialRetry && !isSerialOpened; i++)
                {
                    try
                    {
                        serialPort.Open();
                        isSerialOpened = true;
                    }
                    catch
                    {
                        textBox1.Paste(oldPortName + " 연결 실패, 재시도 중..(" + (i + 1) + "/" + maxConnectSerialRetry + ")\r\n");
                        serialPort.Dispose();
                        Task.Delay(200).Wait();
                    }

                }
                if (isSerialOpened)
                {
                    textBox1.Paste(oldPortName + " 연결 성공\r\n");
                    try
                    {
                        textBox1.Paste("Modbus 전송 중..\r\n");
                        modbusSerialMaster = ModbusSerialMaster.CreateRtu(serialPort);
                        Task.Delay(100).Wait();
                        // velocity mode setup start
                        modbusAndDelay(349, 0); 
                        modbusAndDelay(1, 1); // Velocity mode
                        modbusAndDelay(410, 21845);

                        // velocity control setup start
                        modbusAndDelay(401, 384); // Position? Analog1
                        modbusAndDelay(402, 386); // Position? Analog3
                        modbusAndDelay(403, 64); // VelocityGiven(rpm) (garbage value, when you set the velocity 0 and 1 always printing same value(printing 0 in both cases), and it always 1 lesser than original value in positive value. ex: exact value was 10, it prints 9. negative value seems ok. If you want exact given rpm, use your wrote rpm info)
                        modbusAndDelay(404, 65); // VelocityFeedback(rpm) (i think this is have not problem)
                        modbusAndDelay(405, 1);
                        modbusAndDelay(219, 10);
                        modbusAndDelay(26, 220);
                        modbusAndDelay(27, 250);
                        modbusAndDelay(37, 0);
                        modbusAndDelay(29, 103);
                        modbusAndDelay(4, 600);
                        modbusAndDelay(107, 100);
                        modbusAndDelay(108, 100);
                        modbusAndDelay(109, 0);
                        modbusAndDelay(413, 100);
                        modbusAndDelay(406, 16); // I don't know what it is, I just sniff the modbus packets of velocity control's start button
                        modbusAndDelay(406, 17); // run test program always set the address 406 to 16 right before 17, I don't know why
                        modbusAndDelay(407, 4); // it triggers start the run test, if you want to end the test mode, set it to 0
                        //modbusSerialMaster.ReadHoldingRegistersAsync(slaveAddress, 496, 1).Wait();

                        //// position control setup start
                        //modbusAndDelay(401, 1); // Position Given
                        //modbusAndDelay(402, 2); // Position Feedback
                        //modbusAndDelay(403, 0);
                        //modbusAndDelay(404, 0);
                        //modbusAndDelay(405, 1);
                        //modbusAndDelay(25, 390);
                        //modbusAndDelay(30, 460);
                        //modbusAndDelay(26, 220);
                        //modbusAndDelay(31, 220);
                        //modbusAndDelay(27, 250);
                        //modbusAndDelay(32, 10000);
                        //modbusAndDelay(35, 300);
                        //modbusAndDelay(37, 0);
                        //modbusAndDelay(29, 103);
                        //modbusAndDelay(34, 103);
                        //modbusAndDelay(40, 0);
                        //modbusAndDelay(4, 600);
                        //modbusAndDelay(413, 100);
                        //modbusAndDelay(219, 0); // Velocity(rpm)
                        //modbusAndDelay(107, 90); // AccelrationAndDecelerationTime(ms/Krpm)
                        //modbusAndDelay(235, 10); // Distance(0.1rev)
                        //modbusAndDelay(236, 0); // IntervalTime(ms)
                        //modbusAndDelay(237, 32767); // RepeatTimes
                        //modbusAndDelay(231, 1); // RunningMode(0: PositiveAndNegative, 1:Unidirectional)
                        //modbusAndDelay(406, 16);
                        //modbusAndDelay(406, 17);
                        //modbusAndDelay(407, 4);

                        Task.Delay(50).Wait();
                        textBox1.Paste("Modbus 전송 성공\r\n");
                        commandQueue();
                    }
                    catch
                    {
                        stopCommandQueue();
                        textBox1.Paste("Modbus 전송 실패, " + oldPortName + " 닫는 중..\r\n");
                        modbusSerialMaster?.Dispose();
                        Task.Delay(100).Wait();
                        serialPort?.Close();
                        textBox1.Paste(oldPortName + " 닫기 성공\r\n");
                        Task.Delay(50).Wait();
                        modbusSerialMaster = null;
                        serialPort = null;
                        comboBox1.Items.Clear();
                        oldPortName = "";
                    }
                }
                else
                {
                    textBox1.Paste(oldPortName + " 연결 실패\r\n");
                    serialPort = null;
                    Task.Delay(50).Wait();
                    modbusSerialMaster = null;
                    serialPort = null;
                    comboBox1.Items.Clear();
                    oldPortName = "";
                }

                freezeControls(false);

                trackBar1.Value = 0;
                trackBar1_Scroll(sender, e);
            }
        }

        private void comboBox1_Click(object sender, EventArgs e)
        {
            SetPortNames();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            closeSerialPortAndModbus();
        }

        private async void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                while (serialPort != null && modbusSerialMaster != null && serialPort.IsOpen && checkBox1.Checked)
                {
                    trackBar1.Value = trackBar1.Minimum + (int)Math.Round(expRand() * (trackBar1.Maximum - trackBar1.Minimum));
                    trackBar1_Scroll(sender, e);
                    await Task.Delay(random.Next(30, 3000));
                }
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                checkBox1.Checked = false;
                trackBar1.Maximum = maxPosition;
                trackBar1.Minimum = 0;
                trackBar1.Value = trackBar1.Maximum / 2;
                label1.Text = string.Format("{0:F2} %", (double)trackBar1.Value / (double)trackBar1.Maximum * 100.0);
            }
            else
            {
                checkBox1.Checked = false;
                trackBar1.Maximum = maxRpm;
                trackBar1.Minimum = -maxRpm;
                trackBar1.Value = 0;
                label1.Text = trackBar1.Value + " RPM";
            }
        }
    }
}
