using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Net.NetworkInformation;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolBar;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TrackBar;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TreeView;
using Excel = Microsoft.Office.Interop.Excel;


namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Boolean run = true;
        public Boolean run1 = true;
        public Boolean run2 = true;
        public Boolean run3 = true;

        public Form1()
        {
            InitializeComponent();

            progressBar1.Value = 0;
            start.Enabled = false;
            stop.Enabled = false;
            connect.Enabled = true;
            refresh.Enabled = false;
            checkBox1.Enabled = true;
            filesave.Enabled = false;
            filesavebutton.Enabled = false;
            maskedTextBox1.Enabled = false;
            button1.Enabled = false;
            maskedTextBox2.Enabled = false;
            button2.Enabled = false;
            maskedTextBox11.Enabled = false;
            maskedTextBox3.Enabled = false;
            comboBox1.Enabled = false;
            maskedTextBox12.Enabled = false;
            maskedTextBox4.Enabled = false;
            comboBox2.Enabled = false;
            maskedTextBox13.Enabled = false;
            maskedTextBox5.Enabled = false;
            comboBox3.Enabled = false;
            maskedTextBox14.Enabled = false;
            maskedTextBox6.Enabled = false;
            comboBox4.Enabled = false;
            maskedTextBox15.Enabled = false;
            maskedTextBox10.Enabled = false;
            comboBox8.Enabled = false;
            maskedTextBox16.Enabled = false;
            maskedTextBox9.Enabled = false;
            comboBox7.Enabled = false;
            maskedTextBox17.Enabled = false;
            maskedTextBox8.Enabled = false;
            comboBox6.Enabled = false;
            maskedTextBox18.Enabled = false;
            maskedTextBox7.Enabled = false;
            comboBox5.Enabled = false;
            maskedTextBox25.Enabled = false;
            maskedTextBox24.Enabled = false;
            maskedTextBox23.Enabled = false;
            maskedTextBox22.Enabled = false;
            maskedTextBox21.Enabled = false;
            maskedTextBox20.Enabled = false;
            maskedTextBox19.Enabled = false;
            maskedTextBox1.Enabled = false;
            comboBox16.Enabled = false;
            comboBox15.Enabled = false;
            comboBox14.Enabled = false;
            comboBox13.Enabled = false;
            comboBox12.Enabled = false;
            comboBox11.Enabled = false;
            comboBox10.Enabled = false;
            comboBox9.Enabled = false;
            maskedTextBox32.Enabled = false;
            maskedTextBox31.Enabled = false;
            maskedTextBox30.Enabled = false;
            maskedTextBox29.Enabled = false;
            maskedTextBox26.Enabled = false;
            maskedTextBox28.Enabled = false;
            maskedTextBox27.Enabled = false;
            maskedTextBox33.Enabled = false;
            label65.Visible = false;
            maskedTextBox34.Enabled = false;
            maskedTextBox34.Visible = false;
            label66.Visible = false;
            maskedTextBox35.Enabled = false;
            maskedTextBox35.Visible = false;
            label67.Visible = false;
            maskedTextBox36.Enabled = false;
            maskedTextBox36.Visible = false;
            label68.Visible = false;
            maskedTextBox37.Enabled = false;
            maskedTextBox37.Visible = false;
            label69.Visible = false;
            maskedTextBox38.Enabled = false;
            maskedTextBox38.Visible = false;
            label70.Visible = false;
            maskedTextBox39.Enabled = false;
            maskedTextBox39.Visible = false;
            label71.Visible = false;
            maskedTextBox40.Enabled = false;
            maskedTextBox40.Visible = false;
            label72.Visible = false;
            maskedTextBox41.Enabled = false;
            maskedTextBox41.Visible = false;
            label73.Visible = false;
            maskedTextBox42.Enabled = false;
            maskedTextBox42.Visible = false;
            label74.Visible = false;
            maskedTextBox43.Enabled = false;
            maskedTextBox43.Visible = false;
            label75.Visible = false;
            maskedTextBox44.Enabled = false;
            maskedTextBox44.Visible = false;
            label83.Visible = false;
            maskedTextBox52.Enabled = false;
            maskedTextBox52.Visible = false;
            label82.Visible = false;
            maskedTextBox51.Enabled = false;
            maskedTextBox51.Visible = false;
            label81.Visible = false;
            maskedTextBox50.Enabled = false;
            maskedTextBox50.Visible = false;
            label80.Visible = false;
            maskedTextBox49.Enabled = false;
            maskedTextBox49.Visible = false;
            label79.Visible = false;
            maskedTextBox48.Enabled = false;
            maskedTextBox48.Visible = false;
            label78.Visible = false;
            maskedTextBox47.Enabled = false;
            maskedTextBox47.Visible = false;
            label77.Visible = false;
            maskedTextBox46.Enabled = false;
            maskedTextBox46.Visible = false;
            label76.Visible = false;
            maskedTextBox45.Enabled = false;
            maskedTextBox45.Visible = false;
            label_status.Text = "DISCONNECTED";
            label_status.ForeColor = Color.Red;
            status_acquisitions.Text = "Stopped";
            status_acquisitions.ForeColor = Color.Red;
            textBox1.Enabled = false;
            string[] ports = SerialPort.GetPortNames();
            portselect.Items.AddRange(ports);
        }

        private void exit_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                try
                {
                    serialPort1.Close();
                }
                catch (Exception error)
                {
                    MessageBox.Show(error.Message);
                }
            }
            Close();
        }

        private void filesavebutton_Click(object sender, EventArgs e)
        {
            var fileContent = string.Empty;
            var filePath = string.Empty;

            using (FolderBrowserDialog openFolderDialog = new FolderBrowserDialog())
            {

                if (openFolderDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePath = openFolderDialog.SelectedPath;

                }
            }

            filesave.Text = filePath;
        }

        private void filesave_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void portselect_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void connect_Click(object sender, EventArgs e)
        {
            try
            {
                serialPort1.PortName = portselect.Text;
                serialPort1.BaudRate = 115200;
                serialPort1.Open();

                progressBar1.Value = 100;
                start.Enabled = true;
                stop.Enabled = false;
                connect.Enabled = false;
                refresh.Enabled = true;
                button1.Enabled = true;
                label_status.Text = "CONNECTED";
                label_status.ForeColor = Color.Green;
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
        }

        private void start_Click(object sender, EventArgs e)
        {
            start.Enabled = false;
            stop.Enabled = true;
            checkBox1.Enabled = false;
            filesave.Enabled = false;
            filesavebutton.Enabled = false;
            textBox1.Enabled = false;
            maskedTextBox2.Enabled = false;
            button2.Enabled = false;
            checkBox4.Enabled = false;
            checkBox5.Enabled = false;
            checkBox6.Enabled = false;
            checkBox7.Enabled = false;
            checkBox11.Enabled = false;
            checkBox10.Enabled = false;
            checkBox9.Enabled = false;
            checkBox8.Enabled = false;
            maskedTextBox11.Enabled = false;
            maskedTextBox3.Enabled = false;
            comboBox1.Enabled = false;
            maskedTextBox12.Enabled = false;
            maskedTextBox4.Enabled = false;
            comboBox2.Enabled = false;
            maskedTextBox13.Enabled = false;
            maskedTextBox5.Enabled = false;
            comboBox3.Enabled = false;
            maskedTextBox14.Enabled = false;
            maskedTextBox6.Enabled = false;
            comboBox4.Enabled = false;
            maskedTextBox15.Enabled = false;
            maskedTextBox10.Enabled = false;
            comboBox8.Enabled = false;
            maskedTextBox16.Enabled = false;
            maskedTextBox9.Enabled = false;
            comboBox7.Enabled = false;
            maskedTextBox17.Enabled = false;
            maskedTextBox8.Enabled = false;
            comboBox6.Enabled = false;
            maskedTextBox18.Enabled = false;
            maskedTextBox7.Enabled = false;
            comboBox5.Enabled = false;
            checkBox15.Enabled = false;
            checkBox21.Enabled = false;
            checkBox20.Enabled = false;
            checkBox19.Enabled = false;
            checkBox18.Enabled = false;
            checkBox17.Enabled = false;
            checkBox16.Enabled = false;
            checkBox13.Enabled = false;
            checkBox3.Enabled = false;
            maskedTextBox25.Enabled = false;
            maskedTextBox24.Enabled = false;
            maskedTextBox23.Enabled = false;
            maskedTextBox22.Enabled = false;
            maskedTextBox21.Enabled = false;
            maskedTextBox20.Enabled = false;
            maskedTextBox19.Enabled = false;
            maskedTextBox1.Enabled = false;
            comboBox16.Enabled = false;
            comboBox15.Enabled = false;
            comboBox14.Enabled = false;
            comboBox13.Enabled = false;
            comboBox12.Enabled = false;
            comboBox11.Enabled = false;
            comboBox10.Enabled = false;
            comboBox9.Enabled = false;
            maskedTextBox32.Enabled = false;
            maskedTextBox31.Enabled = false;
            maskedTextBox30.Enabled = false;
            maskedTextBox29.Enabled = false;
            maskedTextBox26.Enabled = false;
            maskedTextBox28.Enabled = false;
            maskedTextBox27.Enabled = false;
            maskedTextBox33.Enabled = false;
            checkBox14.Enabled = false;
            checkBox12.Enabled = false;
            checkBox22.Enabled = false;
            button1.Enabled = false;
            status_acquisitions.Text = "In progress";
            status_acquisitions.ForeColor = Color.Green;

            serialPort1.Close();
            serialPort1.Open();


            if (this.checkBox4.Checked)
            {
                graph.Series.Add(maskedTextBox11.Text + " (" + comboBox1.Text + ")");
                graph.Series[maskedTextBox11.Text + " (" + comboBox1.Text + ")"].ChartType = SeriesChartType.Line;
                graph.Series[maskedTextBox11.Text + " (" + comboBox1.Text + ")"].XValueType = ChartValueType.Time;
            }
            if (this.checkBox5.Checked)
            {
                graph.Series.Add(maskedTextBox12.Text + " (" + comboBox2.Text + ")");
                graph.Series[maskedTextBox12.Text + " (" + comboBox2.Text + ")"].ChartType = SeriesChartType.Line;
                graph.Series[maskedTextBox12.Text + " (" + comboBox2.Text + ")"].XValueType = ChartValueType.Time;
            }
            if (this.checkBox6.Checked)
            {
                graph.Series.Add(maskedTextBox13.Text + " (" + comboBox3.Text + ")");
                graph.Series[maskedTextBox13.Text + " (" + comboBox3.Text + ")"].ChartType = SeriesChartType.Line;
                graph.Series[maskedTextBox13.Text + " (" + comboBox3.Text + ")"].XValueType = ChartValueType.Time;
            }
            if (this.checkBox7.Checked)
            {
                graph.Series.Add(maskedTextBox14.Text + " (" + comboBox4.Text + ")");
                graph.Series[maskedTextBox14.Text + " (" + comboBox4.Text + ")"].ChartType = SeriesChartType.Line;
                graph.Series[maskedTextBox14.Text + " (" + comboBox4.Text + ")"].XValueType = ChartValueType.Time;
            }
            if (this.checkBox8.Checked)
            {
                graph.Series.Add(maskedTextBox18.Text + " (" + comboBox5.Text + ")");
                graph.Series[maskedTextBox18.Text + " (" + comboBox5.Text + ")"].ChartType = SeriesChartType.Line;
                graph.Series[maskedTextBox18.Text + " (" + comboBox5.Text + ")"].XValueType = ChartValueType.Time;
            }
            if (this.checkBox9.Checked)
            {
                graph.Series.Add(maskedTextBox17.Text + " (" + comboBox6.Text + ")");
                graph.Series[maskedTextBox17.Text + " (" + comboBox6.Text + ")"].ChartType = SeriesChartType.Line;
                graph.Series[maskedTextBox17.Text + " (" + comboBox6.Text + ")"].XValueType = ChartValueType.Time;
            }
            if (this.checkBox10.Checked)
            {
                graph.Series.Add(maskedTextBox16.Text + " (" + comboBox7.Text + ")");
                graph.Series[maskedTextBox16.Text + " (" + comboBox7.Text + ")"].ChartType = SeriesChartType.Line;
                graph.Series[maskedTextBox16.Text + " (" + comboBox7.Text + ")"].XValueType = ChartValueType.Time;
            }
            if (this.checkBox11.Checked)
            {
                graph.Series.Add(maskedTextBox15.Text + " (" + comboBox8.Text + ")");
                graph.Series[maskedTextBox15.Text + " (" + comboBox8.Text + ")"].ChartType = SeriesChartType.Line;
                graph.Series[maskedTextBox15.Text + " (" + comboBox8.Text + ")"].XValueType = ChartValueType.Time;
            }
            if (this.checkBox21.Checked)
            {
                graph.Series.Add(maskedTextBox25.Text + " (" + comboBox16.Text + ")");
                graph.Series[maskedTextBox25.Text + " (" + comboBox16.Text + ")"].ChartType = SeriesChartType.Line;
                graph.Series[maskedTextBox25.Text + " (" + comboBox16.Text + ")"].XValueType = ChartValueType.Time;
            }
            if (this.checkBox20.Checked)
            {
                graph.Series.Add(maskedTextBox24.Text + " (" + comboBox15.Text + ")");
                graph.Series[maskedTextBox24.Text + " (" + comboBox15.Text + ")"].ChartType = SeriesChartType.Line;
                graph.Series[maskedTextBox24.Text + " (" + comboBox15.Text + ")"].XValueType = ChartValueType.Time;
            }
            if (this.checkBox19.Checked)
            {
                graph.Series.Add(maskedTextBox23.Text + " (" + comboBox14.Text + ")");
                graph.Series[maskedTextBox23.Text + " (" + comboBox14.Text + ")"].ChartType = SeriesChartType.Line;
                graph.Series[maskedTextBox23.Text + " (" + comboBox14.Text + ")"].XValueType = ChartValueType.Time;
            }
            if (this.checkBox18.Checked)
            {
                graph.Series.Add(maskedTextBox22.Text + " (" + comboBox13.Text + ")");
                graph.Series[maskedTextBox22.Text + " (" + comboBox13.Text + ")"].ChartType = SeriesChartType.Line;
                graph.Series[maskedTextBox22.Text + " (" + comboBox13.Text + ")"].XValueType = ChartValueType.Time;
            }
            if (this.checkBox17.Checked)
            {
                graph.Series.Add(maskedTextBox21.Text + " (" + comboBox12.Text + ")");
                graph.Series[maskedTextBox21.Text + " (" + comboBox12.Text + ")"].ChartType = SeriesChartType.Line;
                graph.Series[maskedTextBox21.Text + " (" + comboBox12.Text + ")"].XValueType = ChartValueType.Time;
            }
            if (this.checkBox16.Checked)
            {
                graph.Series.Add(maskedTextBox20.Text + " (" + comboBox11.Text + ")");
                graph.Series[maskedTextBox20.Text + " (" + comboBox11.Text + ")"].ChartType = SeriesChartType.Line;
                graph.Series[maskedTextBox20.Text + " (" + comboBox11.Text + ")"].XValueType = ChartValueType.Time;
            }
            if (this.checkBox13.Checked)
            {
                graph.Series.Add(maskedTextBox19.Text + " (" + comboBox10.Text + ")");
                graph.Series[maskedTextBox19.Text + " (" + comboBox10.Text + ")"].ChartType = SeriesChartType.Line;
                graph.Series[maskedTextBox19.Text + " (" + comboBox10.Text + ")"].XValueType = ChartValueType.Time;
            }
            if (this.checkBox3.Checked)
            {
                graph.Series.Add(maskedTextBox1.Text + " (" + comboBox9.Text + ")");
                graph.Series[maskedTextBox1.Text + " (" + comboBox9.Text + ")"].ChartType = SeriesChartType.Line;
                graph.Series[maskedTextBox1.Text + " (" + comboBox9.Text + ")"].XValueType = ChartValueType.Time;
            }
            if (this.checkBox12.Checked)
            {
                graph.Series.Add("Switch Climatic Chamber");
                graph.Series["Switch Climatic Chamber"].ChartType = SeriesChartType.Line;
                graph.Series["Switch Climatic Chamber"].XValueType = ChartValueType.Time;
            }
            if (this.checkBox14.Checked)
            {
                graph.Series.Add("Temperature (°C)");
                graph.Series["Temperature (°C)"].ChartType = SeriesChartType.Line;
                graph.Series["Temperature (°C)"].XValueType = ChartValueType.Time;
            }
            if (this.checkBox22.Checked)
            {
                graph.Series.Add("Tension (V)");
                graph.Series["Tension (V)"].ChartType = SeriesChartType.Line;
                graph.Series["Tension (V)"].XValueType = ChartValueType.Time;
            }

            Thread Arduino_data1 = new Thread(new ThreadStart(Arduino_data));
            Arduino_data1.Start();

            if (checkBox1.Checked)
            {
                Thread Excel_data1 = new Thread(new ThreadStart(Excel_data));
                Excel_data1.Start();
            }
        }

        private void stop_Click(object sender, EventArgs e)
        {
            start.Enabled = true;
            stop.Enabled = false;
            checkBox1.Enabled = true;
            checkBox4.Enabled = true;
            checkBox5.Enabled = true;
            checkBox6.Enabled = true;
            checkBox7.Enabled = true;
            checkBox11.Enabled = true;
            checkBox10.Enabled = true;
            checkBox9.Enabled = true;
            checkBox8.Enabled = true;
            checkBox21.Enabled = true;
            checkBox20.Enabled = true;
            checkBox19.Enabled = true;
            checkBox18.Enabled = true;
            checkBox17.Enabled = true;
            checkBox16.Enabled = true;
            checkBox13.Enabled = true;
            checkBox3.Enabled = true;
            label65.Visible = false;
            maskedTextBox34.Visible = false;
            label66.Visible = false;
            maskedTextBox35.Visible = false;
            label67.Visible = false;
            maskedTextBox36.Visible = false;
            label68.Visible = false;
            maskedTextBox37.Visible = false;
            label69.Visible = false;
            maskedTextBox38.Visible = false;
            label70.Visible = false;
            maskedTextBox39.Visible = false;
            label71.Visible = false;
            maskedTextBox40.Visible = false;
            label72.Visible = false;
            maskedTextBox41.Visible = false;
            label73.Visible = false;
            maskedTextBox42.Visible = false;
            label74.Visible = false;
            maskedTextBox43.Visible = false;
            label75.Visible = false;
            maskedTextBox44.Visible = false;
            label83.Visible = false;
            maskedTextBox52.Visible = false;
            label82.Visible = false;
            maskedTextBox51.Visible = false;
            label81.Visible = false;
            maskedTextBox50.Visible = false;
            label80.Visible = false;
            maskedTextBox49.Visible = false;
            label79.Visible = false;
            maskedTextBox48.Visible = false;
            label78.Visible = false;
            maskedTextBox47.Visible = false;
            label77.Visible = false;
            maskedTextBox46.Visible = false;
            label76.Visible = false;
            maskedTextBox45.Visible = false;
            checkBox16.Enabled = true;
            checkBox13.Enabled = true;
            checkBox3.Enabled = true;
            button1.Enabled = true;
            checkBox14.Enabled = true;
            checkBox12.Enabled = true;
            checkBox22.Enabled = true;
            status_acquisitions.Text = "Stopped";
            status_acquisitions.ForeColor = Color.Red;
            run3 = false;
            run2 = false;
            run1 = false;
            run = false;

            if (checkBox4.Checked)
            {
                graph.Series[maskedTextBox11.Text + " (" + comboBox1.Text + ")"].Points.Clear();
            }
            if (checkBox5.Checked)
            {
                graph.Series[maskedTextBox12.Text + " (" + comboBox2.Text + ")"].Points.Clear();
            }
            if (checkBox6.Checked)
            {
                graph.Series[maskedTextBox13.Text + " (" + comboBox3.Text + ")"].Points.Clear();
            }
            if (checkBox7.Checked)
            {
                graph.Series[maskedTextBox14.Text + " (" + comboBox4.Text + ")"].Points.Clear();
            }
            if (checkBox11.Checked)
            {
                graph.Series[maskedTextBox15.Text + " (" + comboBox8.Text + ")"].Points.Clear();
            }
            if (checkBox10.Checked)
            {
                graph.Series[maskedTextBox16.Text + " (" + comboBox7.Text + ")"].Points.Clear();
            }
            if (checkBox9.Checked)
            {
                graph.Series[maskedTextBox17.Text + " (" + comboBox6.Text + ")"].Points.Clear();
            }
            if (checkBox8.Checked)
            {
                graph.Series[maskedTextBox18.Text + " (" + comboBox5.Text + ")"].Points.Clear();
            }
            if (checkBox21.Checked)
            {
                graph.Series[maskedTextBox25.Text + " (" + comboBox16.Text + ")"].Points.Clear();
            }
            if (checkBox20.Checked)
            {
                graph.Series[maskedTextBox24.Text + " (" + comboBox15.Text + ")"].Points.Clear();
            }
            if (checkBox19.Checked)
            {
                graph.Series[maskedTextBox23.Text + " (" + comboBox14.Text + ")"].Points.Clear();
            }
            if (checkBox18.Checked)
            {
                graph.Series[maskedTextBox22.Text + " (" + comboBox13.Text + ")"].Points.Clear();
            }
            if (checkBox17.Checked)
            {
                graph.Series[maskedTextBox21.Text + " (" + comboBox12.Text + ")"].Points.Clear();
            }
            if (checkBox16.Checked)
            {
                graph.Series[maskedTextBox20.Text + " (" + comboBox11.Text + ")"].Points.Clear();
            }
            if (checkBox13.Checked)
            {
                graph.Series[maskedTextBox19.Text + " (" + comboBox10.Text + ")"].Points.Clear();
            }
            if (checkBox3.Checked)
            {
                graph.Series[maskedTextBox1.Text + " (" + comboBox9.Text + ")"].Points.Clear();
            }
            if (checkBox14.Checked)
            {
                graph.Series["Temperature (°C)"].Points.Clear();
            }
            if (checkBox12.Checked)
            {
                graph.Series["Switch Climatic Chamber"].Points.Clear();
            }
            if (checkBox22.Checked)
            {
                graph.Series["Tension (V)"].Points.Clear();
            }
            graph.Series.Clear();
        }

        private void refresh_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                try
                {
                    serialPort1.Close();

                    progressBar1.Value = 0;
                    start.Enabled = false;
                    stop.Enabled = false;
                    connect.Enabled = true;
                    refresh.Enabled = false;
                    button1.Enabled = false;
                    run3 = false;
                    run2 = false;
                    run1 = false;
                    run = false;
                    label_status.Text = "DISCONNECTED";
                    label_status.ForeColor = Color.Red;
                }
                catch (Exception error)
                {
                    MessageBox.Show(error.Message);
                }
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                    serialPort1.Close();
                    run3 = false;
                    run2 = false;
                    run1 = false;
                    run = false;
            }
        }

        private void label_status_Click(object sender, EventArgs e)
        {

        }

        private void status_acquisitions_Click(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                textBox1.Enabled = true;
                filesavebutton.Enabled = true;
                filesave.Enabled = true;
                checkBox15.Enabled = true;
                maskedTextBox2.Enabled = true;
                button2.Enabled = true;
            }
            else
            {
                textBox1.Enabled = false;
                filesavebutton.Enabled = false;
                filesave.Enabled = false;
                checkBox15.Enabled = false;
                maskedTextBox2.Enabled = false;
                button2.Enabled = false;
            }
        }

        private void serialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
        }

        private void ShowData(object sender, EventArgs e)
        {
        }

        private void richTextBox_data_TextChanged(object sender, EventArgs e)
        {
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void graph_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        public void Arduino_data()
        {
            while (run == true)
            {
                string data_serial = serialPort1.ReadLine();
                var split_data_serial = data_serial.Split(';');
                var Date_Time = DateTime.Now;

                graph.Invoke(new MethodInvoker(delegate
                    {
                        if (checkBox4.Checked & stop.Enabled == true)
                        {
                            label68.Visible = true;
                            label68.Text = maskedTextBox11.Text + " (" + comboBox1.Text + ") :";
                            maskedTextBox37.Visible = true;
                            maskedTextBox37.Text = split_data_serial[0];
                            double data_serial1 = Convert.ToDouble(maskedTextBox3.Text) * Convert.ToDouble(split_data_serial[0], System.Globalization.CultureInfo.InvariantCulture);
                            graph.Series[maskedTextBox11.Text + " (" + comboBox1.Text + ")"].Points.AddXY(Date_Time, data_serial1);
                        }
                        if(checkBox5.Checked & stop.Enabled == true)
                        {
                            label69.Visible = true;
                            label69.Text = maskedTextBox12.Text + " (" + comboBox2.Text + ") :";
                            maskedTextBox38.Visible = true;
                            maskedTextBox38.Text = split_data_serial[1];
                            double data_serial2 = Convert.ToDouble(maskedTextBox4.Text) * Convert.ToDouble(split_data_serial[1], System.Globalization.CultureInfo.InvariantCulture);
                            graph.Series[maskedTextBox12.Text + " (" + comboBox2.Text + ")"].Points.AddXY(Date_Time, data_serial2);
                        }
                        if (checkBox6.Checked & stop.Enabled == true)
                        {
                            label70.Visible = true;
                            label70.Text = maskedTextBox13.Text + " (" + comboBox3.Text + ") :";
                            maskedTextBox39.Visible = true;
                            maskedTextBox39.Text = split_data_serial[2];
                            double data_serial3 = Convert.ToDouble(maskedTextBox5.Text) * Convert.ToDouble(split_data_serial[2], System.Globalization.CultureInfo.InvariantCulture);
                            graph.Series[maskedTextBox13.Text + " (" + comboBox3.Text + ")"].Points.AddXY(Date_Time, data_serial3);
                        }
                        if (checkBox7.Checked & stop.Enabled == true)
                        {
                            label71.Visible = true;
                            label71.Text = maskedTextBox14.Text + " (" + comboBox4.Text + ") :";
                            maskedTextBox40.Visible = true;
                            maskedTextBox40.Text = split_data_serial[3];
                            double data_serial4 = Convert.ToDouble(maskedTextBox6.Text) * Convert.ToDouble(split_data_serial[3], System.Globalization.CultureInfo.InvariantCulture);
                            graph.Series[maskedTextBox14.Text + " (" + comboBox4.Text + ")"].Points.AddXY(Date_Time, data_serial4);
                        }
                        if (checkBox11.Checked & stop.Enabled == true)
                        {
                            label72.Visible = true;
                            label72.Text = maskedTextBox15.Text + " (" + comboBox8.Text + ") :";
                            maskedTextBox41.Visible = true;
                            maskedTextBox41.Text = split_data_serial[4];
                            double data_serial5 = Convert.ToDouble(maskedTextBox10.Text) * Convert.ToDouble(split_data_serial[4], System.Globalization.CultureInfo.InvariantCulture);
                            graph.Series[maskedTextBox15.Text + " (" + comboBox8.Text + ")"].Points.AddXY(Date_Time, data_serial5);
                        }
                        if (checkBox10.Checked & stop.Enabled == true)
                        {
                            label73.Visible = true;
                            label73.Text = maskedTextBox16.Text + " (" + comboBox7.Text + ") :";
                            maskedTextBox42.Visible = true;
                            maskedTextBox42.Text = split_data_serial[5];
                            double data_serial6 = Convert.ToDouble(maskedTextBox9.Text) * Convert.ToDouble(split_data_serial[5], System.Globalization.CultureInfo.InvariantCulture);
                            graph.Series[maskedTextBox16.Text + " (" + comboBox7.Text + ")"].Points.AddXY(Date_Time, data_serial6);
                        }
                        if (checkBox9.Checked & stop.Enabled == true)
                        {
                            label74.Visible = true;
                            label74.Text = maskedTextBox17.Text + " (" + comboBox6.Text + ") :";
                            maskedTextBox43.Visible = true;
                            maskedTextBox43.Text = split_data_serial[6];
                            double data_serial7 = Convert.ToDouble(maskedTextBox8.Text) * Convert.ToDouble(split_data_serial[6], System.Globalization.CultureInfo.InvariantCulture);
                            graph.Series[maskedTextBox17.Text + " (" + comboBox6.Text + ")"].Points.AddXY(Date_Time, data_serial7);
                        }
                        if (checkBox8.Checked & stop.Enabled == true)
                        {
                            label75.Visible = true;
                            label75.Text = maskedTextBox18.Text + " (" + comboBox5.Text + ") :";
                            maskedTextBox44.Visible = true;
                            maskedTextBox44.Text = split_data_serial[7];
                            double data_serial8 = Convert.ToDouble(maskedTextBox7.Text) * Convert.ToDouble(split_data_serial[7], System.Globalization.CultureInfo.InvariantCulture);
                            graph.Series[maskedTextBox18.Text + " (" + comboBox5.Text + ")"].Points.AddXY(Date_Time, data_serial8);
                        }
                        if (checkBox21.Checked & stop.Enabled == true)
                        {
                            label83.Visible = true;
                            label83.Text = maskedTextBox25.Text + " (" + comboBox16.Text + ") :";
                            maskedTextBox52.Visible = true;
                            maskedTextBox52.Text = split_data_serial[8];
                            double data_serial9 = Convert.ToDouble(maskedTextBox33.Text) * Convert.ToDouble(split_data_serial[8], System.Globalization.CultureInfo.InvariantCulture);
                            graph.Series[maskedTextBox25.Text + " (" + comboBox16.Text + ")"].Points.AddXY(Date_Time, data_serial9);
                        }
                        if (checkBox20.Checked & stop.Enabled == true)
                        {
                            label82.Visible = true;
                            label82.Text = maskedTextBox24.Text + " (" + comboBox15.Text + ") :";
                            maskedTextBox51.Visible = true;
                            maskedTextBox51.Text = split_data_serial[9];
                            double data_serial10 = Convert.ToDouble(maskedTextBox32.Text) * Convert.ToDouble(split_data_serial[9], System.Globalization.CultureInfo.InvariantCulture);
                            graph.Series[maskedTextBox24.Text + " (" + comboBox15.Text + ")"].Points.AddXY(Date_Time, data_serial10);
                        }
                        if (checkBox19.Checked & stop.Enabled == true)
                        {
                            label81.Visible = true;
                            label81.Text = maskedTextBox23.Text + " (" + comboBox14.Text + ") :";
                            maskedTextBox50.Visible = true;
                            maskedTextBox50.Text = split_data_serial[10];
                            double data_serial11 = Convert.ToDouble(maskedTextBox31.Text) * Convert.ToDouble(split_data_serial[10], System.Globalization.CultureInfo.InvariantCulture);
                            graph.Series[maskedTextBox23.Text + " (" + comboBox14.Text + ")"].Points.AddXY(Date_Time, data_serial11);
                        }
                        if (checkBox18.Checked & stop.Enabled == true)
                        {
                            label80.Visible = true;
                            label80.Text = maskedTextBox22.Text + " (" + comboBox13.Text + ") :";
                            maskedTextBox49.Visible = true;
                            maskedTextBox49.Text = split_data_serial[11];
                            double data_serial12 = Convert.ToDouble(maskedTextBox30.Text) * Convert.ToDouble(split_data_serial[11], System.Globalization.CultureInfo.InvariantCulture);
                            graph.Series[maskedTextBox22.Text + " (" + comboBox13.Text + ")"].Points.AddXY(Date_Time, data_serial12);
                        }
                        if (checkBox17.Checked & stop.Enabled == true)
                        {
                            label79.Visible = true;
                            label79.Text = maskedTextBox21.Text + " (" + comboBox12.Text + ") :";
                            maskedTextBox48.Visible = true;
                            maskedTextBox48.Text = split_data_serial[12];
                            double data_serial13 = Convert.ToDouble(maskedTextBox26.Text) * Convert.ToDouble(split_data_serial[12], System.Globalization.CultureInfo.InvariantCulture);
                            graph.Series[maskedTextBox21.Text + " (" + comboBox12.Text + ")"].Points.AddXY(Date_Time, data_serial13);
                        }
                        if (checkBox16.Checked & stop.Enabled == true)
                        {
                            label78.Visible = true;
                            label78.Text = maskedTextBox20.Text + " (" + comboBox11.Text + ") :";
                            maskedTextBox47.Visible = true;
                            maskedTextBox47.Text = split_data_serial[13];
                            double data_serial14 = Convert.ToDouble(maskedTextBox29.Text) * Convert.ToDouble(split_data_serial[13], System.Globalization.CultureInfo.InvariantCulture);
                            graph.Series[maskedTextBox20.Text + " (" + comboBox11.Text + ")"].Points.AddXY(Date_Time, data_serial14);
                        }
                        if (checkBox13.Checked & stop.Enabled == true)
                        {
                            label77.Visible = true;
                            label77.Text = maskedTextBox19.Text + " (" + comboBox10.Text + ") :";
                            maskedTextBox46.Visible = true;
                            maskedTextBox46.Text = split_data_serial[14];
                            double data_serial15 = Convert.ToDouble(maskedTextBox28.Text) * Convert.ToDouble(split_data_serial[14], System.Globalization.CultureInfo.InvariantCulture);
                            graph.Series[maskedTextBox19.Text + " (" + comboBox10.Text + ")"].Points.AddXY(Date_Time, data_serial15);
                        }
                        if (checkBox3.Checked & stop.Enabled == true)
                        {
                            label76.Visible = true;
                            label76.Text = maskedTextBox1.Text + " (" + comboBox9.Text + ") :";
                            maskedTextBox45.Visible = true;
                            maskedTextBox45.Text = split_data_serial[15];
                            double data_serial16 = Convert.ToDouble(maskedTextBox27.Text) * Convert.ToDouble(split_data_serial[15], System.Globalization.CultureInfo.InvariantCulture);
                            graph.Series[maskedTextBox1.Text + " (" + comboBox9.Text + ")"].Points.AddXY(Date_Time, data_serial16);
                        }
                        if (checkBox14.Checked & stop.Enabled == true)
                        {
                            label66.Visible = true;
                            maskedTextBox35.Visible = true;
                            maskedTextBox35.Text = split_data_serial[16];
                            var data_serial17 = split_data_serial[16];
                            graph.Series["Temperature (°C)"].Points.AddXY(Date_Time, data_serial17);
                        }
                        if (checkBox12.Checked & stop.Enabled == true)
                        {
                            label65.Visible = true;
                            maskedTextBox34.Visible = true;
                            maskedTextBox34.Text = split_data_serial[17];
                            var data_serial18 = split_data_serial[17];
                            graph.Series["Switch Climatic Chamber"].Points.AddXY(Date_Time, data_serial18);
                        }
                        if (checkBox22.Checked & stop.Enabled == true)
                        {
                            label67.Visible = true;
                            maskedTextBox36.Visible = true;
                            maskedTextBox36.Text = split_data_serial[18];
                            var data_serial19 = split_data_serial[18];
                            graph.Series["Tension (V)"].Points.AddXY(Date_Time, data_serial19);
                        }
                    }));
            }
            run = true;
        }

        public void Excel_data()
        {
            while (run1 == true)
            {
                Thread.Sleep(Convert.ToInt32(textBox1.Text) * 60000);

                    Excel.Application xlApp;
                    Excel.Workbook xlWorkBook;
                    Excel.Worksheet xlWorkSheet;
                    object misValue = System.Reflection.Missing.Value;

                    xlApp = new Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                for (int i = 0; i < graph.Series.Count; i++)
                    {
                        xlWorkSheet.Cells[1, 1] = "";//put your column heading here

                        if (checkBox4.Checked)
                        {
                            xlWorkSheet.Cells[1, 2] = string.Format(maskedTextBox11.Text);   //string.Format(maskedTextBox11.Text) + " (" + string.Format(comboBox1.Text) + ")";
                        }
                        if (checkBox5.Checked)
                        {
                            xlWorkSheet.Cells[1, 3] = string.Format(maskedTextBox12.Text);   //string.Format(maskedTextBox12.Text) + " (" + string.Format(comboBox2.Text) + ")";
                        }
                        if (checkBox6.Checked)
                        {
                            xlWorkSheet.Cells[1, 4] = string.Format(maskedTextBox13.Text);   //string.Format(maskedTextBox13.Text) + " (" + string.Format(comboBox3.Text) + ")";
                        }
                        if (checkBox7.Checked)
                        {
                            xlWorkSheet.Cells[1, 5] = string.Format(maskedTextBox14.Text);   //string.Format(maskedTextBox14.Text) + " (" + string.Format(comboBox4.Text) + ")";
                        }
                        if (checkBox11.Checked)
                        {
                            xlWorkSheet.Cells[1, 6] = string.Format(maskedTextBox15.Text);   //string.Format(maskedTextBox15.Text) + " (" + string.Format(comboBox8.Text) + ")";
                        }
                        if (checkBox10.Checked)
                        {
                            xlWorkSheet.Cells[1, 7] = string.Format(maskedTextBox16.Text);   //string.Format(maskedTextBox16.Text) + " (" + string.Format(comboBox7.Text) + ")";
                        }
                        if (checkBox9.Checked)
                        {
                            xlWorkSheet.Cells[1, 8] = string.Format(maskedTextBox17.Text);   //string.Format(maskedTextBox17.Text) + " (" + string.Format(comboBox6.Text) + ")";
                        }
                        if (checkBox8.Checked)
                        {
                            xlWorkSheet.Cells[1, 9] = string.Format(maskedTextBox18.Text);   //string.Format(maskedTextBox18.Text) + " (" + string.Format(comboBox5.Text) + ")";
                        }
                        if (checkBox21.Checked)
                        {
                            xlWorkSheet.Cells[1, 10] = string.Format(maskedTextBox25.Text);
                        }
                        if (checkBox20.Checked)
                        {
                            xlWorkSheet.Cells[1, 11] = string.Format(maskedTextBox24.Text);
                        }
                        if (checkBox19.Checked)
                        {
                            xlWorkSheet.Cells[1, 12] = string.Format(maskedTextBox23.Text);
                        }
                        if (checkBox18.Checked)
                        {
                            xlWorkSheet.Cells[1, 13] = string.Format(maskedTextBox22.Text);
                        }
                        if (checkBox17.Checked)
                        {
                            xlWorkSheet.Cells[1, 14] = string.Format(maskedTextBox21.Text);
                        }
                        if (checkBox16.Checked)
                        {
                            xlWorkSheet.Cells[1, 15] = string.Format(maskedTextBox20.Text);
                        }
                        if (checkBox13.Checked)
                        {
                            xlWorkSheet.Cells[1, 16] = string.Format(maskedTextBox19.Text);
                        }
                        if (checkBox3.Checked)
                        {
                            xlWorkSheet.Cells[1, 17] = string.Format(maskedTextBox1.Text);
                        }
                        if (checkBox14.Checked)
                        {
                            xlWorkSheet.Cells[1, 15] = string.Format("Temperature (°C)");
                        }
                        if (checkBox12.Checked)
                        {
                            xlWorkSheet.Cells[1, 16] = string.Format("Switch Climatic Chamber");
                        }
                        if (checkBox22.Checked)
                        {
                            xlWorkSheet.Cells[1, 17] = string.Format("Tension (V)");
                        }

                    Range rg = (Excel.Range)xlWorkSheet.Cells[1, 1];
                        rg.EntireColumn.NumberFormat = "hh:mm:ss";

                        graph.Invoke(new MethodInvoker(delegate
                        {
                            for (int j = 0; j < graph.Series[i].Points.Count; j++)
                            {
                                xlWorkSheet.Cells[j + 2, 1] = graph.Series[i].Points[j].XValue;

                                if (checkBox4.Checked)
                                {
                                    xlWorkSheet.Cells[j + 2, 2] = graph.Series[maskedTextBox11.Text + " (" + comboBox1.Text + ")"].Points[j].YValues[0];
                                }
                                if (checkBox5.Checked)
                                {
                                    xlWorkSheet.Cells[j + 2, 3] = graph.Series[maskedTextBox12.Text + " (" + comboBox2.Text + ")"].Points[j].YValues[0];
                                }
                                if (checkBox6.Checked)
                                {
                                    xlWorkSheet.Cells[j + 2, 4] = graph.Series[maskedTextBox13.Text + " (" + comboBox3.Text + ")"].Points[j].YValues[0];
                                }
                                if (checkBox7.Checked)
                                {
                                    xlWorkSheet.Cells[j + 2, 5] = graph.Series[maskedTextBox14.Text + " (" + comboBox4.Text + ")"].Points[j].YValues[0];
                                }
                                if (checkBox11.Checked)
                                {
                                    xlWorkSheet.Cells[j + 2, 6] = graph.Series[maskedTextBox15.Text + " (" + comboBox8.Text + ")"].Points[j].YValues[0];
                                }
                                if (checkBox10.Checked)
                                {
                                    xlWorkSheet.Cells[j + 2, 7] = graph.Series[maskedTextBox16.Text + " (" + comboBox7.Text + ")"].Points[j].YValues[0];
                                }
                                if (checkBox9.Checked)
                                {
                                    xlWorkSheet.Cells[j + 2, 8] = graph.Series[maskedTextBox17.Text + " (" + comboBox6.Text + ")"].Points[j].YValues[0];
                                }
                                if (checkBox8.Checked)
                                {
                                    xlWorkSheet.Cells[j + 2, 9] = graph.Series[maskedTextBox18.Text + " (" + comboBox5.Text + ")"].Points[j].YValues[0];
                                }
                                if (checkBox21.Checked)
                                {
                                    xlWorkSheet.Cells[j + 2, 10] = graph.Series[maskedTextBox25.Text + " (" + comboBox16.Text + ")"].Points[j].YValues[0];
                                }
                                if (checkBox20.Checked)
                                {
                                    xlWorkSheet.Cells[j + 2, 11] = graph.Series[maskedTextBox24.Text + " (" + comboBox15.Text + ")"].Points[j].YValues[0];
                                }
                                if (checkBox19.Checked)
                                {
                                    xlWorkSheet.Cells[j + 2, 12] = graph.Series[maskedTextBox23.Text + " (" + comboBox14.Text + ")"].Points[j].YValues[0];
                                }
                                if (checkBox18.Checked)
                                {
                                    xlWorkSheet.Cells[j + 2, 13] = graph.Series[maskedTextBox22.Text + " (" + comboBox13.Text + ")"].Points[j].YValues[0];
                                }
                                if (checkBox17.Checked)
                                {
                                    xlWorkSheet.Cells[j + 2, 14] = graph.Series[maskedTextBox21.Text + " (" + comboBox12.Text + ")"].Points[j].YValues[0];
                                }
                                if (checkBox16.Checked)
                                {
                                    xlWorkSheet.Cells[j + 2, 15] = graph.Series[maskedTextBox20.Text + " (" + comboBox11.Text + ")"].Points[j].YValues[0];
                                }
                                if (checkBox13.Checked)
                                {
                                    xlWorkSheet.Cells[j + 2, 16] = graph.Series[maskedTextBox19.Text + " (" + comboBox10.Text + ")"].Points[j].YValues[0];
                                }
                                if (checkBox3.Checked)
                                {
                                    xlWorkSheet.Cells[j + 2, 17] = graph.Series[maskedTextBox1.Text + " (" + comboBox9.Text + ")"].Points[j].YValues[0];
                                }
                                if (checkBox14.Checked)
                                {
                                    xlWorkSheet.Cells[j + 2, 15] = graph.Series["Temperature (°C)"].Points[j].YValues[0];
                                }
                                if (checkBox12.Checked)
                                {
                                    xlWorkSheet.Cells[j + 2, 16] = graph.Series["Switch Climatic Chamber"].Points[j].YValues[0];
                                }
                                if (checkBox22.Checked)
                                {
                                    xlWorkSheet.Cells[j + 2, 17] = graph.Series["Tension (V)"].Points[j].YValues[0];
                                }
                            }
                        }));
                    }

                    Excel.Range chartRange;

                    Excel.ChartObjects xlCharts = (Excel.ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
                    Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(830, 10, 880, 400);
                    Excel.Chart chartPage = myChart.Chart;

                    chartRange = xlWorkSheet.get_Range("A1", "Q:Q" );//update the range here
                    chartPage.SetSourceData(chartRange, misValue);
                    chartPage.ChartType = Excel.XlChartType.xlLine;
                    xlWorkBook.SaveAs(string.Format(filesave.Text) + "/" + "Graphic_Data_" + DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss") + ".xlsx", Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBook.Close(true, misValue, misValue);
                    xlApp.Quit();

                    releaseObject(xlWorkSheet);
                    releaseObject(xlWorkBook);
                    releaseObject(xlApp);

                    run = false;
                    graph.Invoke(new MethodInvoker(delegate
                    {
                        graph.SaveImage(string.Format(maskedTextBox2.Text) + "/" + "Image_" + DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss") + ".png", ChartImageFormat.Png);
                    }));
                    run = true;

                if (checkBox15.Checked)
                {
                    graph.Invoke(new MethodInvoker(delegate
                    {
                        if (checkBox4.Checked)
                        {
                            graph.Series[maskedTextBox11.Text + " (" + comboBox1.Text + ")"].Points.Clear();
                        }
                        if (checkBox5.Checked)
                        {
                            graph.Series[maskedTextBox12.Text + " (" + comboBox2.Text + ")"].Points.Clear();
                        }
                        if (checkBox6.Checked)
                        {
                            graph.Series[maskedTextBox13.Text + " (" + comboBox3.Text + ")"].Points.Clear();
                        }
                        if (checkBox7.Checked)
                        {
                            graph.Series[maskedTextBox14.Text + " (" + comboBox4.Text + ")"].Points.Clear();
                        }
                        if (checkBox11.Checked)
                        {
                            graph.Series[maskedTextBox15.Text + " (" + comboBox8.Text + ")"].Points.Clear();
                        }
                        if (checkBox10.Checked)
                        {
                            graph.Series[maskedTextBox16.Text + " (" + comboBox7.Text + ")"].Points.Clear();
                        }
                        if (checkBox9.Checked)
                        {
                            graph.Series[maskedTextBox17.Text + " (" + comboBox6.Text + ")"].Points.Clear();
                        }
                        if (checkBox8.Checked)
                        {
                            graph.Series[maskedTextBox18.Text + " (" + comboBox5.Text + ")"].Points.Clear();
                        }
                        if (checkBox21.Checked)
                        {
                            graph.Series[maskedTextBox25.Text + " (" + comboBox16.Text + ")"].Points.Clear();
                        }
                        if (checkBox20.Checked)
                        {
                            graph.Series[maskedTextBox24.Text + " (" + comboBox15.Text + ")"].Points.Clear();
                        }
                        if (checkBox19.Checked)
                        {
                            graph.Series[maskedTextBox23.Text + " (" + comboBox14.Text + ")"].Points.Clear();
                        }
                        if (checkBox18.Checked)
                        {
                            graph.Series[maskedTextBox22.Text + " (" + comboBox13.Text + ")"].Points.Clear();
                        }
                        if (checkBox17.Checked)
                        {
                            graph.Series[maskedTextBox21.Text + " (" + comboBox12.Text + ")"].Points.Clear();
                        }
                        if (checkBox16.Checked)
                        {
                            graph.Series[maskedTextBox20.Text + " (" + comboBox11.Text + ")"].Points.Clear();
                        }
                        if (checkBox13.Checked)
                        {
                            graph.Series[maskedTextBox19.Text + " (" + comboBox10.Text + ")"].Points.Clear();
                        }
                        if (checkBox3.Checked)
                        {
                            graph.Series[maskedTextBox1.Text + " (" + comboBox9.Text + ")"].Points.Clear();
                        }
                        if (checkBox14.Checked)
                        {
                            graph.Series["Temperature (°C)"].Points.Clear();
                        }
                        if (checkBox12.Checked)
                        {
                            graph.Series["Switch Climatic Chamber"].Points.Clear();
                        }
                        if (checkBox22.Checked)
                        {
                            graph.Series["Tension (V)"].Points.Clear();
                        }
                    }));
                }
            }
            run1 = true;
        }

        private void releaseObject(object obj)
        {
                try
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                    obj = null;
                }
                catch (Exception)
                {
                    obj = null;
                }
                finally
                {
                    GC.Collect();
                }
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void maskedTextBox1_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(serialPort1.ReadExisting() != "")
            {
                MessageBox.Show("Information : WiFi is OK.");
            }
            else
            {
                MessageBox.Show("Information : No WiFi connection. Check connection and try to restart PWM of the measurement board.");
            }
        }


        private void maskedTextBox2_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            var fileContent1 = string.Empty;
            var filePath1 = string.Empty;

            using (FolderBrowserDialog openFolderDialog = new FolderBrowserDialog())
            {

                if (openFolderDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePath1 = openFolderDialog.SelectedPath;

                }
            }

            maskedTextBox2.Text = filePath1;
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void maskedTextBox3_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox4_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox5_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox6_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox7_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox8_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox9_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox10_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked)
            {
                maskedTextBox11.Enabled = true;
                maskedTextBox3.Enabled = true;
                comboBox1.Enabled = true;
            }
            else
            {
                maskedTextBox11.Enabled = false;
                maskedTextBox3.Enabled = false;
                comboBox1.Enabled = false;
            }
        }

        private void maskedTextBox11_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked)
            {
                maskedTextBox12.Enabled = true;
                maskedTextBox4.Enabled = true;
                comboBox2.Enabled = true;
            }
            else
            {
                maskedTextBox12.Enabled = false;
                maskedTextBox4.Enabled = false;
                comboBox2.Enabled = false;
            }
        }

        private void maskedTextBox12_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked)
            {
                maskedTextBox13.Enabled = true;
                maskedTextBox5.Enabled = true;
                comboBox3.Enabled = true;
            }
            else
            {
                maskedTextBox13.Enabled = false;
                maskedTextBox5.Enabled = false;
                comboBox3.Enabled = false;
            }
        }

        private void maskedTextBox13_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox7.Checked)
            {
                maskedTextBox14.Enabled = true;
                maskedTextBox6.Enabled = true;
                comboBox4.Enabled = true;
            }
            else
            {
                maskedTextBox14.Enabled = false;
                maskedTextBox6.Enabled = false;
                comboBox4.Enabled = false;
            }
        }

        private void maskedTextBox14_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkBox11_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox11.Checked)
            {
                maskedTextBox15.Enabled = true;
                maskedTextBox10.Enabled = true;
                comboBox8.Enabled = true;
            }
            else
            {
                maskedTextBox15.Enabled = false;
                maskedTextBox10.Enabled = false;
                comboBox8.Enabled = false;
            }
        }

        private void maskedTextBox15_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkBox10_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox10.Checked)
            {
                maskedTextBox16.Enabled = true;
                maskedTextBox9.Enabled = true;
                comboBox7.Enabled = true;
            }
            else
            {
                maskedTextBox16.Enabled = false;
                maskedTextBox9.Enabled = false;
                comboBox7.Enabled = false;
            }
        }

        private void maskedTextBox16_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkBox9_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox9.Checked)
            {
                maskedTextBox17.Enabled = true;
                maskedTextBox8.Enabled = true;
                comboBox6.Enabled = true;
            }
            else
            {
                maskedTextBox17.Enabled = false;
                maskedTextBox8.Enabled = false;
                comboBox6.Enabled = false;
            }
        }

        private void maskedTextBox17_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox8.Checked)
            {
                maskedTextBox18.Enabled = true;
                maskedTextBox7.Enabled = true;
                comboBox5.Enabled = true;
            }
            else
            {
                maskedTextBox18.Enabled = false;
                maskedTextBox7.Enabled = false;
                comboBox5.Enabled = false;
            }
        }

        private void maskedTextBox18_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkBox13_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void checkBox14_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox15_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox12_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            portselect.Items.Clear();
            string[] ports = SerialPort.GetPortNames();
            if (!connect.Enabled == false)
                portselect.Text = "";
            for (int index = 0; index < ports.Length; index++)
            {
                portselect.Items.Add(ports[index]);
            }
        }

        private void checkBox21_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox21.Checked)
            {
                maskedTextBox33.Enabled = true;
                maskedTextBox25.Enabled = true;
                comboBox16.Enabled = true;
            }
            else
            {
                maskedTextBox33.Enabled = false;
                maskedTextBox25.Enabled = false;
                comboBox16.Enabled = false;
            }
        }

        private void checkBox20_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox20.Checked)
            {
                maskedTextBox32.Enabled = true;
                maskedTextBox24.Enabled = true;
                comboBox15.Enabled = true;
            }
            else
            {
                maskedTextBox32.Enabled = false;
                maskedTextBox24.Enabled = false;
                comboBox15.Enabled = false;
            }
        }

        private void checkBox19_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox19.Checked)
            {
                maskedTextBox31.Enabled = true;
                maskedTextBox23.Enabled = true;
                comboBox14.Enabled = true;
            }
            else
            {
                maskedTextBox31.Enabled = false;
                maskedTextBox23.Enabled = false;
                comboBox14.Enabled = false;
            }
        }

        private void checkBox18_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox18.Checked)
            {
                maskedTextBox30.Enabled = true;
                maskedTextBox22.Enabled = true;
                comboBox13.Enabled = true;
            }
            else
            {
                maskedTextBox30.Enabled = false;
                maskedTextBox22.Enabled = false;
                comboBox13.Enabled = false;
            }
        }

        private void checkBox17_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox17.Checked)
            {
                maskedTextBox26.Enabled = true;
                maskedTextBox21.Enabled = true;
                comboBox12.Enabled = true;
            }
            else
            {
                maskedTextBox26.Enabled = false;
                maskedTextBox21.Enabled = false;
                comboBox12.Enabled = false;
            }
        }

        private void checkBox16_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox16.Checked)
            {
                maskedTextBox29.Enabled = true;
                maskedTextBox20.Enabled = true;
                comboBox11.Enabled = true;
            }
            else
            {
                maskedTextBox29.Enabled = false;
                maskedTextBox20.Enabled = false;
                comboBox11.Enabled = false;
            }
        }

        private void checkBox13_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkBox13.Checked)
            {
                maskedTextBox28.Enabled = true;
                maskedTextBox19.Enabled = true;
                comboBox10.Enabled = true;
            }
            else
            {
                maskedTextBox28.Enabled = false;
                maskedTextBox19.Enabled = false;
                comboBox10.Enabled = false;
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            {
                maskedTextBox27.Enabled = true;
                maskedTextBox1.Enabled = true;
                comboBox9.Enabled = true;
            }
            else
            {
                maskedTextBox27.Enabled = false;
                maskedTextBox1.Enabled = false;
                comboBox9.Enabled = false;
            }
        }

        private void maskedTextBox25_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox24_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox23_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox22_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox21_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox20_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox19_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox1_MaskInputRejected_1(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox33_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox32_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox31_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox30_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox26_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox29_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox28_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox27_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void comboBox16_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox15_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox14_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox13_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox12_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox11_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox10_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkBox22_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void maskedTextBox34_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void label65_Click(object sender, EventArgs e)
        {

        }

        private void label68_Click(object sender, EventArgs e)
        {

        }

        private void maskedTextBox37_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void label73_Click(object sender, EventArgs e)
        {

        }

        private void maskedTextBox35_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox36_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void label67_Click(object sender, EventArgs e)
        {

        }

        private void label66_Click(object sender, EventArgs e)
        {

        }

        private void maskedTextBox38_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox39_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox40_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox41_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox42_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox43_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox44_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox52_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox51_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox50_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox49_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox48_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox47_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox46_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox45_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void label69_Click(object sender, EventArgs e)
        {

        }

        private void label70_Click(object sender, EventArgs e)
        {

        }

        private void label71_Click(object sender, EventArgs e)
        {

        }

        private void label72_Click(object sender, EventArgs e)
        {

        }

        private void label74_Click(object sender, EventArgs e)
        {

        }

        private void label75_Click(object sender, EventArgs e)
        {

        }

        private void label83_Click(object sender, EventArgs e)
        {

        }

        private void label82_Click(object sender, EventArgs e)
        {

        }

        private void label81_Click(object sender, EventArgs e)
        {

        }

        private void label80_Click(object sender, EventArgs e)
        {

        }

        private void label79_Click(object sender, EventArgs e)
        {

        }

        private void label78_Click(object sender, EventArgs e)
        {

        }

        private void label77_Click(object sender, EventArgs e)
        {

        }

        private void label76_Click(object sender, EventArgs e)
        {

        }
    }
}