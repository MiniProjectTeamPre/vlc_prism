using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace vlc_prism {
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();
        }

        private string name_file = "";
        private string name_pathfile = "";
        private void button1_Click(object sender, EventArgs e) {
            OpenFileDialog load = new OpenFileDialog();
            load.Filter = "csv files (*.csv)|*.csv";
            load.FilterIndex = 1;
            DialogResult dd;
            try {
                dd = load.ShowDialog();
            } catch (Exception) {
                load.InitialDirectory = @"D:\";
                dd = load.ShowDialog();
            }
            if (dd != DialogResult.OK) return;
            name_file = load.SafeFileName;
            name_pathfile = load.FileName;

            flag_run = true;
            button1.Enabled = false;
            button2.Enabled = false;
        }

        public static void DelaymS(int mS) {
            Stopwatch stopwatchDelaymS = new Stopwatch();
            stopwatchDelaymS.Restart();
            while (mS > stopwatchDelaymS.ElapsedMilliseconds) {
                if (!stopwatchDelaymS.IsRunning) stopwatchDelaymS.Start();
                Application.DoEvents();
            }
            stopwatchDelaymS.Stop();
        }

        private bool flag_run = false;
        private int row = 0;
        private string index = "";
        private string uid = "";
        private string pass = "";
        private string prism = "";
        private string comment = "";
        private string panel = "";
        private string sn = "";
        private int total = 0;
        private int lot_pass = 0;
        private int lot_fail = 0;
        private int prism_pass = 0;
        private int prism_fail = 0;
        private int row_index = 0;
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e) {
            while (true) {
                if (!flag_run) { Thread.Sleep(250); continue; }
                string[] wo_sup = name_file.Split('_');
                if (wo_sup.Length != 7) { MessageBox.Show("file err"); return; }
                string wo = wo_sup[2] + "/" + wo_sup[3];
                string[] data_csv = File.ReadAllLines(name_pathfile);
                row_index = data_csv.Length;
                backgroundWorker1.ReportProgress(1);
                backgroundWorker1.ReportProgress(10);
                backgroundWorker1.ReportProgress(12);
                total = 0;
                lot_pass = 0;
                lot_fail = 0;
                prism_pass = 0;
                prism_fail = 0;
                for (int i = 0; i < data_csv.Length; i++) {
                    row = i;
                    backgroundWorker1.ReportProgress(11);
                    string[] data = data_csv[i].Split(',');
                    if (data[0] == "time_start") continue;
                    total++;
                    index = data[2];
                    uid = data[3].Substring(8, 16);
                    pass = data[4];
                    backgroundWorker1.ReportProgress(2);
                    if (data[4] == "0") {
                        lot_fail++;
                        backgroundWorker1.ReportProgress(3);
                    } else {
                        lot_pass++;
                        backgroundWorker1.ReportProgress(4);
                    }
                    panel = TeamPrecision.PRISM.cSNs.GetPanel_uid_VLC(wo, uid);
                    sn = TeamPrecision.PRISM.cSNs.SaveUID_VLC(wo, panel, uid);
                    backgroundWorker1.ReportProgress(5);
                    string[] stetus = TeamPrecision.PRISM.cSNs.CheckStatusSNv2(sn, wo);
                    comment = stetus[1];
                    if (comment.Contains("ผ่านขั้นตอน FCT1")) {
                        prism_pass++;
                        backgroundWorker1.ReportProgress(6);
                    } else {
                        prism_fail++;
                        backgroundWorker1.ReportProgress(7);
                    }
                    backgroundWorker1.ReportProgress(8);
                    backgroundWorker1.ReportProgress(9);
                    Thread.Sleep(150);
                }
                flag_run = false;
                backgroundWorker1.ReportProgress(0);
            }
        }
        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e) {
            switch (e.ProgressPercentage) {
                case 0:
                    button1.Enabled = true;
                    button2.Enabled = true;
                    break;
                case 1:
                    dataGridView1.Rows.Clear();
                    dataGridView1.Rows.Add(row_index);
                    break;
                case 2:
                    dataGridView1.Rows[row].Cells[0].Value = index;
                    dataGridView1.Rows[row].Cells[1].Value = uid;
                    dataGridView1.Rows[row].Cells[2].Value = pass;
                    break;
                case 3:
                    dataGridView1.Rows[row].Cells[0].Style.ForeColor = Color.Red;
                    dataGridView1.Rows[row].Cells[1].Style.ForeColor = Color.Red;
                    dataGridView1.Rows[row].Cells[2].Style.ForeColor = Color.Red;
                    break;
                case 4:
                    dataGridView1.Rows[row].Cells[0].Style.ForeColor = Color.LimeGreen;
                    dataGridView1.Rows[row].Cells[1].Style.ForeColor = Color.LimeGreen;
                    dataGridView1.Rows[row].Cells[2].Style.ForeColor = Color.LimeGreen;
                    break;
                case 5:
                    dataGridView1.Rows[row].Cells[5].Value = panel;
                    dataGridView1.Rows[row].Cells[6].Value = sn;
                    break;
                case 6:
                    dataGridView1.Rows[row].Cells[3].Value = "PASS";
                    dataGridView1.Rows[row].Cells[3].Style.ForeColor = Color.LimeGreen;
                    dataGridView1.Rows[row].Cells[4].Value = comment;
                    dataGridView1.Rows[row].Cells[4].Style.ForeColor = Color.LimeGreen;
                    dataGridView1.Rows[row].Cells[6].Style.ForeColor = Color.LimeGreen;
                    break;
                case 7:
                    dataGridView1.Rows[row].Cells[3].Value = "FAIL";
                    dataGridView1.Rows[row].Cells[3].Style.ForeColor = Color.Red;
                    dataGridView1.Rows[row].Cells[4].Value = comment;
                    dataGridView1.Rows[row].Cells[4].Style.ForeColor = Color.Red;
                    dataGridView1.Rows[row].Cells[6].Style.ForeColor = Color.Red;
                    break;
                case 8:
                    label1.Text = "TOTAL = " + total.ToString();
                    label2.Text = "LOT_PASS = " + lot_pass.ToString();
                    label3.Text = "LOT_FAIL = " + lot_fail.ToString();
                    label5.Text = "PRISM_PASS = " + prism_pass.ToString();
                    label4.Text = "PRISM_FAIL = " + prism_fail.ToString();
                    break;
                case 9:
                    if (row > 20)
                        dataGridView1.FirstDisplayedScrollingRowIndex = row - 20;
                    break;
                case 10:
                    progressBar1.Maximum = row_index;
                    progressBar1.Value = 0;
                    break;
                case 11:
                    progressBar1.Value++;
                    break;
                case 12:
                    this.Text = name_file;
                    break;
            }
        }

        private void Form1_Load(object sender, EventArgs e) {
            backgroundWorker1.RunWorkerAsync();
        }

        private void button2_Click(object sender, EventArgs e) {
            //StreamWriter swOut = new StreamWriter(name_file.Replace(".csv", "_log.csv"), true);
            string[] save = new string[dataGridView1.RowCount];
            for (int i = 0; i < dataGridView1.RowCount; i++) {
                //string data = "";
                //try { data += dataGridView1.Rows[i].Cells[0].Value.ToString() + ","; } catch { }
                //try { data += dataGridView1.Rows[i].Cells[1].Value.ToString() + ","; } catch { }
                //try { data += dataGridView1.Rows[i].Cells[2].Value.ToString() + ","; } catch { }
                //try { data += dataGridView1.Rows[i].Cells[3].Value.ToString() + ","; } catch { }
                //try { data += dataGridView1.Rows[i].Cells[4].Value.ToString() + ","; } catch { }
                //try { data += dataGridView1.Rows[i].Cells[5].Value.ToString() + ","; } catch { }
                //try { data += dataGridView1.Rows[i].Cells[6].Value.ToString() + ","; } catch { }
                //swOut.WriteLine(data);


                try { save[i] += dataGridView1.Rows[i].Cells[0].Value.ToString() + ","; } catch { }
                try { save[i] += dataGridView1.Rows[i].Cells[1].Value.ToString() + ","; } catch { }
                try { save[i] += dataGridView1.Rows[i].Cells[2].Value.ToString() + ","; } catch { }
                try { save[i] += dataGridView1.Rows[i].Cells[3].Value.ToString() + ","; } catch { }
                try { save[i] += dataGridView1.Rows[i].Cells[4].Value.ToString() + ","; } catch { }
                try { save[i] = save[i].Replace("ผ่านขั้นตอน FCT1 แล้ว", "Passed the FCT1 procedure"); } catch { }
                try { save[i] = save[i].Replace("สถานะ", "stetus"); } catch { }
                try { save[i] += dataGridView1.Rows[i].Cells[5].Value.ToString() + ","; } catch { }
                try { save[i] += dataGridView1.Rows[i].Cells[6].Value.ToString() + ","; } catch { }
            }
            File.WriteAllLines(name_file.Replace(".csv", "_log.csv"), save);
            MessageBox.Show("Save OK");
        }
    }
}
