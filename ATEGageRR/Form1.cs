using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Excel;


namespace ATEGageRR
{
    public partial class Form1 : Form
    {

        #region 參數定義
        string gageDate = string.Empty;
        string Engineer = string.Empty;
        string Model = string.Empty;
        string Machine = string.Empty;
        string excelTemplate = string.Empty;
       // Microsoft.Office.Interop.Excel.Application objExcel;
        

        #endregion



        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Text = "ATE Gage RR Report ,Ver:" + System.Windows.Forms.Application.ProductVersion;
            if (!checkTime("410CD26EBAC024192F32F455445C4834"))
            {
                MessageBox.Show("軟件已過期,請重新申請授權使用!", "Error");
                Environment.Exit(0);
            }
            this.txtExcelTemplate.SetWatermark("雙擊選擇Gage RR Excel 模板");
            this.txtEngineer.SetWatermark("eg:Edward Song");
            this.txtModel.SetWatermark("eg:Cedar-Hsw");
            this.txtMachine.SetWatermark("eg:Cedar-Hsw-1");
            gageDate = getDtpShortValue(dtpDate);           
        }

        #region OpenFile
        /// <summary>
        /// 打開文件,讓textbox顯示文件路徑
        /// </summary>
        /// <param name="textbox"></param>
        public  void OpenFile(System.Windows.Forms.TextBox  textbox)
        {
            textbox.Text = string.Empty;
            OpenFileDialog openfiledialog = new OpenFileDialog();
            openfiledialog.Filter = "Excel 2003 file|*.xls|Excel file|*.xlsx|pdf file|*.pdf|boardview file|*.GR|all file|*.*";
            if (openfiledialog.ShowDialog() == DialogResult.OK)
                textbox.Text = openfiledialog.FileName;
        }
        #endregion

        #region checkTime

        /// <summary>
        /// check time
        /// </summary>
        /// <param name="str">encrpt date</param>
        /// <returns></returns>
        public static bool checkTime(string str)
        {
            DateTime t = Convert.ToDateTime(DES.DesDecrypt(str, "Edward86"));

            if (File.Exists("C:\\Windows\\System32\\" + System.Windows.Forms.Application.ProductName  + ".dll"))
            {
                return false;
            }
            FileInfo[] files = new DirectoryInfo("C:\\").GetFiles();
            int index = 0;
            while (index < files.Length)
            {
                if (DateTime.Compare(files[index].LastAccessTime, t) > 0)
                {
                    FileStream iniStram = File.Create("C:\\Windows\\System32\\" + System.Windows.Forms.Application.ProductName + ".dll");
                    iniStram.Close();
                    return false;
                }
                checked { ++index; }
            }
            return true;

        }
        #endregion

        #region getDtpValue
        /// <summary>
        /// 從DateTimePicker獲取獲取短日期信息
        /// </summary>
        /// <param name="dtp"></param>
        /// <returns></returns>
        public static string getDtpShortValue(DateTimePicker dtp)
        {
            string value = string.Empty;
            try
            {
                value = string.Format("{0:D4}", dtp.Value.Year) + "/" + string.Format("{0:D2}", dtp.Value.Month) + "/" + string.Format("{0:D2}", dtp.Value.Day);

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return value;
        }
        #endregion

        #region KillExcel
        private void KillExcel()
        {
            System.Diagnostics.Process[] excelProcess = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            foreach (System.Diagnostics.Process p in excelProcess)
                p.Kill();
        }
        #endregion


        private void txtExcelTemplate_DoubleClick(object sender, EventArgs e)
        {
            OpenFile(txtExcelTemplate);
        }

        private void btnGo_Click(object sender, EventArgs e)
        {
            
           // checkInput();
            if (!checkInput())
                return;
            this.Enabled = false;
            lstFile.Items.Clear();
           
            //download excelTemplate
            string tempFilePath = System.Windows.Forms.Application.StartupPath + @"\GageRRTemplate.xls";
            if (downLoadFile(tempFilePath))
            {
                excelTemplate = tempFilePath;
            }
            else
            {
                this.Enabled = true;
                return;
            }

            //create folder
            string folderPath = System.Windows.Forms.Application.StartupPath + @"\" + this.txtMachine.Text;
            if (!Directory.Exists(folderPath))
            {
                try
                {
                    Directory.CreateDirectory(folderPath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Create " + folderPath + " Fail,Message:" + ex.Message);
                    this.Enabled = true;
                    return;
                  
                }
            }

           

            do
            {
                gageDate = getDtpShortValue(dtpDate);
                string fileName = folderPath + @"\Gage_R&R_Report_" + Machine + "_" + gageDate.Replace("/", string.Empty) + ".xls";
                createGageRR (gageDate ,Engineer,Model,Machine,fileName );
                dtpDate.Value = dtpDate.Value.AddMonths(3);
                
            } while (dtpDate.Value < DateTime.Now );
            this.Enabled = true;
        }




        private bool createGageRR(string gagedate, string enginner, string model, string machine,string fileName)
        {
            Microsoft.Office.Interop.Excel.Application objExcel = new Microsoft.Office.Interop.Excel.Application();
            Workbook objWorkbook = objExcel.Workbooks.Add(excelTemplate);


            objExcel.Cells[3, 7] = gagedate;
            objExcel.Cells[4, 7] = enginner;
            objExcel.Cells[5, 7] = model;
            objExcel.Cells[6, 7] = machine;

            objExcel.DisplayAlerts = false;
            objExcel.AlertBeforeOverwriting = false;
            objWorkbook.SaveAs(fileName);
            objWorkbook.Close();
            objExcel.SaveWorkspace(fileName);
            objExcel.Quit();
            objExcel = null;
            //objWorkbook.SaveAs(fileName);
            KillExcel();

            // fileName = fileFullPath.Substring(fileFullPath.LastIndexOf("\\") + 1, fileFullPath.Length - fileFullPath.LastIndexOf("\\") - 1);
           // lstFile .Items.Add (fileName.Substring (fileName.LastIndexOf ("\\") +1,fileName.Length -fileName.LastIndexOf ("\\") -1));


            return true;
        }


        private bool  checkInput()
        {
            //check input
            //if (string.IsNullOrEmpty(txtExcelTemplate.Text.Trim()))
            //{
            //    MessageBox.Show("請選擇Excel模板");
            //    return false;
            //}
            //else
            //{
            //    excelTemplate = txtExcelTemplate.Text.Trim();
            //}
            //check engineer
            if (string.IsNullOrEmpty(txtEngineer.Text.Trim()))
            {
                MessageBox.Show("Engineer 不能為空");
                return false;
            }
            else
            {
                Engineer = this.txtEngineer.Text.Trim();
            }
            //check model
            if (string.IsNullOrEmpty(txtModel.Text.Trim()))
            {
                MessageBox.Show("Model 不能為空");
                return false;
            }
            else
            {
                Model = this.txtModel.Text.Trim();
            }
            //check machine
            if (string.IsNullOrEmpty(txtMachine.Text.Trim()))
            {
                MessageBox.Show("Machine 不能為空");
                return false;
            }
            else
            {
                Machine = txtMachine.Text.Trim();
            }
            return true;
        }



        #region getDtpValue
        /// <summary>
        /// 從DateTimePicker獲取獲取短日期信息
        /// </summary>
        /// <param name="dtp"></param>
        /// <returns></returns>
        public static string getDtpShortStringValue(DateTimePicker dtp)
        {
            string value = string.Empty;
            try
            {
                value = string.Format("{0:D4}", dtp.Value.Year) +  string.Format("{0:D2}", dtp.Value.Month) +  string.Format("{0:D2}", dtp.Value.Day);

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return value;
        }
        #endregion

        private void txtModel_TextChanged(object sender, EventArgs e)
        {
            this.txtMachine.Text = this.txtModel.Text;
        }

        private void Form1_DoubleClick(object sender, EventArgs e)
        {

  
                //MessageBox.Show("OK"); 
        }


        private bool downLoadFile(string filePath)
        {

            if (!File.Exists(filePath))
            {
                byte[] template = Properties.Resources.GageRRTemplate;
                FileStream stream = new FileStream(filePath, FileMode.Create);
                try
                {
                    stream.Write(template, 0, template.Length);
                    stream.Close();
                    stream.Dispose();
                    File.SetAttributes(filePath, FileAttributes.Hidden);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Create GageRRTempalte Fail,Message:" + ex.Message);
                    return false;
                }     
          
            }
            return true;
        }
    }
}
