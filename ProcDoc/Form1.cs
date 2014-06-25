using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Windows.Forms;
using System.Configuration;
using System.IO;
using System.Drawing;
using System.Text;
using System.Diagnostics;
using Validator;
namespace ProcDoc
{
    public partial class Form1 : Form
    {
        #region Intialize variables
        NumericTextBox numericTextBox1 = new NumericTextBox();
        string searchPattern = ConfigurationManager.AppSettings["searchPattern"].ToString();
        string outputPath = ConfigurationManager.AppSettings["OutputFilePath"].ToString();
        //assign the path of xml files to read from configuration file
        string InputFilePath = Environment.CurrentDirectory + ConfigurationManager.AppSettings["InputFilePath"].ToString();

        #endregion

        public Form1()
        {
            InitializeComponent();
            // Create an instance of NumericTextBox.
            numericTextBox1.Parent = this.groupBox1;
            //segt back color
            numericTextBox1.BackColor = Color.White;
            //Draw the bounds of the NumericTextBox.
            numericTextBox1.Bounds = new Rectangle(155, 55, 250, 100);
            //set default value for document number
            numericTextBox1.Text = "0000000000";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            pictureBox1.Visible = true;
            //Automated document processing so set the first parameter value as true
            ProcessDocument(true, textBox2.Text);
        }

        private void ProcessDocument(bool auto, string filename = "", string docType = "", string documentNumber = "")
        {
            try
                {

                    Read r = new Read();
                    int pagesRead = r.SplitAndSave(filename, outputPath, searchPattern, auto, docType , documentNumber);
                    pictureBox1.Visible = false;
                //after automatic document processing
                if(auto)
                    MessageBox.Show("No. of Pages PROCESSED = " + pagesRead, "AUTOMATED PROCESS", MessageBoxButtons.OK);
                else
                    //after manual processing of document
                    MessageBox.Show("Manual processing completed", "MANUAL PROCESS", MessageBoxButtons.OK);
                //open
                    Process.Start(Directory.GetCurrentDirectory() + outputPath);
                    Application.Exit();
                }
                catch (Exception ex)
                {
                    // this is the Form1 class
                    foreach (var control in this.Controls)
                    {
                        if (control is CheckBox && (((CheckBox)control).Checked))
                        {
                            //Custom loggin
                            if (((CheckBox)control).Name == "checkBox1")
                                ErrorLogging.Call_Log(ex, !(((CheckBox)control).Checked));
                            //App logging
                            else if (((CheckBox)control).Name == "checkBox2")
                                ErrorLogging.Call_Log(ex, (((CheckBox)control).Checked));
                        }
                    }
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            UpdateSetting("CustomLog", checkBox1.Checked.ToString());
        }
        
        //Update the configuration setting for logging ON/OFF
        private void UpdateSetting(string key, string value)
        {

            Configuration configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            configuration.AppSettings.Settings[key].Value = value;
            try
            {
                configuration.Save();
            }
            catch (Exception ex)
            {
                ErrorLogging.Call_Log(ex, false);
            }
            finally
            {
                ConfigurationManager.RefreshSection("appSettings");
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            UpdateSetting("EventLog", checkBox2.Checked.ToString());
        }

        #region Select the pdf file for manual processing
        private void button2_Click(object sender, EventArgs e)
        {
            string folder = @"\Errors";
            OpenFileDialogue(folder,1);
        }
        #endregion

        #region SUBMIT the file for the  manual processing
        private void button3_Click(object sender, EventArgs e)
        {
            string list = listBox1.SelectedItem.ToString();
            string numeric = numericTextBox1.Text.Trim();
            try
            {
                if (list.Length > 0 && numeric.Length > 0)
                {
                    ProcessDocument(false, textBox1.Text, list, numeric);
                }
                else if ((string.IsNullOrEmpty(list) || string.IsNullOrEmpty(numeric)) ? true : false)
                    {
                        MessageBox.Show("InComplete Submit,\n Closing Application", "InComplete Submit !!!", MessageBoxButtons.OK);
                        Application.Exit();
                    }
            }
            catch (Exception ex)
            {
                ErrorLogging.Call_Log(ex, false);
            }
            finally
            {
                MessageBox.Show("Document Processed successfully.", "Processing closing...", MessageBoxButtons.OK);
                Application.Exit();
            }
        }

        #endregion

        void OpenFileDialogue(string folder,int textboxNo)
        {
            OpenFileDialog theDialog = new OpenFileDialog();
            theDialog.Title = "Open pdf File";
            theDialog.Filter = "pdf files|*.pdf";
            theDialog.InitialDirectory = Directory.GetCurrentDirectory() + folder;
            if ((theDialog.ShowDialog() == DialogResult.OK) && (theDialog.OpenFile()) != null)
            {
                try
                {
                    //FOR MANUAL PROCESSING
                    if (textboxNo == 1)
                    {
                        textBox1.Text = theDialog.FileName;
                    }
                    //FOR AUTOMATED PROCESS
                    else
                    {
                        textBox2.Text = theDialog.FileName;

                        button1.Enabled = true;
                    }
                }
                catch (Exception ex)
                {
                    ErrorLogging.Call_Log(ex, false);
                }
            }
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            string folder = @"\Input";
            OpenFileDialogue(folder,2);
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (textBox1.Text.Length > 0 && textBox1.Text.EndsWith("pdf") && !string.IsNullOrWhiteSpace(numericTextBox1.Text))
                button3.Enabled = true;
        }
    }
}
