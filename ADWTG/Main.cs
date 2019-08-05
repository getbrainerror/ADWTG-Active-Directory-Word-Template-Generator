using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop;

namespace ADWTG
{
    public partial class Main : Form
    {
        private string TemplatePath;
        private string SavePath;
        public Main()
        {
            InitializeComponent();
        }

        private void Main_Load(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.DomainServer == String.Empty)
            {
                MessageBox.Show("Please Conigure this Application before using it.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
        }


        private void GenerateTemplate(string savePath, string templatePath)
        {
            //Yeah not pretty but ¯\_(ツ)_/¯
            try
            {
                //Support for Enviroment Vars like %username%
                object saveAsObj = Environment.ExpandEnvironmentVariables(savePath);
                templatePath = Environment.ExpandEnvironmentVariables(templatePath);

                DirectoryEntry de = new DirectoryEntry(Properties.Settings.Default.DomainServer);
                // Authentication details
                de.AuthenticationType = AuthenticationTypes.FastBind;
                DirectorySearcher DirectorySearcher = new DirectorySearcher(de);
                DirectorySearcher.ClientTimeout = TimeSpan.FromSeconds(30);
                // load all properties
                DirectorySearcher.PropertiesToLoad.Add("*");
                //Use Current logged in SamAccountName as filter            
                DirectorySearcher.Filter = "(sAMAccountName=" + UserPrincipal.Current.SamAccountName + ")";
                SearchResult result = DirectorySearcher.FindOne(); // There should only be one entry
                if (result != null)
                {
                    //Word init
                    //Source: https://www.techrepublic.com/blog/how-do-i/how-do-i-modify-word-documents-using-c/
                    object missing = System.Reflection.Missing.Value;
                    Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                    Microsoft.Office.Interop.Word.Document document = null;
                    if (File.Exists(templatePath))
                    {
                        object file = (object)templatePath;
                        object readOnly = true;
                        object isVisisble = false;
                        document = wordApp.Documents.Open(ref file, ref missing,
                            ref readOnly, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref isVisisble, ref missing, ref missing,
                            ref missing, ref missing
                            );

                        document.Activate();



                        //Loop through each properties

                        foreach (string propname in result.Properties.PropertyNames)
                        {
                            foreach (Object objValue in result.Properties[propname])
                            {
                                this.FindAndReplace(wordApp, "<" + propname.ToLower() + ">", objValue);
                            }
                        }
                        document.SaveAs2(ref saveAsObj, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLTemplate);
                        document.Close();
                    }
                    else
                    {
                        MessageBox.Show("Couldn't find Template File " + templatePath, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }


                }
                else
                {
                    //Didnt find anything, which should be impossible but safe is safe
                    MessageBox.Show("Couldn't find any entries for " + UserPrincipal.Current.SamAccountName, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Unknown Error: " + e.Message , "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);                
            }
                    
        }

        private void FindAndReplace(Microsoft.Office.Interop.Word.Application WordApp, object findText, object replaceVithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchVildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object replace = 2;
            object wrap = 1;

            WordApp.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord, ref matchVildCards, ref matchSoundsLike,
                ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceVithText, ref replace, ref matchKashida,
                ref matchDiacritics, ref matchAlefHamza, ref matchControl);
            }

        private void Main_Shown(object sender, EventArgs e)
        {
            string[] args = Environment.GetCommandLineArgs();
            if (args.Length == 3)
            {
                //1 arg == filename
                //2 arg == templatePath
                //3 args == savePath
                TemplatePath = args[1];
                SavePath = args[2];
                backgroundWorker.RunWorkerAsync();
            } else {				
				 MessageBox.Show("Usage: adwtg.exe templatepath savepath" + e.Message , "Info", MessageBoxButtons.OK, MessageBoxIcon.Error);
				 this.Close();
				 
			}
        }

        private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            GenerateTemplate(SavePath, TemplatePath);
        }

        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.Close();
        }
    }
}
