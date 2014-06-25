using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Configuration;
using System.Xml;

namespace ProcDoc
{
    class Read
    {
        #region Intialize variables
        IEnumerable<string> files; int index = 0; bool manualXMLonly = false;
        string OutputFilePath = Environment.CurrentDirectory + ConfigurationManager.AppSettings["OutputFilePath"].ToString();
        #endregion

        #region CHECKING IF THE PDF FILE IS VALID
        private bool IsValidPdf(string filepath)
        {
            bool Ret = true;

            PdfReader reader = null;

            try
            {
                reader = new PdfReader(filepath);
            }
            catch
            {
                Ret = false;
            }

            return Ret;
        }
        #endregion

        #region splitting, naming, creating directory(s) and saving both type of files i.e pdf and xml

        public int SplitAndSave(string inputPath, string outputPath, string pattern, bool auto, string doctype = "", string docnum = "")
        {
            int NumberOfPages = 0;
            try
            {
                //if the document processing is automatic
                if (auto)
                {
                    //get all the file with searchPattern as *.pdf
                    files = Directory.EnumerateFiles(Path.GetDirectoryName(inputPath), pattern);
                    foreach (var file in files)
                    {
                        //checking if the pdf file is valid
                        if (IsValidPdf(file))
                        {
                            using (PdfReader reader = new PdfReader(file))
                            {
                                for (int pagenumber = 1; pagenumber <= reader.NumberOfPages; pagenumber++)
                                {
                                    //reading the time date stamp in specified format
                                    string timestamp = DateTime.Now.ToString("HHmmss_ddMMyyyy");

                                    //get the filepath to put the output files
                                    string finalOutputPath = Environment.CurrentDirectory + outputPath;

                                    //reading the pdf content,storing in string
                                    string thePage = PdfTextExtractor.GetTextFromPage(reader, pagenumber);

                                    //splitting with line ending and storing in string array
                                    string[] lines = thePage.Split('\n');

                                    //removing the whitespaces
                                    lines = lines.Where(line => !string.IsNullOrWhiteSpace(line)).ToArray();

                                    // reading the doc type in the first line of the array
                                    thePage = lines[0].ToString();
                                    if (!manualXMLonly)
                                    {
                                        //counting till colon is encountered
                                        index = thePage.IndexOf(":");

                                        //reading string till colon to take out the document type
                                        string docType = thePage.Substring(0, index).Trim();

                                        //reading string after colon to get the document number
                                        string docNumber = thePage.Substring(index + 1).Trim();

                                        //naming the file without extension as per naming convention 
                                        string filenameWithoutExt = docType + " " + docNumber + " " + timestamp;

                                        ////replacing the whitespaces if any within the filename with underscore
                                        filenameWithoutExt = filenameWithoutExt.Replace(" ", "_");

                                        //if invalid file naming the file starting with ERROR_
                                        if (docType != "Tax Invoice" && docType != "Credit Note")
                                        {
                                            filenameWithoutExt = "ERROR_" + timestamp;
                                            //create the directory to save invalid file
                                            finalOutputPath = CreateDirectories(Environment.CurrentDirectory + "\\Errors", true);
                                            docType = "";
                                        }
                                        else
                                        {
                                            finalOutputPath = CreateDirectories(finalOutputPath + "\\" + docType, false);
                                        }
                                        //create the files in pdf and xml files

                                        CreatePDF(reader, pagenumber, finalOutputPath, filenameWithoutExt);
                                        CreateXML(docType, docNumber, timestamp, lines, finalOutputPath);
                                    }
                                    else
                                    {
                                        doctype = doctype.Replace(" ", "_");
                                        //finalOutputPath = CreateDirectories(finalOutputPath + "\\" + doctype, false);
                                        CreateXML(doctype, docnum, timestamp, lines, outputPath);
                                    }
                                }
                                NumberOfPages = reader.NumberOfPages;
                            }
                        }

                    }
                }
                //Manual processing of invalid files
                else
                {
                    ManualProcess(inputPath, doctype, docnum);
                }
            }
            catch (NullReferenceException ex)
            {
                ErrorLogging.Call_Log(ex, false);
            }
            return NumberOfPages;
        }
        #endregion

        #region MANUAL PROCESSING OF INVALID FILES
        private void ManualProcess(string sourceFilePath, string documentType, string documentNumber)
        {
            try
            {
                string filetype = ".pdf";
                string getDestinationPath = OutputFilePath + @"\" + documentType;
                string sourcefileName = Path.GetFileNameWithoutExtension(sourceFilePath);
                sourcefileName = documentType + " " + documentNumber + " " + DateTime.Now.ToString("HHmmss.ddMMyyyy");
                sourcefileName = sourcefileName.Replace(" ", "_").Replace(".", "_");
                string finalDestinationPath = CreateDirectories(getDestinationPath, false);
                //COPY THE ERROR FILES WITHIN SPECIFIED PROCESSED FOLDER as pdf
                File.Copy(sourceFilePath, finalDestinationPath + "\\" + sourcefileName + filetype);
                filetype = ".xml"; manualXMLonly = true;
                SplitAndSave(sourceFilePath, finalDestinationPath, Path.GetFileNameWithoutExtension(sourceFilePath) + ".pdf", true, documentType, documentNumber);
            }
            catch (Exception ex)
            {
                ErrorLogging.Call_Log(ex, false);
            }
            finally
            {
                //garbage collector invoked to collect unused resources
                //though not a very good option
                GC.Collect();
                //waiting till unused resources are finalized
                GC.WaitForPendingFinalizers();
                //delete the invalid file from the source location i.e Errors folder
                File.Delete(sourceFilePath);
            }
        }
        #endregion


        #region CONVERT TO XML
        protected string ConvertToXML(Object[] args, string rootName, string elemName)
        {
            string xmlStr = "<" + rootName + ">";
            try
            {
                foreach (Object arg in args)
                {
                    xmlStr += "<" + elemName + ">" + arg.ToString() +
                              "</" + elemName + ">";
                }

                xmlStr += "</" + rootName + ">";
            }
            catch (Exception ex)
            {
                ErrorLogging.Call_Log(ex, false);
            }
            return xmlStr;
        }
        #endregion

        #region CREATE PDF FILES IN SPECIFIED FOLDER LOCATION AND SPECIFIED NAMING CONVENTION
        private void CreatePDF(PdfReader reader, int pagenumber, string finalOutputPath, string filename)
        {
            Document document = new Document();
            filename += ".pdf";
            try
            {
               using( PdfCopy copy = new PdfCopy(document, new FileStream(finalOutputPath + "\\" + filename, FileMode.Create)))
               {

                document.Open();

                copy.AddPage(copy.GetImportedPage(reader, pagenumber));
               }

            }
            catch (Exception ex)
            {
                ErrorLogging.Call_Log(ex, false);
            }
            finally
            {
                document.Close();
            }

        }
        #endregion

        #region CREATE XML FILES IN SPECIFIED FOLDER LOCATION AND SPECIFIED NAMING CONVENTION
        private void CreateXML(string docType, string docNumber, string timestamp, string[] lines, string finalOutputPath)
        {
            string rootname = "root";
            string filename = string.Empty;
            docType = docType.Replace(" ", "_");
            if (docType.Length > 0)
            {
                //naming of file by putting underscore between document type , document number, time date stamp
                // so the file name should document_type_documentnumber_time_date.xml
                filename = docType + "_" + docNumber + "_" + timestamp + ".xml";

                //Create the XmlDocument.
                XmlDocument doc = new XmlDocument();

                try
                {
                    // the xml declaration is recommended, but not mandatory
                    XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
                    XmlElement root = doc.DocumentElement;
                    doc.InsertBefore(xmlDeclaration, root);
                    //set root node name
                    if (docType.Contains("Tax"))
                    {
                        rootname = "TaxInvoice";
                    }
                    else if (docType.Contains("Credit"))
                    {
                        rootname = "CreditNote";
                    }
                    //load the xml file in XmlDocument object
                    doc.LoadXml(ConvertToXML(lines, rootname, "param"));
                    //Save the xml document to a file.
                    doc.Save(finalOutputPath + "\\" + filename);
                }
                catch (Exception ex)
                {
                    ErrorLogging.Call_Log(ex, false);
                }
                finally
                {
                    //release the XmlDocument object
                    doc = null;
                }
            }
        }
        #endregion

        #region CREATE DIRECTORIES SPECIFIED FOLDER LOCATION
        private string CreateDirectories(string path, bool err)
        {
            String newOutputPath = string.Empty;
            DateTime current = DateTime.Now;
            //if the file does not have errors, the file is valid
            if (!err)
                newOutputPath = path + String.Format(@"\{0:yyyy}\{1:MMMM}\{2:dd}", current, current, current);
            else
                newOutputPath = path;
            try
            {
                //create new directory to save respective files
                Directory.CreateDirectory(newOutputPath);
            }
            catch (Exception ex)
            {
                ErrorLogging.Call_Log(ex, false);
            }
            return newOutputPath;
        }
        #endregion

    }
}
