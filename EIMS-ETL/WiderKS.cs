#region Help:  Introduction to the script task
/* The Script Task allows you to perform virtually any operation that can be accomplished in
 * a .Net application within the context of an Integration Services control flow. 
 * 
 * Expand the other regions which have "Help" prefixes for examples of specific ways to use
 * Integration Services features within this script task. */
#endregion


#region Namespaces
using System;
using System.Data;
using Microsoft.SqlServer.Dts.Runtime;
using System.Windows.Forms;
using PdfLib;
using PdfSharp;
using PdfSharp.Windows;
using PdfSharp.Charting;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.IO;
using System.Data.SqlClient;
using System.Linq;
using System.Text.RegularExpressions;
#endregion

namespace ST_5eaa059b187349eead0b26efa46c6f3e
{
    /// <summary>
    /// ScriptMain is the entry point class of the script.  Do not change the name, attributes,
    /// or parent of this class.
    /// </summary>
	[Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute]
    public partial class ScriptMain : Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase
    {
        #region Help:  Using Integration Services variables and parameters in a script
        /* To use a variable in this script, first ensure that the variable has been added to 
         * either the list contained in the ReadOnlyVariables property or the list contained in 
         * the ReadWriteVariables property of this script task, according to whether or not your
         * code needs to write to the variable.  To add the variable, save this script, close this instance of
         * Visual Studio, and update the ReadOnlyVariables and 
         * ReadWriteVariables properties in the Script Transformation Editor window.
         * To use a parameter in this script, follow the same steps. Parameters are always read-only.
         * 
         * Example of reading from a variable:
         *  DateTime startTime = (DateTime) Dts.Variables["System::StartTime"].Value;
         * 
         * Example of writing to a variable:
         *  Dts.Variables["User::myStringVariable"].Value = "new value";
         * 
         * Example of reading from a package parameter:
         *  int batchId = (int) Dts.Variables["$Package::batchId"].Value;
         *  
         * Example of reading from a project parameter:
         *  int batchId = (int) Dts.Variables["$Project::batchId"].Value;
         * 
         * Example of reading from a sensitive project parameter:
         *  int batchId = (int) Dts.Variables["$Project::batchId"].GetSensitiveValue();
         * */

        #endregion

        #region Help:  Firing Integration Services events from a script
        /* This script task can fire events for logging purposes.
         * 
         * Example of firing an error event:
         *  Dts.Events.FireError(18, "Process Values", "Bad value", "", 0);
         * 
         * Example of firing an information event:
         *  Dts.Events.FireInformation(3, "Process Values", "Processing has started", "", 0, ref fireAgain)
         * 
         * Example of firing a warning event:
         *  Dts.Events.FireWarning(14, "Process Values", "No values received for input", "", 0);
         * */
        #endregion

        #region Help:  Using Integration Services connection managers in a script
        /* Some types of connection managers can be used in this script task.  See the topic 
         * "Working with Connection Managers Programatically" for details.
         * 
         * Example of using an ADO.Net connection manager:
         *  object rawConnection = Dts.Connections["Sales DB"].AcquireConnection(Dts.Transaction);
         *  SqlConnection myADONETConnection = (SqlConnection)rawConnection;
         *  //Use the connection in some code here, then release the connection
         *  Dts.Connections["Sales DB"].ReleaseConnection(rawConnection);
         *
         * Example of using a File connection manager
         *  object rawConnection = Dts.Connections["Prices.zip"].AcquireConnection(Dts.Transaction);
         *  string filePath = (string)rawConnection;
         *  //Use the connection in some code here, then release the connection
         *  Dts.Connections["Prices.zip"].ReleaseConnection(rawConnection);
         * */
        #endregion


        /// <summary>
        /// This method is called when this script task executes in the control flow.
        /// Before returning from this method, set the value of Dts.TaskResult to indicate success or failure.
        /// To open Help, press F1.
        /// </summary>
        public void Main()
        {

            DataSet ds = new DataSet();
            DataSet pkgRunStat = new DataSet();
            string strcon = Dts.Variables["User::ConnectionString"].Value.ToString();
            string localPath = Dts.Variables["User::LocalPath"].Value.ToString();
            string serverPath = Dts.Variables["User::ServerPath"].Value.ToString();

            bool isError = false;

            //In below we are storing the before day datetime into variable
            DateTime today = DateTime.Today.AddDays(-1);

            //We are storing todays date by converting it into string in below variable
            string str_dt = today.ToString();

            //using below query we fetch only the wider report entries from the temporary table to generate PDFs
            string query = @"select  distinct FLE.RPT_NME,FLE.FILENAME1, FLE.GENERATED_DATE, REPLACE(FLE.FILENAME1,'" + localPath + "','" + serverPath + "') AS SRC_PATH,  isnull(CASE WHEN FLE.LVL1 = 'CA' AND FLE.RPT_CAT LIKE 'FIS%' THEN 'PAGE:' ELSE SH.PAGE_PARSER_STRING  END ,META.PAGE_SPLIT) PAGE_PARSER_STRING ,a.RPT_HEIGHT,a.RPT_LENGTH,a.RPT_WIDTH ,isnull(CLMN_HDR1,'') CLMN_HDR1,	isnull(CLMN_HDR2,'') CLMN_HDR2,	isnull(CLMN_HDR3,'') CLMN_HDR3,	isnull(CLMN_HDR4,'') CLMN_HDR4  from TB_EBT_NON_STD_RPT A join TB_SPF_FILE_SORT_PDF FLE ON A.STATE_CDE=FLE.LVL1  JOIN TB_SRC_STATE_DIM SH ON FLE.LVL1=SH.STATE_CDE  left JOIN TB_SRC_RPT_METADATA_DTL META on META.RPT_NME=left(FLE.RPT_NME,(len(FLE.RPT_NME)-7)) WHERE (FLE.LVL1 != 'CA' OR (FLE.LVL1 ='CA' AND FLE.RPT_CAT LIKE 'FIS%')) and CONCAT(FLE.LVL1,LVL2,LVL3,FLE.RPT_NME) NOT IN(SELECT CONCAT(LVL1,LVL2,LVL3,RPT_NME) FROM TB_SPF_FILE_SORT_PDF where FILE_EXT = 'PDF')  AND FLE.FILE_EXT = 'TXT'  and FLE.RPT_NME LIKE '%'+A.RPT_NME+'%' ";

            //We are writing all the entries into DSrowCount text file, the same file will be overwritten again in below logics
            File.WriteAllText("F:\\ETL_Package\\DSrowCount.txt", query);

            //Database connections
            using (SqlConnection conn = new SqlConnection(strcon))
            {
                conn.Open();
                SqlCommand command = new SqlCommand(query, conn);
                SqlDataAdapter adapter = new SqlDataAdapter(command);
                adapter.Fill(ds);
                conn.Close();
            }

            //checking the dataset table[0] whether we have fetched any rows using above query,if we have then it will ges inside the loop to generate PDF
            if (ds.Tables[0].Rows.Count > 0)
            {
                //Using foreach loop fetching 
                foreach (DataRow dr in ds.Tables[0].Rows)
                {

                    // converting the variable to string and storing in another variable for script task use
                    string textfilefullpath = dr["SRC_PATH"].ToString();


                    string page_split = dr["PAGE_PARSER_STRING"].ToString();
                    string header1 = dr["CLMN_HDR1"].ToString();
                    string header2 = dr["CLMN_HDR2"].ToString();
                    string header3 = dr["CLMN_HDR3"].ToString();
                    string header4 = dr["CLMN_HDR4"].ToString();
                    

                    // As mentioned above there the DSrowCount file is overwritten with its counts instead of its original entries
                    File.WriteAllText("F:\\ETL_Package\\DSrowCount.txt", ds.Tables[0].Rows.Count.ToString());

                    string datetime = DateTime.Now.ToString("yyyyMMddHHmmss");
                    //entered into try block
                    try
                    {
                        //All the below variable are used for the PDF generation for line space, width and height of the page and etc
                        string line = null;
                        double yPoint = 11;
                        double yPointpage = 11;
                        double yPointincrease = 9;
                        double yPointheader = 12;
                        double rowspacing = 26;
                        double yPointdatafinder = 9.35;
                        double PageHeight = Convert.ToDouble(dr["RPT_HEIGHT"].ToString());
                        double PageWidth = Convert.ToDouble(dr["RPT_WIDTH"].ToString());
                        int width = 50;

                        //checking whether the file available for the entries
                        if (File.Exists(textfilefullpath))
                        {
                            //creating pdf file
                            PdfDocument pdf = new PdfDocument();

                            //adding a first page to it
                            PdfPage pdfPage = pdf.AddPage();

                            //reading the complete file data which is avaiable in path using textfilefullpath variable
                            System.IO.TextReader readFile = new System.IO.StreamReader(textfilefullpath);

                            //in below we are getting the directoryname and file name without extension to save the pdf files later after generation
                            string filename = Path.GetFileNameWithoutExtension(textfilefullpath);
                            string pdfsavefullpath = System.IO.Path.GetDirectoryName(textfilefullpath);

                            /*In below, we are reading only the 1st line of the file, and using regular expression we are counting the no.of pages
                            available in that page , in simple terms calculation the pages in source file*/
                            int count = File.ReadLines(textfilefullpath).Select(lin => Regex.Matches(lin, @"(?i)\b" + page_split + "\b").Count).Sum();
                            int pgNbr = File.ReadLines(textfilefullpath).Select(lin => Regex.Matches(lin, @"(?i)" + page_split + "").Count).Sum();


                            //giving layout design to display
                            pdf.PageLayout = PdfPageLayout.SinglePage;

                            // creating graph object to draw the content to pdf file
                            XGraphics graph = XGraphics.FromPdfPage(pdfPage);

                            //in below we declaring the width, height and other margin sizes for PDF
                            XSize size = new XSize();
                            XSize footersize = new XSize();

                            footersize.Width = XUnit.FromInch(11);
                            footersize.Width = XUnit.FromInch(2);

                            size.Width = XUnit.FromInch(PageWidth);
                            size.Height = XUnit.FromInch(PageHeight);
                            pdfPage.Width = size.Width;
                            pdfPage.Height = size.Height;
                            pdfPage.TrimMargins.Top = XUnit.FromInch(0);
                            pdfPage.TrimMargins.Bottom = XUnit.FromInch(0);
                            pdfPage.TrimMargins.Right = XUnit.FromInch(0);
                            pdfPage.TrimMargins.Left = XUnit.FromInch(0);

                            //will make a entry to DSrow file that we have entered into try
                            File.WriteAllText("F:\\ETL_Package\\DSrow.txt", "Entered into try ");

                            //font size and footersize declaration
                            XFont font = new XFont("Courier New", 9, XFontStyle.Regular);
                            XFont font_footer = new XFont("Arial", 8, XFontStyle.Italic);


                            int COUNTER = 0;
                            int Space = 1;

                            // given while loop as true to initialize the content drawn to the pdf page line by line
                            while (true)
                            {
                                //reading line by line content from source file to draw the same in pdf file
                                line = readFile.ReadLine();

                                /*the below if loop is used to break the loop once we attain the end of the page,we have the empty line at end of the
                                page using that we will get into this loop*/
                                if (line == null)
                                {
                                    line = pgNbr.ToString();

                                    COUNTER = 0;
                                    Space = 1;
                                    break;
                                }

                                //The first line into the PDF will be drawn using below else if loop 
                                else if (line.Contains(page_split))
                                {
                                    /*For the second page of the file will get into this if loop and add a new page to pdf and will allocate size for
                                    the pages*/ 
                                    if (COUNTER >= 1)
                                    {
                                      

                                        pdfPage = pdf.AddPage();

                                        size.Width = XUnit.FromInch(PageWidth);
                                        size.Height = XUnit.FromInch(PageHeight);
                                        pdfPage.Width = size.Width;
                                        pdfPage.Height = size.Height;
                                        pdfPage.TrimMargins.Top = XUnit.FromInch(0);
                                        pdfPage.TrimMargins.Bottom = XUnit.FromInch(0);
                                        pdfPage.TrimMargins.Right = XUnit.FromInch(0);
                                        pdfPage.TrimMargins.Left = XUnit.FromInch(0);


                                        yPoint = yPointpage;
                                        graph.Dispose();
                                        graph = XGraphics.FromPdfPage(pdfPage);

                                        //COUNTER++;
                                        Space = 1;
                                    }

                                    graph.DrawString(line, font, XBrushes.Black, new XRect(3.5, yPoint, width, pdfPage.Height), XStringFormats.TopLeft);
                                    yPoint = yPoint + yPointincrease;
                                    Space++;
                                    COUNTER++;


                                }
                                
                                // the second line of the page will be get into this if loop and drawn to pdf
                                else
                                {
                                   
                                    if (Space == 2)
                                    {

                                        graph.DrawString(line, font, XBrushes.Black, new XRect(3.5, yPoint, width, pdfPage.Height), XStringFormats.TopLeft);
                                        yPoint = yPoint + rowspacing;
                                        
                                        Space++;



                                    }
                                    else
                                    {
                                        //the header of the page will be differentiatie using below condition and will be drawn to pdf of every page below

                                        if (line.Contains(header1))
                                        {
                                            graph.DrawString(line, font, XBrushes.Black, new XRect(3.5, yPoint, width, pdfPage.Height), XStringFormats.TopLeft);
                                            yPoint = yPoint + yPointheader;
                                            Space++;
                                        }
                                        else if (line.Contains(header2) || line.Contains(header3) || line.Contains(header4))
                                        {
                                            graph.DrawString(line, font, XBrushes.Black, new XRect(3.5, yPoint, width, pdfPage.Height), XStringFormats.TopLeft);
                                            yPoint = yPoint + yPointincrease;
                                            Space++;
                                        }
                                        
                                        else

                                        //The rest of the data and main content will be drawn to pdf using this else loop only

                                        {
                                            graph.DrawString(line, font, XBrushes.Black, new XRect(3.5, yPoint, width, pdfPage.Height), XStringFormats.TopLeft);
                                            yPoint = yPoint + yPointdatafinder;
                                            Space++;
                                        }

                                    }

                                }

                            }

                            //here we are giving the directory to save the pdf file
                            string pdfFilename = pdfsavefullpath + "\\" + filename + ".PDF";
                            pdf.Save(pdfFilename);

                            //using this query we are inserting the PDF ENTRIES to the temporary table
                            string insertquery = @"INSERT INTO TB_SPF_FILE_SORT_PDF SELECT '" + pdfFilename.Replace(serverPath, localPath) + "',GENERATED_DATE,GENERATED_TIME,DATETIME,'PDF',LVL1,LVL2,LVL3,RPT_CAT,RPT_NME,GETDATE(),IS_ARCHIVED,VER_NBR FROM TB_SPF_FILE_SORT_PDF  WHERE FILENAME1= '" + textfilefullpath.Replace(serverPath, localPath) + "';DELETE FROM TB_SPF_FILE_SORT_PDF WHERE FILENAME1 =''  AND FILE_EXT='PDF';";

                            File.WriteAllText("F:\\ETL_Package\\DSinsert.txt", insertquery);

                            using (SqlConnection conn = new SqlConnection(strcon))
                            {
                                conn.Open();
                                SqlCommand command = new SqlCommand(insertquery, conn);
                                SqlDataAdapter adapter = new SqlDataAdapter(command);
                                adapter.Fill(ds);
                                conn.Close();
                            }

                        }

                        //If the text file is not available it will enter into this else loop and will generate error log file as file not available
                        else
                        {
                            string Log_Folder = "F:\\ETL_Package\\";
                            // Create Log File for Errors
                            using (StreamWriter sw = File.CreateText(Log_Folder +
                                "ErrorLog_" + datetime + ".log"))
                            {
                                sw.WriteLine("File Not Available:" + textfilefullpath);
                                Dts.TaskResult = (int)ScriptResults.Failure;

                            }
                        }
                    }

                    //if any other issue found then the exception will be catched here and will be written to the different error log file again
                    catch (Exception exception)
                    {
                        string Log_Folder = "F:\\ETL_Package\\";
                        // Create Log File for Errors
                        using (StreamWriter sw = File.CreateText(Log_Folder +
                            "ErrorLog_" + datetime + ".log"))
                        {
                            sw.WriteLine(exception.ToString());
                            Dts.TaskResult = (int)ScriptResults.Failure;

                        }
                    }
                }
            }

            Dts.TaskResult = (int)ScriptResults.Success;
        }

        #region ScriptResults declaration
        /// <summary>
        /// This enum provides a convenient shorthand within the scope of this class for setting the
        /// result of the script.
        /// 
        /// This code was generated automatically.
        /// </summary>
        enum ScriptResults
        {
            Success = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Success,
            Failure = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Failure
        };
        #endregion

    }
}