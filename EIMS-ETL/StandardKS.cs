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

            //dataset ds created for the DB data storage
            DataSet ds = new DataSet();

            // assgined the package variable into new variable by converting it to string for script execution
            string strcon = Dts.Variables["User::ConnectionString"].Value.ToString();
            string localPath = Dts.Variables["User::LocalPath"].Value.ToString();
            string serverPath = Dts.Variables["User::ServerPath"].Value.ToString();

            /*Query to retrieve the entries for PDF generation from PDF table,using this query we remove the CA Reports,wider reports,
             PDF entries which are made for their respective txt entries. This will fetch the Non-CA entries along with CA Internal Reports,because CA internal
            reports also processed only by EIMS ETL not by CA ETLs also the pagination for CA Internal Reports is PAGE: not PAGE as we have for CA Reports.
            In this query PDF is the base table in join with State_dim table and left join with meta data table and Non std pdf table.
            In case if we dont have entry for the state in state_dim table it will not process the PDF but if we dont have entry in meta data table
            for reports it will generate pdf for those reports*/
            string query = @"SELECT distinct FLE.RPT_NME,REPLACE(FLE.FILENAME1,'" + localPath + "','" + serverPath + "') AS SRC_PATH,isnull(CASE WHEN FLE.LVL1 = 'CA' AND FLE.RPT_CAT LIKE 'FIS%' THEN 'PAGE:' ELSE SH.PAGE_PARSER_STRING  END ,isnull(META.PAGE_SPLIT,'')) PAGE_PARSER_STRING,isnull(CLMN_HDR1,'') CLMN_HDR1,	isnull(CLMN_HDR2,'') CLMN_HDR2,	isnull(CLMN_HDR3,'') CLMN_HDR3,	isnull(CLMN_HDR4,'') CLMN_HDR4 FROM TB_SPF_FILE_SORT_PDF FLE JOIN TB_SRC_STATE_DIM SH ON FLE.LVL1=SH.STATE_CDE left JOIN TB_SRC_RPT_METADATA_DTL META on META.RPT_NME=left(FLE.RPT_NME,(len(FLE.RPT_NME)-7))  left join (select distinct RPT_NME from TB_EBT_NON_STD_RPT) A on FLE.RPT_NME LIKE '%'+A.RPT_NME+'%' WHERE (FLE.LVL1 != 'CA' OR (FLE.LVL1 ='CA' AND FLE.RPT_CAT LIKE 'FIS%')) and CONCAT(FLE.LVL1,LVL2,LVL3,FLE.RPT_NME) NOT IN(SELECT CONCAT(LVL1,LVL2,LVL3,RPT_NME) FROM TB_SPF_FILE_SORT_PDF where FILE_EXT = 'PDF') AND FLE.FILE_EXT = 'TXT' and A.RPT_NME is null;";



            /*The above retrieved entries will be Written in the DSrowCount.txt, but it will be overwritten in upcoming logics again with only the
            Count instead of its complete entry*/
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

            /* The dataset stored the entries into it using the above connection, and now we are checking whether we have any entries to process by
             checking the dataset variable table[0] (means table 1) count as below*/

            if (ds.Tables[0].Rows.Count > 0)
            {
                //using foreach loop we are processing by rows by rows (entries by entries)
                foreach (DataRow dr in ds.Tables[0].Rows)
                {

                    //we are storing all the column data into another variable by converting it to string
                    string textfilefullpath = dr["SRC_PATH"].ToString();
                    //string pdfsavefullpath = Dts.Variables["User::DestPath"].Value.ToString();


                    string page_split = dr["PAGE_PARSER_STRING"].ToString();
                    string header1 = dr["CLMN_HDR1"].ToString();
                    string header2 = dr["CLMN_HDR2"].ToString();
                    string header3 = dr["CLMN_HDR3"].ToString();
                    string header4 = dr["CLMN_HDR4"].ToString();
                    //string datafinder = "#";
                    //string page_footer = "";
                    //string date_footer = "";
                    //string emptyLine = "";

                    /*As we mentioned before we are overwritting the DSrowCount file with total count, In order to check how many no of files processed
                    by every ETL run we can check this file*/

                    File.WriteAllText("F:\\ETL_Package\\DSrowCount.txt", ds.Tables[0].Rows.Count.ToString());

                    //using datetime.now function we are getting current datatime in given format and storing it in the datetime variable
                    string datetime = DateTime.Now.ToString("yyyyMMddHHmmss");

                    try
                    {
                        //All the below varibles are for PDF generation which will be used below during pdf generation
                        string line = null;
                        double yPoint = 11;
                        double yPointpage = 11;
                        double yPointincrease = 9;
                        double yPointheader = 12;
                        double rowspacing = 26;
                        //double yPointdata = 10;
                        double yPointdatafinder = 9.35;
                        int width = 50;

                        /*In this if condition it will check the file existence, using the SRC Path assigned to this variable, it will check from DB server
                        via share path of application server */
                        if (File.Exists(textfilefullpath))
                        {
                            // To create a new pdf document, we created the PDF Document object
                            PdfDocument pdf = new PdfDocument();

                            //We are adding a empty page in pdf document using Addpage method
                            PdfPage pdfPage = pdf.AddPage();

                            //using textreader class we are creating an var called readfile with variable of SRC_PATH
                            System.IO.TextReader readFile = new System.IO.StreamReader(textfilefullpath);

                            /*Below method used to get the file without extension and stored it in filename variable and the next variable 
                            will store the only the directory path, these two variables is used to save the pdf file in the same Report folder
                            which has its text file*/
                            string filename = Path.GetFileNameWithoutExtension(textfilefullpath);
                            string pdfsavefullpath = System.IO.Path.GetDirectoryName(textfilefullpath);


                            /*In below using Readlines method we are reading all the lines in textfile for each entry using the src path assigned in 
                             textfilefullpath variable, along with that using regular expression we are catching the pagination in text file
                            and counting the no.of. Pages available in that text file and storing it in pdNbr variable*/
                            int count = File.ReadLines(textfilefullpath).Select(lin => Regex.Matches(lin, @"(?i)\b" + page_split + "\b").Count).Sum();
                            int pgNbr = File.ReadLines(textfilefullpath).Select(lin => Regex.Matches(lin, @"(?i)" + page_split + "").Count).Sum();


                            /* Using pageLayout enum we have given pdf layout as Singlepage
                            In General, one page contains two pages from the source document If the number of pages of the source document
                            can not be divided by 4, the first pages of the output document will
                            each contain only one page from the source document.*/

                            pdf.PageLayout = PdfPageLayout.SinglePage;

                            //Xgraphics used to draw the string on pdf file
                            XGraphics graph = XGraphics.FromPdfPage(pdfPage);

                            //using Xsize class creating size, footersize of pdf and assigning a value to it in below
                            XSize size = new XSize();
                            XSize footersize = new XSize();

                            footersize.Width = XUnit.FromInch(11);
                            footersize.Width = XUnit.FromInch(2);

                            /*Here we have given the Width as 11inch (which is equal to 133 character per line) and height of the page as 8.50 inch
                            Which is 56 lines per page,data after than will get truncated from the page*/
                            size.Width = XUnit.FromInch(11);
                            size.Height = XUnit.FromInch(8.50);

                            pdfPage.Width = size.Width;
                            pdfPage.Height = size.Height;
                            pdfPage.TrimMargins.Top = XUnit.FromInch(0);
                            pdfPage.TrimMargins.Bottom = XUnit.FromInch(0);
                            pdfPage.TrimMargins.Right = XUnit.FromInch(0);
                            pdfPage.TrimMargins.Left = XUnit.FromInch(0);

                            //Writing Text as "Entered into try " on DSrow.txt
                            File.WriteAllText("F:\\ETL_Package\\DSrow.txt", "Entered into try ");

                            //Font and footer font to draw in pdf file
                            XFont font = new XFont("Courier New", 9, XFontStyle.Regular);
                            XFont font_footer = new XFont("Arial", 8, XFontStyle.Italic);

                            int COUNTER = 0;
                            int Space = 1;

                            while (true)
                            {
                                /* using Readline method we are reading the 1st line of the txt file, to draw the string in PDF file,
                                For the second loop it will read the second line of the text line*/
                                line = readFile.ReadLine();

                                /* the below if condition is the break condition of the above while loop, this condition will become True 
                                 when the txt file comes to its Last Page*/
                                if (line == null)
                                {
                                    line = pgNbr.ToString();

                                    COUNTER = 0;
                                    Space = 1;
                                    break;
                                }

                                /* the below else if condition check the line has page_split in it, it will have so it will enter into this condition
                               and draw the 1st line in the pdf page from text file*/ 
                                else if (line.Contains(page_split))
                                {
                                    /*Once the 1st page is generated, for second page it will get into this loop and create the new pdf
                                     page with the below given graphics then get into its respective loop to draw the line of that page*/
                                    if (COUNTER >= 1)
                                    {

                                        pdfPage = pdf.AddPage();

                                        size.Width = XUnit.FromInch(11);
                                        size.Height = XUnit.FromInch(8.50);
                                        pdfPage.Width = size.Width;
                                        pdfPage.Height = size.Height;
                                        pdfPage.TrimMargins.Top = XUnit.FromInch(0);
                                        pdfPage.TrimMargins.Bottom = XUnit.FromInch(0);
                                        pdfPage.TrimMargins.Right = XUnit.FromInch(0);
                                        pdfPage.TrimMargins.Left = XUnit.FromInch(0);


                                        yPoint = yPointpage;
                                        graph.Dispose();
                                        graph = XGraphics.FromPdfPage(pdfPage);

                                        
                                        Space = 1;
                                    }

                                    //The above if counter condition will not satisfy so directly it comes here and print the 1st line in PDF
                                    
                                    
                                    graph.DrawString(line, font, XBrushes.Black, new XRect(3.5, yPoint, width, pdfPage.Height), XStringFormats.TopLeft);
                                   
                                    /*ypoint increment is used to draw the content line by line, if we didnt give y point increment then all
                                    the line of file from text will written in PDF at single line again and again*/
                                    yPoint = yPoint + yPointincrease;

                                    Space++;
                                    COUNTER++;


                                }
                                //For second loop this else block will execute 
                                else

                                {
                                    /*In second loop this if block will execute first, since space is incremented in before loop
                                     so the second line will be draw to pdf using this*/

                                    if (Space == 2)
                                    {

                                        graph.DrawString(line, font, XBrushes.Black, new XRect(3.5, yPoint, width, pdfPage.Height), XStringFormats.TopLeft);
                                        yPoint = yPoint + rowspacing;
                                        //line = "";
                                        //line = line.Insert(0,"\n");
                                        //graph.DrawString(line, font, XBrushes.Black, new XRect(3.5, yPoint, width, pdfPage.Height), XStringFormats.TopLeft);
                                        //yPoint = yPoint + rowspacing;
                                        Space++;



                                    }
                                    else
                                    {
                                        /*During the third and fourth loop until main data, which has header section will get into this loop*/
                                        if (line.Contains(header1))
                                        {
                                            graph.DrawString(line, font, XBrushes.Black, new XRect(3.5, yPoint, width, pdfPage.Height), XStringFormats.TopLeft);
                                            /*Since it is the header section, we use increment with ypointheader which has higher value than pointincrease
                                             variable, the reason is the header section should has extra Y space from previous line from others*/
                                            
                                            yPoint = yPoint + yPointheader;
                                            Space++;
                                        }
                                        else if (line.Contains(header2) || line.Contains(header3) || line.Contains(header4))
                                        {
                                            graph.DrawString(line, font, XBrushes.Black, new XRect(3.5, yPoint, width, pdfPage.Height), XStringFormats.TopLeft);
                                            yPoint = yPoint + yPointincrease;
                                            Space++;
                                        }
                                       
                                        //The Main data will drawn using the below else loop 
                                        else


                                        {
                                            graph.DrawString(line, font, XBrushes.Black, new XRect(3.5, yPoint, width, pdfPage.Height), XStringFormats.TopLeft);
                                            yPoint = yPoint + yPointdatafinder;
                                            Space++;
                                        }

                                    }

                                }

                            }

                            /*below variable define generated pdf file save path which we have derived before  */
                            string pdfFilename = pdfsavefullpath + "\\" + filename + ".PDF";
                            pdf.Save(pdfFilename);

                            //SQL Query to make PDF entries for the generated pdf files.
                            string insertquery = @"INSERT INTO TB_SPF_FILE_SORT_PDF SELECT '" + pdfFilename.Replace(serverPath, localPath) + "',GENERATED_DATE,GENERATED_TIME,DATETIME,'PDF',LVL1,LVL2,LVL3,RPT_CAT,RPT_NME,GETDATE(),IS_ARCHIVED,VER_NBR FROM TB_SPF_FILE_SORT_PDF  WHERE FILENAME1= '" + textfilefullpath.Replace(serverPath, localPath) + "';DELETE FROM TB_SPF_FILE_SORT_PDF WHERE FILENAME1 =''  AND FILE_EXT='PDF';";

                            //We will write the above query into DSinsert text file for each entry execution
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


                        else
                        {
                            /*In case if the file not exist for any entries it will create the error log file here along with time,
                           in some case the pdf will not generate that case we can check here whether that file is available for not  */

                            string Log_Folder = "F:\\ETL_Package\\";
                            
                            using (StreamWriter sw = File.CreateText(Log_Folder +
                                "ErrorLog_" + datetime + ".log"))
                            {
                                sw.WriteLine("File Not Available:" + textfilefullpath);
                                Dts.TaskResult = (int)ScriptResults.Failure;

                            }
                        }
                    }
                    catch (Exception exception)
                    {
                        //In case any exception occurs during the pdf generation then it will catch here
                        string Log_Folder = "F:\\ETL_Package\\";

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