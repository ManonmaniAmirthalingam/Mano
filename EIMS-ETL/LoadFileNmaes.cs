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
using System.IO;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Text;
#endregion

namespace ST_b8d10399ec2d43478b05cae7b6289ba9
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
            // TODO: Add your code here

            string datetime = DateTime.Now.ToString("yyyyMMddHHmmss");
            /* Declaring the datetime variable and assgined the current datetime in the
             string format*/

            try
            {
                //creating the dataset called ds, to store the retrieved data from database
                DataSet ds = new DataSet()

                //Storing the Variable into the new string variable again
                string strcon = Dts.Variables["User::ConnectionString"].Value.ToString();
                string localPath = Dts.Variables["User::LocalPath"].Value.ToString();
                string serverPath = Dts.Variables["User::ServerPath"].Value.ToString();
                
                /*We are creating the array called Str_Filesort_PDF and we directly read the entries from text file,
                each line of the text file will store in each array index*/
                string[] str_Filesort_PDF = System.IO.File.ReadAllLines(@"F:\FileName_PDF.txt");

                //Creating an empty array
                string[] str_Filesort_PDF_SPLIT = new string[] { };

                //creating a Var variable to store all kind of datatype and assigning it to the new created string list
                var str_Filesort_PDF_SPLIT_after = new List<string>();
                string str_Filesort_PDF_DATE = "";
                string str_RPT_Name = "";
                int i = 0;

                if (str_Filesort_PDF.Length > 0)
                {
                    StringBuilder sb = new StringBuilder();
                    foreach (string val_PDF in str_Filesort_PDF)
                    {
                        
                        str_Filesort_PDF_SPLIT = val_PDF.Split('|');
                        /*We are spliting the entry using | demiliter and stroing it in split varibale as below,
                        this split will remove the time date value attached for each entry and store only the complete file path data*/

                       
                        str_Filesort_PDF_SPLIT_after.Add(str_Filesort_PDF_SPLIT[0].ToString());
                        // we are saving the captured value in "str_Filesort_PDF_SPLIT_after" array varibale with the index value

                       
                        str_RPT_Name = str_Filesort_PDF_SPLIT[0].Substring(str_Filesort_PDF_SPLIT[0].LastIndexOf('\\') + 1, ((str_Filesort_PDF_SPLIT[0].Length) - 5) - str_Filesort_PDF_SPLIT[0].LastIndexOf('\\'));
                        /*We are retrieving  the Rpt_name and storing it in "str_RPT_Name" Variable,
                        we use substring method to get the report name, the start index will be the LastIndex of // where it goes to the report name
                       start we will add +1 here to exclude the / before report name.  We give length by calculating as follows 
                       ((str_Filesort_PDF_SPLIT[0].Length) - 5) - str_Filesort_PDF_SPLIT[0].LastIndexOf('\\')   , (total length -5(these -5 
                       will exclude the .txt at end and // at begining of the entry both // will be considered as the single value )) and the by subtracting 
                       the values from start to lastindexof (\\).
                       */


                        str_Filesort_PDF_DATE = str_Filesort_PDF_SPLIT_after[i].ToString();
                        //In above we have assigned the array value 0 of str_Filesort_PDF_SPLIT_after to str_Filesort_PDF_DATE

                        str_Filesort_PDF_DATE = str_Filesort_PDF_DATE.Substring(str_Filesort_PDF_DATE.Length - 11, 11).Substring(0, 7);
                        /* Using 1st substring function we have retrieved only _220217.txt with follow 2nd substring we are capturing only _220217
                        and storing it it "str_Filesort_PDF_DATE" variable*/


                        string converteddate = str_Filesort_PDF_DATE.Substring(1, 2) + "-" + str_Filesort_PDF_DATE.Substring(3, 2) + "-" + str_Filesort_PDF_DATE.Substring(5, 2);
                        /*using substring and operator we are created the manual date format and storing it in the converted date variable
                         the value store lik 22-02-17, Remember the _ operator before date will be neglected by using the substring start index as 1*/
                        
                        DateTime fromDateValue;
                        var formats = new[] { "yy-MM-dd" };
                        //created an array called format with date format

                        if ((DateTime.TryParseExact(converteddate, formats, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out fromDateValue)) && (str_Filesort_PDF_DATE.IndexOf("_") > -1) &&
                        str_Filesort_PDF_SPLIT_after[i].IndexOf("-") > -1 && Regex.Matches(str_RPT_Name, "[a-zA-Z]").Count > 0)
                        /*in this we have used Datetime.TryParseExact Method to check whether the parsing format is correct as we required if it is correct
                         then it will return true else it will return false */

                        {

                            
                            string QUERY = "INSERT INTO TB_SPF_FILE_SORT_PDF(FILENAME1,PRCS_DTE) VALUES(" + "'" + str_Filesort_PDF_SPLIT_after[i].Replace(serverPath, localPath) + "'" + ", GETDATE()" + ")";
                            //inserted query to insert the entry into filenmae column and the prcs_dte column will be updated here with getdate method
                            
                            using (SqlConnection conn = new SqlConnection(strcon))

                            {
                                conn.Open();
                                SqlCommand command = new SqlCommand(QUERY, conn);
                                SqlDataAdapter adapter = new SqlDataAdapter(command);
                                adapter.Fill(ds);
                                conn.Close();
                            }
                        }
                        else
                        {

                            sb.AppendLine(str_Filesort_PDF_SPLIT_after[i].ToString());
                            /*if the format is not as we regularize then it will add that entry to sb variable in stringbuilder,
                             in some case the report name and dateformat will not be in correct format in such case it will be filtered 
                            here and will not be processed further by not making entry in pdf table*/

                        }
                        i++;
                        // using this increment it will follow for all the lines in the file
                    }

                    string Log_Folder = "F:\\ETL_Package\\";
                    // Create Log File for Errors

                    using (StreamWriter sw = File.CreateText(Log_Folder + "BadFiles_" + datetime + ".log"))
                        //Using streamWriter we are creating creating the text in the above folder with the name Badfiles_datetime.log
                        //If file exist it will write on the same file if not it will create the new file
                    {
                        sw.WriteLine(sb);
                        // it will write the data into file from sb Variable (string builder) which will be updated when the file format is not correct

                    }

                }
            }
            catch (Exception exception)
            {

                string Log_Folder = "F:\\ETL_Package\\";
                // Create Log File for Errors
                using (StreamWriter sw = File.CreateText(Log_Folder + "ErrorLog_" + datetime + ".log"))
                {
                    sw.WriteLine(exception.ToString());
                    //In these we will catch the exception and we will write it into the ErrorLog file

                    Dts.TaskResult = (int)ScriptResults.Failure;

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