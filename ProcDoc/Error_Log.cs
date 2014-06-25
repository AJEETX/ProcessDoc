using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace ProcDoc
{
    internal static class ErrorLogging
    {
       public static void Call_Log(Exception ex,bool AppLog)
        {
                bool bReturnLog = false;

                ErrorLog.LogFilePath = ErrorLog.GetApplicationPath() + "ErrorLogFile.txt";
                //false for writing log entry to customized text file
                bReturnLog = ErrorLog.ErrorRoutine(AppLog, ex);

                if (false == bReturnLog)
                    MessageBox.Show("Unable to write a log");
        }
    }
}
