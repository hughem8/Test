/*  Created By:     Mike Hughes
 *  Created On:     01/19/16
 *  
 *  Purpose:        This program will verify the MAPI plugin is running.
 *  Mods:
 *                  01/19/16 - MTH - Created.
 *    TEST GIT Change
 */


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.IO;
using System.Net.Mail;
using System.Data;
using System.Data.SqlClient;
using System.Threading;
using System.Configuration;

namespace MapiVerification
{
    public class MainMapiVerify
    {
        static void Main(string[] args)
        {
            bool blnSuccess = false;
            //Check to see if an instance is already running.
            var me = Process.GetCurrentProcess();

            var arrProcesses = Process.GetProcessesByName(me.ProcessName);
            if (arrProcesses.Length <= 1)
            {
                VerifyLastHeartBeat objVerifyLastHeartBeat = new VerifyLastHeartBeat();
                blnSuccess = objVerifyLastHeartBeat.CheckMapi();

            }
        }
    }
    
    class VerifyLastHeartBeat
    {
        private string m_strAppName = string.Empty;
        private string m_strEmailServer = string.Empty;
        private string m_strLogFullFilePathName = string.Empty;
        private string m_strLogFilePath = string.Empty;
        private string m_strMailTo = string.Empty;
        private string m_strMailFrom = string.Empty;
        private string m_strMailCC = string.Empty;
        private string m_strMailBody = string.Empty;
        private string m_strMailSubject = string.Empty;
        private string m_strSourceFilePathName = string.Empty;
        private int m_intSleepIntervalMS = 30000; //Default.
        private DateTime m_dteProgramEndDateTime = DateTime.MinValue;
        private string m_strEndTime = string.Empty;
        private int m_intTolerance = 60; //Default.

        public bool CheckMapi()
        {
            string strFrom = "CheckMapi";
            bool blnSuccess = false;
           
            string strMessage = string.Empty;
            string strTemp = string.Empty;

            try
            {

                blnSuccess = GetAppSettings();
                if(blnSuccess)
                {

                    strMessage = "Application Started at " + DateTime.Now.ToLongTimeString();
                    WriteLogMessage(strFrom, "INFO", strMessage);

                    // See if the end time is reached.  if not, process, else log & end program.
                    strTemp = DateTime.Today.ToString("MM/dd/yyyy") + " " + EndTime;
                    m_dteProgramEndDateTime = DateTime.Parse(strTemp);
                    while (DateTime.Now <= m_dteProgramEndDateTime)
                    {

                        CompareMapiFile();

                        strMessage = "Application sleeping...";
                        WriteLogMessage(strFrom, "INFO", strMessage);

                        // pause application according to 'IntervalMilliSec'
                        Thread.Sleep(SleepIntervalMS);
                    }


                }

                return true;
            }
            catch (Exception ex)
            {
                WriteLogMessage(strFrom, ex.Source, ex.Message);
                return false;
            }
        }
        private void CompareMapiFile()
        {
            string strFrom = "CompareMapiFile";
            int intDiffMinutes = 0;
            DateTime dteNow = DateTime.Now;
            DateTime dteFile = DateTime.MinValue;
            string strMessage = string.Empty;

            try
            {
                dteFile = File.GetLastWriteTime(SourceFilePathName);

                intDiffMinutes = dteNow.Subtract(dteFile).Minutes;

                //if((intDiffMinutes>Tolerance)&&(File.Exists(SourceFilePathName + @"\MapiHeartBeat.txt")==true))
                if(intDiffMinutes>Tolerance)
                {
                    MailBody = MailBody.Replace("[DELAY]", intDiffMinutes.ToString());
                    SendEmail();

                   strMessage = "SendEmail() called. Body=" + MailBody;
                }else
                {
                    strMessage = intDiffMinutes.ToString() + " minutes difference were in tolerance = " + Tolerance.ToString() + ". No Action Needed.";
                }
                WriteLogMessage(strFrom, "INFO", strMessage);
            }
            catch (Exception ex)
            {
                WriteLogMessage(strFrom, ex.Source, ex.Message);
            }
        }
        private bool GetAppSettings()
        {
            string strFrom = "GetAppSettings";
            bool blnSucess = false;
            int intTemp = 30000;

            try
            {

                AppName = System.Configuration.ConfigurationManager.AppSettings["AppName"];
                EmailServer = ConfigurationManager.AppSettings["EmailServer"];
                MailTo = ConfigurationManager.AppSettings["MailTo"];
                MailFrom = ConfigurationManager.AppSettings["MailFrom"];
                MailSubject = ConfigurationManager.AppSettings["MailSubject"];
                MailCC = ConfigurationManager.AppSettings["MailCC"];
                MailBody = ConfigurationManager.AppSettings["MailBody"];
                LogFilePath = ConfigurationManager.AppSettings["LogFilePath"];
                blnSucess = Int32.TryParse(ConfigurationManager.AppSettings["SleepIntervalMS"], out intTemp);
                if (intTemp != 0)
                    SleepIntervalMS = intTemp;
                blnSucess = Int32.TryParse(ConfigurationManager.AppSettings["Tolerance"], out intTemp);
                if (intTemp != 0)
                    Tolerance = intTemp;
                EndTime = ConfigurationManager.AppSettings["EndTime"];
                SourceFilePathName = ConfigurationManager.AppSettings["SourceFilePathName"];


                return true;
            }
            catch (Exception ex)
            {
                WriteLogMessage(strFrom, ex.Source, ex.Message);
                return false;
            }

        }
        public bool SendEmail()
        {
            string strFrom = "SendEmail";
            
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient(EmailServer);

                mail.From = new MailAddress(MailFrom);
                mail.To.Add(MailTo);
                mail.CC.Add(MailCC);
                mail.Subject = MailSubject;
                mail.Body = MailBody;
                mail.Priority = MailPriority.High;

                SmtpServer.Send(mail);

                mail.Dispose(); //Dispose of the mail object to prevent attachment file locking.
                return true;
            }
            catch (Exception ex)
            {
                WriteLogMessage(strFrom, ex.Source, ex.Message);
                SendEmailWithAttachment(ex.Message, LogFullFilePathName);
                return false;
            }
        }
        public bool SendEmailWithAttachment(string strEmailBody, string strAttachmentFilePath)
        {
            string strFrom = "SendEmailWithAttachment";

            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient(EmailServer);
                mail.From = new MailAddress(m_strMailFrom);
                mail.To.Add(m_strMailTo);
                mail.Subject = m_strMailSubject;
                mail.Body = strEmailBody;

                System.Net.Mail.Attachment attachment;
                attachment = new System.Net.Mail.Attachment(strAttachmentFilePath);
                mail.Attachments.Add(attachment);

                SmtpServer.Send(mail);
                mail.Dispose(); //Dispose of the mail object to prevent file attachment locking.

                return true;
            }
            catch (Exception ex)
            {
                WriteLogMessage(strFrom, ex.Source, ex.Message);
                SendEmailWithAttachment(ex.Message, LogFullFilePathName);
                return false;
            }
        }
        public bool WriteLogMessage(string strFrom, string strSource, string strMessage)
        {
            string strFullLogFilePath = string.Empty;

            try
            {
                StringBuilder sb = new StringBuilder();

                //Add the Application name and date values to the log file path.
                strFullLogFilePath = LogFilePath + @"\" + DateTime.Today.Year.ToString() + @"\";

                //Create the directory if it doesn't exist.
                if (Directory.Exists(strFullLogFilePath) == false)
                    Directory.CreateDirectory(strFullLogFilePath);

                strFullLogFilePath = strFullLogFilePath + AppName + DateTime.Today.ToString("yyyyMMdd") + ".log";

                LogFullFilePathName = strFullLogFilePath;

                //Use SysDateTime|AppName|From|Source|Message
                sb.AppendFormat("{0,-25}", DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt"));
                sb.AppendFormat("{0,-25}", "|" + AppName);
                sb.AppendFormat("{0,-25}", "|" + strFrom);
                sb.AppendFormat("{0,-25}", "|" + strSource);

                strMessage = strMessage.Replace("\n", "");
                strMessage = strMessage.Replace("\r", "");
                sb.AppendFormat("{0,-255}", "|" + strMessage);

                using (StreamWriter sw = new StreamWriter(strFullLogFilePath, true)) //Create the file if it doesn't exist, and append to it.
                {
                    sw.WriteLine(sb.ToString());
                }

                return true;

            }
            catch
            {
                return false;
            }

        }
        public string AppName
        {
            get
            {
                return m_strAppName;
            }
            set
            {
                m_strAppName = value;
            }
        }

       public int SleepIntervalMS
        {
           get
            {
                return m_intSleepIntervalMS;
            }
           set
            {
                m_intSleepIntervalMS = value;
            }
        }
       public int Tolerance
       {
           get
           {
               return m_intTolerance;
           }
           set
           {
               m_intTolerance = value;
           }
       }
       public string EndTime
       {
           get
           {
               return m_strEndTime;
           }
           set
           {

               m_strEndTime = value;
           }
       }
        public string EmailServer
        {
            get
            {
                return m_strEmailServer;
            }
            set
            {

                m_strEmailServer = value;
            }
        }
        public string SourceFilePathName
        {
            get
            {
                return m_strSourceFilePathName;
            }
            set
            {

                m_strSourceFilePathName = value;
            }
        }
        public string LogFullFilePathName
        {
            get
            {
                return m_strLogFullFilePathName;
            }
            set
            {

                m_strLogFullFilePathName = value;
            }
        }
        public string LogFilePath
        {
            get
            {
                return m_strLogFilePath;
            }
            set
            {

                m_strLogFilePath = value;
            }
        }
        public string MailSubject
        {
            get
            {
                return m_strMailSubject;
            }
            set
            {

                m_strMailSubject = value;
            }
        }
        public string MailTo
        {
            get
            {
                return m_strMailTo;
            }
            set
            {

                m_strMailTo = value;
            }
        }
        public string MailFrom
        {
            get
            {
                return m_strMailFrom;
            }
            set
            {

                m_strMailFrom = value;
            }
        }
        public string MailCC
        {
            get
            {
                return m_strMailCC;
            }
            set
            {

                m_strMailCC = value;
            }
        }
        public string MailBody
        {
            get
            {
                return m_strMailBody;
            }
            set
            {

                m_strMailBody = value;
            }
        }
    }
}

