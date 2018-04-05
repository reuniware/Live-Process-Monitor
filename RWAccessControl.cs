using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using System.Net.Mail;
using RWAccessControl.Properties;
using System.Threading;
using System.IO;
using System.Reflection;
using System.Diagnostics;

namespace RWAccessControl
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        protected override void OnShown(EventArgs e)
        {
            if (Settings.Default.Visible == false)
            {
                base.OnShown(e);
                this.Hide();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (Settings.Default.SendEmail == true)
            {
                string Subject = "Login at : " + DateTime.Now.ToString();
                string Body = "Login at : " + DateTime.Now.ToString() + " - Username = " + Environment.UserName;
                SendEmail(Subject, Body, "");
            }

            FileLog("Email has been sent");

            Application.ApplicationExit += new EventHandler(this.OnApplicationExit);

            if (Settings.Default.StartProcessMonitor == true)
            {
                int duration = Settings.Default.ProcessMonitorDuration; // en secondes
                StartProcessMonitor(duration);
            }

            Application.Exit();
        }


        private List<string> lstProcesses = new List<string>();
        private bool firstRunDone = false;
        private Process[] processes = null;
        private string processData = "";
        private string[] tab = null;
        private string name = "";
        private int id = -1;
        private bool found = false;
        private bool killNewProcess = false; // Force killing of any new process (apart "cmd") ; a shortcut to cmd 
                                             // should be created on the desktop prior to launching this app
        bool StopProcessMonitor = false;
        private void StartProcessMonitor(int duration)
        {
            StopProcessMonitor = false;
            Stopwatch sw = new Stopwatch();
            sw.Start();
            Thread t = new Thread(ProcessMonitor);
            t.Start();
            while (sw.ElapsedMilliseconds < (duration * 1000))
            {
                Thread.Sleep(250);
            }
            StopProcessMonitor = true;
            return;
        }

        private void ProcessMonitor()
        {
            FileLog("Process Monitor Started");
            while (StopProcessMonitor == false)
            {
                if (firstRunDone == false)
                {
                    processes = Process.GetProcesses();
                    foreach (Process p in processes)
                    {
                        processData = p.ProcessName + "#" + p.Id;
                        lstProcesses.Add(processData);
                        //FileLog(DateTime.Now.ToString() + " : Initial process = " + processData);
                    }
                    firstRunDone = true;
                }
                else
                {
                    processes = Process.GetProcesses();
                    // pour chaque process sauvegardé, rechercher s'il existe toujours en mémoire
                    foreach (string processData in lstProcesses)
                    {
                        tab = processData.Split('#');
                        name = tab[0];
                        id = int.Parse(tab[1]);
                        found = false;
                        foreach (Process p in processes)
                        {
                            if (p.ProcessName.Equals(name) && (p.Id == id))
                            {
                                found = true;
                                break;
                            }
                        }
                        if (found == false)
                        {
                            //FileLog(DateTime.Now.ToString() + " : Process not in memory anymore = " + processData);
                            lstProcesses.Remove(processData);
                            break;
                        }
                    }

                    // pour chaque process en mémoire, rechercher s'il existe dans la liste des process sauvegardés
                    foreach (Process p in processes)
                    {
                        processData = p.ProcessName + "#" + p.Id;
                        if (!lstProcesses.Contains(processData))
                        {
                            lstProcesses.Add(processData);
                            FileLog(DateTime.Now.ToString() + " : New process detected = " + processData);
                            if (killNewProcess == true)
                            {
                                if (!p.ProcessName.ToLower().Trim().Equals("cmd") && !p.ProcessName.ToLower().Trim().Equals("conhost"))
                                {
                                    try
                                    {
                                        p.Kill();
                                    }
                                    catch (Exception)
                                    {

                                    }
                                }
                            }
                        }
                    }
                }
                Thread.Sleep(250);
            }
            FileLog("Process Monitor Stopped");
            SendEmail("Process Monitor Logs.", "Please check log file attached to this e-mail.", getPathToLogFile());
        }

        private string getPathToLogFile()
        {
            DateTime dt = DateTime.Now;
            string timestamp = dt.Year.ToString("D4") + dt.Month.ToString("D2") + dt.Day.ToString("D2");
            string workingDirDoc = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string pathToFile = workingDirDoc + @"\" + timestamp + "_rw_access_control.txt";
            return pathToFile;
        }

        private void FileLog(string str)
        {
            DateTime dt = DateTime.Now;
            string pathToFile = getPathToLogFile();
            FileStream fs = null;
            StreamWriter sw = null;
            try
            {
                fs = new FileStream(pathToFile, FileMode.Append);
                sw = new StreamWriter(fs);
                sw.WriteLine(dt.ToString() + " - " + str);
                sw.Close();
                fs.Close();
            }
            catch (Exception)
            {
                try { sw.Close(); } catch (Exception) { }
                try { fs.Close(); } catch (Exception) { }
            }
        }

        private void OnApplicationExit(object sender, EventArgs e)
        {
        }

        private void SendEmail(string subject, string body, string pathToAttachment)
        {
            try
            {
                var client = new SmtpClient
                {
                    Host = Settings.Default.SmtpHost,
                    Port = Settings.Default.SmtpPort,//465,
                    Credentials = new NetworkCredential(Base64Decode(Settings.Default.Login), Base64Decode(Settings.Default.Password))
                };

                string exMsg = "";
                Attachment data = null;
                try
                {
                    if (!pathToAttachment.Trim().Equals(""))
                    {
                        data = new Attachment(pathToAttachment);
                    }
                }
                catch (Exception ex)
                {
                    exMsg = ex.Message;
                }

                var msg = new MailMessage
                {
                    From = new MailAddress(Base64Decode(Settings.Default.EmailFrom)),
                    To = { new MailAddress(Base64Decode(Settings.Default.EmailTo)) },
                    Subject = subject,
                    Body = body,
                };

                try
                {
                    msg.Attachments.Add(data);
                }
                catch (Exception)
                {
                    //exMsg += ex.Message;
                }

                client.Send(msg);
            }
            catch (Exception)
            {
                //Console.WriteLine("Erreur générale");
            }
        }

        public static string Base64Decode(string base64EncodedData)
        {
            var base64EncodedBytes = Convert.FromBase64String(base64EncodedData);
            return Encoding.UTF8.GetString(base64EncodedBytes);
        }

    }
}
