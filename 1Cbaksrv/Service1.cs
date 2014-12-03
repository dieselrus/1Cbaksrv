using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Configuration;

namespace _1Cbaksrv
{

    public partial class BuckUper1C : ServiceBase
    {
        // Create a new FileSystemWatcher and set its properties.
        FileSystemWatcher watcher = new FileSystemWatcher();
        
        private INI _ini = new INI(AppDomain.CurrentDomain.BaseDirectory + "\\Task.ini");
        public String [,] lstTask;

        private Timer tm = new Timer();

        string sSource;
        string sLog;
        //string sEvent;

        public BuckUper1C()
        {
            InitializeComponent();

            sSource = "BackUper1C";
            //sLog = "Application";
            sLog = "BackUper1C";
            //sEvent = "Sample Event";

            if (!System.Diagnostics.EventLog.SourceExists(sSource))
            {
                System.Diagnostics.EventLog.CreateEventSource(sSource, sLog);
            }

            eventLog1.Source = sSource;
            eventLog1.Log = sLog;
      
        }

        // Define the event handlers.
        private void OnChanged(object source, FileSystemEventArgs e)
        {
            // Specify what is done when a file is changed, created, or deleted.
            // Console.WriteLine("File: " + e.FullPath + " " + e.ChangeType);
            // eventLog1.WriteEntry("FilwWatcher OnChanged");
            readTasksName();
        }

        protected override void OnStart(string[] args)
        {
            eventLog1.WriteEntry("In OnStart");

            watcher.Path = AppDomain.CurrentDomain.BaseDirectory;
            //watcher.Path = System.Configuration.Con ConfigurationManager.AppSettings[AppDomain.CurrentDomain.BaseDirectory + "\\Task.ini"];
            watcher.EnableRaisingEvents = true;
            // Add event handlers.
            watcher.Changed += new FileSystemEventHandler(OnChanged);
            //watcher.Created += new FileSystemEventHandler(OnChanged);
            //watcher.Deleted += new FileSystemEventHandler(OnChanged);
            //watcher.Renamed += new RenamedEventHandler(OnRenamed);

            readTasksName();

            this.tm = new System.Timers.Timer(60000);  // 30000 milliseconds = 30 seconds
            this.tm.AutoReset = true;
            this.tm.Elapsed += new System.Timers.ElapsedEventHandler(this.timer_Elapsed);
            this.tm.Start();
        }

        private void timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            //BuckUper1C.Main(); // my separate static method for do work
            //eventLog1.WriteEntry("Timer tick");
            //readTasksName();

            for (int i = 0; i < lstTask.GetLength(0); i++ )
            {
                //eventLog1.WriteEntry("TaskName = " + lstTask[0, i]);
                //eventLog1.WriteEntry("Hour = " + lstTask[1, i]);
                //eventLog1.WriteEntry("Minute = " + lstTask[2, i]);

                // Hour
                if (Convert.ToInt32(lstTask[1, i]) == DateTime.Now.Hour)
                {
                    //eventLog1.WriteEntry("Hour = " + lstTask[1, i]);                 
                    // Minute
                    if (Convert.ToInt32(lstTask[2, i]) == DateTime.Now.Minute)
                    {
                        //eventLog1.WriteEntry("Minute = " + lstTask[2, i]);

                        // Формируем строку с параметрами запуска
                        string arg = " DESIGNER /DisableStartupMessages ";
                        string patch = "";
                  
                        // 1C путь файлу 1C
                        patch = _ini.IniReadValue(lstTask[0, i], "file1CPath");

                        // Путь к архивам
                        arg += " /DumpIB " + _ini.IniReadValue(lstTask[0, i], "backupPath") + "\\" + lstTask[0, i] + "_" + DateTime.Now.ToString().Replace(" ", "_").Replace(":", "-") + ".dt";
                        // 1C пользователь
                        arg += " /N\"" + _ini.IniReadValue(lstTask[0, i], "user1C") + "\"";
                        // 1C пароль
                        arg += " /P\"" + _ini.IniReadValue(lstTask[0, i], "password1C") + "\"";

                        // вид базы
                        if (_ini.IniReadValue(lstTask[0, i], "FileOrServer") == "1")
                        {
                            // 1C путь к базе
                            arg += " /F" + _ini.IniReadValue(lstTask[0, i], "base1CPath");
                        }
                        else if (_ini.IniReadValue(lstTask[0, i], "FileOrServer") == "2")
                        {
                            // 1C сервер
                            arg += " /S" + _ini.IniReadValue(lstTask[0, i], "server1C");
                            // 1C база
                            arg += "\\" + _ini.IniReadValue(lstTask[0, i], "base1C");
                        }

                        //eventLog1.WriteEntry("arg = " + arg); 

                        // запуск процесса
                        startBackUp1C(patch, arg, lstTask[0, i]);
                    }
                }
            }
        }

        protected override void OnStop()
        {
            eventLog1.WriteEntry("In OnStop");
            this.tm.Stop();
            this.tm = null;
        }

        protected override void OnContinue()
        {
            eventLog1.WriteEntry("In OnContinue.");
        }

        // Чтение названия задач, секций файла
        private void readTasksName()
        {
            String[] lstTaskName = _ini.GetSectionNames();
            lstTask = new String[3, lstTaskName.Count()];

            //eventLog1.WriteEntry("Task count = " + lstTaskName.Count().ToString());

            int Count = 0;

            foreach (String tsk in lstTaskName)
            {
                //eventLog1.WriteEntry(tsk);  
                readTaskINI(tsk, Count);
                Count++;
            }
        }

        // Запуск 1С с параметрами
        private void startBackUp1C(string patch, string arg, string tsk)
        {
            System.Diagnostics.Process proc = new System.Diagnostics.Process();
            proc.StartInfo.FileName = patch;
            proc.StartInfo.Arguments = arg;
            proc.EnableRaisingEvents = true;
            //подписываемся на событие завершения процесса
            proc.Exited += new EventHandler(endBackUpProcess); // своя процедура обработки
            
            proc.Start();
            // Ожидание завершения
            proc.WaitForExit();

            if (proc.ExitCode == 0)
            {
                eventLog1.WriteEntry("Архивироание бызы " + tsk + " было успешно завершено в " + proc.ExitTime.ToString()); 
            }
            else
            {
                eventLog1.WriteEntry("Архивироание бызы " + tsk + " было прервано в " + proc.ExitTime.ToString());
            }
        }

        private void endBackUpProcess(object sender, System.EventArgs e)
        {
            //eventHandled = true;
            //Console.WriteLine("Exit time:    {0}\r\n" +
            //    "Exit code:    {1}\r\nElapsed time: {2}", proc.ExitTime, proc.ExitCode, elapsedTime);
            //eventLog1.WriteEntry("====");
        }

        // Чтение параметров задачи
        private void readTaskINI(String Task, int Count)
        {
            /*
            txtTaskName.Text = Task;
            // Путь к архивам
            txtBackUpPath.Text = _ini.IniReadValue(Task, "backupPath");
            // 1C путь файлу 1C
            txt1CFilePath.Text = _ini.IniReadValue(Task, "file1CPath");
            // Activ
            if (_ini.IniReadValue(Task, "activ") == "True")
            {
                chbActiv.Checked = true;
            }
            else
            {
                chbActiv.Checked = false;
            }
            // 1C путь к базе
            txtBasePath.Text = _ini.IniReadValue(Task, "base1CPath");
            // 1C сервер
            txt1CServer.Text = _ini.IniReadValue(Task, "server1C");
            // 1C база
            txt1CBase.Text = _ini.IniReadValue(Task, "base1C");
            // 1C пользователь
            txt1CUser.Text = _ini.IniReadValue(Task, "user1C");
            // 1C пароль
            txt1CPassword.Text = _ini.IniReadValue(Task, "password1C");
            // час
            numHour.Value = Convert.ToInt32(_ini.IniReadValue(Task, "taskHour"));
            // минута
            numMin.Value = Convert.ToInt32(_ini.IniReadValue(Task, "taskMin"));
            // вид базы
            if (_ini.IniReadValue(Task, "FileOrServer") == "1")
            {
                rbtn1CFile.Checked = true;
            }
            else if (_ini.IniReadValue(Task, "FileOrServer") == "2")
            {
                rbtn1CServer.Checked = true;
            }
            */

            if (_ini.IniReadValue(Task, "activ") == "True")
            {
                //eventLog1.WriteEntry("TaskName = " + Task);
                //eventLog1.WriteEntry("Hour = " + _ini.IniReadValue(Task, "taskHour"));
                //eventLog1.WriteEntry("Minute = " + _ini.IniReadValue(Task, "taskMin"));

                lstTask[0, Count] = Task;
                lstTask[1, Count] = _ini.IniReadValue(Task, "taskHour");
                lstTask[2, Count] = _ini.IniReadValue(Task, "taskMin");
            }
        
        }
    }

    class INI
    {
        //Использование может быть следующим. Записываем значение в файл: 
        //INI ini = new INI("Путь_к_файлу"); ini.IniWriteValue("Test_block","Key","Value");

        //Теперь в нашем файле есть значение Key, которое равно Value. Теперь считаем его: 
        //string value = ini.IniReadValue("Test_block","Key");

        public string pathINI;

        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

        // Second Method
        [DllImport("kernel32")]
        static extern int GetPrivateProfileString(string Section, int Key, string Value, [MarshalAs(UnmanagedType.LPArray)] byte[] Result, int Size, string FileName);

        // Third Method
        [DllImport("kernel32")]
        static extern int GetPrivateProfileString(int Section, string Key, string Value, [MarshalAs(UnmanagedType.LPArray)] byte[] Result, int Size, string FileName);

        public INI(string INIPath)
        {
            pathINI = INIPath;
        }

        public void IniWriteValue(string Section, string Key, string Value)
        {
            if (!Directory.Exists(Path.GetDirectoryName(pathINI)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(pathINI));
                //eventLog1.WriteEntry(tsk);
            }

            if (!File.Exists(pathINI))
            {
                using (File.Create(pathINI)) {
                    //eventLog1.WriteEntry("");
                };
            }

            WritePrivateProfileString(Section, Key, Value, this.pathINI);
        }

        public string IniReadValue(string Section, string Key)
        {
            StringBuilder temp = new StringBuilder(255);
            int i = GetPrivateProfileString(Section, Key, "", temp, 255, this.pathINI);
            return temp.ToString();
        }

        public string[] GetSectionNames()
        {
            //    Sets the maxsize buffer to 500, if the more
            //    is required then doubles the size each time.
            for (int maxsize = 500; true; maxsize *= 2)
            {
                //    Obtains the information in bytes and stores
                //    them in the maxsize buffer (Bytes array)
                byte[] bytes = new byte[maxsize];
                int size = GetPrivateProfileString(0, "", "", bytes, maxsize, this.pathINI);

                // Check the information obtained is not bigger
                // than the allocated maxsize buffer - 2 bytes.
                // if it is, then skip over the next section
                // so that the maxsize buffer can be doubled.
                if (size < maxsize - 2)
                {
                    // Converts the bytes value into an ASCII char. This is one long string.
                    //string Selected = Encoding.ASCII.GetString(bytes, 0, size - (size > 0 ? 1 : 0));

                    // русский
                    string Selected = Encoding.Default.GetString(bytes, 0, size - (size > 0 ? 1 : 0));
                    // sRet = System.Text.Encoding.Default.GetString(bRet, 0, i).TrimEnd((char)0);
                    //string Selected = System.Text.Encoding.Default.GetString(bRet, 0, i).TrimEnd((char)0);

                    // Splits the Long string into an array based on the "\0"
                    // or null (Newline) value and returns the value(s) in an array
                    return Selected.Split(new char[] { '\0' });
                }
            }
        }
    }
}
