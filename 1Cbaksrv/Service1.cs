using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace _1Cbaksrv
{
    public partial class BuckUper1C : ServiceBase
    {
        public BuckUper1C()
        {
            InitializeComponent();

            if (!System.Diagnostics.EventLog.SourceExists("1C BackUper"))
            {
                System.Diagnostics.EventLog.CreateEventSource(
                    "1C BackUper", "MyNewLog");
            }
            eventLog1.Source = "1C BackUper";
            eventLog1.Log = "MyNewLog";
        }

        protected override void OnStart(string[] args)
        {
            eventLog1.WriteEntry("In OnStart");
        }

        protected override void OnStop()
        {
            eventLog1.WriteEntry("In OnStop");
        }

        protected override void OnContinue()
        {
            eventLog1.WriteEntry("In OnContinue.");
        }
    }
}
