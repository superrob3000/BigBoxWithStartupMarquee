using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Xml.XPath;

namespace BigBoxWithStartupMarquee
{
    public partial class StartupMarqueeForm : Form
    {
        [DllImport("User32.dll", CallingConvention = CallingConvention.StdCall, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool ShowWindow([In] IntPtr hWnd, [In] int nCmdShow);


        Process ps_mplayer = null;
        Process ps_bigbox = null;
        public StartupMarqueeForm()
        {
            InitializeComponent();

            //Check if BigBox is already running
            {
                bool BigBoxRunning = false;
                Process[] Processes = System.Diagnostics.Process.GetProcesses();
                for (int i = 0; i < Processes.Length; i++)
                {
                    if (Processes[i].ProcessName.StartsWith("LaunchBox"))
                    {
                        BigBoxRunning = true;
                    }
                }

                if (BigBoxRunning)
                {
                    MessageBox.Show("An instance of BigBox is still running.");
                    this.Close();
                    return;
                }
            }

            // Start bigbox
            ps_bigbox = new Process();
            ps_bigbox.StartInfo.UseShellExecute = false;
            ps_bigbox.StartInfo.RedirectStandardInput = false;
            ps_bigbox.StartInfo.RedirectStandardOutput = false;
            ps_bigbox.StartInfo.CreateNoWindow = true;
            ps_bigbox.StartInfo.UserName = null;
            ps_bigbox.StartInfo.Password = null;
            ps_bigbox.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
            ps_bigbox.StartInfo.FileName = Path.GetDirectoryName(Application.ExecutablePath).ToString() + "/BigBox.exe";
            ps_bigbox.Start();

            //Check to see if we are already running
            {
                Process[] Processes = System.Diagnostics.Process.GetProcesses();
                int pcount = 0;
                for (int i = 0; i < Processes.Length; i++)
                {
                    if (Processes[i].ProcessName.StartsWith("BigBoxWithStartupMarquee"))
                    {
                        pcount++;
                    }
                }

                if(pcount >1)
                {
                    this.Close();
                    return;
                }
            }

            try
            {
                //Get BigBox settings from XML file
                string xml_path = Path.GetDirectoryName(Application.ExecutablePath).ToString() + "/Data/BigBoxSettings.xml";
                XDocument xSettingsDoc;
                xSettingsDoc = XDocument.Load(xml_path);

                MarqueeMonitorIndex = xSettingsDoc
                .XPathSelectElement("/LaunchBox/BigBoxSettings")
                .Element("MarqueeMonitorIndex")
                .Value;

                if ((Convert.ToInt32(MarqueeMonitorIndex) < 0) || (Convert.ToInt32(MarqueeMonitorIndex) > Screen.AllScreens.GetUpperBound(0)))
                {
                    this.Close();
                    return;
                }

                //Get Omega settings from XML file
                xml_path = Path.GetDirectoryName(Application.ExecutablePath).ToString() + "/Data/OmegaSettings.xml";
                xSettingsDoc = XDocument.Load(xml_path);

                MarqueeWidth = xSettingsDoc
                .XPathSelectElement("/OmegaSettings")
                .Element("MarqueeWidth")
                .Value;

                MarqueeHeight = xSettingsDoc
                .XPathSelectElement("/OmegaSettings")
                .Element("MarqueeHeight")
                .Value;

                MarqueeStretch = xSettingsDoc
                .XPathSelectElement("/OmegaSettings")
                .Element("MarqueeStretch")
                .Value;

                MarqueeVerticalAlignment = xSettingsDoc
                .XPathSelectElement("/OmegaSettings")
                .Element("MarqueeVerticalAlignment")
                .Value;
            }
            catch
            {
                this.Close();
                return;
            }

            Screen marquee = Screen.AllScreens[Convert.ToInt32(MarqueeMonitorIndex)];

            ///Set size and location of this form
            this.Width = Convert.ToInt32(MarqueeWidth);
            this.Height = Convert.ToInt32(MarqueeHeight);

            //Center the form Horizontally, Align the Form vertically as per MarqueeVerticalAlignment
            this.Location = new Point(
                marquee.Bounds.Location.X + ((marquee.Bounds.Size.Width - this.Width) / 2),
                marquee.Bounds.Location.Y + (MarqueeVerticalAlignment.Equals("Center") ? ((marquee.Bounds.Size.Height - this.Height) / 2) : 0));

            ps_mplayer = new Process();

            ps_mplayer.StartInfo.UseShellExecute = false;
            ps_mplayer.StartInfo.RedirectStandardInput = false;
            ps_mplayer.StartInfo.RedirectStandardOutput = false;
            ps_mplayer.StartInfo.CreateNoWindow = true;
            ps_mplayer.StartInfo.UserName = null;
            ps_mplayer.StartInfo.Password = null;
            ps_mplayer.StartInfo.WindowStyle = ProcessWindowStyle.Minimized; //Minimize the cmd window

            //Path to mplayer            
            ps_mplayer.StartInfo.FileName = Path.GetDirectoryName(Application.ExecutablePath).ToString() + "/ThirdParty/MPlayer/mplayer.exe";

            //Choose random video from Launchbox/Videos/StartupMarquee
            var rand = new Random();
            var files = Directory.GetFiles(Path.GetDirectoryName(Application.ExecutablePath).ToString() + "/Videos/StartupMarquee");
            if (files.Length == 0)
            {
                this.Close();
                return;
            }

            //-wid will tell MPlayer to show output inisde our panel
            ps_mplayer.StartInfo.Arguments = "\"" + files[rand.Next(files.Length)] + "\" ";

            if (MarqueeStretch.Equals("Fill"))
                ps_mplayer.StartInfo.Arguments += " -aspect " + this.Width + ":" + this.Height;

            ps_mplayer.StartInfo.Arguments += " -colorkey 0x00000000 -noborder -nosound -loop 0 -nomouseinput -noconsolecontrols -wid " + (int)panel1.Handle;


            //Wait for BigBox
            bool BigBoxStarted = false;
            int count = 0;
            while (!BigBoxStarted)
            {
                int milliseconds = 500;
                Thread.Sleep(milliseconds);
                {
                    Process[] Processes = System.Diagnostics.Process.GetProcesses();
                    for (int i = 0; i < Processes.Length; i++)
                    {
                        if (Processes[i].ProcessName.StartsWith("LaunchBox"))
                        {
                            BigBoxStarted = true;
                        }
                    }
                }

                if (BigBoxStarted)
                {
                    //Start mplayer
                    ps_mplayer.Start();


                    IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;
                    ShowWindow(handle, 6);

                    milliseconds = 2000;
                    Thread.Sleep(milliseconds);
                }
                else
                {
                    if (++count > 15)
                    {
                        this.Close();
                        return;
                    }
                }

            }
        }

        private string MarqueeMonitorIndex;
        private string MarqueeWidth;
        private string MarqueeHeight;
        private string MarqueeStretch;
        private string MarqueeVerticalAlignment;


        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                //Make sure Mplayer is exited when form is closed
                ps_mplayer.Kill();
            }
            catch { }
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

    }
}
