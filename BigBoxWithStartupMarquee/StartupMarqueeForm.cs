using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
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

        String ExecutablePath;

        public StartupMarqueeForm()
        {
            InitializeComponent();

            ExecutablePath = Directory.GetParent(Directory.GetParent(Path.GetDirectoryName(Application.ExecutablePath).ToString()).ToString()).ToString();

            if(!File.Exists(Path.Combine(ExecutablePath, "BigBox.exe")))
            {
                //If we aren't in the LaunchBox folder then just force
                //the path (useful for debugging).
                //                ExecutablePath = "C:/Users/Administrator/LaunchBox/";

//                this.Close();
//                return;
            }


            //Check to see if we are already running
            {
                Process[] Processes = System.Diagnostics.Process.GetProcesses();
                int pcount = 0;
                for (int i = 0; i < Processes.Length; i++)
                {
                    if (Processes[i].ProcessName.Contains("StartupMarquee"))
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

            Screen marquee;

            try
            {
                //Get BigBox settings from XML file
                string xml_path = ExecutablePath + "/Data/BigBoxSettings.xml";
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

                marquee = Screen.AllScreens[Convert.ToInt32(MarqueeMonitorIndex)];


                //Use BigBox default Marquee settings
                MarqueeStretchImages = xSettingsDoc
                .XPathSelectElement("/LaunchBox/BigBoxSettings")
                .Element("MarqueeStretchImages")
                .Value;

                MarqueeScreenCompatibilityMode = xSettingsDoc
                .XPathSelectElement("/LaunchBox/BigBoxSettings")
                .Element("MarqueeScreenCompatibilityMode")
                .Value;

                if (MarqueeStretchImages.Equals("true"))
                    MarqueeStretch = "Fill";
                else
                    MarqueeStretch = "Preserve Aspect Ratio";

                MarqueeWidth = Convert.ToString(marquee.Bounds.Size.Width);

                switch (MarqueeScreenCompatibilityMode)
                {
                    case "TopHalfCutOff":
                        MarqueeHeight = Convert.ToString(marquee.Bounds.Size.Height / 2);
                        MarqueeVerticalAlignment = "Bottom";
                        break;
                    case "TopTwoThirdsCutOff":
                        MarqueeHeight = Convert.ToString(marquee.Bounds.Size.Height / 3);
                        MarqueeVerticalAlignment = "Bottom";
                        break;

                    case "TopAndBottomOneThirdCutOff":
                        MarqueeHeight = Convert.ToString(marquee.Bounds.Size.Height / 3);
                        MarqueeVerticalAlignment = "Center";
                        break;

                    case "BottomHalfCutOff":
                        MarqueeHeight = Convert.ToString(marquee.Bounds.Size.Height / 2);
                        MarqueeVerticalAlignment = "Top";
                        break;
                    case "BottomTwoThirdsCutOff":
                        MarqueeHeight = Convert.ToString(marquee.Bounds.Size.Height / 3);
                        MarqueeVerticalAlignment = "Top";
                        break;


                    case "HalfSizeStretched":
                    case "ThirdSizeStretched":
                        //Currently not supported

                    case "None":
                    default:
                        MarqueeHeight = Convert.ToString(marquee.Bounds.Size.Height);
                        MarqueeVerticalAlignment = "Top";
                        break;

                }
            }
            catch
            {
                this.Close();
                return;
            }


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
            ps_mplayer.StartInfo.FileName = ExecutablePath + "/ThirdParty/MPlayer/mplayer.exe";

            //Choose random video from Launchbox/Videos/StartupMarquee
            var rand = new Random();
            var files = Directory.GetFiles(ExecutablePath + "/Videos/StartupMarquee");
            if (files.Length == 0)
            {
                this.Close();
                return;
            }

            //-wid will tell MPlayer to show output inisde our panel
            String file = files[rand.Next(files.Length)];

            //Check for static image
            if (Path.GetExtension(file).Equals(".jpg"))
            {
                //static images only work from current directory.
                Directory.SetCurrentDirectory(ExecutablePath + "/Videos/StartupMarquee");
                ps_mplayer.StartInfo.Arguments = "\"" + "mf://" + Path.GetFileName(file) + "\" " + "-mf type=jpg ";
            }
            else if (Path.GetExtension(file).Equals(".png"))
            {
                //static images only work from current directory.
                Directory.SetCurrentDirectory(ExecutablePath + "/Videos/StartupMarquee");
                ps_mplayer.StartInfo.Arguments = "\"" + "mf://" + Path.GetFileName(file) + "\" " + "-mf type=png ";
            }
            else
            {
                ps_mplayer.StartInfo.Arguments = "\"" + file + "\" ";
            }

            if (MarqueeStretch.Equals("Fill"))
                ps_mplayer.StartInfo.Arguments += " -aspect " + this.Width + ":" + this.Height;

            ps_mplayer.StartInfo.Arguments += " -colorkey 0x00000000 -noborder -nosound -loop 0 -nomouseinput -noconsolecontrols -wid " + (int)panel1.Handle;

            //Start mplayer
            ps_mplayer.Start();


            IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;
            ShowWindow(handle, 6);

            Thread.Sleep(1500); //ms
        }

        //BigBox settings
        private string MarqueeMonitorIndex;
        private string MarqueeStretchImages;
        private string MarqueeScreenCompatibilityMode;

        //Omega settings
        private string MarqueeWidth;
        private string MarqueeHeight;
        private string MarqueeStretch;
        private string MarqueeVerticalAlignment;


        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                //Make sure Mplayer is exited when form is closed
                if(ps_mplayer != null)
                    ps_mplayer.Kill();
            }
            catch { }
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void StartupMarqueeForm_Shown(object sender, EventArgs e)
        {
            Task.Factory.StartNew(() => { SetFocusToBigBox(); });
        }

        [DllImport("user32.dll")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        private IntPtr GetHandleWindow(string title)
        {
            return FindWindow(null, title);
        }

        private void SetFocusToBigBox()
        {
            IntPtr handle = IntPtr.Zero;
            AutomationElement element = null;
            int timeout = 0;

            //Find the BigBox Intro Video window
            while (timeout++ < 15)
            {
                Thread.Sleep(250);

                //This is kind of fragile. It will break if they change
                //the window title in the future.
                handle = GetHandleWindow("LaunchBox Big Box Startup Video");

                if (handle != IntPtr.Zero)
                {
                    element = AutomationElement.FromHandle(handle);
                    break;
                }
            }

            //Set focus to BigBox so that intro video can be skipped with a button press.                    
            if (element != null)
            {
                element.SetFocus();
            }
        }
    }
}
