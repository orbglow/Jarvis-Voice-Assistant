using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Windows.Forms;
using System.Threading.Tasks;
using CoreAudioApi;
using System.Speech.Synthesis;
using System.Speech.Recognition;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.IO.Ports;
using System.Xml;

namespace JVB
{
    public partial class Form1 : Form
    {

       // SerialPort port = new SerialPort("COM7", 9600, Parity.None, 8, StopBits.One);


        string condition;
        string temp;

        //////////////////////////////////////////Volume/////////////////////////////////////////////
        private const int APPCOMMAND_VOLUME_MUTE = 0x80000;
        private const int APPCOMMAND_VOLUME_UP = 0xA0000;
        private const int APPCOMMAND_VOLUME_DOWN = 0x90000;
        private const int WM_APPCOMMAND = 0x319;
        private MMDevice device;


        [DllImport("user32.dll")]
        public static extern IntPtr SendMessageW(IntPtr hWnd, int Msg,
            IntPtr wParam, IntPtr lParam);



        //This is a replacement for Cursor.Position in WinForms
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern bool SetCursorPos(int x, int y);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        public const int MOUSEEVENTF_LEFTDOWN = 0x02;
        public const int MOUSEEVENTF_LEFTUP = 0x04;


        Boolean wake = false;
        DateTime now = DateTime.Now;
        SpeechSynthesizer s = new SpeechSynthesizer();
        Choices list = new Choices();

        public Boolean search = false;

        public Form1()
        {
            
            SpeechRecognitionEngine rec = new SpeechRecognitionEngine();

            // list.Add(File.ReadAllLines(@"D:\M\Commands\DB.txt"));


            list.Add(new string[] { 

               //Jarvis
               "Hey Jarvis", "Jarvis", "Hello", "Hi", "Salam", "Salam Jarvis", "Hello Beautiful", "What can i do for you", "Yes Sir", "How can i help you",
               "Good morning sir", "Good afternoon sir", "Good evening sir", "Quiet","Be quiet", "Stop talking", "Shut up", "Kari nakon", "Hiss",
               "Saket bash", "Restart", "Update", "Saket sho", "Harf Nazan", "dont do anything", "Type", "Search for", "shut down", "weather",
               
               "Volume up", "Volume down", "seda ro kam kon", "seda ro ziad kon",
               "What time is it", "What is today",

               //Farsi-test
               "Type Twitter", "Type pizza", "Type Fast Food", "Pizza", "Fast food", "Type Corona",
               "Type Perspolis", "Type Esteghlal", "che kar hayee mitoni anjam bedi?",


               //Programs
               "Close it", "bebandesh", "Minimize", "Normal", "Maximize",
               "Open chrome", "Chrome o baz kon", "Open Firefox", "Firefox o baz kon", "Telegram o baz kon",
               "Open Edge", "Open microsoft edge", "edge o baz kon",
               "Play music", "Close spotify", "Play", "Pause", "Next song", "ahang badi", "previous song ", "ahang ghabli",
               "Open Telegram", "Close Telegram", "Next chat",

               //Sites
               "Whats on twitter", "Open twitter", "twittero baz kon", "Order a pizza", "filmaye jadido be ar",
               "serialaye jadido be ar",

               //Sites Interactions
               "next tweet", "tweet badee", "tweet bad", "i like that", "like kon", "likesh kon", "tweet kon", "tweetesh kon", "befrest", "tamoome",
               "i wanna tweet something", "mikham tweet konam",

                });


            Grammar gr = new Grammar(new GrammarBuilder(list));

            try
            {
                rec.RequestRecognizerUpdate();
                rec.LoadGrammar(gr);
                rec.SpeechRecognized += rec_SpeachRecognized;
                rec.SetInputToDefaultAudioDevice();
                rec.RecognizeAsync(RecognizeMode.Multiple);
            }
            catch { return; }


            InitializeComponent();


            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();
            device = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia);
            tbMaster.Value = (int)(device.AudioEndpointVolume.MasterVolumeLevelScalar * 100);
            device.AudioEndpointVolume.OnVolumeNotification += new AudioEndpointVolumeNotificationDelegate(AudioEndpointVolume_OnVolumeNotification);
        }

        public String GetWeather(String input)
        {
            String query = String.Format("https://query.yahooapis.com/v1/public/yql?q=select * from weather.forecast where woeid in (select woeid from geo.places(1) where text='Iran, Tehran')&format=xml&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys");
            XmlDocument wData = new XmlDocument();
            try
            {
                wData.Load(query);
            }
            catch
            {

                return "No Internet";
            }
            XmlNamespaceManager manager = new XmlNamespaceManager(wData.NameTable);
            manager.AddNamespace("yweather", "http://xml.weather.yahoo.com/ns/rss/1.0");

            XmlNode channel = wData.SelectSingleNode("query").SelectSingleNode("results").SelectSingleNode("channel");
            XmlNodeList nodes = wData.SelectNodes("query/results/channel");
            try
            {
                temp = channel.SelectSingleNode("item").SelectSingleNode("yweather:condition", manager).Attributes["temp"].Value;
                condition = channel.SelectSingleNode("item").SelectSingleNode("yweather:condition", manager).Attributes["text"].Value;

                if (input == "temp")
                {
                    return temp;
                }

                if (input == "cond")
                {
                    return condition;
                }
            }
            catch
            {
                return "Error Reciving data";
            }
            return "error";
        }


        void AudioEndpointVolume_OnVolumeNotification(AudioVolumeNotificationData data)
        {
            if (this.InvokeRequired)
            {
                object[] Params = new object[1];
                Params[0] = data;
                this.Invoke(new AudioEndpointVolumeNotificationDelegate(AudioEndpointVolume_OnVolumeNotification), Params);
            }
            else
            {
                tbMaster.Value = (int)(data.MasterVolume * 100);
            }
        }

        private void tbMaster_Scroll(object sender, ScrollEventArgs e)
        {
             device.AudioEndpointVolume.MasterVolumeLevelScalar = ((float)tbMaster.Value / 100.0f);
        }


       /* public static void killprog(string s)
        {
           System.Diagnostics.Process[] procs = null;

            try
            {
                procs = Process.GetProcessesByName(s);
                Process prog = procs[0];

                if (!prog.HasExited)
                {
                    prog.Kill();
                }
            }

            finally
            {
                if (procs != null)
                {
                    foreach (Process p in procs)
                    {
                        p.Dispose();
                    }
                }
            }
            procs = null;
        } */
        
        public void say(string h)

        {
            s.SpeakAsync(h);
            wake = false;
            label1.Text = "وضعیت : غیرفعال";
            textBox2.AppendText(h + "\n");
            
        }


        //////////////////////////////////////////////////Multiple Answers///////////////////////////////////////////

        string[] GreetingsA = new string[3] { "Hello Beautiful", "What can i do for you", "How can i help you", };

        public string Greetings_A()
        {
            Random r = new Random();

            return GreetingsA[r.Next(3)];
        }
        
      /*  string[] Greetings1 = new string[3] { "Good morning sir", "How can i help you?", "What can i do for you?", };

        public string Greetings_morning()
        {
            Random r = new Random();

            return Greetings1[r.Next(3)];
        }

        string[] Greetings2 = new string[3] { "Good afternoon sir", "What can i do for you?", "How can i help you?", };

        public string Greetings_afternoon()
        {
            Random r = new Random();

            return Greetings2[r.Next(3)];
        }

        string[] Greetings3 = new string[3] { "Good evening sir", "How can i help you?", "What can i do for you", };

        public string Greetings_evening()
        {
            Random r = new Random();

            return Greetings3[r.Next(3)];
        }  */
 

        private void rec_SpeachRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            string r = e.Result.Text;

            if(search)
            {
                Process.Start("https://www.google.com/search?q="+r);
                search = false;
            }


            if (r == "Hey Jarvis" || r == "Jarvis" || r == "Hello" || r == "Hi" || r == "Salam" || r == "Salam Jarvis")
            {

                say(Greetings_A());
                label1.Text = "وضعیت : فعال";
                wake = true;
            }


            /* if (r == "Jarvis" || r == "Hello" || r == "Hi" || r == "Salam" || r == "Salam Jarvis")
             {
                 say(Greetings_A());
                 wake = true;
                 label1.Text = "وضعیت : فعال";

                 if (now.Hour >= 5 && now.Hour < 12)
                 {
                     //  say("Good morning");
                     say(Greetings_morning());
                 }

                 if (now.Hour >= 12 && now.Hour < 18)
                 {
                     //  say("Good afternoon");
                     say(Greetings_afternoon());
                 }

                 if (now.Hour >= 18 && now.Hour < 24)
                 {
                     // say("Good evening");
                     say(Greetings_evening());
                 }

                 if (now.Hour < 5)
                 {
                     say("It's getting late");
                 }

             } */


            if (r == "Quiet" || r == "Be quiet" || r == "Stop talking" || r == "Hiss" || r == "Kari nakon" ||
            r == "Harf Nazan" || r == "Saket bash" || r == "Saket sho" || r == "dont do anything" || r == "Shut up")
            {
                wake = false;
                s.SpeakAsyncCancelAll();
                label1.Text = "وضعیت : غیرفعال";
            }

            if (r == "Restart")
               {
                Application.Restart();
                Environment.Exit(0);
            }

            {
                if (r == "thats it" || r == "befrest" || r == "tamoome" || r == "tweetesh kon" || r == "tweet kon")
                {
                    SendKeys.Send("^{ENTER}");
                }
            }

            if (wake == true && search == false)

            {

                if (r == "Search for")
                    {
                        search = true;
                    }


                //Farsi-test

                {
                    if (r == "Type Fast Food")
                    {
                        SendKeys.Send("فست فود");
                    }
                }

                {
                    if (r == "Type pizza")
                    {
                        SendKeys.Send("پیتزا");
                    }
                }
                
                {
                    if (r == "Type Twitter")
                    {
                        SendKeys.Send("توییتر");
                    }
                }

                {
                    if (r == "Type Corona")
                    {
                        SendKeys.Send("کرونا");
                    }
                }

                {
                    if (r == "Type Perspolis")
                    {
                        SendKeys.Send("پرسپولیس");
                    }
                }

                {
                    if (r == "Type Esteghlal")
                    {
                        SendKeys.Send("استقلال");
                    }
                }

                {
                    if (r == "che kar hayee mitoni anjam bedi?")
                    {
                        say("right now i'm limited, but in the future i will be able to do a lot of things ");
                    }
                }
      
                {
                    if (r == "shut down")
                    {
                        say("Shutting Down");
                        int sleepTime = 1000;
                        Task.Delay(sleepTime).Wait();
                        this.Close();
                    }
                }



                if (r == "Volume down" || r == "seda ro kam kon")
                {
                    SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle,
                    (IntPtr)APPCOMMAND_VOLUME_DOWN);
                    SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle,
                    (IntPtr)APPCOMMAND_VOLUME_DOWN);
                    SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle,
                    (IntPtr)APPCOMMAND_VOLUME_DOWN);
                }

                if (r == "Volume up" || r == "seda ro ziad kon")
                {
                    SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle,
                    (IntPtr)APPCOMMAND_VOLUME_UP);
                    SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle,
                    (IntPtr)APPCOMMAND_VOLUME_UP);
                    SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle,
                    (IntPtr)APPCOMMAND_VOLUME_UP);
                }


                if (r == "What time is it")
                {
                    say(DateTime.Now.ToString("h:mm tt"));
                }

                if (r == "What is today")
                {
                    say(DateTime.Now.ToString("M/d/yyyy"));
                }

                
                /////////////////////////////////////Programs/////////////////////////////////////

                if (r == "Close it" || r == "bebandesh")
                {
                    SendKeys.Send("^{W}");
                }

                {
                    if (r == "Switch application")
                    {
                        SendKeys.Send("%{TAB}");
                    }
                }

                if (r == "Minimize")
                {
                    this.WindowState = FormWindowState.Minimized;
                }

                if (r == "Normal")
                {
                    this.WindowState = FormWindowState.Normal;
                }

                if (r == "Maximize")
                {
                    this.WindowState = FormWindowState.Maximized;
                }

                if (r == "Open chrome" || r == "Chrome o baz kon")
                {
                    Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe");
                }

                if (r == "Open Edge" || r == "Open microsoft edge" || r == "edge o baz kon")
                {
                    Process.Start(@"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe");
                }

                if (r == "Open Firefox" || r == "Firefox o baz kon")
                {
                    Process.Start(@"C:\Program Files\Mozilla Firefox\firefox.exe");
                }

                if (r == "Play music")
                {
                    Process.Start(@"C:\Users\****Spotify.exe");
                    int sleepTime = 500;
                    Task.Delay(sleepTime).Wait();
                    SendKeys.Send(" ");
                }

                if (r == "Play" || r == "Pause")
                {
                    SendKeys.Send(" ");
                }

                if (r == "Next song" || r == "ahang badi")
                {
                    SendKeys.Send("^{RIGHT}");
                }

                if (r == "previous song " || r == "ahang ghabli")
                {
                    SendKeys.Send("^{LEFT}");
                }

               /* if (r == "Close spotify")
                {
                   killprog("Spotify.exe");
                } */


                /////////////////////////////////////Sites/////////////////////////////////////


                if (r == "Whats on twitter" || r == "Open twitter" || r == "twittero baz kon")
                {
                    Process.Start("https://www.twitter.com");
                }

                if (r == "filmaye jadido be ar")
                {
                    Process.Start("http://google.com");
                    int sleepTime = 3000;
                    Task.Delay(sleepTime).Wait();
                    SendKeys.Send("New Movies");
                    SendKeys.Send("{ENTER}");
                    System.Threading.Thread.Sleep(2000);
                }

                if (r == "serialaye jadido be ar")
                {
                    Process.Start("http://google.com");
                    int sleepTime = 3000;
                    Task.Delay(sleepTime).Wait();
                    SendKeys.Send("New TV Shows");
                    SendKeys.Send("{ENTER}");
                    System.Threading.Thread.Sleep(2000);
                    say("i heard that shameless is a good tv show.");
                }

                if (r == "Order a pizza")
                {
                    Process.Start("https://snappfood.ir");
                    int sleepTime = 2000;
                    Task.Delay(sleepTime).Wait();
                     /////
                }


                /////////////////////////////////////Sites Interactions/////////////////////////////////////

                {
                    if (r == "next tweet" || r == "tweet badee" || r == "tweet bad")
                    {
                        SendKeys.Send("j");
                    }
                }

                {
                    if (r == "like kon" || r == "likesh kon" || r == "i like that")
                    {
                        SendKeys.Send("l");
                    }
                }

                {
                    if (r == "i wanna tweet something" || r == "mikham tweet konam")
                    {
                        SendKeys.Send("n");
                    }
                }

            }

            textBox1.AppendText(r + "\n");

        }

        private void label2_Click(object sender, EventArgs e)
        {
            Form about = new About();
            about.ShowDialog();
        }

    } 
}
