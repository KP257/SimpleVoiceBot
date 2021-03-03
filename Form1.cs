using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Speech.Synthesis;
using System.Speech.Recognition;
using System.Diagnostics;
using System.Xml;

namespace Serena
{
    public partial class Serena : Form
    {

        SpeechSynthesizer s = new SpeechSynthesizer();

        Boolean wake = false;
        String temp;
        String condition;
        Choices list = new Choices();

        String[] greetings = new String[3] {"yes", "hello", "now what" };


        public String GreetingsAction() {
            Random r = new Random();
            return greetings[r.Next(3)];

           
        }
        public Serena()
        {
             InitializeComponent();
            s.SelectVoiceByHints(VoiceGender.Female);

            SpeechRecognitionEngine rec = new SpeechRecognitionEngine();

            list.Add(new String[] { "hello", "how are you", "what time is it", "what is the date", "open google", "sleep", "wake", "restart", "open steam", "close steam"
            , "weather report", "serena", "minimize", "maximize", "play", "pause", "next"});
            Grammar gr = new Grammar(new GrammarBuilder(list));

            try
            {
                rec.RequestRecognizerUpdate();
                rec.LoadGrammar(gr);
                rec.SpeechRecognized += rec_SpeechRecognize;
                rec.SetInputToDefaultAudioDevice();
                rec.RecognizeAsync(RecognizeMode.Multiple);
            }
            catch (Exception)
            {

                return;
            }
           
        }

       

        private void rec_SpeechRecognize(object sender, SpeechRecognizedEventArgs e)
        {
            String r = e.Result.Text;


            if (r == "serena")
            {
                wake = true;
                say(GreetingsAction());
            }

            if (r == "wake")
            {
                wake = true;
                label3.Text = "State: Awake!";
            }
            if (r == "sleep")
            {
                wake = false;
                label3.Text = "State: Sleep!";
            }

            if (wake == true)
            {

                //max and mind
                if (r == "minimize")
                {
                    this.WindowState = FormWindowState.Minimized;
                }
                if (r == "maximize")
                {
                    this.WindowState = FormWindowState.Normal;
                }
                //restart
                if (r == "restart" || r == "update")
                {
                    restart();
                }

                //time
                if (r == "what time is it")
                {
                    say(DateTime.Now.ToString("h:mm tt"));
                }
                //date
                if (r == "what is the date")
                {
                    say(DateTime.Now.ToString("M/d/yyyy"));
                }

                // open and close
                if (r == "open google")
                {
                    say("opening now");
                    Process.Start("https://www.google.co.za/?gfe_rd=cr&ei=ZhZfVfqLOeuo8wfn2IGgDw&gws_rd=ssl");
                }

                if (r == "open steam")
                {
                    say("opening now");
                    Process.Start(@"C:\Program Files (x86)\Steam\Steam.exe");
                }

                if (r == "close steam")
                {
                    say("closed");
                    KillProg("steam.exe"); 
                }

                if (r == "weather report")
                {
                    say(GetWeather("condition" + "temp"));
                }
                //play and pause music
                if (r == "open youtube")
                {
                    Process.Start("https://www.youtube.com/?gl=ZA");
                }
                if (r == "play" || r == "pause")
                {
                    SendKeys.Send(" ");
                }
                if (r == "next")
                {
                }
                    SendKeys.Send("");

                //weather
                if (r == "weather report")
                {
                    say(GetWeather("condition" + "temp"));
                }

            }
            txtInput.AppendText(r + "\n");
        }


        public void say(String h)
        {
            s.Speak(h);
            wake = false;
            txtOutput.AppendText(h + "\n");
        }

        public void restart()
        {
            Process.Start(@"C:\Users\sebas\source\repos\Serena");
            Environment.Exit(0);
        }

        public void KillProg(String s)
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
            catch (Exception)
            {

                say("The app you want to close is not open");
            }
            finally 
            {
                if (procs != null) 
                {
                    foreach (Process p in procs)
                    {
                        p.Dispose();
                    }
                    {

                    }
                }
            }
             procs = null;
        }


        public String GetWeather(String input)
        {
            String query = String.Format("https://query.yahooapis.com/v1/public/yql?q=select * from weather.forecast where woeid in (select woeid from geo.places(1) where text='johannesburg, state')&format=xml&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys");
            XmlDocument wData = new XmlDocument();
            try
            {
                wData.Load(query);
            }
            catch (Exception)
            {

                MessageBox.Show("No Internet connection");
                return "No internet";
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





        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
