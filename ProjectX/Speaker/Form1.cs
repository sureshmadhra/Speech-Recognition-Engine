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
using System.Threading;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Xml; 


namespace Speaker
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
          
        }
        

        StreamReader CommandsReader = new StreamReader(@"C:\users\" + Environment.UserName.ToString()+@"\documents\commands.txt");
        SpeechSynthesizer sSynth = new SpeechSynthesizer();
        PromptBuilder pBuilder = new PromptBuilder();
        SpeechRecognitionEngine sRecognize = new SpeechRecognitionEngine();
        string name = "Microsoft Zira Desktop"; 
        Point lastPosition;
        string temp;
        string condition;
        string Humidity;
        string wind;
         
        private void Form1_Load(object sender, EventArgs e)
        {

            initializeSpeach();          
            
        }

        public String GetWeather(String input)
        {
            String query = String.Format("https://query.yahooapis.com/v1/public/yql?q=select * from weather.forecast where woeid in (select woeid from geo.places(1) where text='surat, gujarat')&format=xml&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys");
            XmlDocument wData = new XmlDocument();
            wData.Load(query);

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
                if (input == "win")
                {
                    return wind;
                }
                
            }
            catch
            {
                return "Error Reciving data";
            }
            return "error";
        }

        GrammarBuilder gbuilder = new GrammarBuilder();
     public void initializeSpeach()
     {
            Choices sList = new Choices();
            try
            {
                gbuilder.Append(new Choices(System.IO.File.ReadAllLines(@"C:\users\" + Environment.UserName.ToString()+@"\documents\commands.txt")));
            }
            catch { MessageBox.Show("The 'Commands' file must not contain empty lines.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); pr.StartInfo.FileName = @"C:\users\" + Environment.UserName.ToString()+@"\documents\commands.txt"; pr.Start(); Application.Exit(); return; }
            Grammar gr = new Grammar(gbuilder);
            try
            {
                sRecognize.UnloadAllGrammars();
                sRecognize.RecognizeAsyncCancel();
                sRecognize.RequestRecognizerUpdate();
                sRecognize.LoadGrammar(gr);
                sRecognize.SpeechRecognized += sRecognize_SpeechRecognized;
                sRecognize.SetInputToDefaultAudioDevice();
                sRecognize.RecognizeAsync(RecognizeMode.Multiple);
               
            }

            catch
            {
                MessageBox.Show("Grammar Builder Error"); 
                return;
            }
           
        }

     
     bool start = false;
         
     Process pr = new Process(); 
     public void lockComputer()
     {

         System.Diagnostics.Process.Start(@"C:\WINDOWS\system32\rundll32.exe", "user32.dll,LockWorkStation");
         return; 
     }

         public void speakText(string textSpeak)
        {
            sRecognize.RecognizeAsyncCancel();
            sRecognize.RecognizeAsyncStop();
            pBuilder.ClearContent();
            pBuilder.AppendText(textSpeak.ToString());
           sSynth.SelectVoice(name);
            sSynth.SpeakAsync(pBuilder);
            sRecognize.RecognizeAsyncCancel();
            sRecognize.RecognizeAsyncStop();
            sRecognize.RecognizeAsync(RecognizeMode.Multiple);
        }
         bool exitCondition = false;
        private void sRecognize_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
     {
         

           

         if(exitCondition)
         {
             Thread.Sleep(100);
             if (e.Result.Text == "yes")
             {
                 sSynth.SpeakAsyncCancelAll();
                 sRecognize.RecognizeAsyncCancel();
                 Application.Exit();

                 return;
             }
             else { exitCondition = false;  speakText("Exit Cancelled"); return; }
         }
         switch (e.Result.Text)
         {
             case "hello":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 speakText("Hello Master");
                 break;

             case "what is the meaning of life":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 speakText("Try and be nice to people, avoid eating fat, read a good book every now and then, get some walking in, and try to live together in peace and harmony with people of all creeds and nations");
                 break;
             case "ok google":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 speakText("Ver funny. I mean, not funny “ha-ha”, but funny.");
                 break;

             case "invisible":
                 if (ActiveForm.Visible == true)
                 {
                     listBox2.Items.Add(e.Result.Text.ToString());
                    ActiveForm.ShowInTaskbar = false;  
                     
                     ActiveForm.Hide(); 
                     
                                   
                     speakText("I am now invisible. You can access me by clicking on the icon down here in the tray.");
                     break;
                 }
                 else
                 {
                       
                     break; 
                 }
             case "hi":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 speakText("Hello, Master" );
                 break;
            
             case "weather":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 speakText(GetWeather("cond"));
                 break;
             
             
             
             case "temperature":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 speakText(GetWeather("temp"));
                 break;
             case "wind":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 speakText(GetWeather("win"));
                 break;
             case "humidity":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 speakText(GetWeather("hum"));
                 break;
             case "hi cortana":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 speakText(" I think you’ve got the wrong assistant");
               

                 break;
             case "do you sleep":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 speakText(" I don’t need much sleep, but it’s nice of you to ask");


                 break;

             case "exit":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 speakText("Goodbye!!!!!!!!!!!!!!!!!!! ,,,       ");
                 Application.Exit();
                 
                 break;
             
             
             case "normal":
                 if (this.WindowState == FormWindowState.Minimized)
                 {
                     speakText("I'm sorry,but I can't raise the window when it is minimized.");

                 }
                 speakText("Going back.");
                 ActiveForm.WindowState = FormWindowState.Normal;

                 break;
             
             case "add words":
                 speakText("Just type in the word to add to my dictionary and press enter. ");
                 new Form2().ShowDialog();


                 break;
          
             
             case "thank you":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 speakText("You're welcome");
                 break;
             case "How are you":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 speakText("I am fine what about you");
                 break;
             
             case "What is your name":
                 speakText("My name is Microsoft Zira Desktop" );
                 break;
             
             
             case "stop talking":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 cancelSpeech();
                 break;
             case "Who owns you":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 speakText("Nobody own me");

                 break;


             case "lock computer":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 speakText("Done");
                 lockComputer();
                 break;
             case "what is the date":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 speakText("Today is" + DateTime.Today.ToShortDateString());
                 
                 break;
             case "what time is it":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 speakText("It is" + DateTime.Now.ToString("h:mm tt"));

                 break;

             case "be quiet":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 cancelSpeech();
                 sSynth.SpeakAsyncCancelAll();
                 break;

             case "internet":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 speakText("One moment.");
                 Process pr = new Process();
                 pr.StartInfo.FileName = "http://www.google.com/";
                 pr.Start();
                 break; 
             
             
             case "help":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 bool working = true; 
                 while (working)
                 {
                     try
                     {
                         listBox1.Show();
                      listBox1.Items.Add(CommandsReader.ReadLine());
                     }
                     catch { working = false; break; }

                 }
                 speakText("Here they are... ,, , There are bunch!");
                 working = false;
                 
                 
                 break;
             case "maximize":
                 
                 try
                 {
                     ActiveForm.WindowState = FormWindowState.Maximized;
                     listBox2.Items.Add(e.Result.Text.ToString());
                     break;
                 }
                 catch
                 {
                     listBox2.Items.Add(e.Result.Text.ToString()+ "[FAILED]") ;
                     break;
                 }

             case "minimize":
                 speakText("OK");
                 ActiveForm.WindowState = FormWindowState.Minimized;
                 
                 break;
             case "commands":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 bool worrking = true; 
                 while (worrking)
                 {
                     try
                     {
                         listBox1.Show();
                      listBox1.Items.Add(CommandsReader.ReadLine());
                     }
                     catch { worrking = false; break; }

                 }

                 break;
             case "eject drive":
                 OpenCloseCD();
                 listBox2.Items.Add(e.Result.Text.ToString());
                 
                 break; 
             case "what's up":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 speakText("Oh, not much. How about you?");
                 break;

             case "note":
                
                 break;
             case "youtube":
                 listBox2.Items.Add(e.Result.Text.ToString());
                 speakText("One moment.");
                 Process pr1 = new Process();
                 pr1.StartInfo.FileName = "http://www.youtube.com/";
                 pr1.Start();


                 break; 

                 
          
         }


       
           
     }
        
    
       
        private void button3_Click(object sender, EventArgs e)
        {
            sRecognize.RecognizeAsyncCancel();
            sRecognize.RequestRecognizerUpdate();
            sRecognize.UnloadAllGrammars();
            
             
        }
        
        public void OpenCloseCD()
        {
            EjectMedia.Eject(@"\\.\D:");

            
            
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            pBuilder.ClearContent(); 
            sRecognize.RecognizeAsyncCancel(); 
        }

        private void addWordsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            f2.Show(); 
            
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            
        }

        private void pictureBox1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                lastPosition = e.Location;
            }
            else { return; } 
        }

        private void pictureBox1_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                this.Left += e.X - lastPosition.X;
                this.Top += e.Y - lastPosition.Y;
            }
        }

        private void pictureBox1_MouseUp(object sender, MouseEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            sSynth.SpeakAsyncCancelAll(); 
        }







        internal void cancelSpeech()
        {
            sSynth.SpeakAsyncCancelAll(); 
        }

        

        private void pictureBox1_Click(object sender, EventArgs e)
        {
           
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            EjectMedia.Eject(@"\\.\D:");
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {

            

        }
    }
}
