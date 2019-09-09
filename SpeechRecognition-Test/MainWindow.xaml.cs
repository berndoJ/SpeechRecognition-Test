using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.Globalization;

namespace SpeechRecognition_Test
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        SpeechSynthesizer speech_synth = new SpeechSynthesizer();
        SpeechRecognitionEngine sr_engine = new SpeechRecognitionEngine(new CultureInfo("de-AT"));

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // From example code on MS-website

            // Grammar
            Choices sr_choices = new Choices();
            sr_choices.Add(new string[] { "schalte die lichter ein", "öffne word", "öffne excel", "öffne powerpoint", "öffne outlook" });
            GrammarBuilder sr_grb = new GrammarBuilder();
            sr_grb.Append(sr_choices);
            Grammar sr_gr = new Grammar(sr_grb);

            // Create and load a dictation grammar.  
            sr_engine.LoadGrammar(sr_gr);

            // Add a handler for the speech recognized event.  
            sr_engine.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(sr_engine_speech_recognized);
            sr_engine.SpeechDetected += new EventHandler<SpeechDetectedEventArgs>(sr_engine_speech_detected);

            // Configure input to the speech recognizer.  
            sr_engine.SetInputToDefaultAudioDevice();

            // Start asynchronous, continuous speech recognition.  
            sr_engine.RecognizeAsync(RecognizeMode.Multiple);

            gui_output_box.Text += "Started speech recognition.\n";

            speech_synth.Rate = -10;

            speech_synth.SelectVoiceByHints(VoiceGender.Male, VoiceAge.NotSet);
        }

        private void sr_engine_speech_detected(object sender, SpeechDetectedEventArgs e)
        {
            gui_output_box.Text += $"DETECTED\n";
        }

        private void sr_engine_speech_recognized(object sender, SpeechRecognizedEventArgs ev)
        {
            string cmd = ev.Result.Text;
            switch (cmd)
            {
                case "schalte die lichter ein":
                    speech_synth.SpeakAsync("Die Lichter im Zimmer wurden eingeschalten.");
                    break;
                case "öffne word":
                    System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE");
                    break;
                case "öffne excel":
                    System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE");
                    break;
                case "öffne powerpoint":
                    System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE");
                    break;
                case "öffne outlook":
                    System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE");
                    break;
            }
        }
    }
}
