using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.Timers;

//== The class is utilizing both recognition and synthesis classes of system.speech
namespace SimpleSpeechRecognition
{
    public partial class MainForm : Form
    {
        //The speech recognition engine
        SpeechRecognitionEngine recognizer = new SpeechRecognitionEngine();

        //The synthesizer object
        SpeechSynthesizer synthesizer = new SpeechSynthesizer();

        //The dictation grammar that will be used when in dictation mode
        DictationGrammar dictationGrammar;

        //Define timers for system close and sleep commands
        System.Windows.Forms.Timer closeTimer;
        System.Timers.Timer sleepTimer;

        //Flag to indicate if recognition is already active 
        bool alreadyActive = false;

        //Flag to determine if dictation mode started
        bool dictationPrompt = false;

        public MainForm()
        {
            InitializeComponent();
            
            //Instantiate a new dictationGrmmar that will be used when in dictation mode
            dictationGrammar = new DictationGrammar();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            //Start by loading all needed grammar for our recognizer
            LoadBasicGrammar();

            //Start all recognition states
            StartRecognition();

            //Indicate that the system is in sleeping mode until user speaks
            lblStatus.Text = "Sleeping";
            lblStatus.ForeColor = Color.Blue;
        }

        //Define the basic grammar needed for the recognizer
        private void LoadBasicGrammar()
        {
            //Define the choices of grammar, in this case we have three choices:
            //==1. Dyna X Wake up => this will start the recognition process
            //==2. Start Dictation => this will the set the dictation mode
            //==3. Close System => this will shut down the system
            Choices choices = new Choices(new string[] { "Dyna X Wake up", "Start Dictation", "Close System" });

            //We need to use grammarBuilder to add the choices to the grammar's dictionary
            GrammarBuilder grammarBuilder = new GrammarBuilder(choices);
            Grammar grammar = new Grammar(grammarBuilder);

            //Remove all previous grammars (if any) and load the new grammars we added
            recognizer.UnloadAllGrammars();
            recognizer.LoadGrammar(grammar);

            //Remove dictation events and add recognition events since the recognizer is currently in sleep state
            recognizer.SpeechRecognized -= recognizer_SpeechRecognized_Dictation;
            recognizer.SpeechRecognized += recognizer_SpeechRecognized;

            //Set zero max alternates because we want only specific commands
            recognizer.MaxAlternates = 0;
        }

        //First time recognition call
        private void StartRecognition()
        {
            //Define preference synthesizer voice, female/male/neutral/..etc
            synthesizer.SelectVoiceByHints(VoiceGender.Female, VoiceAge.Senior);

            try
            {
                //recognizer should receive input from the default audio device
                recognizer.SetInputToDefaultAudioDevice();

                //Add event in case recognizer didn't recognize words in dictionary
                recognizer.SpeechRecognitionRejected += recognizer_SpeechRecognitionRejected;

                //Start recognition mode asyncronously, set mode to multiple because we don't want recognizer to stop
                recognizer.RecognizeAsync(RecognizeMode.Multiple);
                recognizer.Recognize();
            }
            catch
            {
                //In case of any error do nothing
                return;
            }
        }


        //Check recognized words from the recognizer 
        private void recognizer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            //Make sure system is listening
            if (e.Result.Text != "Triple X Wake up" && lblStatus.Text != "Listening") return;

            //Check if the recognized word actually matches our defined grammar
            switch (e.Result.Text)
            {
                case "Dyna X Wake up":
                    //If recognizer is already active speak to user to let them know
                    if (alreadyActive)
                        synthesizer.SpeakAsync("Already active boss!");
                    else
                        synthesizer.SpeakAsync("Yes boss!");

                    //Activate the recognizer if not already active
                    lblStatus.Text = "Listening";
                    lblStatus.ForeColor = Color.Green;
                    lblMode.Text = "Recognition";
                    alreadyActive = true;
                    break;

                case "Start Dictation":
                    //Must be in listening mode before starting dictation
                    if (lblStatus.Text != "Listening") return;

                    synthesizer.SpeakAsync("Dictation Started.");

                    //Since it is dictation mode, we need to use only the dictation grammar
                    recognizer.UnloadAllGrammars();
                    recognizer.LoadGrammar(dictationGrammar);

                    //Remove previous recognizer event and add dication event
                    recognizer.SpeechRecognized -= recognizer_SpeechRecognized;
                    recognizer.SpeechRecognized += recognizer_SpeechRecognized_Dictation;

                    lblMode.Text = "Dictation";
                    lblMode.ForeColor = Color.Orange;
                    break;

                case "Close System":
                    if (lblStatus.Text != "Listening") return;

                    //Inform the user that system is stopping
                    synthesizer.SpeakAsync("Goodbye sir.");

                    //Stop the recognizer
                    recognizer.RecognizeAsyncStop();

                    //Set the timer to close after a couple of seconds
                    closeTimer = new System.Windows.Forms.Timer();
                    closeTimer.Interval = 1200;
                    closeTimer.Tick += new EventHandler(timer_Tick);
                    closeTimer.Start();

                    break;
            }
        }

        //In case recognizer didn't recognize our defined words
        private void recognizer_SpeechRecognitionRejected(object sender, SpeechRecognitionRejectedEventArgs e)
        {
            //Recgonizer must be active to 
            if (alreadyActive)
                //Prompt voice message to user from the synthesizer
                synthesizer.Speak("Sorry i cannot recognize this word.");
        }

        //Check for activation of dictation mode
        private void recognizer_SpeechRecognized_Dictation(object sender, SpeechRecognizedEventArgs e)
        {
            //Clear current textbox contents
            richTextBox1.Text = "";

            //Check if spoken word is dictation mode
            if (e.Result.Text != "")
            {
                if (e.Result.Text.Contains("dictation"))
                {
                    //Dictation mode activated
                    dictationPrompt = true;
                    synthesizer.SpeakAsync("Please say Deactivate if you want to stop dictation.");
                    return;
                }
                else if (e.Result.Text.ToLower() == "deactivate" && dictationPrompt == true)
                {
                    //Dictation mode deactivated
                    synthesizer.SpeakAsync("Dictation mode deactivated");
                    LoadBasicGrammar();
                    lblMode.Text = "Recognition";
                    lblMode.ForeColor = Color.Black;
                    dictationPrompt = false;
                    return;
                }

                //Loop through all recognized words in e.Result.Words 
                foreach (RecognizedWordUnit word in e.Result.Words)
                {
                    //Append recognized words to the textbox
                    richTextBox1.Text += word.Text + " ";
                }

                //Make the synthesizer speaks the written text
                synthesizer.SpeakAsync(richTextBox1.Text);
            }
        }

        //Once timer tick is called, the application will close
        void timer_Tick(object sender, EventArgs e)
        {
            Application.Exit();
        }

    }
}