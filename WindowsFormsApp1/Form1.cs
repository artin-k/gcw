using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using Word =  Microsoft.Office.Interop.Word;


using NAudio.Wave;

using System.Linq;
using System.Threading.Tasks;
using System.Drawing;
using System.Drawing.Printing;
using System.Drawing.Text;

using System.Runtime.InteropServices;
using System.Text;
using System.Diagnostics;
using System.Threading;
using System.Media;
using Timer = System.Windows.Forms.Timer;
using Microsoft.Office.Interop.Word;

//cursor tutorial

namespace WindowsFormsApp1
{



    public partial class Form1 : Form
    {
        private bool isAlarmActive = false; // Flag to track if the alarm has already sounded
        public int[] savedCarentPosiotion = new int[10];
        public int countBmark = 1;
        string fontFilePath;
        string filePathData;
        string[] fontNames;

        private object lastSelectedItem;
        string pressedChar;
        string wordName;
        string wordPath;
        string userVoicePath;
        string mainData;
        Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
        int projecktNummber;

        //string DocumentsPath;
        static int newVoice = 0;
        //static int newFile = 0;
        private WaveInEvent waveIn;
        private WaveFileWriter writer;        
        private IWavePlayer waveOutDevice;
        private AudioFileReader audioFileReader;
       

        string textBoxValue;
        private bool isRecording = false;
        private bool isWaitingForKey = false;

        private string outputFilePath;

        private Timer beepTimer;

        List<string> allDocxFiles = new List<string>();
        private List<string> dataList = new List<string>();
        private List<string> alphaList = new List<string>();
        private List<string> nameList = new List<string>();
        private List<string> voiceTagList = new List<string>();
        private List<string> dateList = new List<string>();
        private List<string> fontList = new List<string>();
        //02D1
        private AudioFileReader audioFile;
        private WaveOutEvent outputDevice;

        
        string soundMapping = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "soundDictionary");
        
        private Dictionary<char, string> soundMappings;

        // opening save the style in the word 
        //paging 
        //inset add break page 
        //
        public Form1()
        {
            this.KeyPreview = true;
            InitializeComponent();
            
            InitializeTimer();
            InitializeSoundMappings();
            this.MainrichTextBox.PreviewKeyDown += new PreviewKeyDownEventHandler(textBox1_PreviewKeyDown);
           
        }

        private void InitializeTimer() //this make a timer for alarm group
        {
            beepTimer = new Timer();
            beepTimer.Interval = 2000; // 2 seconds
            beepTimer.Tick += BeepTimer_Tick;
            
        }
        

        //83899518
     

        private void InitializeSoundMappings()
        {
            soundMappings = new Dictionary<char, string>
        { 
        {'ا', Path.Combine(soundMapping, "alef.mp3")},
        {'ب', Path.Combine(soundMapping, "be.mp3")},
        {'پ', Path.Combine(soundMapping, "pe.mp3")},
        {'ت', Path.Combine(soundMapping, "te.mp3")},
        {'ث', Path.Combine(soundMapping, "se.mp3")},
        {'ج', Path.Combine(soundMapping, "je.mp3")},
        {'چ', Path.Combine(soundMapping, "che.mp3")},
        {'ح', Path.Combine(soundMapping, "hhe.mp3")},
        {'خ', Path.Combine(soundMapping, "khe.mp3")},
        {'د', Path.Combine(soundMapping, "dal.mp3")},
        {'ذ', Path.Combine(soundMapping, "dal_ze.mp3")},
        {'ر', Path.Combine(soundMapping, "re.mp3")},
        {'ز', Path.Combine(soundMapping, "ze.mp3")},
        {'ژ', Path.Combine(soundMapping, "zhe.mp3")},
        {'س', Path.Combine(soundMapping, "sse.mp3")},
        {'ش', Path.Combine(soundMapping, "she.mp3")},
        {'ص', Path.Combine(soundMapping, "sad.mp3")},
        {'ض', Path.Combine(soundMapping, "zad.mp3")},
        {'ط', Path.Combine(soundMapping, "ta.mp3")},
        {'ظ', Path.Combine(soundMapping, "za.mp3")},
        {'ع', Path.Combine(soundMapping, "ain.mp3")},
        {'غ', Path.Combine(soundMapping, "ghain.mp3")},
        {'ف', Path.Combine(soundMapping, "fe.mp3")},
        {'ق', Path.Combine(soundMapping, "ghaf.mp3")},
        {'ک', Path.Combine(soundMapping, "kaf.mp3")},
        {'گ', Path.Combine(soundMapping, "gaf.mp3")},
        {'ل', Path.Combine(soundMapping, "lam.mp3")},
        {'م', Path.Combine(soundMapping, "mim.mp3")},
        {'ن', Path.Combine(soundMapping, "non.mp3")},
        {'و', Path.Combine(soundMapping, "ve.mp3")},
        {'ه', Path.Combine(soundMapping, "he.mp3")},
        {'ی', Path.Combine(soundMapping, "ye.mp3")},
        {'إ', Path.Combine(soundMapping, "alf_hamze.mp3")},
        {'ؤ', Path.Combine(soundMapping, "ve_hamze.mp3")},
        {'ئ', Path.Combine(soundMapping, "ye_hamze.mp3")},

        {'a', Path.Combine(soundMapping, "a.mp3")},
        {'b', Path.Combine(soundMapping, "b.mp3")},
        {'c', Path.Combine(soundMapping, "c.mp3")},
        {'d', Path.Combine(soundMapping, "d.mp3")},
        {'e', Path.Combine(soundMapping, "e.mp3")},
        {'f', Path.Combine(soundMapping, "f.mp3")},
        {'g', Path.Combine(soundMapping, "g.mp3")},
        {'h', Path.Combine(soundMapping, "h.mp3")},
        {'i', Path.Combine(soundMapping, "i.mp3")},
        {'j', Path.Combine(soundMapping, "j.mp3")},
        {'k', Path.Combine(soundMapping, "k.mp3")},
        {'l', Path.Combine(soundMapping, "l.mp3")},
        {'m', Path.Combine(soundMapping, "m.mp3")},
        {'n', Path.Combine(soundMapping, "n.mp3")},
        {'o', Path.Combine(soundMapping, "o.mp3")},
        {'p', Path.Combine(soundMapping, "p.mp3")},
        {'q', Path.Combine(soundMapping, "q.mp3")},
        {'r', Path.Combine(soundMapping, "r.mp3")},
        {'s', Path.Combine(soundMapping, "s.mp3")},
        {'t', Path.Combine(soundMapping, "t.mp3")},
        {'u', Path.Combine(soundMapping, "u.mp3")},
        {'v', Path.Combine(soundMapping, "v.mp3")},
        {'w', Path.Combine(soundMapping, "w.mp3")},
        {'x', Path.Combine(soundMapping, "x.mp3")},
        {'y', Path.Combine(soundMapping, "y.mp3")},
        {'z', Path.Combine(soundMapping, "z.mp3")},
        {'0', Path.Combine(soundMapping, "0.mp3")},
        {'1', Path.Combine(soundMapping, "1.mp3")},
        {'2', Path.Combine(soundMapping, "2.mp3")},
        {'3', Path.Combine(soundMapping, "3.mp3")},
        {'4', Path.Combine(soundMapping, "4.mp3")},
        {'5', Path.Combine(soundMapping, "5.mp3")},
        {'6', Path.Combine(soundMapping, "6.mp3")},
        {'7', Path.Combine(soundMapping, "7.mp3")},
        {'8', Path.Combine(soundMapping, "8.mp3")},
        {'9', Path.Combine(soundMapping, "9.mp3")},
        // Add more mappings here
        };
    }




        private void Form1_Load(object sender, EventArgs e)
        {            
            PositionGroupBoxes();



            wordPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "words");
            userVoicePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "user_voices");
            filePathData = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "files");
            mainData = Path.Combine(filePathData, "data.txt");
            fontFilePath = Path.Combine(filePathData, "fonts.txt");


            try
            {
                //make sure they exist and if they dont makeing them
                if (!Directory.Exists(filePathData))
                {
                    Console.WriteLine("Folder not found: " + filePathData);
                    Directory.CreateDirectory(filePathData);
                }

                if (!File.Exists(mainData))
                {
                    Console.WriteLine("file not found: " + mainData);
                    File.WriteAllText(mainData, "");
                }

                if (!File.Exists(fontFilePath))
                {
                    Console.WriteLine("file not found: " + fontFilePath);
                    File.WriteAllText(fontFilePath, "");
                }

                // Check if the folder exists
                if (!Directory.Exists(wordPath))
                {
                    Console.WriteLine("Folder not found: " + wordPath);
                    Directory.CreateDirectory(wordPath);
                }

                if (!Directory.Exists(userVoicePath))
                {
                    Console.WriteLine("folder users not founded but now is builded" + userVoicePath);
                    Directory.CreateDirectory(userVoicePath);
                }

            }

            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message);
            }




            try
            {
                ReadFileIntoList(mainData, fontFilePath);
                Console.WriteLine("****************" + projecktNummber);
                MessageBox.Show("data loaded seccessfully");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"an error occurred banana:{ex.Message}");
            }

            List<int> fontSize = new List<int> { 8, 9, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72 };
 
            foreach (int size in fontSize)
            {
                fontSizeComboBox.Items.Add(size.ToString());
            }

        }


        private void ReadFileIntoList(string filePath,string fontFilePath)
        {
            allDocxFiles.Clear();
            dataList.Clear();
            alphaList.Clear();
            nameList.Clear();
            voiceTagList.Clear();
            dateList.Clear();
            fontList.Clear();

            char delimiter = '|';

            try
            {

                if (Directory.Exists(wordPath))
                {
                    string[] docxFiles = Directory.GetFiles(wordPath, "*.docx", SearchOption.AllDirectories);
                    allDocxFiles = new List<string>(docxFiles); // Initialize and add found files
                }
                else
                {
                    Console.WriteLine($"Directory {wordPath} does not exist.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred BANANA : {ex.Message}");
            }

            if (File.Exists(fontFilePath))
            {
                fontNames = File.ReadAllLines(fontFilePath);

            }      
            else
            {
                MessageBox.Show("Font list file not found!");
            }



                if (File.Exists(filePath))
            {
                string[] lines = File.ReadAllLines(filePath);
                dataList.AddRange(lines);
            }
            else
            {
                MessageBox.Show("file is empty");
                return;
            }

            if (dataList.Count > 0)
            {
                foreach (string line in dataList)
                {
                    Console.WriteLine(line);
                    string[] substring = line.Split(delimiter);
                    if (substring.Length >= 4)
                    {
                        alphaList.Add(substring[0]);
                        nameList.Add(substring[1]);
                        voiceTagList.Add(substring[2]);
                        dateList.Add(substring[3]);
                    }
                }
            }
            else
            {
                MessageBox.Show("no data available ");
                return;
            }

            // Parse dateList to DateTime and sort indices based on the dates
            List<int> sortedIndices = dateList
                .Select((date, index) => new { Date = DateTime.Parse(date), Index = index })
                .OrderBy(x => x.Date)
                .Select(x => x.Index)
                .ToList();

            // Reorder all lists based on the sorted indices
            alphaList = sortedIndices.Select(index => alphaList[index]).ToList();
            nameList = sortedIndices.Select(index => nameList[index]).ToList();
            voiceTagList = sortedIndices.Select(index => voiceTagList[index]).ToList();
            dateList = sortedIndices.Select(index => dateList[index]).ToList();


            // Print sorted lists for verification
            for (int i = 0; i < alphaList.Count; i++)
            {
                Console.WriteLine($"{alphaList[i]} | {nameList[i]} | {voiceTagList[i]} | {dateList[i]}");
            }



            //at the end fill up the font list 
            PopulateFontComboBox();

        }




        public class Printer
        {
            private PrintDocument printDocument = new PrintDocument();
            private string documentText;

            public Printer(string text)
            {

                documentText = text;
                printDocument.PrintPage += new PrintPageEventHandler(PrintPage);

            }

            public void Print()
            {
                try
                {
                    printDocument.Print();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occurred while printing: " + ex.Message);
                }
            }

            private void PrintPage(object sender, PrintPageEventArgs ev)
            {
                // Set the font and location for printing
                System.Drawing.Font printFont = new System.Drawing.Font("Arial", 12);
                float leftMargin = ev.MarginBounds.Left;
                float topMargin = ev.MarginBounds.Top;

                // Draw the string on the page
                ev.Graphics.DrawString(documentText, printFont, Brushes.Black, leftMargin, topMargin);
            }
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            PositionGroupBoxes();
        }

        private void PositionGroupBoxes()
        {
            int margin = 10;

            // Top-left corner
            groupBox2.Location = new System.Drawing.Point(margin, margin);

            // Top-right corner
            groupBox1.Location = new System.Drawing.Point(ClientSize.Width - groupBox1.Width - margin, margin);

            // Bottom-left corner
            groupBox3.Location = new System.Drawing.Point(margin, ClientSize.Height - groupBox3.Height - margin);

            // Bottom-right corner
            groupBox4.Location = new System.Drawing.Point(ClientSize.Width - groupBox4.Width - margin, ClientSize.Height - groupBox4.Height - margin);
        }



        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (soundMappings.ContainsKey(e.KeyChar))
            {
                PlayMp3(soundMappings[e.KeyChar]);
            }
        }

        private void PlayMp3(string filePath)
        {
            // Dispose existing objects if they are not null
            if (waveOutDevice != null)
            {
                waveOutDevice.Dispose();
                waveOutDevice = null;
            }
            if (audioFileReader != null)
            {
                audioFileReader.Dispose();
                audioFileReader = null;
            }

            // Initialize the WaveOut and AudioFileReader
            waveOutDevice = new WaveOut();
            audioFileReader = new AudioFileReader(filePath);

            // Register event handler for playback stopped
            waveOutDevice.PlaybackStopped += OnPlaybackStopped;

            // Initialize and play  
            waveOutDevice.Init(audioFileReader);
            waveOutDevice.Play();
        }

        // Event handler for playback stopped
        private void OnPlaybackStopped(object sender, StoppedEventArgs e)
        {
            // Dispose WaveOut and AudioFileReader to release the file handle
            if (waveOutDevice != null)
            {
                waveOutDevice.Dispose();
                waveOutDevice = null;
            }
            if (audioFileReader != null)
            {
                audioFileReader.Dispose();
                audioFileReader = null;
            }
        }

        private void textBox1_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            int currentPosition = MainrichTextBox.SelectionStart;

            if (e.KeyCode == Keys.Left && currentPosition > 0)
            {
                try
                {
                    Console.WriteLine("flag left");
                    char previousChar = MainrichTextBox.Text[currentPosition - 1];
                    Console.WriteLine(previousChar);
                    if (soundMappings.ContainsKey(previousChar))
                    {
                        PlayMp3(soundMappings[previousChar]);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occurred: " + ex.Message);
                }
            }
            else if (e.KeyCode == Keys.Right && currentPosition < MainrichTextBox.Text.Length + 1)
            {
                try
                {
                    Console.WriteLine("flag right");
                    char nextChar = MainrichTextBox.Text[currentPosition];
                    Console.WriteLine(nextChar);
                    if (soundMappings.ContainsKey(nextChar))
                    {
                        PlayMp3(soundMappings[nextChar]);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occurred: " + ex.Message);
                }
            }
        }


        private void startparBtn_Click(object sender, EventArgs e)
        {
            string text = MainrichTextBox.Text;
            int cursorPos = MainrichTextBox.SelectionStart;

            // Find the [PARA] marker before the current cursor
            int previousMarker = text.LastIndexOf("[PARA]", cursorPos - 1);

            if (previousMarker >= 0)
            {
                MainrichTextBox.SelectionStart = previousMarker;
            }
            else
            {
                // If no marker found, go to start
                MainrichTextBox.SelectionStart = 0;
            }

            MainrichTextBox.ScrollToCaret();
            MainrichTextBox.Focus();
        }



        private void endparBtn_Click(object sender, EventArgs e)
        {
            string text = MainrichTextBox.Text;
            int cursorPos = MainrichTextBox.SelectionStart;

            int lastMarkerBeforeCursor = text.LastIndexOf("[PARA]", cursorPos - 1);

            if (lastMarkerBeforeCursor >= 0)
            {
                // Move cursor to the END of that marker
                int newPosition = lastMarkerBeforeCursor + "[PARA]".Length;
                MainrichTextBox.SelectionStart = newPosition;
                MainrichTextBox.ScrollToCaret();
                MainrichTextBox.Focus();
            }
            else
            {
                // No marker found, go to start
                MainrichTextBox.SelectionStart = 0;
                MainrichTextBox.ScrollToCaret();
                MainrichTextBox.Focus();
            }
        }





        private void startsentBtn_Click(object sender, EventArgs e)
        {
            MainrichTextBox.Focus();
            // Get the current cursor position
            int cursorPosition = MainrichTextBox.SelectionStart;

            // Find the position of the previous period before the cursor
            int startOfSentence = MainrichTextBox.Text.LastIndexOf('.', cursorPosition - 1);

            // If the cursor is at the beginning of a sentence or there is no period found
            if (cursorPosition == startOfSentence + 1 || startOfSentence == -1)
            {
                // Find the period before the current sentence
                startOfSentence = MainrichTextBox.Text.LastIndexOf('.', startOfSentence - 1);

                // If another period is found, move to the character after it
                if (startOfSentence != -1)
                {
                    startOfSentence += 1; // Move to the first character after the period (and space)
                }
                else
                {
                    // If no previous period is found, move to the start of the text
                    startOfSentence = 0;
                }
            }
            else
            {
                // Move to the first character after the found period (and space)
                startOfSentence += 1;
            }

            // Set the cursor to the beginning of the sentence
            MainrichTextBox.SelectionStart = startOfSentence;
            MainrichTextBox.SelectionLength = 0;

            // Ensure the TextBox has focus so the cursor is visible
            MainrichTextBox.Focus();
        }


        private void Recording()
        {
            try
            {
                waveIn = new WaveInEvent();
                waveIn.WaveFormat = new WaveFormat(44100, 1);
                waveIn.DataAvailable += OnDataAvailable;
                waveIn.RecordingStopped += onRecordingStopped;


                writer = new WaveFileWriter(outputFilePath, waveIn.WaveFormat);

                waveIn.StartRecording();

                label1.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred while starting the recording: " + ex.Message);
            }
        }

        private void stopRecording()
        {
            try
            {
                if (waveIn != null)
                {
                    waveIn.StopRecording();
                    label1.Visible = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred while stopping the recording: " + ex.Message);
            }
        }



        private void button7_Click(object sender, EventArgs e)
        {
            string bookChar = "ˑ";
            int textPlace = MainrichTextBox.SelectionStart;
            MainrichTextBox.Text = MainrichTextBox.Text.Substring(0, textPlace) + bookChar + MainrichTextBox.Text.Substring(textPlace);
            MainrichTextBox.SelectionStart = textPlace;
            MainrichTextBox.Focus();

            MessageBox.Show("push a key");

            isWaitingForKey = true; // <---- Activate waiting mode
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // Stop default newline
                int position = MainrichTextBox.SelectionStart;
                MainrichTextBox.Text = MainrichTextBox.Text.Insert(position, "[PARA]\n");
                MainrichTextBox.SelectionStart = position + "[PARA]\n".Length;
            }

            if (isWaitingForKey)
            {
                if (e.KeyCode >= Keys.A && e.KeyCode <= Keys.Z)
                {
                    HandleAlphabetKeyDown(e.KeyCode);
                }
                else
                {
                    MessageBox.Show("Wrong key, try again!");
                }
            }
        }

        private void HandleAlphabetKeyDown(Keys key)
        {
            if (!isRecording)
            {
                isRecording = true;
                isWaitingForKey = false; // <--- Exit waiting mode after key
                MainrichTextBox.ReadOnly = true;

                Recording();

                int textPlace = MainrichTextBox.SelectionStart;
                MainrichTextBox.Text = MainrichTextBox.Text.Substring(0, textPlace) + key + MainrichTextBox.Text.Substring(textPlace);
                MainrichTextBox.SelectionStart = textPlace + 1;
                MainrichTextBox.Focus();
            }

            
        }

        private void textBox1_KeyUp(object sender, KeyEventArgs e)
        {
            if (isRecording)
            {
                stopRecording();
                isRecording = false;
                MainrichTextBox.ReadOnly = false;
            }
        }



        private void exitBtn_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }





        private void Form1_keyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsLetter(e.KeyChar))
            {
                string pressedCharOpen = e.KeyChar.ToString().ToUpper();
                bool keyFound = false;

                for (int i = 0; i < alphaList.Count; i++) // Loop through all items in alphaList
                {
                    if (pressedCharOpen == alphaList[i].ToUpper())
                    {
                        try
                        {
                            keyFound = true; // Mark that the key was found
                            string addressToOpen = allDocxFiles[i];
                            PlayMp3(voiceTagList[i]);
                            MainrichTextBox.Text = openWordDocument(addressToOpen);
                            textBox2.Text = nameList[i].Replace(".docx", "");
                            keyPressTcs.TrySetResult(true);

                            break; // Exit the loop since we found the key
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"an error occured{ex.Message}");
                        }

                    }
                }

                if (!keyFound)
                {
                    MessageBox.Show("Wrong key, try again!");
                }

                Console.WriteLine("keydown" + pressedCharOpen);
            }

            this.KeyPress -= new KeyPressEventHandler(Form1_keyPress);

        }



        // Dummy implementations of the methods to make the code compile



        private void OnDataAvailable(object sender, WaveInEventArgs e)
        {
            if (writer != null)
            {
                writer.Write(e.Buffer, 0, e.BytesRecorded);
                writer.Flush();
            }
        }

        private void onRecordingStopped(object sender, StoppedEventArgs e)
        {
            if (writer != null)
            {
                writer.Dispose();
                writer = null;
            }
            if (waveIn != null)
            {
                waveIn.Dispose();
                waveIn = null;
            }
            if (e.Exception != null)
            {
                MessageBox.Show("An error occurred during recording: " + e.Exception.Message);
            }
            else
            {
                MessageBox.Show("Recording saved in " + outputFilePath);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {

            if (isAlarmActive)
            {
                // Stop the alarm
                beepTimer.Stop();
                isAlarmActive = false;
            }
            else
            {
                // Start the alarm
                beepTimer.Start();
                isAlarmActive = true;
            }


            this.gotoBtn.Visible = !this.gotoBtn.Visible;
            this.bookMarkBtn.Visible = !this.bookMarkBtn.Visible;
        }

        private void gotoBtn_Click(object sender, EventArgs e)
        {
            this.startparBtn.Visible = !this.startparBtn.Visible;
            this.fontGroupBtn.Visible = !this.fontGroupBtn.Visible;
            endsentBtn.Visible = !endsentBtn.Visible;
            this.endparBtn.Visible = !this.endparBtn.Visible;
            startsentBtn.Visible = !startsentBtn.Visible;
            this.gotoBmark.Visible = !this.gotoBmark.Visible;
            this.searchBtn.Visible = !this.searchBtn.Visible;
            this.insertBtn.Visible = !this.insertBtn.Visible;
        }

        private void gotoBmark_Click(object sender, EventArgs e)
        {
            string Bname;
            textBoxValue = MainrichTextBox.Text;
            for (int i = 0; i < textBoxValue.Length; i++)
            {
                if (i.ToString() == "ˑ")
                {
                    Bname = (i++).ToString();
                    listBox1.Items.Add(Bname);

                }
            }


        }


        private void CreateWordDocument(string filePath)
        {

            Word.Application wordApp = null;
            Word.Document wordDoc = null;
            try
            {
                wordApp = new Word.Application();
                wordDoc = wordApp.Documents.Add();

                wordDoc.Content.Text = MainrichTextBox.Text;
                wordDoc.SaveAs2(filePath);
                MessageBox.Show($"Word document created: {filePath}");
            }
            catch (Exception er)
            {
                MessageBox.Show($"An error occurred: {er.Message}");
            }
            finally
            {
                if (wordDoc != null)
                {
                    wordDoc.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDoc);
                }
                if (wordApp != null)
                {
                    wordApp.Quit(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                }
            }



        }


        private void Form1_keyDown(object sender, KeyEventArgs e)
        {
            Console.WriteLine("flag key down");

            if (e.KeyCode >= Keys.A && e.KeyCode <= Keys.Z)
            {
                pressedChar = e.KeyCode.ToString();
                Console.WriteLine("keydown" + pressedChar);
                handelRecording(e.KeyCode);

            }
            else
            {
                MessageBox.Show("Wrong key, try again!");
            }


        }

        private void handelRecording(Keys key)
        {
            if (!isRecording)
            {
                isRecording = true;


                Recording();

            }

        }

        private void Form1_keyUp(object sender, KeyEventArgs e)
        {

            if (isRecording)
            {
                stopRecording();
                isRecording = false;

            }
            keyPressTcs.TrySetResult(true);
        }

        private TaskCompletionSource<bool> keyPressTcs;


        private async System.Threading.Tasks.Task WaitForKeyPressAsync()
        {
            keyPressTcs = new TaskCompletionSource<bool>();
            await keyPressTcs.Task;
        }


        // Ensure pressedChar is initialized
        private void saveBtn_Click(object sender, EventArgs e)
        {

            if (string.IsNullOrWhiteSpace(textBox2.Text))
            {
                DialogResult result = MessageBox.Show("The file is new. Do you want to save it?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (result == DialogResult.Yes)
                {
                    textBox2.Enabled = true;
                    textBox2.Focus();

                    this.textBox2.KeyDown -= new System.Windows.Forms.KeyEventHandler(this.textBox2_KeyDown);
                    this.textBox2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox2_KeyDown);

                }
                else
                {
                    MessageBox.Show("You didn't specify a file name.");
                }
            }

            string wordFileName = $"{textBox2.Text}.docx";
            //string wordFilePath = Path.Combine(DocumentsPath, "wordFiles", wordFileName);
            
            //string voiceFilePath = Path.Combine(DocumentsPath, "subDatas", "voices", wordFileName);

            //bool wordFileExists = File.Exists(wordFilePath);
            //bool voiceFileExists = File.Exists(voiceFilePath);
            /*
            if (wordFileExists && voiceFileExists)
            {
                SaveTextToWordFile(wordFilePath, MainrichTextBox.Text);
                MessageBox.Show("Data saved successfully.");
            }
            else
            {
                DialogResult result = MessageBox.Show("The Word document or voice file doesn't exist. Do you want to create a new file?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (result == DialogResult.Yes)
                {
                    // Assuming saveAsBtn_Click is intended to handle creating new files
                    saveAsBtn_Click(sender, e);
                }
            }
            */
        }

        private void SaveTextToWordFile(string filePath, string text)
        {
            Word.Application wordApp = new Word.Application();
            Word.Document wordDoc = null;

            try
            {
                wordDoc = wordApp.Documents.Add();
                wordDoc.Content.Text = text;
                wordDoc.SaveAs2(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving document: {ex.Message}");
            }
            finally
            {
                wordDoc?.Close();
                wordApp.Quit();
            }
        }

        private void PopulateListBoxOpen() //to be change haie
        {
            listBox1.Items.Clear();

            for (int i = 0; i < allDocxFiles.Count; i++)
            {

                listBox1.Items.Add(Path.GetFileName(allDocxFiles[i]));
            }
            this.listBox1.KeyDown -= new System.Windows.Forms.KeyEventHandler(this.listBox1_keyDown);
            this.listBox1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.listBox1_keyDown);
        }

        private void ListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            // Get the selected index
            int index = listBox1.SelectedIndex;

            // Check if an item is selected
            if (index != -1)
            {
                string selectedFileName = listBox1.SelectedItem.ToString();

                Console.WriteLine("Selected index: " + index);
                Console.WriteLine("Selected item: " + selectedFileName);

                // Check if the index is within the range of voiceTagList
                if (index >= 0 && index < voiceTagList.Count)
                {
                    // Replace ".docx" with an empty string in the file name
                    string fileNameWithoutExtension = selectedFileName.Replace(".docx", "");

                    string voiceTag = Path.Combine(userVoicePath, $"{fileNameWithoutExtension}.wav");

                    Console.WriteLine("Playing voice tag: " + voiceTag);
                 
                    try
                    {
                        // Check if the voice tag file exists
                        if (File.Exists(voiceTag))
                        { 
                                PlayMp3(voiceTag);
                        }
                        else
                        {
                            MessageBox.Show("Voice file doesn't exist", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    catch
                    {
                        MessageBox.Show("voice doesnt exist");
                    }

                }
                else
                {
                    Console.WriteLine("Index out of range for voiceTagList");
                }
            }
            else
            {
                Console.WriteLine("No item is selected");
            }
                 
        }



        private void listBox1_keyDown(object sender, KeyEventArgs e)
        {

            if (e.KeyCode == Keys.Delete)
            {
                int index = listBox1.SelectedIndex;

                if (index != -1) // Check if an item is selected
                {
                    string wordName = listBox1.SelectedItem.ToString();

                    // Ensure index is within valid range for allDocxFiles
                    if (index >= 0 && index < allDocxFiles.Count)
                    {
                        string fileNameWithoutExtension = wordName.Replace(".docx", "");

                        string voiceTag = Path.Combine(userVoicePath, $"{fileNameWithoutExtension}.wav");
                        string wordFilePath = Path.Combine(wordPath, wordName);

                        // Try to delete the voice file
                        if (File.Exists(voiceTag))
                        {
                            try
                            {
                                                              
                                // Delete the voice file
                                File.Delete(voiceTag);
                                MessageBox.Show($"Voice file deleted: {voiceTag}");
                            }
                            catch (IOException ex)
                            {
                                MessageBox.Show($"Error: The file is in use and cannot be deleted. Please close any program that might be using it.\nDetails: {ex.Message}");
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Error deleting voice file: {ex.Message}");
                            }
                        }
                        else
                        {
                            MessageBox.Show("Voice file does not exist.");
                        }

                        // Try to delete the Word file
                        if (File.Exists(wordFilePath))
                        {
                            try
                            {
                                // Delete the Word file
                                File.Delete(wordFilePath);
                                MessageBox.Show($"Word file deleted: {wordFilePath}");
                            }
                            catch (IOException ex)
                            {
                                MessageBox.Show($"Error: The file is in use and cannot be deleted. Please close any program that might be using it.\nDetails: {ex.Message}");
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Error deleting Word file: {ex.Message}");
                            }
                        }
                        else
                        {
                            MessageBox.Show("Word file does not exist.");
                        }

                        // Update the dataList by removing the item at the current index
                        try
                        {
                            dataList.RemoveAt(index);

                            // Save updated dataList to file
                            using (StreamWriter writer = new StreamWriter(filePathData, false)) // false to overwrite the file
                            {
                                foreach (var item in dataList)
                                {
                                    writer.WriteLine(item); // Write each item on a new line
                                }
                            }
                            Console.WriteLine("Data list updated and saved.");

                            // Remove the item from the ListBox
                            listBox1.Items.RemoveAt(index);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Error updating data list: {ex.Message}");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Index out of range. Please select a valid item.");
                    }
                }
                else
                {
                    MessageBox.Show("No item selected. Please select an item to delete.");
                }
            }


            if (e.KeyCode == Keys.Enter) { //open the file 
                int index = listBox1.SelectedIndex;

                if (!string.IsNullOrEmpty(MainrichTextBox.Text))
                {
                    
                    DialogResult result = MessageBox.Show("Are you sure you want to discard the changes?", "Alert", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        // Execute the logic for discarding changes
                        // (e.g., clear the TextBox, reset the form, etc.)
                    }
                    else
                    {

                        this.saveBtn.Click -= saveBtn_Click;
                        this.saveBtn.Click += saveBtn_Click;
                    }
                }


                if (index != ListBox.NoMatches)// Check if an item is clicked
                {
                    Console.WriteLine(index);
                    wordName = listBox1.SelectedItem.ToString();
                    
                    if (index >= 0 && index < allDocxFiles.Count)
                    {
                        wordName.Replace(".docx", "");
                        string addressToOpen = Path.Combine(wordPath, wordName);
                        var wordApp = new Microsoft.Office.Interop.Word.Application();
                        Document doc = wordApp.Documents.Open(addressToOpen, ReadOnly: false, Visible: false);
                        doc.Range().Copy(); // Standard copy preserves editable RTF

                        // Get data from clipboard as RTF
                        IDataObject data = Clipboard.GetDataObject();
                        if (data.GetDataPresent(DataFormats.Rtf))
                        {
                            string rtfContent = (string)data.GetData(DataFormats.Rtf);
                            // Load full editable formatted content into RichTextBox
                            MainrichTextBox.Rtf = rtfContent;
                        }

                        // Cleanup
                        doc.Close(false);
                        textBox2.Text = wordName;
                    }
                    else
                    {
                        Console.WriteLine("Index out of range for voiceTagList");
                    }

                }
            }

            if (e.KeyCode == Keys.Space)
            {
                bool removableDriveFound = false;

                DriveInfo[] drives = DriveInfo.GetDrives();
                foreach (DriveInfo drive in drives)
                {
                    if (drive.DriveType == DriveType.Removable && drive.IsReady)
                    {
                        removableDriveFound = true;
                        int index = listBox1.SelectedIndex;

                        if (index >= 0 && index < nameList.Count)
                        {
                            Console.WriteLine($"Found removable drive: {drive.Name}");

                            string sourceFilePath = Path.Combine(wordPath, nameList[index]);
                            string destinationFilePath = Path.Combine(drive.Name, nameList[index]);

                            try
                            {
                                File.Copy(sourceFilePath, destinationFilePath, overwrite: true);
                                MessageBox.Show($"File copied to {destinationFilePath}");
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Error copying file: {ex.Message}");
                            }
                        }
                        else
                        {
                            MessageBox.Show("No file selected or invalid selection.");
                        }
                    }
                }

                if (!removableDriveFound)
                {
                    MessageBox.Show("Cannot find any external storages");
                }
            }
            // Detect external storage drives

            if(e.KeyCode == Keys.Back)
            {
                MessageBox.Show("enter the new name of the file");
                textBox2.Enabled = true;
                textBox2.Focus();
                this.textBox2.KeyDown -= new System.Windows.Forms.KeyEventHandler(this.textBox_keyDown_Rename);
                this.textBox2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox_keyDown_Rename);
                
            }

        }


        private void textBox_keyDown_Rename(object sender, KeyEventArgs e)
        {
            if (string.IsNullOrEmpty(textBox2.Text))
            {
                return;
            }
            if(e.KeyCode == Keys.Enter)
            {
                string wordname = listBox1.SelectedItem.ToString();
                
                string oldWordFile = Path.Combine(wordPath, wordname);
                string newWordFile = Path.Combine(wordPath, $"{textBox2.Text}.docx");

                wordname = wordname.Replace(".docx", "");
                string oldVoiceTag = Path.Combine(userVoicePath, $"{wordname}0.wav");
                string newVoiceTag = Path.Combine(userVoicePath, "voices", $"{textBox2.Text}0.wav");

                try
                {
                    File.Move(oldWordFile, newWordFile);
                    File.Move(oldVoiceTag, newVoiceTag);
                    MessageBox.Show("rename was seccesful");

                    this.textBox2.KeyDown -= new System.Windows.Forms.KeyEventHandler(this.textBox_keyDown_Rename);
                    textBox2.Enabled = false;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"error while renameing the files: {ex.Message}");
                }
                
            }
      

        }
        private string openWordDocument(string filePath)
        {
            Word.Application wordApp = null;
            Word.Document wordDoc = null;
            string textBoxContent = string.Empty;
            try
            {
                wordApp = new Word.Application();
                wordDoc = wordApp.Documents.Open(filePath);
                //wordApp.Visible = true;
                textBoxContent = wordDoc.Content.Text;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"an error occurred:{ex.Message}");
            }
            finally
            {
                wordDoc.Close();
                wordApp.Quit();

            }
            return textBoxContent;
        }

        

        private void BeepTimer_Tick(object sender, EventArgs e)
        {
            SystemSounds.Beep.Play(); // Play the beep sound
        }

        private void fileMagament_Click(object sender, EventArgs e)
        {

            if (isAlarmActive)
            {
                // Stop the alarm
                beepTimer.Stop();
                isAlarmActive = false;
                
                listBox1.Items.Clear();
            }
            else
            {
                // Start the alarm
                beepTimer.Start();
                isAlarmActive = true;
                this.listBox1.SelectedIndexChanged += new EventHandler(this.ListBox1_SelectedIndexChanged);
                listBox1.Items.Clear();
                PopulateListBoxOpen();
                listBox1.Focus();
            }
            
            DialogResult result = MessageBox.Show("do you want to use random accsess or regular access", "info");

            ToggleButtonVisibilityFileManage();
            ManageFiles();
        }


       
        private void ToggleButtonVisibilityFileManage()
        {
            this.gotoBtn.Visible = !this.gotoBtn.Visible;
            this.saveAsBtn.Visible = !this.saveAsBtn.Visible;
            this.printBtn.Visible = !this.printBtn.Visible;
            this.savePdfBtn.Visible = !this.savePdfBtn.Visible;
            this.searchBtn.Visible = !this.searchBtn.Visible;
        }

        private void ManageFiles()
        {
            
            this.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.Form1_keyPress);
            this.KeyPress -= new System.Windows.Forms.KeyPressEventHandler(this.Form1_keyPress);
        }



        private void textBox2_keyPress(object sender, KeyPressEventArgs e)
        {
            if (soundMappings.ContainsKey(e.KeyChar))
            {
                PlayMp3(soundMappings[e.KeyChar]);
            }
        }

        private void printBtn_Click(object sender, EventArgs e)
        {
            Printer printer = new Printer(MainrichTextBox.Text);
            printer.Print();
        }



        private void saveAsBtn_Click(object sender, EventArgs e)
        {
            MessageBox.Show("clicked");
            listBox1.Focus();// Focus on the list box


            textBox2.Enabled = true;
            textBox2.Focus();

            this.textBox2.KeyDown -= new System.Windows.Forms.KeyEventHandler(this.textBox2_KeyDown);
            this.textBox2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox2_KeyDown);
        }

        private async void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && textBox2.Text != null)
            {
                DateTime currentDateTime = DateTime.Now;
                string dateTimeString = currentDateTime.ToString("yyyy-MM-dd HH:mm:ss");

                wordName = textBox2.Text;

                
                textBox2.Enabled = false;
                // MessageBox.Show("Press a key", "Press key", MessageBoxButtons.OK, MessageBoxIcon.Information);




                

                

                
                // Attach new handlers
                this.KeyDown += new KeyEventHandler(Form1_keyDown);
                

                outputFilePath = Path.Combine(userVoicePath, $"{wordName}.wav");
                this.KeyUp += new KeyEventHandler(Form1_keyUp);

                this.Enabled = true;
                await WaitForKeyPressAsync();
                 

                if (string.IsNullOrEmpty(pressedChar) || !char.IsLetterOrDigit(pressedChar[0]))
                {
                    Console.WriteLine("Ignored key: " + pressedChar);
                    return;
                }

                // Check if any item in alphaList starts with pressedChar followed by a possible index
                if (alphaList.Any(item => item.StartsWith(pressedChar)))
                {
                    // Find the last index of items starting with pressedChar in alphaList
                    int lastIndex = alphaList.FindLastIndex(item => item.StartsWith(pressedChar));

                    // Extract the current index from the last matching item
                    string lastItem = alphaList[lastIndex];
                    int startIndex = pressedChar.Length;
                    int currentIndex = 0;

                    // Check if there is an index to extract
                    if (lastItem.Length > startIndex && lastItem[startIndex] == '(' && lastItem.EndsWith(")"))
                    {
                        string indexString = lastItem.Substring(startIndex + 1, lastItem.Length - startIndex - 2);
                        if (int.TryParse(indexString, out int parsedIndex))
                        {
                            currentIndex = parsedIndex;
                        }
                    }

                    // Append (currentIndex + 1) to pressedChar
                    pressedChar = $"{pressedChar}({currentIndex + 1})";
                }

                // Add pressedChar to alphaList if it's not already there
                if (!alphaList.Contains(pressedChar))
                {
                    alphaList.Add(pressedChar);
                }
                Console.WriteLine("Before saving char: " + pressedChar);



                // Write data to file with appended pressedChar
                try
                {
                    using (StreamWriter writer = new StreamWriter(filePathData, true))
                    {
                        writer.WriteLine($"{pressedChar}|{wordName}.docx|{outputFilePath}|{dateTimeString}");
                    }

                    // Re-enable the form


                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error saving data: {ex.Message}");
                    this.Enabled = true;
                }
                finally
                {
                    this.KeyDown -= new KeyEventHandler(Form1_keyDown); ;
                    this.KeyUp -= new KeyEventHandler(Form1_keyUp);
                    pressedChar = string.Empty;


                    textBox2.Text = wordName;
                    this.textBox2.KeyDown -= new System.Windows.Forms.KeyEventHandler(this.textBox2_KeyDown);

                }

            }
        }

        
        private void fontGroupBtn_Click(object sender, EventArgs e)
        {
            if (isAlarmActive)
            {
                // Stop the alarm
                beepTimer.Stop();
                isAlarmActive = false;
                removeItemsListBox();
                listBox1.Items.Clear();
            }
            else
            {
                // Start the alarm
                beepTimer.Start();
                isAlarmActive = true;

               
            }

            insertBtn.Visible = !insertBtn.Visible;
            gotoBtn.Visible = !gotoBtn.Visible;
            searchBtn.Visible = !searchBtn.Visible;
            fileManagment.Visible = !fileManagment.Visible;
            fontComboBox.Visible = !fontComboBox.Visible;
            fontSizeComboBox.Visible = !fontSizeComboBox.Visible;
            changeColor.Visible = !changeColor.Visible;
            BIUbtn.Visible = !BIUbtn.Visible;
            PBObtn.Visible = !PBObtn.Visible;
            statusBtn.Visible = !statusBtn.Visible;
            
        }
        
        private void removeItemsListBox()
        {
            
            for(int i = 0; i <= listBox1.Items.Count - 1; i++)
            {
                listBox1.Items.Remove(i);
            }
            
        }




        private void changeColor_Click(object sender, EventArgs e)
        {
            MessageBox.Show("enter chararcter between ORGBCMYW");
            this.KeyPress += new KeyPressEventHandler(this.Form1_keyPress_color);
        }

        private void Form1_keyPress_color(object sender , KeyPressEventArgs e)
        {
            
            string getColor = e.KeyChar.ToString().ToUpper();
            switch (getColor)
            {
                case "O":
                    MainrichTextBox.ForeColor = System.Drawing.Color.Orange;
                    MessageBox.Show("color changed");
                    break;
                case "R":

                    MainrichTextBox.ForeColor = System.Drawing.Color.Red;
                    break;
                case "G":

                    MainrichTextBox.ForeColor = System.Drawing.Color.Green;
                    break;
                case "B":

                    MainrichTextBox.ForeColor = System.Drawing.Color.Blue;
                    break;
                case "W":

                    MainrichTextBox.ForeColor = System.Drawing.Color.White;
                    break;
                case "M":

                    MainrichTextBox.ForeColor = System.Drawing.Color.Maroon;
                    break;
                case "Y":

                    MainrichTextBox.ForeColor = System.Drawing.Color.Yellow;
                    break;
                case "C":

                    MainrichTextBox.ForeColor = System.Drawing.Color.Black;
                    break;
                default:
                    MessageBox.Show("invalid color key");
                    break;
            }
            //keyPressTcs.SetResult(true);
            this.KeyPress -= new KeyPressEventHandler(this.Form1_keyPress_color);

        }

        private void BIUbtn_Click(object sender, EventArgs e)
        {
            this.KeyPress += new KeyPressEventHandler(this.Form1_keyPress_style);
        }

        private void Form1_keyPress_style(object sender , KeyPressEventArgs e)
        {
            string getStyle = e.KeyChar.ToString().ToUpper();
            switch (getStyle)
            {
                case "B":
                    ToggleFontStyle(FontStyle.Bold);
                    break;
                case "I":
                    ToggleFontStyle(FontStyle.Italic);
                    break;
                case "U":
                    ToggleFontStyle(FontStyle.Underline);
                    break;
                default:
                    MessageBox.Show("wrong style key");
                    break;
            }
            this.KeyPress -= new KeyPressEventHandler(this.Form1_keyPress_style);

        }
        private void ToggleFontStyle(FontStyle style)
        {
            System.Drawing.Font currentFont = this.MainrichTextBox.Font;
            FontStyle newFontStyle;

            // Check if the style is already applied
            if (currentFont.Style.HasFlag(style))
            {
                // Remove the style if it's already applied
                newFontStyle = currentFont.Style & ~style;
            }
            else
            {
                // Add the style if it's not applied
                newFontStyle = currentFont.Style | style;
            }

            this.MainrichTextBox.Font = new System.Drawing.Font(currentFont, newFontStyle);
        }



        private void PopulateFontComboBox()
        {
            fontComboBox.Items.Clear();
            
            foreach (string fontName in fontNames)
            {
                // Optional: verify that the font is installed
                if (FontFamily.Families.Any(f => f.Name.Equals(fontName, StringComparison.OrdinalIgnoreCase)))
                {
                    fontComboBox.Items.Add(fontName);
                }
                else
                {
                    Console.WriteLine($"Font not installed: {fontName}");
                }

                if (fontComboBox.Items.Count > 0)
                {
                    fontComboBox.SelectedIndex = 0; // Select first font
                }

            }
        }


        private void FontComboBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            e.DrawBackground();

            if (e.Index >= 0)
            {
                string fontName = fontComboBox.Items[e.Index].ToString();
                using (System.Drawing.Font font = new System.Drawing.Font(fontName, e.Font.Size, FontStyle.Regular, GraphicsUnit.Point))
                {
                    e.Graphics.DrawString(fontName, font, Brushes.Black, e.Bounds);
                }
            }

            e.DrawFocusRectangle();
        }


        private void FontComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (fontComboBox.SelectedItem != null)
            {
                string selectedFont = fontComboBox.SelectedItem.ToString();
                ChangeSelectedFontFamily(selectedFont);
            }
        }


        private void fontSizeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (float.TryParse(fontSizeComboBox.SelectedItem?.ToString(), out float selectedSize))
            {
                ChangeSelectedFontSize(selectedSize);
            }
            else
            {
                MessageBox.Show("error");
            }
        }


        private void ChangeSelectedFontFamily(string newFontFamily)
        {
            if (MainrichTextBox.SelectionFont != null)
            {
                System.Drawing.Font currentFont = MainrichTextBox.SelectionFont;
                System.Drawing.Font newFont = new System.Drawing.Font(newFontFamily, currentFont.Size, currentFont.Style);
                MainrichTextBox.SelectionFont = newFont;
            }
        }

        private void ChangeSelectedFontSize(float newSize)
        {
            if (MainrichTextBox.SelectionFont != null)
            {
                System.Drawing.Font currentFont = MainrichTextBox.SelectionFont;
                System.Drawing.Font newFont = new System.Drawing.Font(currentFont.FontFamily, newSize, currentFont.Style);
                MainrichTextBox.SelectionFont = newFont;
            }
        }


        private void saveFontFile(string selectedFont)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(fontFilePath, true))
                {
                    writer.WriteLine($"{pressedChar}|{selectedFont}|{outputFilePath}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while saving the font file: {ex.Message}");
            }
            finally
            {
                this.KeyDown -= Form1_keyDown;
                this.KeyUp -= Form1_keyUp;
            }
        }

       
        private void ComboBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Suppress keypresses for alphanumeric keys
            if (char.IsLetterOrDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void statusBtn_Click(object sender, EventArgs e)
        {
            MessageBox.Show($"{MainrichTextBox.Font}{MainrichTextBox.Font.Style}{MainrichTextBox.ForeColor}");
        }

        private void endsentBtn_Click(object sender, EventArgs e)
        {
                // Get the current cursor position
                int cursorPosition = MainrichTextBox.SelectionStart;

                // Find the position of the next period after the cursor
                int endOfSentence = MainrichTextBox.Text.IndexOf('.', cursorPosition);

                // If a period is found
                if (endOfSentence != -1)
                {
                // Move to the character after the period (accounting for the period and space)
                MainrichTextBox.SelectionStart = endOfSentence + 1;

                    // If there is a space after the period, move the cursor after the space
                    if (endOfSentence + 1 < MainrichTextBox.Text.Length && MainrichTextBox.Text[endOfSentence + 1] == ' ')
                    {
                    MainrichTextBox.SelectionStart++;
                    }
                }
                else
                {
                // If no period is found, move to the end of the text
                MainrichTextBox.SelectionStart = MainrichTextBox.Text.Length;
                }

            // Ensure the TextBox has focus so the cursor is visible
            MainrichTextBox.Focus();
            
        }

        private void settingBtn_Click(object sender, EventArgs e)
        {
            setting_form sf = new setting_form();
            sf.Show();

        }

        private void rtlBtn_Click(object sender, EventArgs e)
        {
            MainrichTextBox.RightToLeft = MainrichTextBox.RightToLeft == RightToLeft.Yes ? RightToLeft.No : RightToLeft.Yes;
        }

        private void aligenmentBtn_Click(object sender, EventArgs e)
        {
            switch (MainrichTextBox.SelectionAlignment)
            {
                case HorizontalAlignment.Left:
                    MainrichTextBox.SelectionAlignment = HorizontalAlignment.Center;
                    break;

                case HorizontalAlignment.Center:
                    MainrichTextBox.SelectionAlignment = HorizontalAlignment.Right;
                    break;

                case HorizontalAlignment.Right:
                    MainrichTextBox.SelectionAlignment = HorizontalAlignment.Left;
                    break;
            }
        }

        private void btnSetSpacing_Click(object sender, EventArgs e)
        {
            if (float.TryParse(spacingValue.Text, out float lineSpacing))
            {
                RichTextBoxHelper.SetLineSpacing(MainrichTextBox, lineSpacing); // Use the value entered
            }
            else
            {
                MessageBox.Show("Please enter a valid number.");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var wordApp = new Microsoft.Office.Interop.Word.Application();
            Document doc = wordApp.Documents.Open(@"C:\Users\Artin\Desktop\New Microsoft Word Document (6).docx", ReadOnly: false, Visible: false);

            // Select all content in Word document
            doc.Range().Copy(); // Standard copy preserves editable RTF

            // Get data from clipboard as RTF
            IDataObject data = Clipboard.GetDataObject();
            if (data.GetDataPresent(DataFormats.Rtf))
            {
                string rtfContent = (string)data.GetData(DataFormats.Rtf);
                // Load full editable formatted content into RichTextBox
                MainrichTextBox.Rtf = rtfContent;
            }

            // Cleanup
            doc.Close(false);

        }
        private void ExportPagesToWord(List<string> pages)
        {

                var wordApp = new Word.Application();
                wordApp.Visible = false;
                var doc = wordApp.Documents.Add();

                string[] sections = MainrichTextBox.Text.Split(new string[] { "[PAGE_BREAK]" }, StringSplitOptions.None);

                foreach (string section in sections)
                {
                    Word.Paragraph para = doc.Content.Paragraphs.Add();
                    para.Range.Text = section.Trim();

                    if (section != sections.Last())
                    {
                        para.Range.InsertBreak(Word.WdBreakType.wdPageBreak);
                    }
                }

                doc.SaveAs2(@"C:\Users\Artin\Desktop\MyDocumentWithPageBreaks.docx");
                doc.Close();
                wordApp.Quit();
       

        }

        private List<string> GetPagesFromRichTextBox()
        {
            List<string> pages = new List<string>();

            string allText = MainrichTextBox.Text;

            // Example: split every 1000 characters (dummy pagination)
            int charsPerPage = 1000;
            for (int i = 0; i < allText.Length; i += charsPerPage)
            {
                string pageText = allText.Substring(i, Math.Min(charsPerPage, allText.Length - i));
                pages.Add(pageText);
            }

            return pages;
        }


        private void exportBtn_Click(object sender, EventArgs e)
        {
            // Step 1: Get your pages from your UI (from RichTextBox(es))
            List<string> pages = GetPagesFromRichTextBox();

            // Step 2: Export them to Word
            ExportPagesToWord(pages);

            MessageBox.Show("Exported to Word successfully!");
        }

        private void spacingValue_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnInsertPageBreak_Click(object sender, EventArgs e)
        {
            int pos = MainrichTextBox.SelectionStart;
            MainrichTextBox.Text = MainrichTextBox.Text.Insert(pos, "[PAGE_BREAK]\n");
        }


    }
}


    

    



