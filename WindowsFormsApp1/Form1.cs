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
        
        string fontFilePath;
        string filePathData;
        private object lastSelectedItem;
        string pressedChar;
        string wordName;
        string wordPath;
        Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
        int projecktNummber;
        string DocumentsPath;
        static int newVoice = 0;
        //static int newFile = 0;
        private WaveInEvent waveIn;
        private WaveFileWriter writer;
        string soundMappingFiles;
        private IWavePlayer waveOutDevice;
        private AudioFileReader audioFileReader;
        private Dictionary<char, string> soundMappings;
        private Button saveBtn;
        string textBoxValue;
        //string projektName;
        //string projektPath;
        private string outputFilePath;
        string mainFolderPath;
        string wordFilePath;
        private string mainFolder = "subDatas";
        private Timer beepTimer;
        private string[] subfolder = { "voices", "datafiles", "sound mapping", "pdf folder" };
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

        // opening save the style in the word 
        //paging 
        //inset add break page 
        //
        public Form1()
        {
            this.KeyPreview = true;
            InitializeComponent();
            PopulateFontComboBox();
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
        {'ا', @"C:\Users\Artin\Documents\voice files project\alef.mp3"},
        {'ب', @"C:\Users\Artin\Documents\voice files project\be.mp3"},
        {'پ', @"C:\Users\Artin\Documents\voice files project\pe.mp3"},
        {'ت', @"C:\Users\Artin\Documents\voice files project\te.mp3"},
        {'ث', @"C:\Users\Artin\Documents\voice files project\se.mp3"},
        {'ج', @"C:\Users\Artin\Documents\voice files project\je.mp3"},
        {'چ', @"C:\Users\Artin\Documents\voice files project\che.mp3"},
        {'ح', @"C:\Users\Artin\Documents\voice files project\hhe.mp3"},
        {'خ', @"C:\Users\Artin\Documents\voice files project\khe.mp3"},
        {'د', @"C:\Users\Artin\Documents\voice files project\dal.mp3"},
        {'ذ', @"C:\Users\Artin\Documents\voice files project\dal_ze.mp3"},
        {'ر', @"C:\Users\Artin\Documents\voice files project\re.mp3"},
        {'ز', @"C:\Users\Artin\Documents\voice files project\ze.mp3"},
        {'ژ', @"C:\Users\Artin\Documents\voice files project\zhe.mp3"},
        {'س', @"C:\Users\Artin\Documents\voice files project\sse.mp3"},
        {'ش', @"C:\Users\Artin\Documents\voice files project\she.mp3"},
        {'ص', @"C:\Users\Artin\Documents\voice files project\sad.mp3"},
        {'ض', @"C:\Users\Artin\Documents\voice files project\zad.mp3"},
        {'ط', @"C:\Users\Artin\Documents\voice files project\ta.mp3"},
        {'ظ', @"C:\Users\Artin\Documents\voice files project\za.mp3"},
        {'ع', @"C:\Users\Artin\Documents\voice files project\ain.mp3"},
        {'غ', @"C:\Users\Artin\Documents\voice files project\ghain.mp3"},
        {'ف', @"C:\Users\Artin\Documents\voice files project\fe.mp3"},
        {'ق', @"C:\Users\Artin\Documents\voice files project\ghaf.mp3"},
        {'ک', @"C:\Users\Artin\Documents\voice files project\kaf.mp3"},
        {'گ', @"C:\Users\Artin\Documents\voice files project\gaf.mp3"},
        {'ل', @"C:\Users\Artin\Documents\voice files project\lam.mp3"},
        {'م', @"C:\Users\Artin\Documents\voice files project\mim.mp3"},
        {'ن', @"C:\Users\Artin\Documents\voice files project\non.mp3"},
        {'و', @"C:\Users\Artin\Documents\voice files project\ve.mp3"},
        {'ه', @"C:\Users\Artin\Documents\voice files project\he.mp3"},
        {'ی', @"C:\Users\Artin\Documents\voice files project\ye.mp3"},
        {'إ', @"C:\Users\Artin\Documents\voice files project\alf_hamze.mp3"},
        {'ؤ', @"C:\Users\Artin\Documents\voice files project\ve_hamze.mp3"},
        {'ئ', @"C:\Users\Artin\Documents\voice files project\ye_hamze.mp3"},


        {'a', @"C:\Users\Artin\Documents\voice files project\a.mp3"},
        {'b', @"C:\Users\Artin\Documents\voice files project\b.mp3"},
        {'c', @"C:\Users\Artin\Documents\voice files project\c.mp3"},
        {'d', @"C:\Users\Artin\Documents\voice files project\d.mp3"},
        {'e', @"C:\Users\Artin\Documents\voice files project\e.mp3"},
        {'f', @"C:\Users\Artin\Documents\voice files project\f.mp3"},
        {'g', @"C:\Users\Artin\Documents\voice files project\g.mp3"},
        {'h', @"C:\Users\Artin\Documents\voice files project\h.mp3"},
        {'i', @"C:\Users\Artin\Documents\voice files project\i.mp3"},
        {'j', @"C:\Users\Artin\Documents\voice files project\j.mp3"},
        {'k', @"C:\Users\Artin\Documents\voice files project\k.mp3"},
        {'l', @"C:\Users\Artin\Documents\voice files project\l.mp3"},
        {'m', @"C:\Users\Artin\Documents\voice files project\m.mp3"},
        {'n', @"C:\Users\Artin\Documents\voice files project\n.mp3"},
        {'o', @"C:\Users\Artin\Documents\voice files project\o.mp3"},
        {'p', @"C:\Users\Artin\Documents\voice files project\p.mp3"},
        {'q', @"C:\Users\Artin\Documents\voice files project\q.mp3"},
        {'r', @"C:\Users\Artin\Documents\voice files project\r.mp3"},
        {'s', @"C:\Users\Artin\Documents\voice files project\s.mp3"},
        {'t', @"C:\Users\Artin\Documents\voice files project\t.mp3"},
        {'u', @"C:\Users\Artin\Documents\voice files project\u.mp3"},
        {'v', @"C:\Users\Artin\Documents\voice files project\v.mp3"},
        {'w', @"C:\Users\Artin\Documents\voice files project\w.mp3"},
        {'x', @"C:\Users\Artin\Documents\voice files project\x.mp3"},
        {'y', @"C:\Users\Artin\Documents\voice files project\y.mp3"},
        {'z', @"C:\Users\Artin\Documents\voice files project\z.mp3"},

        {'1', @"C:\Users\Artin\Documents\voice files project\1.mp3"},
        {'2', @"C:\Users\Artin\Documents\voice files project\2.mp3"},
        {'3', @"C:\Users\Artin\Documents\voice files project\3.mp3"},
        {'4', @"C:\Users\Artin\Documents\voice files project\4.mp3"},
        {'5', @"C:\Users\Artin\Documents\voice files project\5.mp3"},
        {'6', @"C:\Users\Artin\Documents\voice files project\6.mp3"},
        {'7', @"C:\Users\Artin\Documents\voice files project\7.mp3"},
        {'8', @"C:\Users\Artin\Documents\voice files project\8.mp3"},
        {'9', @"C:\Users\Artin\Documents\voice files project\9.mp3"},
        // Add more mappings here
        };
        }







        private void Form1_Load(object sender, EventArgs e)
        {
            
            PositionGroupBoxes();

            DocumentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            mainFolderPath = Path.Combine(DocumentsPath, mainFolder);

            fontFilePath = Path.Combine(DocumentsPath, "subDatas", "dataFiles", "fontList.txt");
            filePathData = Path.Combine(DocumentsPath, "subDatas", "dataFiles", "data.txt");
            soundMappingFiles = Path.Combine(DocumentsPath, "subDatas", "voices");
            wordPath = Path.Combine(DocumentsPath, "wordFiles");
            // Create a new instance of the Word application

            try
            {

                ReadFileIntoList(filePathData, fontFilePath);
                projecktNummber = File.ReadAllLines(filePathData).Length;
                Console.WriteLine("****************" + projecktNummber);
                MessageBox.Show("data loaded seccessfully");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"an error occurred :{ex.Message}");
            }


            if (!Directory.Exists(mainFolder))
            {
                Directory.CreateDirectory(mainFolderPath);

                foreach (var folder in subfolder)
                {
                    Directory.CreateDirectory(Path.Combine(mainFolderPath, folder));
                }
                Console.WriteLine("the first run ! folders maided seccessfully");
            }
            else
            {
                Console.WriteLine("already exist!");
            }

            List<int> fontSize = new List<int> { 8, 9, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72 };
            //size of fonts 

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
                Console.WriteLine($"An error occurred: {ex.Message}");
            }



            

           

            if (File.Exists(filePath))
            {
                string[] lines = File.ReadAllLines(filePath);
                dataList.AddRange(lines);
            }
            else
            {
                MessageBox.Show("file does not exist");
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
                    char previousChar = MainrichTextBox.Text[currentPosition];
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
                    char nextChar = MainrichTextBox.Text[currentPosition - 1];
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
            // Get the current cursor position
            int cursorPosition = MainrichTextBox.SelectionStart;

            // Find the start of the current paragraph (the first newline before the cursor)
            int currentParagraphStart = MainrichTextBox.Text.LastIndexOf(Environment.NewLine, cursorPosition - 1);

            if (currentParagraphStart > 0)
            {
                // Find the start of the previous paragraph (the newline before the current paragraph start)
                int previousParagraphStart = MainrichTextBox.Text.LastIndexOf(Environment.NewLine, currentParagraphStart - 1);

                // If a previous paragraph was found, move the cursor to the character after the previous paragraph's newline
                // If no previous newline, move to the start of the text
                MainrichTextBox.SelectionStart = previousParagraphStart >= 0 ? previousParagraphStart + Environment.NewLine.Length : 0;
            }
            else
            {
                // If the cursor is in the first paragraph, move to the start of the text
                MainrichTextBox.SelectionStart = 0;
            }

            // Scroll to the cursor position and focus the textbox
            MainrichTextBox.ScrollToCaret();
            MainrichTextBox.Focus();
        }





        private void endparBtn_Click(object sender, EventArgs e)
        {

            // Get the current cursor position
            int cursorPosition = MainrichTextBox.SelectionStart;

            // Find the start of the current paragraph (the first newline before the cursor)
            int currentParagraphStart = MainrichTextBox.Text.LastIndexOf(Environment.NewLine, cursorPosition - 1);

            if (currentParagraphStart > 0)
            {
                // Find the start of the previous paragraph (another newline before the current paragraph)
                int previousParagraphStart = MainrichTextBox.Text.LastIndexOf(Environment.NewLine, currentParagraphStart - 1);

                // Move the cursor to the start of the previous paragraph
                MainrichTextBox.SelectionStart = previousParagraphStart >= 0 ? previousParagraphStart + Environment.NewLine.Length : 0;

                // Scroll to the cursor position
                MainrichTextBox.ScrollToCaret();
                MainrichTextBox.Focus();
            }
            else
            {
                // If no previous paragraph, move to the very start of the text
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



        public int[] savedCarentPosiotion = new int[10];
        public int countBmark = 1;





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

        private bool isRecording = false;
        private bool isWaitingForKey = false;

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
                    // Handle the "Yes" button click (e.g., save the file)
                    // Attach the event handler for the saveAsBtn.Click event here
                    this.saveAsBtn.Click += saveAsBtn_Click;
                }
                else
                {
                    MessageBox.Show("You didn't specify a file name.");
                }
            }

            string wordFileName = $"{textBox2.Text}.docx";
            string wordFilePath = Path.Combine(DocumentsPath, "wordFiles", wordFileName);
            
            string voiceFilePath = Path.Combine(DocumentsPath, "subDatas", "voices", wordFileName);

            bool wordFileExists = File.Exists(wordFilePath);
            bool voiceFileExists = File.Exists(voiceFilePath);

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

                    string voiceTag = Path.Combine(DocumentsPath, "subDatas", "voices", $"{fileNameWithoutExtension}.wav");

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

                        string voiceTag = Path.Combine(DocumentsPath, "subDatas", "voices", $"{fileNameWithoutExtension}.wav");
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
                                // Attempt to close the file if it is open elsewhere
                                

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
                        // Only add the event handler once, or ensure it's added only if it's not already attached
                        this.saveBtn.Click -= saveBtn_Click; // Remove if already attached
                        this.saveBtn.Click += saveBtn_Click;
                    }
                }


                if (index != ListBox.NoMatches)// Check if an item is clicked
                {
                    Console.WriteLine(index);
                    wordName = listBox1.SelectedItem.ToString();
                    // Check if the index is within the range of voiceTagList
                    if (index >= 0 && index < allDocxFiles.Count)
                    {
                        wordName.Replace(".docx", "");
                        string addressToOpen = Path.Combine(DocumentsPath, "wordFiles", wordName);
                        MainrichTextBox.Text = openWordDocument(addressToOpen);
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
                string oldVoiceTag = Path.Combine(mainFolderPath, "voices", $"{wordname}0.wav");
                string newVoiceTag = Path.Combine(mainFolderPath, "voices", $"{textBox2.Text}0.wav");

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
            string wordFolder = Path.Combine(DocumentsPath, "wordFiles");
            this.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.Form1_keyPress);
            this.KeyPress -= new System.Windows.Forms.KeyPressEventHandler(this.Form1_keyPress);
        }



        private void starttextBtn_Click(object sender, EventArgs e)
        {

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

/*
    private void savePdfBtn_Click(object sender, EventArgs e)
    {
        try
        {
            // Ensure the output directory exists
            string outputDir = Path.Combine(DocumentsPath, "subDatas", "pdf folder");
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // File path to save the PDF
            string outputFile = Path.Combine(outputDir, $"{textBox2.Text}.pdf");

            // Create a PDF document
            Document document = new Document();
            using (FileStream fs = new FileStream(outputFile, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                PdfWriter writer = PdfWriter.GetInstance(document, fs);
                document.Open();

                // Use a font that supports Arabic
                BaseFont baseFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, 12);

                // Add content to the document using the specified font
                Paragraph paragraph = new Paragraph(MainrichTextBox.Text, font)
                {
                    Alignment = Element.ALIGN_RIGHT // Align text to the right for Arabic
                };
                document.Add(paragraph);

                document.Close();
            }

            MessageBox.Show("PDF created successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"An error occurred while creating the PDF: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
*/

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
                wordFilePath = Path.Combine(DocumentsPath, "wordFiles", $"{wordName}.docx");

                CreateWordDocument(wordFilePath);

                MessageBox.Show("Press a key", "Press key", MessageBoxButtons.OK, MessageBoxIcon.Information);


                // Detach previous handlers to avoid multiple subscriptions


                this.Enabled = false;
                // Attach new handlers
                this.KeyDown += new KeyEventHandler(Form1_keyDown);
                textBox2.Enabled = false;

                outputFilePath = Path.Combine(mainFolderPath, "voices", $"{wordName}.wav");
                this.KeyUp += new KeyEventHandler(Form1_keyUp);

                this.Enabled = true;
                await WaitForKeyPressAsync();
                //this.Enabled = true;





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
                // Check if the selected item has changed


                removeItemsListBox();
                populateFontListBox();
                listBox1.Focus();
               

                // Update the last selected item
               
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

        private void populateFontListBox()
        {
            listBox1.Items.Clear();
            fontList.Clear();
            using (StreamReader reader = new StreamReader(fontFilePath, true))
            {
                while (!reader.EndOfStream)
                {
                    string item = reader.ReadLine();
                    fontList.Add(item);
                }
            }



            foreach (string font in fontList)
            {
                string[] fonts = font.Split('|');
                for (int i = 0; i < fonts.Length; i++)
                {
                    if (i % 3 == 1) // Check if the index is odd
                    {
                        listBox1.Items.Add(fonts[i]);
                        // Add the odd-numbered element to your ListBox (listBox1.Items.Add(...))
                    }
                }

            }

            this.listBox1.KeyDown += new KeyEventHandler(this.listBox_KeyDown_Font);

        }

        private void listBox_KeyDown_Font(object sender, KeyEventArgs e)
        {
            int index = listBox1.SelectedIndex;

            if (index != -1) // Check if an item is selected
            {
                if (e.KeyCode == Keys.Delete)
                
                    // Ensure index is within valid range for allDocxFiles
                    if (index >= 0 && index < allDocxFiles.Count)
                    {
                        string voiceName = listBox1.SelectedIndex.ToString();


                        string voiceTag = Path.Combine(DocumentsPath, "subDatas", "voices", $"{voiceName}.wav");
                        

                        // Try to delete the voice file
                        if (File.Exists(voiceTag))
                        {
                            try
                            {
                                // Attempt to close the file if it is open elsewhere


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


                        try
                        {
                            fontList.RemoveAt(index);

                            // Save updated dataList to file
                            using (StreamWriter writer = new StreamWriter(fontFilePath, false)) // false to overwrite the file
                            {
                                foreach (var item in fontList)
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
                        MessageBox.Show("No item selected. Please select an item to delete.");
                    }
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
            InstalledFontCollection installedFonts = new InstalledFontCollection();
            var fontFamilies = installedFonts.Families.OrderBy(f => f.Name);

            foreach (var font in fontFamilies)
            {
                fontComboBox.Items.Add(font.Name);
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

        private void FontComboBox_MeasureItem(object sender, MeasureItemEventArgs e)
        {
            if (e.Index >= 0)
            {
                string fontName = fontComboBox.Items[e.Index].ToString();
                using (System.Drawing.Font font = new System.Drawing.Font(fontName, 10, FontStyle.Regular, GraphicsUnit.Point))
                {
                    SizeF size = e.Graphics.MeasureString(fontName, font);
                    e.ItemHeight = (int)size.Height;
                    e.ItemWidth = (int)size.Width;
                }
            }
        }
        /*
        private void FontComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (fontList.Count >= 10)
            {
                MessageBox.Show("you cant add more than 10 fonts in your list");
                return;
            }

            if (fontComboBox.SelectedItem != null)
            {
                getComboItem(fontComboBox.SelectedItem.ToString());
            }
        }

        private async void getComboItem(string selectedFont)
        {
            // Add the selected font to the fontList
            if (!fontList.Contains(selectedFont))
            {
                fontList.Add(selectedFont);
            }

            // Set the output file path
            

            // Ensure KeyDown and KeyUp event handlers are added only once
            this.KeyDown -= Form1_keyDown;
            this.KeyDown += Form1_keyDown;
            outputFilePath = Path.Combine(mainFolderPath, "voices", $"{selectedFont}.wav");
            
            this.KeyUp -= Form1_keyUp;
            this.KeyUp += Form1_keyUp;


            await WaitForKeyPressAsync();
            this.KeyDown -= Form1_keyDown;
            this.KeyUp -= Form1_keyUp;

            saveFontFile(selectedFont);
           
        }
*/

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

   
        private void fontSizeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            updateFontSize();

            if (fontSizeComboBox.SelectedItem != null)
            {
                string selectedSize = fontSizeComboBox.SelectedItem.ToString();
                MessageBox.Show($"Selected font: {selectedSize}");
            }
        }

        private void updateFontSize()
        {
            if(int.TryParse(fontSizeComboBox.Text, out int newSize))
            {
                MainrichTextBox.Font = new System.Drawing.Font(MainrichTextBox.Font.FontFamily, newSize);
            }
            else
            {
                MessageBox.Show("enter valid number");
            }
        }

        private void fontSizeComboBox_keyPress(object sender , KeyPressEventArgs e)
        {
        
            if(!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }

            if(e.KeyChar == (char)Keys.Enter)
            {
                updateFontSize();
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
            Document doc = wordApp.Documents.Open(@"C:\Users\Artin\Desktop\sprechen außland.docx", ReadOnly: false, Visible: false);

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


    

    



