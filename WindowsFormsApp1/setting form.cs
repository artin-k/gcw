using iText.StyledXmlParser.Jsoup.Safety;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class setting_form : Form
    {

        public setting_form()
        {
            InitializeComponent();
            LoadFonts();
        }

        private void setting_form_Load(object sender, EventArgs e)
        {          
              
            
            string folderPath = AppDomain.CurrentDomain.BaseDirectory;

            try
            {
                string filePath = Path.Combine(folderPath, fgdfgdf );

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


            private void ReadFileIntoList(string filePath, string fontFilePath)
            {
                allDocxFiles.Clear();


                char delimiter = '-';

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


        }

        private void LoadFonts()
        {
            using (InstalledFontCollection fontsCollection = new InstalledFontCollection())
            {
                foreach (FontFamily font in fontsCollection.Families)
                {
                    listBoxAllFonts.Items.Add(font.Name);
                }
            }
        }

        private void ButtonAdd_Click(object sender, EventArgs e)
        {

            // Retrieve the selected items from both ListBoxes
            var selectedFontName = listBoxAllFonts.SelectedItem?.ToString();
            var selectedFontSize = listBoxSize.SelectedItem != null ? (int)listBoxSize.SelectedItem : 0;

            if (!string.IsNullOrEmpty(selectedFontName) && selectedFontSize > 0)
            {
                listBoxSelectedFonts.Items.Add(new FontItem { Name = selectedFontName, Size = selectedFontSize });
                listBoxAllFonts.Items.Remove(listBoxAllFonts.SelectedItem);
            }
            else
            {
                MessageBox.Show("Please select both a font and a size.");

            }


          
        }

        private void ButtonRemove_Click(object sender, EventArgs e)
        {
            if (listBoxSelectedFonts.SelectedItem != null)
            {
                listBoxAllFonts.Items.Add(listBoxSelectedFonts.SelectedItem);
                listBoxSelectedFonts.Items.Remove(listBoxSelectedFonts.SelectedItem);
            }
        }

        private void ButtonSave_Click(object sender, EventArgs e)
        {
            SaveSelectedFontsToFile();
        }

        private void SaveSelectedFontsToFile()
        {
            // Get the path to the directory where the program is installed
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;

            // Combine the base directory with the file name
            string filePath = Path.Combine(baseDirectory, "selected_fonts.txt");

            try
            {
                // Use StreamWriter with the append parameter set to true
                using (StreamWriter writer = new StreamWriter(filePath, true))
                {
                    if (listBoxSelectedFonts.Items.Count == 0)
                    {
                        MessageBox.Show("No fonts selected!");
                        return;
                    }

                    foreach (FontItem font in listBoxSelectedFonts.Items)
                    {
                        writer.WriteLine(font.Name + " - " + font.Size);
                    }
                }

                MessageBox.Show("Selected fonts appended to " + filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred while saving the file: " + ex.Message);
            }


        }
        private class FontItem
        {
            public string Name { get; set; }
            public int Size { get; set; }

            public override string ToString()
            {
                return Name + " - " + Size;
            }
        }

    }
}
