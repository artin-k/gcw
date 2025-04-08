using iText.StyledXmlParser.Jsoup.Safety;
using Microsoft.Office.Interop.Word;
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
        string[] selectedFonts;
        string filePath;

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
                filePath = Path.Combine(folderPath, "selected_fonts.txt");

                if (File.Exists(filePath))
                {                    
                    selectedFonts = File.ReadAllLines(filePath);
                    listBoxSelectedFonts.Items.AddRange(selectedFonts);
                }
                else
                {
                    MessageBox.Show("File not found: " + filePath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"an error occurred :{ex.Message}");
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
            int countItem = listBoxSelectedFonts.Items.Count;
            // Prevent adding more than 9 fonts
            if (countItem > 9)
            {
                MessageBox.Show("you can not select more then 9 fonts ");
                return;
            }

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
