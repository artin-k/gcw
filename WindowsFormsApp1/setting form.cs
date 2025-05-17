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
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            string folderPath = Path.Combine(basePath, "files");
            try
            {
                filePath = Path.Combine(folderPath, "fonts.txt");
                Console.WriteLine(filePath);

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
            try
            {
                listBoxAllFonts.Items.Clear();

                // Get all installed fonts
                using (var fonts = new InstalledFontCollection())
                {
                    // Add font names sorted alphabetically
                    var fontNames = fonts.Families
                        .Select(f => f.Name)
                        .OrderBy(f => f)
                        .ToArray();

                    listBoxAllFonts.Items.AddRange(fontNames);
                }

                // Configure ListBox for multiple selection
                listBoxAllFonts.SelectionMode = SelectionMode.MultiExtended;
                listBoxAllFonts.Sorted = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading fonts: {ex.Message}");
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
            try
            {
                // Get all items from the ListBox
                var allFonts = listBoxSelectedFonts.Items
                    .OfType<string>()  // Cast to string
                    .ToList();

                // Check if list is empty
                if (allFonts.Count == 0)
                {
                    MessageBox.Show("The font list is empty!", "Warning",
                                  MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                File.WriteAllLines(filePath, allFonts);  // Overwrite file each time

                MessageBox.Show($"Successfully saved {allFonts.Count} fonts to:\n{filePath}",
                              "Success",
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Save Error: {ex.Message}");
            }
        }
    }
    // Add this in a new file (e.g., FontItem.cs) or at the top of your form class
    public class FontItem
    {
        public string Name { get; set; }
        public float Size { get; set; }

        // Optional: Override ToString() for better ListBox display
        public override string ToString() => $"{Name} ({Size}pt)";
    }
}


