using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

//using
//must use these three libraries
using pokaLibrary;
using Microsoft.Office.Interop.Excel;
using System.Data;

namespace PokaTest
{
    public partial class Form1 : Form
    {
        string fileName;
        PokaClass NewXl;
        Range r1;

        public Form1()
        {
            InitializeComponent();

            //Load the default
            comboFontSize.Text = "11";
            comboFontName.Text = "Arial";
            comboFontStyle.SelectedIndex = 0;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Wolud you like to cancel operations", "cancel auto fillter", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
            {
                NewXl.xlCurrentSheet.AutoFilterMode = false;
                r1 = NewXl.xlCurrentSheet.UsedRange;
                DataConverter c1 = new DataConverter(r1);
                System.Data.DataTable T1 = c1.GetDataTable();
                dataGridView1.DataSource = T1;
            }

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Select the current Ranges
            r1 = NewXl.xlCurrentSheet.UsedRange;

            //Connvert it to data table
            string chkTxt = textBox2.Text.Trim();
            System.Data.DataTable T1;
            if (chkTxt.Count() > 0) 
                // Search by two criterias
                T1 = NewXl.SearchMe(r1, (PokaClass.fields)comboBox1.SelectedIndex , textBox1.Text,textBox2.Text);
            else 
                //search by one criteria
                T1 = NewXl.SearchMe(r1, (PokaClass.fields)comboBox1.SelectedIndex , textBox1.Text);
            
            //Show it in the gridView
            dataGridView1.Visible = true;
            dataGridView1.DataSource = T1;
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Open the file dialog 
            openFileDialog1.ShowDialog();
            openFileDialog1.Filter = "(*.txt)|*.txt";

            if (openFileDialog1.FileName != null)
            {
                fileName = openFileDialog1.FileName;
                NewXl = new PokaClass(fileName);
                textBox3.Text = NewXl.currentSheetIndex.ToString();
                groupBox1.Enabled = true;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            openFileDialog1.Filter = "(*.txt)|*.txt";

            if (openFileDialog1.FileName != null)
            {
                fileName = openFileDialog1.FileName;
                NewXl = new PokaClass(fileName);
                groupBox1.Enabled = true;
                textBox3.Text = NewXl.currentSheetIndex.ToString();

                InitialControls();
                this.Show();
            }
        }

        private void InitialControls() {
            // Form items
            //1. Load fonts
            foreach (FontFamily font in System.Drawing.FontFamily.Families)
                comboFontName.Items.Add(font.Name.ToString());

            //Load comboSearchBy
            DataConverter convert = new DataConverter(NewXl.xlCurrentSheet.UsedRange);
            comboBox1.Items.Clear();
            comboBox1.Items.AddRange(convert.GetArray(DataConverter.Special.Columns));
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            try
            {
                NewXl.ChangeSheet(Convert.ToInt16(textBox3.Text));
                MessageBox.Show("Changed to Sheet# " + NewXl.currentSheetIndex);
                
            }
            catch(PokaException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void textBox3_KeyDown(object sender, KeyEventArgs e)
        {
            //
            short sh;
            if (Int16.TryParse(textBox3.Text, out sh))
            {

            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("Wolud you like to exit and save", "Exit and save", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
                NewXl.xlWorkBook.Close(XlSaveAction.xlSaveChanges); //save changes
            else //else close without changing
                NewXl.xlWorkBook.Close(XlSaveAction.xlDoNotSaveChanges);
        }

        private void btnColor_Click(object sender, EventArgs e)
        {
            if (r1 == null) {
                MessageBox.Show("Specify elements first! >> change the index of the search combo");
                return;
            }

            colorDialog1.ShowDialog();
            NewXl.ChangeColor(colorDialog1.Color);
        }

        private void btnFont_Click(object sender, EventArgs e)
        {
            try
            {
                FontStyling newFont = new FontStyling();
                newFont.Size = Convert.ToInt32(comboFontSize.Text);
                newFont.fStyle = (FontStyling.fontStyle)comboFontStyle.SelectedIndex;
                newFont.FontName = comboFontName.Text;

                //Set All
                // Call the function
                NewXl.ChangeFont(newFont);
            }

            catch {
                MessageBox.Show("Wrong values !");
            }
                       
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FontStyling s = new FontStyling(NewXl.xlCurrentSheet.UsedRange);
        }
    }
}
