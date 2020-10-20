using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;

namespace pokaLibrary
{
    public abstract class MyExlAssistant
    {
        // Varaibles
        public int sheetsCount;
        public int currentSheetIndex; //Index of current sheet
        string filePath; //FilePath varaible
        protected  string FilePath
        {
            set { filePath = value; }
            get { return filePath; }
        }


        //List of sheets
        protected Dictionary<int, Worksheet> sheetsList = new Dictionary<int, Worksheet>(); //has a list of sheets
        protected Worksheet SheetsList
        {
            //set { imgList  = value; }
            get { return sheetsList[currentSheetIndex]; }

        }

        StringBuilder dir = new StringBuilder(); //directory of the sheets

        //Excel Objects
        private Application XlApp = null;
        public Workbook xlWorkBook = null;
        public Worksheet xlCurrentSheet = null;
        object misValue = System.Reflection.Missing.Value;

        // Functions
        //Constructors
        //default constructor
        public MyExlAssistant()
        { /*NOP*/ }

        public MyExlAssistant(string _FilePath)
        {
            InitializeVars(_FilePath);
            PathPrepare(filePath);
            CollectSheets();
            //Open();
        }
        // End Constructors

        //Functions
        // 1. Setting all varaibles
        public void InitializeVars(string Path)
        {
            FilePath = Path;
        }

        // 2. Streaming with memory
        public void Open()
        {
            XlApp = new Application();
            xlWorkBook = XlApp.Workbooks.Open(filePath, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue
                , misValue, misValue, misValue, misValue, misValue, misValue);
            sheetsCount = xlWorkBook.Sheets.Count;

            xlWorkBook.Activate();
            XlApp.Visible = true;
            //xlCurrentSheet = sheetsList[currentSheetIndex];
        }

        // 3. Prepare the path
        private void PathPrepare(string Path)
        {
            //1. splitting the path of the image
            string[] splitter = filePath.Split('\\');

            //2. rebuilding the path
            for (int i = 0; i < splitter.Count() - 1; i++)
            {
                dir.Append(splitter[i] + '\\');
            }
        }

        // 3. Collect Sheets
        private void CollectSheets()
        {
            try
            {
                //1. Check if the dir
                if (!Directory.Exists(dir.ToString()))
                    throw new DirectoryNotFoundException();

                // check if the file not found
                if (!File.Exists(filePath))
                    throw new FileNotFoundException();
                //2. inilaize the xlcomponents
                Open();

                //3. Clearing the sheetsList
                sheetsList.Clear();

                //4. list the dictionary
                for (int i = 1; i < sheetsCount; i++)
                    //a. collect all sheets
                    sheetsList.Add(i, xlWorkBook.Sheets[i]);

                // 5. Current sheet
                currentSheetIndex = 1;
                xlCurrentSheet = sheetsList[currentSheetIndex];
            }
            catch (Exception ex)
            {
                //System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        // Destructing
        //Destruct "Assiatant" object from memory
        public void Dispose()
        {
            GC.Collect(GC.GetGeneration(this), GCCollectionMode.Forced);
            GC.Collect();
        }

        //Destructor
        ~MyExlAssistant()
        {
            //xlCurrentSheet = null;
            //XlApp.Quit();
            //xlWorkBook.Close(misValue, misValue, misValue);
            if(xlCurrentSheet != null)
                Marshal.ReleaseComObject(xlCurrentSheet);
            if(xlWorkBook != null)
                Marshal.ReleaseComObject(xlWorkBook);
            if(XlApp != null)
                Marshal.ReleaseComObject(XlApp);
            XlApp= null;
            
            this.Dispose();
        }

    }
}
