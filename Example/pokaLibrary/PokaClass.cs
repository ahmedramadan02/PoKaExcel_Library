using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

using System.Xml;
using System.Data;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace pokaLibrary
{
    public class PokaClass : MyExlAssistant
    {
        public PokaClass(string _filePath)
            : base(_filePath)
        {

        }

        //Functions
        //1. Searching
        public enum fields
        {
            Profile,
            Desc,
            Px,
            Py,
            Pz,
            Code,
            e1,
            e2,
            e3,
            e4,
            e5,
        }

        //a. Searching by one criteria
        public System.Data.DataTable SearchMe(Range currentRanges, fields field, string criteria1)
        {
            // get the index of the column to be filtered
            int fieldIndex = Convert.ToInt32(field) + 1;

            // Auto-filter with a certain criteria
            bool stat = currentRanges.AutoFilter(fieldIndex.ToString(), "=" + criteria1,
                                                 XlAutoFilterOperator.xlAnd, Type.Missing, Type.Missing);

            //Convert the resulting range to data table to be viewed in the gridView 
            Range filteredRanges = currentRanges.SpecialCells(XlCellType.xlCellTypeVisible, true);
            DataConverter convert = new DataConverter(filteredRanges);

            //Return the data
            return convert.GetDataTable();
        }

        //Searching by two criteria
        public System.Data.DataTable SearchMe(Range currentRanges, fields field, string criteria1, string criteria2)
        {
            // get the index of the column to be filtered
            int fieldIndex = Convert.ToInt32(field) + 1;

            // Auto-filter by two criterias
            bool result = currentRanges.AutoFilter(fieldIndex.ToString(), "=" + criteria1,
                                                    XlAutoFilterOperator.xlAnd, "=" + criteria2, Type.Missing);

            //Convert the resulting range to data table to be viewed in the gridView 
            Range filteredRanges = currentRanges.SpecialCells(XlCellType.xlCellTypeVisible, true);
            DataConverter convert = new DataConverter(currentRanges);

            //return data
            return convert.GetDataTable();
        }

        //2. Change Sheet
        public Worksheet ChangeSheet(int current)
        {
            if (current > sheetsCount || current < 1)
                throw new PokaException("Sheet No. ERROR");

            //Change the Index 
            currentSheetIndex = current;

            // load the current sheet
            xlCurrentSheet = null;
            xlCurrentSheet = sheetsList[currentSheetIndex];
            return xlCurrentSheet; //return must be out of try ==> varaibles life cycle

        }

        //3. Formating operations
        // ChangeColor
        public void ChangeColor(Color color)
        {
            //Change Color
            xlCurrentSheet.UsedRange.Cells.Font.Color = color;
        }

        public void ChangeFont(FontStyling myFont)
        {
            // Font name
            xlCurrentSheet.UsedRange.Cells.Font.Name = myFont.FontName ;

            //Change size
            xlCurrentSheet.UsedRange.Cells.Style.Font.Size = myFont.Size;

            //Font Styles
            switch (myFont.fStyle)
            {
                case FontStyling.fontStyle.Bold:
                    xlCurrentSheet.UsedRange.Cells.Font.Bold = true;
                    break;
                case FontStyling.fontStyle.Italic:
                    xlCurrentSheet.UsedRange.Cells.Font.Italic = true;
                    break;
                case FontStyling.fontStyle.Underline:
                    xlCurrentSheet.UsedRange.Cells.Font.Underline = true;
                    break;
                default:
                    xlCurrentSheet.UsedRange.Cells.Font.Underline = false;
                    xlCurrentSheet.UsedRange.Cells.Font.Italic = false;
                    xlCurrentSheet.UsedRange.Cells.Font.Bold = false;
                    break;
            }
        }
    }

    //Font styling
    //
    public class FontStyling
    {
        //Vars with default values !
        private double size = 11;
        private string fontName = "Arial";
        private fontStyle fstyle = fontStyle.None;

        public FontStyling(){
        }

        //(Optional !)
        public FontStyling(Range _usedRange) {
            SetDefaultValues(_usedRange);
        }

        public double Size
        {
            //Get accessor
            get { return size; }

            //Set accessor
            set
            {
                int result;
                if (int.TryParse(Convert.ToString(value), out result))
                    size = value; // Default Value
                else
                    size = 11;
            }
        }

        public string FontName
        {
            //get accessor
            get { return fontName; }

            //Set accessor
            set
            {
                if (value == null || value == "")
                    fontName = "Arial"; // Default Value
                else
                    fontName = value;
            }
        }

        public fontStyle fStyle
        {
            //Get accessor
            get { return fstyle; }

            //Set accessor
            set
            {
                fontStyle result;
                if (Enum.TryParse<fontStyle>(Convert.ToString(value) ,out result))
                    fstyle = value; // Default Value
                else
                    fstyle = fontStyle.None;
            }
        }

        //Enum
        public enum fontStyle
        {
            None,
            Bold,
            Italic,
            Underline,
        }

        private void SetDefaultValues(Range r){
        //1. Get the Default Values
            fontName = r.Cells.Font.Name;
            size = r.Cells.Font.Size;
            #region Check style
		        fontStyle f;
                if (!Enum.TryParse<fontStyle>(Convert.ToString(r.Cells.Font.FontStyle),out f))
                    fstyle = fontStyle.None; 
	        #endregion
            
        //2.Save to the Xml
            XmlWrite();
        }

        XmlDocument xDoc;
        private void XmlRead() {
            xDoc = new XmlDocument();
            xDoc.Load("../../DefaultValues.xml");
            XmlNodeList xFont = xDoc.GetElementsByTagName("FontStyling");
            fontName = xFont[0].FirstChild.InnerText;
            size = Convert.ToInt16(xFont[0].FirstChild.NextSibling.InnerText);
            fstyle = (fontStyle) Enum.Parse(typeof(fontStyle),xFont[0].FirstChild.NextSibling.NextSibling.InnerText);
        }

        private void XmlWrite() {
            xDoc = new XmlDocument();
            xDoc.Load("../../DefaultValues.xml");
            XmlNodeList xFont = xDoc.GetElementsByTagName("FontStyling");
            xFont[0].FirstChild.InnerText = fontName;
            xFont[0].FirstChild.NextSibling.InnerText = size.ToString();
            xFont[0].FirstChild.NextSibling.NextSibling.InnerText = fstyle.ToString();

            //Saving 
            xDoc.Save("../../DefaultValues.xml");
        }
    }

}

