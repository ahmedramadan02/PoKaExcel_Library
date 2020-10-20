using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Data;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Threading;

namespace pokaLibrary
{
    public class DataConverter
    {
        private Range range;
        private System.Data.DataTable table;
        object[,] arr;

        int Cols;
        int Rows;

        public DataConverter(Range r) {
            range = r;
            table = new System.Data.DataTable();
            arr = range.Value2;

            Cols = range.Columns.Count;
            Rows = range.Rows.Count;
        }

        public System.Data.DataTable GetDataTable(){
            // Constructing the data table
            for (int i = 0; i < Cols; i++)
                table.Columns.Add(arr[1, i+1].ToString());

           // You can't access filtered values Expecting access the areas
           foreach (Range r in range.Areas)
           {
               // Get the current range
               arr = r.Value2; 

               //Reget the row and columns
               Cols = r.Columns.Count;
               Rows = r.Rows.Count;

               // Assigning the values
               for (int i = 0; i < Rows; i++)
               {

                   DataRow Row = table.NewRow();
                   for (int j = 0; j < Cols; j++)
                   {
                       if (arr[i+1, j+1] == null)
                       {
                           //table.Rows[j][i] = "null"; 
                           Row[j] = "--";
                           continue;
                       }
                       Row[j] = arr[i +1 , j+1 ].ToString(); //Tables.Rows can't be directly filled !
                   }
                   table.Rows.Add(Row);
               }
           }
           
           table.Rows[0].Delete();
            return table;
        }


        // 
        public string[] GetArray() {
            int arrRows = arr.GetLowerBound(0);
            int arrCols = arr.GetUpperBound(1);
            
            int myarrIndex = arrCols * arrRows;
            string[] myarr = new string[myarrIndex];
            int counter = 0;

                // Rearrenging the values
                    for (int i = 0; i < arrRows; i++)
                    {
                        for (int j = 0; j < arrCols; j++)
                        {
                            if (arr[i + 1, j + 1] == null)
                            {
                                //table.Rows[j][i] = "null"; 
                                myarr[counter] = "--";
                                counter++;
                                continue;
                            }
                            myarr[counter] = arr[i + 1, j + 1].ToString();
                            counter++;
                        }
                    }

            return myarr;
        }

        public string[] GetArray(Special specialArray) { 
            string[] myarr;

            switch (specialArray)
            {
                case Special.Columns:
                    myarr = new string[range.Columns.Count];
                    for (int i = 0; i < myarr.Length - 1; i++)
                    {
                        if (arr[1, i + 1] == null) continue; //if null break;
                        myarr[i] = arr[1, i + 1].ToString();
                    }
                    return myarr;
                    break;

                case Special.Rows:
                    myarr = new string[range.Rows.Count];
                    for (int i = 0; i < myarr.Length - 1; i++)
                    {
                        if (arr[i + 1,1] == null) continue;
                        myarr[i] = arr[i + 1, 1].ToString();
                    }
                    return myarr;
                    break;

                default:
                    return null;
                    break;
            } 
        }

        public enum Special { 
            Columns,
            Rows,
        }

    }
}
