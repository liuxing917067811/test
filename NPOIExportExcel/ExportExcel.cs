using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel ;
using System.Data;
using System.Windows.Forms;

namespace NPOIExportExcel
{
    public static class ExportExcel
    {
        public static  DataSet ListViewToDataSet(ListView lst)
        {
            int i, j;
            DataSet ds = new DataSet();
            ds.Tables.Add(new DataTable("sheet1"));
            DataTable dt = ds.Tables[0];
            DataRow datarow;

            for(i = 0; i < lst.Columns.Count; i++)
            {
                dt.Columns.Add(new DataColumn(lst.Columns[i].Text.Trim()));
            }

            for (i = 0; i < lst.Items.Count; i++) 
            {
                datarow = dt.NewRow();
                for (j = 0; j < lst.Columns.Count; j++)
                {
                    datarow[j] = lst.Items[i].SubItems[j].Text.Trim();
                }
                dt.Rows.Add(datarow);
            }


            return ds;
        }

        
        public static HSSFWorkbook BuildWorkbook(DataTable dt)
        {
            int i,j,k;
            HSSFWorkbook book = new HSSFWorkbook();
            var sheet = book.CreateSheet("Sheet1");
            var drow = sheet.CreateRow(0);

            for (k = 0; k < dt.Columns.Count; k++)
            {
                var cell = drow.CreateCell(k);
                cell.SetCellValue(dt.Columns[k].ToString());

            }

            for(i=0; i < dt.Rows.Count; i++)
            {
                  drow = sheet.CreateRow(i+1);
                for (j = 0; j < dt.Columns.Count; j++)
                {
                    var cell = drow.CreateCell(j);
                    cell.SetCellValue(dt.Rows[i][j].ToString());
                }

            }



            return book;
        }

        public static void ExportToExcel(string filepath, DataTable dt) 
        {
            HSSFWorkbook newBook = BuildWorkbook(dt);
            using (var fs = File.OpenWrite(filepath))
            {
                newBook.Write(fs);
            }
            
        }


    }
}
