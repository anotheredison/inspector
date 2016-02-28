using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace Inspector
{
    class ExcelWriter
    {
        private string file;
        private Excel.Application app;
        private Excel.Workbook wBook;
        private Excel.Worksheet mysheet;
        object misValue = System.Reflection.Missing.Value;
        private int curRow;

        public ExcelWriter(string file)
        {
            this.file = file;
            this.curRow = 1;
        }

        public bool init()
        {
            try
            {
                app = new Excel.Application();
                if (app == null)
                {
                    return false;
                }

                app.Visible = false;
                wBook = app.Workbooks.Add(misValue);
                mysheet = (Excel.Worksheet)wBook.Worksheets.get_Item(1);

                return true;
            }
            catch (Exception ex)
            {
            }

            return false;
        }

        public void WriteLine(string[] cols)
        {
            for (int i = 0; i < cols.Length; i++)
            {
                mysheet.Cells[curRow, i + 1] = cols[i];
            }

            curRow++;
        }

        public void SaveAndClose()
        {
            wBook.SaveAs(this.file);
            wBook.Close(true, misValue, misValue);
            app.Quit();

            releaseObject(mysheet);
            releaseObject(wBook);
            releaseObject(app);
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
