using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Data.Analysis;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace Esclavitud
{
    internal class analisisDatos
    {
        public analisisDatos() { }

        public static void REPs(Excel.ListObject table)
        {
            Excel.Range rango = table.DataBodyRange;
            int filas = rango.Rows.Count;
            int columnas = rango.Columns.Count;

            string[] UUID = new string[filas-1];
            string[] FP = new string[filas-1];
            string[] TOTAL = new string[filas-1];

            for (int fila = 1; fila <= filas-1; fila++)
            {
                UUID[fila - 1]      = (rango.Cells[fila, 11] as Excel.Range).Value;
                FP[fila - 1]        = (rango.Cells[fila, 15] as Excel.Range).Value;
                TOTAL[fila - 1]     = (rango.Cells[fila, 35] as Excel.Range).Value;
            }
            var a = new StringDataFrameColumn("UUID", UUID);
            var b = new StringDataFrameColumn("FormaPago", FP);
            var c = new StringDataFrameColumn("Total", TOTAL);

            DataFrame df = new DataFrame(a, b, c);

            MessageBox.Show(df.ToString());
        }
    }
}
