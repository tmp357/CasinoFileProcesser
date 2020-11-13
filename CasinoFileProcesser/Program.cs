using CasinoFileProcesser.Modal;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CasinoFileProcesser
{
    class Program
    {
        static void Main(string[] args)
        {
            Processer();
        }

        private static void Processer()
        {
            System.Data.DataTable addressList = new System.Data.DataTable();
            string pathName = @"C:\Projects\Casino\Data.xlsx";
            string sheetName = "Sheet1";

            using (OleDbConnection connection =
                    new OleDbConnection((pathName.TrimEnd().ToLower().EndsWith("x"))
                    ? "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + pathName + "';" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'"
                    : "provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + pathName + "';Extended Properties=Excel 8.0;"))
            {
                OleDbDataAdapter data = new OleDbDataAdapter(string.Format("select * from [{0}$]", sheetName), connection);

                data.Fill(addressList);

                List<MatrixList> matrix = GetMatrixTableData();

                if (addressList.Rows.Count > 0)
                {
                    foreach (DataRow row in addressList.Rows)
                    {
                        
                    }
                }
            }
        }

        private static List<MatrixList> GetMatrixTableData()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            string pathName = @"C:\Projects\Casino\Matrix.xlsx";
            string sheetName = "Sheet2";

            using (OleDbConnection connection =
                    new OleDbConnection((pathName.TrimEnd().ToLower().EndsWith("x"))
                    ? "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + pathName + "';" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'"
                    : "provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + pathName + "';Extended Properties=Excel 8.0;"))
            {
                OleDbDataAdapter data = new OleDbDataAdapter(string.Format("select * from [{0}$]", sheetName), connection);

                data.Fill(dt);                
            }

            List<MatrixList> matrixList = new List<MatrixList>();

            if (dt.Rows.Count > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    matrixList.Add(new MatrixList()
                    {
                        Segment = row[0].ToString(),
                        Monthly = int.Parse(row[1].ToString()),
                        MediaCoupon = int.Parse(row[2].ToString()),
                        MediaCategory = row[3].ToString(),
                        FP = int.Parse(row[4].ToString()),
                        Valid = row[5].ToString()
                    });
                }
            }

            return matrixList;

        }
    }
}
