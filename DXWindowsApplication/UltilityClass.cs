using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXWindowsApplication
{
    class UltilityClass
    {
        //read excel 
        public static DataSet loadDataFromExcelToDataSet(String path, String extension)
        {
            DataSet ds = new DataSet();
            String connString = "";

            switch (extension)
            {
                case ".xls"://excel 97-03
                    connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1};'";
                    break;
                case ".xlsx": // excel 2010 and later
                    connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
                    break;

            }
            //create connect string
            connString = String.Format(connString, path, "Yes");

            OleDbConnection connExcel = new OleDbConnection(connString);
            OleDbCommand command = new OleDbCommand();
            //OleDbDataAdapter adapter = new OleDbDataAdapter();

            //command.Connection = connExcel;

            connExcel.Open();

            DataTable sheetNames = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            DataTable table = new DataTable();

            for (int i = 0; i < sheetNames.Rows.Count; i++)
            {
                String sheetName = sheetNames.Rows[i]["TABLE_NAME"].ToString();
                table = makeDataTableFromSheetName(connExcel, sheetName);
                ds.Tables.Add(table);
            }

            connExcel.Close();
            return ds;
        }

        public static DataTable makeDataTableFromSheetName(OleDbConnection connExcel, String sheetName)
        {
            DataTable table = new DataTable();
            String command = "SELECT * FROM [" + sheetName + "]";
            OleDbDataAdapter adapter = new OleDbDataAdapter(command, connExcel);
            adapter.Fill(table);

            return table;
        }

        // phuong thuc xu ly diem danh
        public static void absence(ref DataSet list, DataSet checkList)
        {
            int n = checkList.Tables[0].Rows.Count;
            for (int i = 0; i < n; i++)
            {
                //get student id from check list
                String idCheck = checkList.Tables[0].Rows[i][0].ToString();
                //get student id from list
                int index = findStudentById(list, idCheck);
                //check
                if (index != -1)
                    list.Tables[0].Rows[index][2] = "Di Hoc";
                //
            }

        }

        //find id
        public static int findStudentById(DataSet list, String studentID)
        {
            int n = list.Tables[0].Rows.Count;
            for (int i = 0; i < n; i++)
            {
                //get student id from list
                String id = list.Tables[0].Rows[i][0].ToString();
                if (id.Equals(studentID))
                    return i;
                //
            }
            return -1;
        }

    }
}
