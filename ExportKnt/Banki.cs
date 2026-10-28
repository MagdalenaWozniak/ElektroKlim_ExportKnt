using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportKnt
{
    class Banki
    {
        public string nazwa { get; set; } 
        public string nrRachunku { get; set; }
        public string waluta { get; set; } 
        public int domyslny { get; set; }

        Banki(string Bnk_Nazwa, string RkB_NrRachunku, string Knt_Symbol, string RkB_Domyslny)
        {
            this.nazwa = Bnk_Nazwa;
            this.nrRachunku = RkB_NrRachunku;
            this.waluta = Knt_Symbol;
            this.domyslny = Convert.ToInt32(RkB_Domyslny);
        }


        public static List<Banki> SelectListRowColumn(string query, string srcTableName, string connection_string)
        {
            List<string> str = new List<string>();
            List<Banki> myList = new List<Banki>();
            string sttr = "";
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    adapter.SelectCommand = new SqlCommand(query, connection);
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet, srcTableName);
                    foreach (DataRow row in (InternalDataCollectionBase)dataSet.Tables[0].Rows)
                    {
                        foreach (DataColumn column in (InternalDataCollectionBase)dataSet.Tables[0].Columns)
                        {
                            str.Add(row[column].ToString());

                            //MessageBox.Show(column + ": " + row[column].ToString());
                        }
                        myList.Add(new Banki(str[0], str[1], str[2], str[3]));
                        str.Clear();

                    }
                }

            }
            catch (Exception ex)
            {
                //MessageBox.Show(sttr = sttr + "Bład SQL:" + ex.ToString() + "\n");
            }
            return myList;
        }
    }
}
