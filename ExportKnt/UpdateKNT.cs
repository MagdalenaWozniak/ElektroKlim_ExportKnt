using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportKnt
{
    class UpdateKNT
    {
        public int GIDTyp { get; set; } 
        public int GIDNumer { get; set; } 
        public string akronim { get; set; } 
        public int WF_ID { get; set; } 
        public string firma { get; set; }

        UpdateKNT(string GIDTyp, string GIDNumer, string akronim, string WF_id, string firma)
        {
            this.GIDTyp = Convert.ToInt32(GIDTyp);
            this.GIDNumer = Convert.ToInt32(GIDNumer);
            this.akronim = akronim;
            this.WF_ID = Convert.ToInt32(WF_id);
            this.firma = firma;

        }

        public static List<UpdateKNT> SelectListRowColumn(string query, string srcTableName, string connection_string)
        {
            List<string> str = new List<string>();
            List<UpdateKNT> myList = new List<UpdateKNT>();
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
                        myList.Add(new UpdateKNT(str[0], str[1], str[2], str[3], str[4]));
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
