using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportKnt
{
    class KNT
    {
        

        public int gidTyp { get; set; }
        public int gidFirma { get; set; }
        public int gidNumer { get; set; }
        public int gidLp { get; set; }
        public string akronim { get; set; }
        public string nazwa1 { get; set; }
        public string nazwa2 { get; set; }
        public string nazwa3 { get; set; }
        public string miasto { get; set; }
        public string ulica { get; set; }
        public string kodP { get; set;}
        public string nrRachunku { get; set;}
        public string kraj { get; set; }
        public string telefon1 { get; set; }
        public string telefon2 { get; set; }
        public string eMail { get; set; }
        public string nip { get; set;}
        public string nipPrefiks { get; set; }
        public string pesel { get; set;}
        public string regon { get; set; }
        public int termin { get; set; }
        public string waluta { get; set; }
        public int podzialPlat { get; set; }
        public string url { get; set; }
        public int typ { get; set; }

        KNT(string Knt_GIDTYP, string Knt_GIDFirma, string Knt_GIDNumer, string Knt_GIDLp, string Knt_Akronim, string Knt_Nazwa1, string Knt_Nazwa2, string Knt_Nazwa3, string Knt_Miasto, string Knt_Ulica, 
             string Knt_KodP, string RkB_NrRachunku, string KPC_NazwaPL, string Knt_Telefon1, string Knt_Telefon2, string Knt_EMail, string Knt_Nip, string Knt_NipPrefiks, string Knt_Pesel, string Knt_Regon, 
             string Knt_TerminPlZak, string Knt_Symbol, string Knt_PodzialPlat, string Knt_URL, string Knt_Typ)
        {
            this.gidTyp = Convert.ToInt32(Knt_GIDTYP);
            this.gidFirma = Convert.ToInt32(Knt_GIDFirma);
            this.gidNumer = Convert.ToInt32(Knt_GIDNumer);
            this.gidLp = Convert.ToInt32(Knt_GIDLp);
            this.akronim = Knt_Akronim;
            this.nazwa1 = Knt_Nazwa1;
            this.nazwa2 = Knt_Nazwa2;
            this.nazwa3 = Knt_Nazwa3;
            this.miasto = Knt_Miasto;
            this.ulica = Knt_Ulica;
            this.kodP = Knt_KodP;
            this.nrRachunku = RkB_NrRachunku;
            this.kraj = KPC_NazwaPL;
            this.telefon1 = Knt_Telefon1;
            this.telefon2 = Knt_Telefon2;
            this.eMail = Knt_EMail;
            this.nip = Knt_Nip;
            this.nipPrefiks = Knt_NipPrefiks;
            this.pesel = Knt_Pesel;
            this.regon = Knt_Regon;
            this.termin = Convert.ToInt32(Knt_TerminPlZak);
            this.waluta = Knt_Symbol;
            this.podzialPlat = Convert.ToInt32(Knt_PodzialPlat);
            this.url = Knt_URL;
            this.typ = Convert.ToInt32(Knt_Typ);

        }





        public static List<KNT> SelectListRowColumn(string query, string srcTableName, string connection_string)
        {
            List<string> str = new List<string>();
            List<KNT> myList = new List<KNT>();
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
                        myList.Add(new KNT(str[0], str[1], str[2], str[3], str[4], str[5], str[6], str[7], str[8], str[9], str[10], str[11], str[12], str[13], str[14], str[15], str[16], str[17], str[18], str[19],
                                           str[20], str[21], str[22], str[23], str[24]));
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


        public static string Update(string sql, string connection_string)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    connection.Open();
                    adapter.UpdateCommand = new SqlCommand(sql, connection);
                    adapter.UpdateCommand.ExecuteNonQuery();
                    //MessageBox.Show("Row updated !! ");
                }
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }
            return sql;
        }


        public static string DeleteFrom(string sql, string connection_string)       //Użycie delete i inserta nie jest optymalne, ale w niektórych przypadkach pozycje nie miały atrybutu. Po testach zmienić na update
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    connection.Open();
                    adapter.DeleteCommand = new SqlCommand(sql, connection);
                    adapter.DeleteCommand.ExecuteNonQuery();
                    //MessageBox.Show("Row deleted!! ");
                }
            }
            catch (Exception ex)
            {
               // MessageBox.Show(ex.ToString(), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return sql;
        }

        public static string ConnectToSQL(string connection_string, string query)
        {
            string str = "";
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    adapter.SelectCommand = new SqlCommand(query, connection);
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet, "Query");
                    if (dataSet.Tables.Count == 0)
                    {
                        return "";
                    }
                    foreach (DataRow row in (InternalDataCollectionBase)dataSet.Tables[0].Rows)
                    {
                        str = row[0].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                str = str + "Bład SQL:" + ex.ToString() + "\n";
            }
            return str;
        }

        public static string InsertInto(string sql, string connection_string)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    connection.Open();
                    adapter.InsertCommand = new SqlCommand(sql, connection);
                    adapter.InsertCommand.ExecuteNonQuery();
                    //MessageBox.Show("Row inserted !! ");
                }
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }
            return sql;
        }
    }
}
