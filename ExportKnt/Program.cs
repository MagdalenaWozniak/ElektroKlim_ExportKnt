using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExportKnt
{
    static class Program
    {
        /// <summary>
        /// Główny punkt wejścia dla aplikacji.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Form1());

            //string connection_string = @"Data Source = EKSRV04\COMARCH; Initial Catalog = ERPXL_E; User ID = ComarchAPI; Password=hT@S+NQZBT";
            // string connection_string2 = @"Data Source = EKSRV02\WORKFLOW; Initial Catalog = Elektroklim_Shared; User ID = ComarchIntegracja; Password=JKLH@!g2wDSFSKJ!ANHoq217P8n6G1h";
            //string connection_string2 = @"Data Source = 172.51.1.7; Initial Catalog = Elektroklim_Shared; User ID = ComarchIntegracja; Password=JKLH@!g2wDSFSKJ!ANHoq217P8n6G1h";

            string connection_string = @"Data Source = EKSRV04CLONE\COMARCH; Initial Catalog = ERPXL_UP; User ID = ComarchAPI; Password=hT@S+NQZBT";
            string connection_string2 = @"Data Source = EKSRV03\WORKFLOW; Initial Catalog = Elektroklim_Shared; User ID = ComarchIntegracja; Password=JKLH@!g2wDSFSKJ!ANHoq217P8n6G1h";

            TextWriter tw2 = new StreamWriter(@"C:\OneSupport\ExportKNT\log.txt", append: true);
            tw2.WriteLine();
            tw2.WriteLine();
            tw2.WriteLine("Hello " + DateTime.Now.ToString("f"));

            /*---------------------------------------------------------------UPDATE-------------------------------------------------------------------*/

            string updateKntQuery = "SELECT [GIDTyp] ,[GIDNumer] ,[akronim] ,[WF_ID] ,[firma] FROM [CDN].[UpdateKnt]";
            List<UpdateKNT> updateKontrahenci = UpdateKNT.SelectListRowColumn(updateKntQuery, "cdn.UpdateKnt", connection_string);

            foreach(UpdateKNT uKnt in updateKontrahenci)
            {
                try
                {
                    tw2.WriteLine("Update z tabeli: " + uKnt.firma.Trim()+"_"+uKnt.akronim.Trim());
                    string ukntQuery = "if exists(select * from cdn.kntKarty   join cdn.RachunkiBankowe on RkB_ObiFirma = Knt_GIDFirma and RkB_ObiLp = Knt_GIDLp and RkB_ObiNumer = Knt_GIDNumer and " 
                            + " RkB_ObiTyp = Knt_GIDTyp  where Knt_GIDTyp = " + uKnt.GIDTyp + " and Knt_GIDNumer = " + uKnt.GIDNumer + "  and RkB_Domyslny = 1) " 
                            + " SELECT Knt_GIDTYP, Knt_GIDFirma, Knt_GIDNumer, Knt_GIDLp, Knt_Akronim, Knt_Nazwa1, Knt_Nazwa2, Knt_Nazwa3, Knt_Miasto, Knt_Ulica, Knt_KodP, RkB_NrRachunku, KPC_NazwaPL, "
                            + " Knt_Telefon1, Knt_Telefon2, Knt_EMail, Knt_Nip, Knt_NipPrefiks, Knt_Pesel, Knt_Regon, Knt_TerminPlZak, Knt_Symbol, Knt_PodzialPlat, Knt_URL, Knt_Typ FROM cdn.KntKarty "
                            + " join cdn.RachunkiBankowe on RkB_ObiFirma = Knt_GIDFirma and RkB_ObiLp = Knt_GIDLp and RkB_ObiNumer = Knt_GIDNumer and RkB_ObiTyp = Knt_GIDTyp "
                            + " join cdn.KrajeCelne on Knt_Kraj = KPC_Kod "
                            + " WHERE Knt_GIDTyp="+uKnt.GIDTyp+" and Knt_GIDNumer="+uKnt.GIDNumer +" and RkB_Domyslny = 1 "
                            + " else SELECT Knt_GIDTYP, Knt_GIDFirma, Knt_GIDNumer, Knt_GIDLp, Knt_Akronim, Knt_Nazwa1, Knt_Nazwa2, Knt_Nazwa3, Knt_Miasto, Knt_Ulica, Knt_KodP, '', KPC_NazwaPL, "
                            + " Knt_Telefon1, Knt_Telefon2, Knt_EMail, Knt_Nip, Knt_NipPrefiks, Knt_Pesel, Knt_Regon, Knt_TerminPlZak, Knt_Symbol, Knt_PodzialPlat, Knt_URL, Knt_Typ FROM cdn.KntKarty "
                            + " join cdn.KrajeCelne on Knt_Kraj = KPC_Kod "
                            + " WHERE Knt_GIDTyp = " + uKnt.GIDTyp + " and Knt_GIDNumer = " + uKnt.GIDNumer;

                    List<KNT> kontrahent = KNT.SelectListRowColumn(ukntQuery, "cdn.kntKarty", connection_string);

                    if(kontrahent.Count==1)
                    { 
                        try
                        {

                            tw2.WriteLine("Update kontrahenta: " + uKnt.akronim+" z WF_ID: "+uKnt.WF_ID);
                            int typ = 0;
                            switch (kontrahent[0].typ)
                            {
                                case 8:
                                    {
                                        typ = 2;
                                        break;
                                    }
                                case 16:
                                    {
                                        typ = 1;
                                        break;
                                    }
                                case 24:
                                    {
                                        typ = 3;
                                        break;
                                    }
                                default:
                                    {
                                        tw2.WriteLine("Dla kontrahenta: " + kontrahent[0].akronim + " jest błędny typ " + kontrahent[0].typ);
                                        typ = 8;
                                        break;
                                    }
                            }

                            int ceidg = 0;
                            if (kontrahent[0].regon != "")
                                ceidg = 1;
                            string akronim = uKnt.firma.Trim() +"_" + uKnt.akronim.Trim();
                            //if (knt.akronim.Substring(0, 5) != "EKSP_")
                            //  akronim = "EKSP_" + knt.akronim;

                            string wersjaQuery = "SELECT  RecordVersion FROM dbo.tb_Contractors WHERE ID=" + uKnt.WF_ID;//Acronym = '" + akronim + "'";
                            //tw2.WriteLine("Wersja query: " + wersjaQuery);
                            int wersja = Convert.ToInt32(KNT.ConnectToSQL(connection_string2, wersjaQuery).Trim())+1;
                            tw2.WriteLine("Wersja: " + wersja);

                            string numerKRSQuery = "Select Atr_Wartosc  from cdn.atrybuty join cdn.KntKarty on Knt_GIDNumer = Atr_ObiNumer AND Knt_GIDTyp = Atr_Obityp where Atr_AtkId = 55 " +
                                                    " AND Knt_GIDTyp = " + uKnt.GIDTyp + "  AND Knt_GIDNumer = " + uKnt.GIDNumer;
                            string numerKRS = KNT.ConnectToSQL(connection_string, numerKRSQuery);
                            tw2.WriteLine("Numer KRS: " + numerKRS);
                           
                            if (kontrahent[0].nipPrefiks.Contains("PL"))
                                kontrahent[0].nipPrefiks = "";



                            string kntUpdate = "UPDATE [dbo].[tb_Contractors] SET [ContractorName] = '" + kontrahent[0].nazwa1 + " " + kontrahent[0].nazwa2 + " " + kontrahent[0].nazwa3 + "' ,[ContractorType] = " + typ
                                                + "  ,[Country] = '" + kontrahent[0].kraj + "' ,[City] = '" + kontrahent[0].miasto + "' ,[Street] = '" + kontrahent[0].ulica + "' ,[PostalCode] = '" + kontrahent[0].kodP + "'"
                                                + "  ,[AccountNumber] = '" + kontrahent[0].nrRachunku + "' ,[PhoneNumber1] = '" + kontrahent[0].telefon1 + "' ,[PhoneNumber2] = '" + kontrahent[0].telefon2 + "'"
                                                + "  ,[Email] = '" + kontrahent[0].eMail + "' ,[TIN] = '" + kontrahent[0].nipPrefiks + kontrahent[0].nip + "',[PersonalIdentityNumber] = '" + kontrahent[0].pesel + "'"
                                                + " ,[RecordVersion] = " + wersja + " ,[ModifiedBy] = 'ERP' ,[ModifiedDate] = '" + DateTime.Now.ToString("yyyy - MM - dd HH: mm: ss.fff") + "' ,[ExtID] = " + kontrahent[0].gidNumer
                                                + "  ,[CEIDG] = " + ceidg + " ,[REGON] = '" + kontrahent[0].regon + "' ,[WWW] = '" + kontrahent[0].url + "' ,[PaymentDays] = " + kontrahent[0].termin + " ,[CurrencyCode] = '" + kontrahent[0].waluta + "'"
                                                + " ,[SplitPayment] = " + kontrahent[0].podzialPlat + ", [KRS]='" + numerKRS + "' WHERE ID=" + uKnt.WF_ID;//Acronym='" + akronim + "'";
                            KNT.Update(kntUpdate, connection_string2);

                            string bnkQuery = "SELECT Bnk_Nazwa, RkB_NrRachunku, Knt_Symbol, RkB_Domyslny FROM cdn.KntKarty "
                                                + " join cdn.RachunkiBankowe on RkB_ObiFirma = Knt_GIDFirma and RkB_ObiLp = Knt_GIDLp and RkB_ObiNumer = Knt_GIDNumer and RkB_ObiTyp = Knt_GIDTyp "
                                                + " join cdn.Banki on RkB_BnkFirma = Bnk_GIDFirma and RkB_BnkLp = Bnk_GIDLp and RkB_BnkNumer = Bnk_GIDNumer and RkB_BnkTyp = Bnk_GIDTyp "
                                                + " WHERE Knt_GIDTyp =" + kontrahent[0].gidTyp + " and KNT_GIDFirma =" + kontrahent[0].gidFirma + " and Knt_GIDNumer = " + kontrahent[0].gidNumer + " and Knt_GIDLp = " + kontrahent[0].gidLp;

                            List<Banki> banki = Banki.SelectListRowColumn(bnkQuery, "cdn.kntKarty", connection_string);

                            foreach (Banki bnk in banki)
                            {
                                try
                                {

                                    string czyIstniejeQuery = "if exists(SELECT * FROM [dbo].[tb_ContractorAccounts] where ContractorID = "+uKnt.WF_ID+" and AccountNumber = '"+bnk.nrRachunku+"') SELECT 'TAK' else SELECT 'NIE'";
                                    string czyIstnieje = KNT.ConnectToSQL(connection_string2, czyIstniejeQuery);

                                    if (czyIstnieje == "TAK")
                                    {
                                        tw2.WriteLine("Modyfikacja nr rachunku: " + bnk.nrRachunku);
                                        string updateBnk = "UPDATE [dbo].[tb_ContractorAccounts] SET [AccountName] = '" + bnk.nazwa + "' , State=0 ,[CurrencyCode] = '" + bnk.waluta + "' ,[IsMajor] = " + bnk.domyslny
                                                            + " ,[ModifiedDate] = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "' ,[ModifiedBy] = 'ERP' WHERE ContractorID = " + uKnt.WF_ID
                                                            + " and AccountNumber = '" + bnk.nrRachunku + "'";
                                        //tw2.WriteLine("Update bnk: " + updateBnk);
                                        KNT.Update(updateBnk, connection_string2);

                                    }
                                    else
                                    {
                                        tw2.WriteLine("Dodawanie nr rachunku: " + bnk.nrRachunku);
                                        string bnkInsert = "INSERT INTO [dbo].[tb_ContractorAccounts] ([ContractorID] ,[AccountName] ,[AccountNumber] ,[CurrencyCode] ,[State] ,[IsMajor] ,[CreatedDate] ,[CreatedBy] ,[ModifiedDate] ,[ModifiedBy]) "
                                                            + " VALUES (" + uKnt.WF_ID + " , '" + bnk.nazwa + "' , '" + bnk.nrRachunku + "' , '" + bnk.waluta + "' , 0 , " + bnk.domyslny + " , '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "' , 'ERP' , '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "' ,"
                                                            + " 'ERP' )";

                                        KNT.InsertInto(bnkInsert, connection_string2);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    tw2.WriteLine("Błąd modyfikacji rachunku: " + bnk.nrRachunku + "\n" + ex.Message);
                                }



                            }
                            banki.Clear();

                            string updateState = "update [dbo].[tb_ContractorAccounts] set state =1  WHERE (YEAR(ModifiedDate) != " + DateTime.Now.ToString("yyyy") + "  or MONTH(ModifiedDate) != " + DateTime.Now.ToString("MM")
                                                + " or DAY(ModifiedDate) != " + DateTime.Now.ToString("dd") + " or DATEPART(Hour, ModifiedDate) != " + DateTime.Now.ToString("HH") + ") and ContractorID = " + uKnt.WF_ID;
                            KNT.Update(updateState, connection_string2);

                           
                        }
                        catch (Exception e)
                        {
                            tw2.WriteLine("Błąd updatu kontrahenta: " + uKnt.akronim + "\n" + e.Message);
                        }

                    }


                }
                catch(Exception e)
                {
                    tw2.WriteLine("Błąd updatu kontrahenta: " + uKnt.akronim + "\n" + e.Message);
                }
            }


            string deleteTable = "delete from cdn.UpdateKnt";
            KNT.DeleteFrom(deleteTable, connection_string);



            /*------------------------------------------------------------INSERT--------------------------------------------------------------------*/

            string kntQuery = "SELECT Knt_GIDTYP, Knt_GIDFirma, Knt_GIDNumer, Knt_GIDLp, Knt_Akronim, Knt_Nazwa1, Knt_Nazwa2, Knt_Nazwa3, Knt_Miasto, Knt_Ulica, Knt_KodP, RkB_NrRachunku, KPC_NazwaPL, "
                            + " Knt_Telefon1, Knt_Telefon2, Knt_EMail, Knt_Nip, Knt_NipPrefiks, Knt_Pesel, Knt_Regon, Knt_TerminPlZak, Knt_Symbol, Knt_PodzialPlat, Knt_URL, Knt_Typ FROM cdn.KntKarty "
                            + " join cdn.Atrybuty on Knt_GIDFirma = Atr_ObiFirma and Knt_GIDLp = Atr_ObiLp and Knt_GIDNumer = Atr_ObiNumer and Knt_GIDTyp = Atr_ObiTyp "
                            + " join cdn.AtrybutyKlasy on Atk_ID = Atr_AtkId "
                            + " join cdn.RachunkiBankowe on RkB_ObiFirma = Knt_GIDFirma and RkB_ObiLp = Knt_GIDLp and RkB_ObiNumer = Knt_GIDNumer and RkB_ObiTyp = Knt_GIDTyp "
                            + " join cdn.KrajeCelne on Knt_Kraj = KPC_Kod "
                            + " WHERE AtK_Nazwa = 'Export WF' and Atr_Wartosc = 'NIE' and RkB_Domyslny = 1";

            List<KNT> kontrahenci = KNT.SelectListRowColumn(kntQuery, "cdn.kntKarty", connection_string);

           

            foreach (KNT knt in kontrahenci)
            {
                bool czyIstnieje = CzyIstnieje("EKI", knt, tw2, connection_string, connection_string2);
                if(czyIstnieje==false)
                    dodajKnt("EKI", knt, tw2, connection_string, connection_string2);

            }

            kontrahenci.Clear();



            string kntQuerybezBnk = "SELECT Knt_GIDTYP, Knt_GIDFirma, Knt_GIDNumer, Knt_GIDLp, Knt_Akronim, Knt_Nazwa1, Knt_Nazwa2, Knt_Nazwa3, Knt_Miasto, Knt_Ulica, Knt_KodP, '' , KPC_NazwaPL, "
                           + " Knt_Telefon1, Knt_Telefon2, Knt_EMail, Knt_Nip, Knt_NipPrefiks, Knt_Pesel, Knt_Regon, Knt_TerminPlZak, Knt_Symbol, Knt_PodzialPlat, Knt_URL, Knt_Typ FROM cdn.KntKarty "
                           + " join cdn.Atrybuty on Knt_GIDFirma = Atr_ObiFirma and Knt_GIDLp = Atr_ObiLp and Knt_GIDNumer = Atr_ObiNumer and Knt_GIDTyp = Atr_ObiTyp "
                           + " join cdn.AtrybutyKlasy on Atk_ID = Atr_AtkId " 
                           + " join cdn.KrajeCelne on Knt_Kraj = KPC_Kod "
                           + " WHERE AtK_Nazwa = 'Export WF' and Atr_Wartosc = 'NIE'";

            kontrahenci = KNT.SelectListRowColumn(kntQuerybezBnk, "cdn.kntKarty", connection_string);



            foreach (KNT knt in kontrahenci)
            {
                bool czyIstnieje = CzyIstnieje("EKI", knt, tw2, connection_string, connection_string2);
                if (czyIstnieje == false)
                    dodajKnt("EKI", knt, tw2, connection_string, connection_string2);

            }

            kontrahenci.Clear();

        /*    kntQuery = "SELECT Knt_GIDTYP, Knt_GIDFirma, Knt_GIDNumer, Knt_GIDLp, Knt_Akronim, Knt_Nazwa1, Knt_Nazwa2, Knt_Nazwa3, Knt_Miasto, Knt_Ulica, Knt_KodP, RkB_NrRachunku, KPC_NazwaPL, "
                           + "  Knt_Telefon1, Knt_Telefon2, Knt_EMail, Knt_Nip, Knt_NipPrefiks, Knt_Pesel, Knt_Regon, Knt_TerminPlZak, Knt_Symbol, Knt_PodzialPlat, Knt_URL, Knt_Typ FROM cdn.KntKarty "
                           + "  join cdn.Atrybuty as a1 on Knt_GIDFirma = a1.Atr_ObiFirma and Knt_GIDLp = a1.Atr_ObiLp and Knt_GIDNumer = a1.Atr_ObiNumer and Knt_GIDTyp = a1.Atr_ObiTyp and a1.Atr_AtkId = 42 "
                           + "  join cdn.Atrybuty as a2 on Knt_GIDFirma = a2.Atr_ObiFirma and Knt_GIDLp = a2.Atr_ObiLp and Knt_GIDNumer = a2.Atr_ObiNumer and Knt_GIDTyp = a2.Atr_ObiTyp and a2.Atr_AtkId = 43 "
                           + "  join cdn.RachunkiBankowe on RkB_ObiFirma = Knt_GIDFirma and RkB_ObiLp = Knt_GIDLp and RkB_ObiNumer = Knt_GIDNumer and RkB_ObiTyp = Knt_GIDTyp "
                           + "  join cdn.KrajeCelne on Knt_Kraj = KPC_Kod "
                           + "  WHERE a1.Atr_Wartosc = 'TAK' and a2.Atr_Wartosc = '-' and RkB_Domyslny = 1";

            kontrahenci = KNT.SelectListRowColumn(kntQuery, "cdn.kntKarty", connection_string);

            foreach (KNT knt in kontrahenci)
            {
                dodajKnt("EKSP", knt, tw2, connection_string, connection_string2);

            }

            kontrahenci.Clear();


            kntQuerybezBnk = "SELECT Knt_GIDTYP, Knt_GIDFirma, Knt_GIDNumer, Knt_GIDLp, Knt_Akronim, Knt_Nazwa1, Knt_Nazwa2, Knt_Nazwa3, Knt_Miasto, Knt_Ulica, Knt_KodP, '', KPC_NazwaPL, "
                          + "  Knt_Telefon1, Knt_Telefon2, Knt_EMail, Knt_Nip, Knt_NipPrefiks, Knt_Pesel, Knt_Regon, Knt_TerminPlZak, Knt_Symbol, Knt_PodzialPlat, Knt_URL, Knt_Typ FROM cdn.KntKarty "
                          + "  join cdn.Atrybuty as a1 on Knt_GIDFirma = a1.Atr_ObiFirma and Knt_GIDLp = a1.Atr_ObiLp and Knt_GIDNumer = a1.Atr_ObiNumer and Knt_GIDTyp = a1.Atr_ObiTyp and a1.Atr_AtkId = 42 "
                          + "  join cdn.Atrybuty as a2 on Knt_GIDFirma = a2.Atr_ObiFirma and Knt_GIDLp = a2.Atr_ObiLp and Knt_GIDNumer = a2.Atr_ObiNumer and Knt_GIDTyp = a2.Atr_ObiTyp and a2.Atr_AtkId = 43 "
                          + "  join cdn.KrajeCelne on Knt_Kraj = KPC_Kod "
                          + "  WHERE a1.Atr_Wartosc = 'TAK' and a2.Atr_Wartosc = '-' ";

            kontrahenci = KNT.SelectListRowColumn(kntQuerybezBnk, "cdn.kntKarty", connection_string);

            foreach (KNT knt in kontrahenci)
            {
                dodajKnt("EKSP", knt, tw2, connection_string, connection_string2);

            }

            kontrahenci.Clear();


            kntQuery = "SELECT Knt_GIDTYP, Knt_GIDFirma, Knt_GIDNumer, Knt_GIDLp, Knt_Akronim, Knt_Nazwa1, Knt_Nazwa2, Knt_Nazwa3, Knt_Miasto, Knt_Ulica, Knt_KodP, RkB_NrRachunku, KPC_NazwaPL, "
                          + "  Knt_Telefon1, Knt_Telefon2, Knt_EMail, Knt_Nip, Knt_NipPrefiks, Knt_Pesel, Knt_Regon, Knt_TerminPlZak, Knt_Symbol, Knt_PodzialPlat, Knt_URL, Knt_Typ FROM cdn.KntKarty "
                          + "  join cdn.Atrybuty as a1 on Knt_GIDFirma = a1.Atr_ObiFirma and Knt_GIDLp = a1.Atr_ObiLp and Knt_GIDNumer = a1.Atr_ObiNumer and Knt_GIDTyp = a1.Atr_ObiTyp and a1.Atr_AtkId = 42 "
                          + "  join cdn.Atrybuty as a2 on Knt_GIDFirma = a2.Atr_ObiFirma and Knt_GIDLp = a2.Atr_ObiLp and Knt_GIDNumer = a2.Atr_ObiNumer and Knt_GIDTyp = a2.Atr_ObiTyp and a2.Atr_AtkId = 44 "
                          + "  join cdn.RachunkiBankowe on RkB_ObiFirma = Knt_GIDFirma and RkB_ObiLp = Knt_GIDLp and RkB_ObiNumer = Knt_GIDNumer and RkB_ObiTyp = Knt_GIDTyp "
                          + "  join cdn.KrajeCelne on Knt_Kraj = KPC_Kod "
                          + "  WHERE a1.Atr_Wartosc = 'TAK' and a2.Atr_Wartosc = '-' and RkB_Domyslny = 1";

            kontrahenci = KNT.SelectListRowColumn(kntQuery, "cdn.kntKarty", connection_string);

            foreach (KNT knt in kontrahenci)
            {
                dodajKnt("EKPW", knt, tw2, connection_string, connection_string2);

            }

            kontrahenci.Clear();


            kntQuerybezBnk = "SELECT Knt_GIDTYP, Knt_GIDFirma, Knt_GIDNumer, Knt_GIDLp, Knt_Akronim, Knt_Nazwa1, Knt_Nazwa2, Knt_Nazwa3, Knt_Miasto, Knt_Ulica, Knt_KodP,'', KPC_NazwaPL, "
                          + "  Knt_Telefon1, Knt_Telefon2, Knt_EMail, Knt_Nip, Knt_NipPrefiks, Knt_Pesel, Knt_Regon, Knt_TerminPlZak, Knt_Symbol, Knt_PodzialPlat, Knt_URL, Knt_Typ FROM cdn.KntKarty "
                          + "  join cdn.Atrybuty as a1 on Knt_GIDFirma = a1.Atr_ObiFirma and Knt_GIDLp = a1.Atr_ObiLp and Knt_GIDNumer = a1.Atr_ObiNumer and Knt_GIDTyp = a1.Atr_ObiTyp and a1.Atr_AtkId = 42 "
                          + "  join cdn.Atrybuty as a2 on Knt_GIDFirma = a2.Atr_ObiFirma and Knt_GIDLp = a2.Atr_ObiLp and Knt_GIDNumer = a2.Atr_ObiNumer and Knt_GIDTyp = a2.Atr_ObiTyp and a2.Atr_AtkId = 44 "
                          + "  join cdn.KrajeCelne on Knt_Kraj = KPC_Kod "
                          + "  WHERE a1.Atr_Wartosc = 'TAK' and a2.Atr_Wartosc = '-' ";

            kontrahenci = KNT.SelectListRowColumn(kntQuerybezBnk, "cdn.kntKarty", connection_string);

            foreach (KNT knt in kontrahenci)
            {
                dodajKnt("EKPW", knt, tw2, connection_string, connection_string2);

            }

            kontrahenci.Clear();*/

            kntQuery = "SELECT Knt_GIDTYP, Knt_GIDFirma, Knt_GIDNumer, Knt_GIDLp, Knt_Akronim, Knt_Nazwa1, Knt_Nazwa2, Knt_Nazwa3, Knt_Miasto, Knt_Ulica, Knt_KodP, RkB_NrRachunku, KPC_NazwaPL, "
                          + "  Knt_Telefon1, Knt_Telefon2, Knt_EMail, Knt_Nip, Knt_NipPrefiks, Knt_Pesel, Knt_Regon, Knt_TerminPlZak, Knt_Symbol, Knt_PodzialPlat, Knt_URL, Knt_Typ FROM cdn.KntKarty "
                          + "  join cdn.Atrybuty as a1 on Knt_GIDFirma = a1.Atr_ObiFirma and Knt_GIDLp = a1.Atr_ObiLp and Knt_GIDNumer = a1.Atr_ObiNumer and Knt_GIDTyp = a1.Atr_ObiTyp and a1.Atr_AtkId = 42 "
                          + "  join cdn.Atrybuty as a2 on Knt_GIDFirma = a2.Atr_ObiFirma and Knt_GIDLp = a2.Atr_ObiLp and Knt_GIDNumer = a2.Atr_ObiNumer and Knt_GIDTyp = a2.Atr_ObiTyp and a2.Atr_AtkId = 45 "
                          + "  join cdn.RachunkiBankowe on RkB_ObiFirma = Knt_GIDFirma and RkB_ObiLp = Knt_GIDLp and RkB_ObiNumer = Knt_GIDNumer and RkB_ObiTyp = Knt_GIDTyp "
                          + "  join cdn.KrajeCelne on Knt_Kraj = KPC_Kod "
                          + "  WHERE a1.Atr_Wartosc = 'TAK' and a2.Atr_Wartosc = '-' and RkB_Domyslny = 1";

            kontrahenci = KNT.SelectListRowColumn(kntQuery, "cdn.kntKarty", connection_string);

            foreach (KNT knt in kontrahenci)
            {
                bool czyIstnieje = CzyIstnieje("WABE", knt, tw2, connection_string, connection_string2);
                if (czyIstnieje == false)
                    dodajKnt("WABE", knt, tw2, connection_string, connection_string2);

            }

            kontrahenci.Clear();


            kntQuerybezBnk = "SELECT Knt_GIDTYP, Knt_GIDFirma, Knt_GIDNumer, Knt_GIDLp, Knt_Akronim, Knt_Nazwa1, Knt_Nazwa2, Knt_Nazwa3, Knt_Miasto, Knt_Ulica, Knt_KodP, '', KPC_NazwaPL, "
                          + "  Knt_Telefon1, Knt_Telefon2, Knt_EMail, Knt_Nip, Knt_NipPrefiks, Knt_Pesel, Knt_Regon, Knt_TerminPlZak, Knt_Symbol, Knt_PodzialPlat, Knt_URL, Knt_Typ FROM cdn.KntKarty "
                          + "  join cdn.Atrybuty as a1 on Knt_GIDFirma = a1.Atr_ObiFirma and Knt_GIDLp = a1.Atr_ObiLp and Knt_GIDNumer = a1.Atr_ObiNumer and Knt_GIDTyp = a1.Atr_ObiTyp and a1.Atr_AtkId = 42 "
                          + "  join cdn.Atrybuty as a2 on Knt_GIDFirma = a2.Atr_ObiFirma and Knt_GIDLp = a2.Atr_ObiLp and Knt_GIDNumer = a2.Atr_ObiNumer and Knt_GIDTyp = a2.Atr_ObiTyp and a2.Atr_AtkId = 45 "
                          + "  join cdn.KrajeCelne on Knt_Kraj = KPC_Kod "
                          + "  WHERE a1.Atr_Wartosc = 'TAK' and a2.Atr_Wartosc = '-' ";

            kontrahenci = KNT.SelectListRowColumn(kntQuerybezBnk, "cdn.kntKarty", connection_string);

            foreach (KNT knt in kontrahenci)
            {
                bool czyIstnieje = CzyIstnieje("WABE", knt, tw2, connection_string, connection_string2);
                if (czyIstnieje == false)
                    dodajKnt("WABE", knt, tw2, connection_string, connection_string2);

            }

            kontrahenci.Clear();


            kntQuery = "SELECT Knt_GIDTYP, Knt_GIDFirma, Knt_GIDNumer, Knt_GIDLp, Knt_Akronim, Knt_Nazwa1, Knt_Nazwa2, Knt_Nazwa3, Knt_Miasto, Knt_Ulica, Knt_KodP, RkB_NrRachunku, KPC_NazwaPL, "
                         + "  Knt_Telefon1, Knt_Telefon2, Knt_EMail, Knt_Nip, Knt_NipPrefiks, Knt_Pesel, Knt_Regon, Knt_TerminPlZak, Knt_Symbol, Knt_PodzialPlat, Knt_URL, Knt_Typ FROM cdn.KntKarty "
                         + "  join cdn.Atrybuty as a1 on Knt_GIDFirma = a1.Atr_ObiFirma and Knt_GIDLp = a1.Atr_ObiLp and Knt_GIDNumer = a1.Atr_ObiNumer and Knt_GIDTyp = a1.Atr_ObiTyp and a1.Atr_AtkId = 42 "
                         + "  join cdn.Atrybuty as a2 on Knt_GIDFirma = a2.Atr_ObiFirma and Knt_GIDLp = a2.Atr_ObiLp and Knt_GIDNumer = a2.Atr_ObiNumer and Knt_GIDTyp = a2.Atr_ObiTyp and a2.Atr_AtkId = 47 "
                         + "  join cdn.RachunkiBankowe on RkB_ObiFirma = Knt_GIDFirma and RkB_ObiLp = Knt_GIDLp and RkB_ObiNumer = Knt_GIDNumer and RkB_ObiTyp = Knt_GIDTyp "
                         + "  join cdn.KrajeCelne on Knt_Kraj = KPC_Kod "
                         + "  WHERE a1.Atr_Wartosc = 'TAK' and a2.Atr_Wartosc = '-' and RkB_Domyslny = 1";

            kontrahenci = KNT.SelectListRowColumn(kntQuery, "cdn.kntKarty", connection_string);

            foreach (KNT knt in kontrahenci)
            {
                bool czyIstnieje = CzyIstnieje("WD", knt, tw2, connection_string, connection_string2);
                if (czyIstnieje == false)
                    dodajKnt("WD", knt, tw2, connection_string, connection_string2);

            }

            kontrahenci.Clear();


            kntQuerybezBnk = "SELECT Knt_GIDTYP, Knt_GIDFirma, Knt_GIDNumer, Knt_GIDLp, Knt_Akronim, Knt_Nazwa1, Knt_Nazwa2, Knt_Nazwa3, Knt_Miasto, Knt_Ulica, Knt_KodP, '', KPC_NazwaPL, "
                          + "  Knt_Telefon1, Knt_Telefon2, Knt_EMail, Knt_Nip, Knt_NipPrefiks, Knt_Pesel, Knt_Regon, Knt_TerminPlZak, Knt_Symbol, Knt_PodzialPlat, Knt_URL, Knt_Typ FROM cdn.KntKarty "
                          + "  join cdn.Atrybuty as a1 on Knt_GIDFirma = a1.Atr_ObiFirma and Knt_GIDLp = a1.Atr_ObiLp and Knt_GIDNumer = a1.Atr_ObiNumer and Knt_GIDTyp = a1.Atr_ObiTyp and a1.Atr_AtkId = 42 "
                          + "  join cdn.Atrybuty as a2 on Knt_GIDFirma = a2.Atr_ObiFirma and Knt_GIDLp = a2.Atr_ObiLp and Knt_GIDNumer = a2.Atr_ObiNumer and Knt_GIDTyp = a2.Atr_ObiTyp and a2.Atr_AtkId = 47 "
                          + "  join cdn.KrajeCelne on Knt_Kraj = KPC_Kod "
                          + "  WHERE a1.Atr_Wartosc = 'TAK' and a2.Atr_Wartosc = '-' ";

            kontrahenci = KNT.SelectListRowColumn(kntQuerybezBnk, "cdn.kntKarty", connection_string);

            foreach (KNT knt in kontrahenci)
            {
                bool czyIstnieje = CzyIstnieje("WD", knt, tw2, connection_string, connection_string2);
                if (czyIstnieje == false)
                    dodajKnt("WD", knt, tw2, connection_string, connection_string2);

            }

            kontrahenci.Clear();





            tw2.Close();
        }


        public static bool CzyIstnieje(string firma, KNT knt, TextWriter tw2, string connection_string, string connection_string2)
        {
            try
            {
                string czyAkronimQuery = "if exists (SELECT ID from dbo.tb_Contractors where Acronym='" + firma + "_" + knt.akronim + "') SELECT ID from dbo.tb_Contractors where Acronym='" + firma + "_" + knt.akronim + "' else SELECT 0";
                string czyAkronim = KNT.ConnectToSQL(connection_string2, czyAkronimQuery);
                if (Convert.ToInt32(czyAkronim) != 0)
                {
                    tw2.WriteLine("Znaleziono kontrahenta po akronimie " + firma + "_" + knt.akronim + " : " + czyAkronim);
                    if (firma == "EKI")
                    {
                        string atrUpdate = "UPDATE e SET [Atr_Wartosc] = 'TAK' FROM [CDN].[Atrybuty] as e   join cdn.AtrybutyKlasy on Atk_ID = Atr_AtkId WHERE AtK_Nazwa = 'Export WF' and Atr_ObiFirma =" + knt.gidFirma + " and Atr_ObiLp = " + knt.gidLp +
                                        " and Atr_ObiNumer =" + knt.gidNumer + " and Atr_ObiTyp = " + knt.gidTyp;
                        KNT.Update(atrUpdate, connection_string);

                        atrUpdate = "UPDATE e SET [Atr_Wartosc] = '" + czyAkronim + "' FROM [CDN].[Atrybuty] as e   join cdn.AtrybutyKlasy on Atk_ID = Atr_AtkId WHERE AtK_Nazwa = 'WF_KontrahentID' and Atr_ObiFirma =" + knt.gidFirma + " and Atr_ObiLp = " + knt.gidLp +
                                        " and Atr_ObiNumer =" + knt.gidNumer + " and Atr_ObiTyp = " + knt.gidTyp;
                        KNT.Update(atrUpdate, connection_string);
                    }
                    else
                    {
                        string atrUpdate = "UPDATE e SET [Atr_Wartosc] = '" + czyAkronim + "' FROM [CDN].[Atrybuty] as e   join cdn.AtrybutyKlasy on Atk_ID = Atr_AtkId WHERE AtK_Nazwa = 'WF_KontrahentID_" + firma + "' and Atr_ObiFirma =" + knt.gidFirma + " and Atr_ObiLp = " + knt.gidLp +
                                        " and Atr_ObiNumer =" + knt.gidNumer + " and Atr_ObiTyp = " + knt.gidTyp;
                        KNT.Update(atrUpdate, connection_string);
                    }

                    return true;
                }

                string czyNIPQuery = "if exists (SELECT ID from dbo.tb_Contractors where TIN='" + knt.nip + "' and CompanyID='"+firma+"') SELECT ID from dbo.tb_Contractors where TIN='" + knt.nip + "' and CompanyID='"+firma+"' else SELECT 0";
                string czyNIP = KNT.ConnectToSQL(connection_string2, czyNIPQuery);

                if (Convert.ToInt32(czyNIP) != 0 && knt.nip!="")
                {
                    tw2.WriteLine("Znaleziono kontrahenta po NIPie " + knt.nip + " : " + czyNIP);
                    if (firma == "EKI")
                    {
                        string atrUpdate = "UPDATE e SET [Atr_Wartosc] = 'TAK' FROM [CDN].[Atrybuty] as e   join cdn.AtrybutyKlasy on Atk_ID = Atr_AtkId WHERE AtK_Nazwa = 'Export WF' and Atr_ObiFirma =" + knt.gidFirma + " and Atr_ObiLp = " + knt.gidLp +
                                        " and Atr_ObiNumer =" + knt.gidNumer + " and Atr_ObiTyp = " + knt.gidTyp;
                        KNT.Update(atrUpdate, connection_string);

                        atrUpdate = "UPDATE e SET [Atr_Wartosc] = '" + czyNIP + "' FROM [CDN].[Atrybuty] as e   join cdn.AtrybutyKlasy on Atk_ID = Atr_AtkId WHERE AtK_Nazwa = 'WF_KontrahentID' and Atr_ObiFirma =" + knt.gidFirma + " and Atr_ObiLp = " + knt.gidLp +
                                        " and Atr_ObiNumer =" + knt.gidNumer + " and Atr_ObiTyp = " + knt.gidTyp;
                        KNT.Update(atrUpdate, connection_string);
                    }
                    else
                    {
                        string atrUpdate = "UPDATE e SET [Atr_Wartosc] = '" + czyNIP + "' FROM [CDN].[Atrybuty] as e   join cdn.AtrybutyKlasy on Atk_ID = Atr_AtkId WHERE AtK_Nazwa = 'WF_KontrahentID_" + firma + "' and Atr_ObiFirma =" + knt.gidFirma + " and Atr_ObiLp = " + knt.gidLp +
                                        " and Atr_ObiNumer =" + knt.gidNumer + " and Atr_ObiTyp = " + knt.gidTyp;
                        KNT.Update(atrUpdate, connection_string);
                    }

                    return true;
                }

                tw2.WriteLine("Nie znaleziono kontrahenta");
                return false;
            }
            catch (Exception e)
            {
                tw2.WriteLine("Błąd szukania kontrahenta: "+e.Message);
                return false;
            }
        }

        public  static void dodajKnt(string firma, KNT knt, TextWriter tw2, string connection_string, string connection_string2)
        {
            try
            {

                tw2.WriteLine("Dodawanie kontrahenta: " + firma+"_"+knt.akronim);
                int typ = 0;
                switch (knt.typ)
                {
                    case 8:
                        {
                            typ = 2;
                            break;
                        }
                    case 16:
                        {
                            typ = 1;
                            break;
                        }
                    case 24:
                        {
                            typ = 3;
                            break;
                        }
                    default:
                        {
                            tw2.WriteLine("Dla kontrahenta: " + knt.akronim + " jest błędny typ " + knt.typ);
                            typ = 8;
                            break;
                        }
                }

                int ceidg = 0;
                if (knt.regon != "")
                    ceidg = 1;
                string akronim = firma+"_" + knt.akronim;
                //if (knt.akronim.Substring(0, 5) != "EKSP_")
                //  akronim = "EKSP_" + knt.akronim;

                string numerKRSQuery = "Select Atr_Wartosc  from cdn.atrybuty join cdn.KntKarty on Knt_GIDNumer = Atr_ObiNumer AND Knt_GIDTyp = Atr_Obityp where Atr_AtkId = 55 " +
                                                    " AND Knt_GIDTyp = " + knt.gidTyp  + "  AND Knt_GIDNumer = " + knt.gidNumer;
                string numerKRS = KNT.ConnectToSQL(connection_string, numerKRSQuery);
                tw2.WriteLine("Numer KRS: " + numerKRS);

                if (knt.nipPrefiks.Contains("PL"))
                    knt.nipPrefiks = "";

                string kntInsert = "INSERT INTO [dbo].[tb_Contractors] ([Acronym] ,[CompanyID] ,[ContractorName]  ,[ContractorType] ,[FirstName] ,[SecondName] ,[LastName] ,[Country] ,[City] "
                                + " ,[Street] ,[BuildingNumber] ,[LocalNumber] ,[PostalCode] ,[BankName] ,[AccountNumber] ,[PhoneNumber1] ,[PhoneNumber2] ,[Email] ,[TIN] ,[PersonalIdentityNumber] "
                                + " ,[Post] ,[CorrespondenceCity]  ,[CorrespondenceStreet] ,[CorrespondenceBuildingNumber] ,[CorrespondenceLocalNumber]  ,[CorrespondencePostalCode]  ,[CorrespondencePost] "
                                + " ,[RecordVersion]  ,[RecordState] ,[CreatedBy] ,[CreatedDate] ,[ModifiedBy] ,[ModifiedDate] ,[ExtID] ,[AgreementSymbols] ,[KRS] ,[KRSCourtName] ,[CEIDG] ,[REGON] "
                                + " ,[WWW] ,[PaymentDays] ,[CurrencyCode]  ,[Segments] ,[DataProtection] ,[SplitPayment]) VALUES ('" + akronim + "' , '"+firma+"','" + knt.nazwa1.Replace("'","''") + " " + knt.nazwa2.Replace("'","''") + " " + knt.nazwa3.Replace("'","''") + "' ,"
                                + " " + typ + " , '','','', '" + knt.kraj + "' , '" + knt.miasto + "' , '" + knt.ulica + "' , '', '' ,'" + knt.kodP + "' , 'Rachunek' , '" + knt.nrRachunku + "' , '" + knt.telefon1 + "' , "
                                + " '" + knt.telefon2 + "' , '" + knt.eMail + "' , '" + knt.nipPrefiks + knt.nip + "' , '" + knt.pesel + "' , '', '', '', '', '', '', '', 1 , 0, 'ERP', '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "' ,"
                                + " 'ERP' ,'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "' , " + knt.gidNumer + " ,'', '"+numerKRS+"', '', " + ceidg + " ,'" + knt.regon + "' ,'" + knt.url + "'," + knt.termin + " ,'" + knt.waluta + "' , "
                                + " '' , 0 , " + knt.podzialPlat + ")";
                tw2.WriteLine(kntInsert);
                KNT.InsertInto(kntInsert, connection_string2);



                string idKntQuery = "SELECT TOP (1) ID FROM [dbo].[tb_Contractors] WHERE  Acronym ='" + akronim + "'"; //ExtID = " + knt.gidNumer + " and
                string idKnt = KNT.ConnectToSQL(connection_string2, idKntQuery);
                tw2.WriteLine("Id dodanego kontrahenta: " + idKnt);

                if (idKnt != "")
                {
                    if (firma == "EKI")
                    {
                        string atrUpdate = "UPDATE e SET [Atr_Wartosc] = 'TAK' FROM [CDN].[Atrybuty] as e   join cdn.AtrybutyKlasy on Atk_ID = Atr_AtkId WHERE AtK_Nazwa = 'Export WF' and Atr_ObiFirma =" + knt.gidFirma + " and Atr_ObiLp = " + knt.gidLp +
                                        " and Atr_ObiNumer =" + knt.gidNumer + " and Atr_ObiTyp = " + knt.gidTyp;
                        KNT.Update(atrUpdate, connection_string);

                        atrUpdate = "UPDATE e SET [Atr_Wartosc] = '" + idKnt + "' FROM [CDN].[Atrybuty] as e   join cdn.AtrybutyKlasy on Atk_ID = Atr_AtkId WHERE AtK_Nazwa = 'WF_KontrahentID' and Atr_ObiFirma =" + knt.gidFirma + " and Atr_ObiLp = " + knt.gidLp +
                                        " and Atr_ObiNumer =" + knt.gidNumer + " and Atr_ObiTyp = " + knt.gidTyp;
                        KNT.Update(atrUpdate, connection_string);
                    }
                    else
                    {
                        string atrUpdate = "UPDATE e SET [Atr_Wartosc] = '" + idKnt + "' FROM [CDN].[Atrybuty] as e   join cdn.AtrybutyKlasy on Atk_ID = Atr_AtkId WHERE AtK_Nazwa = 'WF_KontrahentID_"+firma+"' and Atr_ObiFirma =" + knt.gidFirma + " and Atr_ObiLp = " + knt.gidLp +
                                        " and Atr_ObiNumer =" + knt.gidNumer + " and Atr_ObiTyp = " + knt.gidTyp;
                        KNT.Update(atrUpdate, connection_string);
                    }


                    string bnkQuery = "SELECT Bnk_Nazwa, RkB_NrRachunku, Knt_Symbol, RkB_Domyslny FROM cdn.KntKarty "
                                    + " join cdn.RachunkiBankowe on RkB_ObiFirma = Knt_GIDFirma and RkB_ObiLp = Knt_GIDLp and RkB_ObiNumer = Knt_GIDNumer and RkB_ObiTyp = Knt_GIDTyp "
                                    + " join cdn.Banki on RkB_BnkFirma = Bnk_GIDFirma and RkB_BnkLp = Bnk_GIDLp and RkB_BnkNumer = Bnk_GIDNumer and RkB_BnkTyp = Bnk_GIDTyp "
                                    + " WHERE Knt_GIDTyp =" + knt.gidTyp + " and KNT_GIDFirma =" + knt.gidFirma + " and Knt_GIDNumer = " + knt.gidNumer + " and Knt_GIDLp = " + knt.gidLp;

                    List<Banki> banki = Banki.SelectListRowColumn(bnkQuery, "cdn.kntKarty", connection_string);

                    foreach (Banki bnk in banki)
                    {
                        try
                        {
                            tw2.WriteLine("Dodawanie nr rachunku: " + bnk.nrRachunku);
                            string bnkInsert = "INSERT INTO [dbo].[tb_ContractorAccounts] ([ContractorID] ,[AccountName] ,[AccountNumber] ,[CurrencyCode] ,[State] ,[IsMajor] ,[CreatedDate] ,[CreatedBy] ,[ModifiedDate] ,[ModifiedBy]) "
                                                + " VALUES (" + idKnt + " , '" + bnk.nazwa + "' , '" + bnk.nrRachunku + "' , '" + bnk.waluta + "' , 0 , " + bnk.domyslny + " , '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "' , 'ERP' , '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "' ,"
                                                + " 'ERP' )";

                            KNT.InsertInto(bnkInsert, connection_string2);
                        }
                        catch (Exception ex)
                        {
                            tw2.WriteLine("Błąd dodawania rachunku: " + bnk.nrRachunku + "\n" + ex.Message);
                        }



                    }
                    banki.Clear();


                }
                else
                {
                    tw2.WriteLine("Błąd dodawania kont bankowych: kontrahent " + knt.akronim + " nie został dodany");
                }
            }
            catch (Exception e)
            {
                tw2.WriteLine("Błąd dodawania kontrahenta: " + knt.akronim + "\n" + e.Message);
            }
        }
    }
}
