
using CDNBase;
using System;
using ADODB;
using System.Collections.Generic;
using static ImportDekretuExcel.CSV;
//using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using ExcelDataReader;

namespace ImportDekretuExcel
{
    public partial class Form1 : Form
    {
        // public static string sKnt_Kod = "OSTREWKA";
       /* public static string Operator = "ADMIN";
        public static string Haslo = "Platforma2020!";
        public static string Firma = "PPBW_PPBW";
        public static string Baza = "[CDN_PPBW_PPBW].[CDN].";*/

         public static string sKnt_Kod = "!NIEOKREŚLONY!";
         public static string Operator = "ADMIN";
        public static string Haslo = "NE!w0rk12";
         public static string Firma = "Pomet";
       //  public static string Baza = "[CDN_PREZENTACJA_KH_1].[CDN].";
       protected static CDNBase.Application Application = null;
       protected static ILogin Login = null;
        public Form1()
        {

            InitializeComponent();
        }

        static protected void LogowanieAutomatyczne()
        {
            object[] hPar = new object[] {
                         0,  0,   1,  0,  0,   0,  0,    0,   0,   0,   0,   0,   0,   0,  0,  1, 0 };    // do jakich modułów się logujemy
            /* Kolejno: KP, KH, KHP, ST, FA, MAG, PK, PKXL, CRM, ANL, DET, BIU, SRW, ODB, KB, KBP, HAP*/
            // katalog, gdzie jest zainstalowana Optima (bez ustawienia tej zmiennej nie zadziała, chyba że program odpalimy z katalogu O!)
            System.Environment.CurrentDirectory = @"C:\Program Files (x86)\Comarch ERP Optima";
            // tworzymy nowy obiekt apliakcji
            Application = new CDNBase.Application();
            //Application.HASPServerCommandLine = "41941-FE3B2-002::5000139271";
            // blokujemy
            Application.LockApp(1);//(513, 5000, null, null, null, null);
            // logujemy się do podanej Firmy, na danego operatora, do podanych modułów
            Login = Application.Login(vUser: Operator, vChk: Haslo, vFirm: Firma, _ModulKP: hPar[0], _ModulKH: hPar[1], _ModulKHP: hPar[2], _ModulST: hPar[3], _ModulFA: hPar[4], _ModulMAG: hPar[5], _ModulPK: hPar[6], _ModulPKXL: hPar[7], _ModulCRM: hPar[8], _ModulANL: hPar[9], _ModulDET: hPar[10], _ModulBIU: hPar[11], _ModulSRW: hPar[12], _ModulOBD: hPar[13], _ModulKB: hPar[14], _ModulKBP: hPar[15], _ModulHAP: hPar[16]);

        }

        protected static void Wylogowanie()
        {
            // niszczymy Login
            Login = null;
            // odblokowanie (wylogowanie) O!
            Application.UnlockApp();
            // niszczymy obiekt Aplikacji
            Application = null;
        }


        void WyjscieZApki(string a)
        {
            Cursor.Current = Cursors.Default;
            Wylogowanie();
            if (a == "done")
            {
                MessageBox.Show("Zakończono przetwarzanie pliku.", "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Generowanie dokumentów zakończyło się niepowodzeniem.", "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            return;
        }


        private void Button1_Click(object sender, EventArgs e)
        {


            #region excel5
            Cursor.Current = Cursors.WaitCursor;
            if (tB1.Text == "nie wybrano pliku")
            {
                MessageBox.Show("Nie wybrano pliku arkusza.", "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string filePath = tB1.Text;
            DataTable dataTable;
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader reader;
                reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);

                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                };

                var dataSet = reader.AsDataSet(conf);
                dataTable = dataSet.Tables[0];
            }

            //Przypisywanie danych z datasetu do listy przyszłych pozycji.
            //List<ONE032Data> listaPozycjiPz = new List<ONE032Data>();
            if (dataTable.Rows.Count == 0)
            {
                MessageBox.Show("Pusty arkusz", "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Error);//błąd
                return;
            }





            #endregion

            #region dzialanie na optimie
            #region logowanie
            //Logowanie do O! bez wyświetlania okienka logowania
            Cursor.Current = Cursors.WaitCursor;

            //JakaBaza("klient");
            if (String.IsNullOrEmpty(numerDok.Text.ToString()))
            {
                MessageBox.Show("Proszę uzupełnić pole 'Numer dok'", "Komunikat",
                                                   MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {

                try
                {
                    LogowanieAutomatyczne();
                }
                catch (COMException eo)
                {
                    if (eo.InnerException != null)
                        MessageBox.Show("Błąd logowania: " + eo.InnerException.Message, "Komunikat",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else
                        MessageBox.Show("Błąd logowania: " + eo.Message, "Komunikat",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (Exception ei)
                {
                    if (ei.InnerException != null)
                        MessageBox.Show("Błąd logowania: " + ei.InnerException.Message, "Komunikat",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else
                        MessageBox.Show("Błąd logowania: " + ei.Message, "Komunikat",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                //MessageBox.Show("Zalogowano ", "Komunikat",
                // MessageBoxButtons.OK, MessageBoxIcon.Error);
                #endregion



                CDNBase.AdoSession Sesja = Login.CreateSession();
                #endregion

                dodawanieDekretu(dataTable, Sesja);



                //zapisujemy
                //Sesja.Save();

                WyjscieZApki("done");
            }
            //ActiveForm.Dispose();
        }

        
        void WybraniePliku()
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog()
            {
                InitialDirectory = Directory.GetCurrentDirectory(),
                DefaultExt = "xlsx",
                Filter = "Pliki xlsx (*.xlsx)|*.xlsx|Pliki xls (*.xls)|*.xls|Wszystkie pliki (*.*)|*.*",
            };
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                tB1.Text = openFileDialog1.FileName;
                //MessageBox.Show(path);
            }
            return;
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            WybraniePliku();
        }


        private void dodawanieDekretu(DataTable dataTable, AdoSession sesja)
        {

            try
            {
                
                    //MessageBox.Show(records[0].Opis);
                    var rOtwartyOkres = sesja.CreateObject("CDN.Okresy").Item("OOb_Symbol='" + Convert.ToDateTime(dataTable.Rows[dataTable.Rows.Count - 1][1]).Year + "'");

                    int ID_OkresuObr = rOtwartyOkres.ID;

                    int k = 0;
                    pBar1.Visible = true;
                    // Set Minimum to 1 to represent the first file being copied.
                    pBar1.Minimum = 1;
                    // Set Maximum to the total number of files to copy.
                    pBar1.Maximum = dataTable.Rows.Count + 1;
                    // Set the initial value of the ProgressBar.
                    pBar1.Value = 1;
                    // Set the Step property to a value of 1 to represent each file being copied.
                    pBar1.Step = 1;


                    try
                    {

                    CDNBase.AdoSession Sesja = Login.CreateSession();
                    object dummy = Type.Missing;
                    for (int i=0;i<dataTable.Rows.Count; i++)
                        {
                        if(dataTable.Rows[i][2].ToString()!="")
                                {
                        string czyIstniejeQuery = "if exists (SELECT * FROM cdn.Konta  WHERE ACC_Numer= '" + dataTable.Rows[i][2].ToString() + "') SELECT 'TAK' else SELECT 'NIE'";


                        if (Sesja.Connection.Execute(czyIstniejeQuery, out dummy, 0).GetString().Trim() == "NIE")
                        {
                            if (dataTable.Rows[i][2].ToString().Contains("208"))
                            {
                                czyIstniejeQuery = "if exists (SELECT * FROM cdn.Konta  WHERE ACC_Numer= '" + dataTable.Rows[i][2].ToString().Substring(0, dataTable.Rows[i][2].ToString().IndexOf("-") + 3) + "') SELECT 'TAK' else SELECT 'NIE'";
                               // MessageBox.Show(czyIstniejeQuery);

                                if (Sesja.Connection.Execute(czyIstniejeQuery, out dummy, 0).GetString().Trim() == "TAK")
                                {
                                    string kntKod = dataTable.Rows[i][2].ToString().Substring(dataTable.Rows[i][2].ToString().IndexOf("-") + 4);
                                 //   MessageBox.Show("KntKod: " + kntKod);
                                    string kntNazwaQuery = " if exists(SELECT * FROM cdn.Kontrahenci WHERE Knt_Kod = '" + kntKod + "') SELECT Knt_Nazwa1 FROM cdn.Kontrahenci WHERE Knt_Kod = '" + kntKod + "' else SELECT 'NIE'";
                                    string kntNazwa = Sesja.Connection.Execute(kntNazwaQuery, out dummy, 0).GetString().Trim();
                                    if (kntNazwa != "NIE")
                                    {
                                        string kntIDQuery = "  SELECT Knt_KntID FROM cdn.Kontrahenci WHERE Knt_Kod = '" + kntKod + "' ";
                                        string kntID = Sesja.Connection.Execute(kntIDQuery, out dummy, 0).GetString().Trim();
                                        ADDKonto(dataTable.Rows[i][2].ToString().Substring(0, dataTable.Rows[i][2].ToString().IndexOf("-") + 3).Replace("-", "."), kntKod, kntNazwa, 0, 1, Convert.ToInt32(kntID), Sesja, ID_OkresuObr);
                                    }
                                    else
                                    {
                                        MessageBox.Show("Brak kontrahenta: " + kntNazwa + " oraz konta " + dataTable.Rows[i][2].ToString() + ". Zakończono działanie dodatku", "Komunikat",
                                                             MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        pBar1.Visible = false;
                                        return;

                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Proszę dodać konto " + dataTable.Rows[i][2].ToString().Substring(0, dataTable.Rows[i][2].ToString().IndexOf("-") + 3) + ". Zakończono działanie dodatku", "Komunikat",
                                                              MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    pBar1.Visible = false;
                                    return;
                                }

                            }
                            else
                            {

                                MessageBox.Show("Proszę dodać konto " + dataTable.Rows[i][2].ToString() + ". Zakończono działanie dodatku", "Komunikat",
                                                              MessageBoxButtons.OK, MessageBoxIcon.Error);
                                pBar1.Visible = false;
                                return;
                            }
                        }
                    }


                        if (dataTable.Rows[i][3].ToString() != "")
                        {
                           string  czyIstniejeQuery = "if exists (SELECT * FROM cdn.Konta  WHERE ACC_Numer= '" + dataTable.Rows[i][3].ToString() + "') SELECT 'TAK' else SELECT 'NIE'";


                            if (Sesja.Connection.Execute(czyIstniejeQuery, out dummy, 0).GetString().Trim() == "NIE")
                            {
                                MessageBox.Show("Proszę dodać konto " + dataTable.Rows[i][3].ToString() + ". Zakończono działanie dodatku", "Komunikat",
                                                              MessageBoxButtons.OK, MessageBoxIcon.Error);
                                pBar1.Visible = false;
                                return;
                            }
                        }
                        }



                        
                        var rDekret = Sesja.CreateObject("CDN.DEKRETY").AddNew(null);

                        // Id dziennika 
                        rDekret.OObID = ID_OkresuObr;

                        string IdDziennik_Query = "SELECT Dzi_DziId FROM cdn.Dzienniki WHERE Dzi_OObId = " + ID_OkresuObr + " and Dzi_Symbol = 'PK KANTYNY'";

                        var wynikQuery = Sesja.Connection.Execute(IdDziennik_Query, out dummy, 0);

                        int ID_Dziennik = 0;

                        if (!wynikQuery.EOF)
                            ID_Dziennik = Convert.ToInt32(wynikQuery.GetString());
                        else
                        {
                            MessageBox.Show("Brak dziennika inne dla okresu obrachunkowego z pliku", "Komunikat",
                                                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }


                        rDekret.DziID = ID_Dziennik;

                        rDekret.Bufor = 1;
                        rDekret.DataDok = Convert.ToDateTime(dataTable.Rows[dataTable.Rows.Count - 1][1]);
                        rDekret.Dokument = numerDok.Text.ToString(); //"NowyDok";//"Rozliczenie projektów (konta pozabilansowe) " + dataDok;//rDekret.DataDok.ToString("yyyy-MM-dd");
                        //rDekret.Opis = "Opis";//"Rozliczenie projektów (konta pozabilansowe) " + dataDok;//rDekret.DataDok.ToString("yyyy-MM-dd");






                        var rPozycje = rDekret.Elementy;

                    for (int i=0; i < dataTable.Rows.Count; i++)
                        {
                            pBar1.PerformStep();
                            try
                            {

                                var rPozycja = rPozycje.AddNew();
                            if (dataTable.Rows[i][2].ToString() != "")
                            {
                                var rKontaWN = Sesja.CreateObject("CDN.Konta").Item("ACC_OOBid =" + ID_OkresuObr + " and  ACC_Numer='" + dataTable.Rows[i][2].ToString() + "'");

                                rPozycja.KontoWN = rKontaWN;
                            }


                            if (dataTable.Rows[i][3].ToString() != "")
                            {

                                var rKontaMA = Sesja.CreateObject("CDN.Konta").Item("ACC_OOBid =" + ID_OkresuObr + " and ACC_Numer='" + dataTable.Rows[i][3].ToString() + "'");

                                rPozycja.KontoMA = rKontaMA;
                            }


                                rPozycja.Kwota = dataTable.Rows[i][4].ToString()!="" ? Math.Round(Convert.ToDouble(dataTable.Rows[i][4].ToString()), 2, MidpointRounding.AwayFromZero) : Math.Round(Convert.ToDouble(dataTable.Rows[i][4].ToString()), 2, MidpointRounding.AwayFromZero);
                            


                                rPozycja.Opis = dataTable.Rows[i][6].ToString();//dataTable.Rows[i][7].ToString() + " / " + dataTable.Rows[i][10].ToString() + " / " + dataTable.Rows[i][11].ToString();
                            if (dataTable.Rows[i][7].ToString() != "")
                            {

                                string katQuery = "SELECT [Kat_KatID] FROM [CDN].[Kategorie] WHERE Kat_KodSzczegol = '" + dataTable.Rows[i][7] + "'";
                                wynikQuery = Sesja.Connection.Execute(katQuery, out dummy, 0);

                                

                                if (!wynikQuery.EOF)
                                    rPozycja.KatID = Convert.ToInt32(wynikQuery.GetString());
                            }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Błąd dla pozycji " + dataTable.Rows[i][6].ToString() + " : " + ex.Message, "Komunikat",
                                                             MessageBoxButtons.OK, MessageBoxIcon.Error);
                                // return;
                            }


                        }
                        pBar1.Visible = false;
                        Sesja.Save();
                        Sesja.Dispose();

                    }

                    catch (Exception es)
                    {
                    if (es.ToString().Contains("niezbilansowanego"))
                    {
                        MessageBox.Show("Błąd w dodawaniu dekretu: Suma kwot Winien nie jest równa sumie kwot ma", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                        MessageBox.Show("Błąd w dodawaniu dekretu  : " + Convert.ToString(es), "Komunikat",
                                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                
               
            }
            catch (Exception e)
            {
                 if(e.ToString().Contains("niezbilansowanego"))
                 {
                     MessageBox.Show("Błąd w dodawaniu dekretu: Suma kwot Winien nie jest równa sumie kwot ma", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Error);
                 }
                 else
                    MessageBox.Show("Błąd w dodawaniu dekretu: " + Convert.ToString(e), "Komunikat",
                                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void ADDKonto(string Ojciec, string Syn, string Nazwa, int Slownikowe, int STyp, int SlowID, AdoSession Sesja, int idObr)
        {

            // Ojciec - Numer konkta nadrzendnego.
            // Syn    - Numer dodawanego konata.
            // Nazwa  - Nazwa konta dodawanego.
            // Slownikowe - czy slownikowe 
            // STyp   - Typ slownika 
            // SlowID - ID w slowniku.

            var rKonto = Sesja.CreateObject("CDN.Konta").AddNew();

            var rKontoOj = Sesja.CreateObject("CDN.Konta").Item("ACC_NumerIDx ='" + Ojciec + "'");

            //WScript.Echo(rKontoOj.Id)
            //WScript.Echo(rKontoOj.CzyMaKsiegowania)
            //WScript.Echo()

           
                rKonto.ParId = rKontoOj.Id;

                rKonto.Nazwa = Nazwa;
                rKonto.Segment = Syn;

                rKonto.Numer = Syn;

                rKonto.NumerIdx = Syn;

                rKonto.OkresId = idObr;    	// UWAGA! Podaj prawidłowy ID okresu obrachunkowego //
                
                    rKonto.Slownik = Slownikowe;

                    rKonto.SlownikTyp = STyp;

                    rKonto.SlownikId = SlowID;

                

                //WScript.Echo(rKonto.NumerIdx)
                //WScript.Echo(rKonto.Segment)
                //WScript.Echo(rKonto.Poziom)
                //rKonto.ParId = rKontoOj.Id
                //        rKonto.ParentACC

                Sesja.Save();

    

            //try{
            //} catch( e ) { f.WriteLine("Raport: " + e.descripton ) }

        }
    }
}
