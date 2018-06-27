﻿using System;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace jpk_mag
{
    class Program
    {
        static void Main(string[] args)
        {
            string DP1, DK1;
            string DP2 = " 00:00:00";
            string DK2 = " 23:59:59";
            Console.WriteLine("");
            Console.WriteLine("**** PROGRAM DO GENEROWANIA PLIKU JPK DLA MAGAZYNU ****");
            Console.WriteLine("");
            Console.WriteLine("Podaj datę początkową w postaci (RRRR-MM-DD) np.: 2017-05-01 i naciśnij ENTER");
            DP1 = Console.ReadLine();
            Console.WriteLine("");
            Console.WriteLine("Podaj datę końcową w postaci (RRRR-MM-DD) np.: 2017-05-31 i naciśnij ENTER");
            DK1 = Console.ReadLine();
            Console.WriteLine("");
            Console.WriteLine("");
            string dataPoczatku = DP1 + DP2;
            string dataKonca = DK1 + DK2;
            string fullpath = @"C:\operacje.xls";
            //string fullpath = @"C:\ksapbefg.mdb";
            string xlsVersion = "Excel 8.0";

            DataSet ds = new DataSet();
            DataTable table = new DataTable();
            string strConn = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='{1};HDR=YES'", fullpath, xlsVersion);  //excel
            //string strConn = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Persist Security Info = False; ", fullpath);    //access
            using (OleDbConnection dbConnection = new OleDbConnection(strConn))
            {
                Console.WriteLine("Wczytywanie danych źródłowych z systemu magazynowego rozpoczęte");
                using (OleDbDataAdapter dbAdapter = new OleDbDataAdapter("SELECT * FROM [operacje] WHERE data_oper>=#" + dataPoczatku + "# AND data_oper<=#" + dataKonca + "#", dbConnection))
                    dbAdapter.Fill(table);
            }

            try
            {
                ds.Tables.Add(table);
                Console.WriteLine("Wczytywanie danych źródłowych z systemu magazynowego zakończone");
            }

            catch (Exception Ex)
            {
                Console.WriteLine(Ex.ToString());
            }

            finally
            {
                //Console.WriteLine("Wczytywanie danych źródłowych z systemu magazynowego zakończone");
            }

            string lokalizacjaPlikuXML = @"C:\TEMP\jpk_mag.xml";
            FileInfo InformacjaOPliku = new FileInfo("C:\\TEMP\\jpk_mag.xml");

            string XML_linia1 = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>";
            string XML_linia2 = "<tns:JPK xmlns:etd=\"http://crd.gov.pl/xml/schematy/dziedzinowe/mf/2016/01/25/eD/DefinicjeTypy/\" xmlns:kck=\"http://crd.gov.pl/xml/schematy/dziedzinowe/mf/2013/05/23/eD/KodyCECHKRAJOW/\" xmlns:tns=\"http://jpk.mf.gov.pl/wzor/2016/03/09/03093/\" > ";
            string XML_linia3 = "      <tns:Naglowek>";
            string XML_linia4 = "            <tns:KodFormularza kodSystemowy=\"JPK_MAG (1)\" wersjaSchemy =\"1-0\">JPK_MAG</tns:KodFormularza>";
            string XML_linia5 = "            <tns:WariantFormularza>1</tns:WariantFormularza>";
            string XML_linia6 = "            <tns:CelZlozenia>1</tns:CelZlozenia>";
            string XML_linia7 = "            <etd:DataWytworzeniaJPK>" + DateTime.Now.ToString().Substring(0, 10) + "T" + DateTime.Now.ToString().Substring(11) + "</etd:DataWytworzeniaJPK>";
            string XML_linia8 = "            <etd:DataOd>" + DP1 + "</etd:DataOd>";
            string XML_linia9 = "            <etd:DataDo>" + DK1 + "</etd:DataDo>";
            string XML_linia10 = "            <etd:DomyslnyKodWaluty>PLN</etd:DomyslnyKodWaluty>";
            string XML_linia11 = "            <etd:KodUrzedu>1449</etd:KodUrzedu>";
            string XML_linia12 = "      </tns:Naglowek>";
            string XML_linia13 = "      <Podmiot1>";
            string XML_linia14 = "            <IdentyfikatorPodmiotu>";
            string XML_linia15 = "                  <etd:NIP>5250006124</etd:NIP>";
            string XML_linia16 = "                  <etd:PelnaNazwa>KRAJOWA SZKOŁA ADMINISTRACJI PUBLICZNEJ im.Prezydenta Rzeczypospolitej Polskiej Lecha Kaczyńskiego</etd:PelnaNazwa>";
            string XML_linia17 = "                  <etd:REGON>006472421</etd:REGON>";
            string XML_linia18 = "            </IdentyfikatorPodmiotu>";
            string XML_linia19 = "            <AdresPodmiotu>";
            string XML_linia20 = "                  <tns:KodKraju>PL</tns:KodKraju>";
            string XML_linia21 = "                  <tns:Wojewodztwo>mazowieckie</tns:Wojewodztwo>";
            string XML_linia22 = "                  <tns:Powiat>WARSZAWSKI</tns:Powiat>";
            string XML_linia23 = "                  <tns:Gmina>CENTRUM</tns:Gmina>";
            string XML_linia24 = "                  <tns:Ulica>WAWELSKA</tns:Ulica>";
            string XML_linia25 = "                  <tns:NrDomu>56</tns:NrDomu>";
            string XML_linia26 = "                  <tns:Miejscowosc>WARSZAWA</tns:Miejscowosc>";
            string XML_linia27 = "                  <tns:KodPocztowy>00-922</tns:KodPocztowy>";
            string XML_linia28 = "                  <tns:Poczta>WARSZAWA</tns:Poczta>";
            string XML_linia29 = "            </AdresPodmiotu>";
            string XML_linia30 = "      </Podmiot1>";
            string XML_linia31 = "      <Magazyn>1</Magazyn>";

            if (File.Exists(lokalizacjaPlikuXML))
            {
                File.Delete(lokalizacjaPlikuXML);
            }

            try
            {
                Console.WriteLine("Generowanie nagłówka pliku JPK rozpoczęte");
                StreamWriter plikXML = new StreamWriter(@"C:\TEMP\jpk_mag.xml", true);

                plikXML.WriteLine(XML_linia1);
                plikXML.WriteLine(XML_linia2);
                plikXML.WriteLine(XML_linia3);
                plikXML.WriteLine(XML_linia4);
                plikXML.WriteLine(XML_linia5);
                plikXML.WriteLine(XML_linia6);
                plikXML.WriteLine(XML_linia7);
                plikXML.WriteLine(XML_linia8);
                plikXML.WriteLine(XML_linia9);
                plikXML.WriteLine(XML_linia10);
                plikXML.WriteLine(XML_linia11);
                plikXML.WriteLine(XML_linia12);
                plikXML.WriteLine(XML_linia13);
                plikXML.WriteLine(XML_linia14);
                plikXML.WriteLine(XML_linia15);
                plikXML.WriteLine(XML_linia16);
                plikXML.WriteLine(XML_linia17);
                plikXML.WriteLine(XML_linia18);
                plikXML.WriteLine(XML_linia19);
                plikXML.WriteLine(XML_linia20);
                plikXML.WriteLine(XML_linia21);
                plikXML.WriteLine(XML_linia22);
                plikXML.WriteLine(XML_linia23);
                plikXML.WriteLine(XML_linia24);
                plikXML.WriteLine(XML_linia25);
                plikXML.WriteLine(XML_linia26);
                plikXML.WriteLine(XML_linia27);
                plikXML.WriteLine(XML_linia28);
                plikXML.WriteLine(XML_linia29);
                plikXML.WriteLine(XML_linia30);
                plikXML.WriteLine(XML_linia31);
                plikXML.Close();

                Console.WriteLine("Generowanie nagłówka pliku JPK zakończone");
            }

            catch (Exception Ex)
            {
                Console.WriteLine(Ex.ToString());
            }

            finally
            {
                //Console.WriteLine("Generowanie nagłówka pliku JPK zakończone");
            }

            int liczbaRW_RWWartosc = 0;
            int liczabRW_RWWiersz = 0;
            int liczabRW_RWCtrl = 0;
            double sumaRW = 0;
            int liczbaPZ = 0;
            int liczbaMM = 0;
            int liczbaWZ = 0;
            int liczbaBledow = 0;
            

            Console.WriteLine("Generowanie danych pliku JPK rozpoczęte");
            StreamWriter sw = null;
            sw = new StreamWriter(lokalizacjaPlikuXML, true);
            DataTable dt = ds.Tables[0];

            foreach (DataRow row in dt.Rows)
            {
                object[] array = row.ItemArray;

                if (liczbaRW_RWWartosc == 0 && (array[2].ToString().Contains("RW") || array[2].ToString().Contains("Rw") || array[2].ToString().Contains("rW") || array[2].ToString().Contains("rw")))
                {                    
                    sw.Write("      <tns:RW>", FileMode.Append);
                    sw.WriteLine();
                    sw.Write("            <tns:RWWartosc>", FileMode.Append);
                    sw.WriteLine();
                    sw.Write("                  <tns:NumerRW>" + (array[2].ToString()).Substring(3) + "</tns:NumerRW>", FileMode.Append);
                    sw.WriteLine();
                    sw.Write("                  <tns:DataRW>" + (array[10].ToString()).Substring(0, 10) + "</tns:DataRW>", FileMode.Append);
                    sw.WriteLine();
                    sw.Write("                  <tns:WartoscRW>" + (array[6].ToString()).Replace(",", ".") + "</tns:WartoscRW>", FileMode.Append);
                    sw.WriteLine();
                    sw.Write("                  <tns:DataWydaniaRW>" + (array[10].ToString()).Substring(0, 10) + "</tns:DataWydaniaRW>", FileMode.Append);
                    sw.WriteLine();
                    sw.Write("            </tns:RWWartosc>", FileMode.Append);
                    sw.Flush();
                    liczbaRW_RWWartosc++;
                }

                else if (liczbaRW_RWWartosc > 0 && (array[2].ToString().Contains("RW") || array[2].ToString().Contains("Rw") || array[2].ToString().Contains("rW") || array[2].ToString().Contains("rw")))
                {                    
                    sw.Write("            <tns:RWWartosc>", FileMode.Append);
                    sw.WriteLine();
                    sw.Write("                  <tns:NumerRW>" + (array[2].ToString()).Substring(3) + "</tns:NumerRW>", FileMode.Append);
                    sw.WriteLine();
                    sw.Write("                  <tns:DataRW>" + (array[10].ToString()).Substring(0, 10) + "</tns:DataRW>", FileMode.Append);
                    sw.WriteLine();
                    sw.Write("                  <tns:WartoscRW>" + (array[6].ToString()).Replace(",", ".") + "</tns:WartoscRW>", FileMode.Append);
                    sw.WriteLine();
                    sw.Write("                  <tns:DataWydaniaRW>" + (array[10].ToString()).Substring(0, 10) + "</tns:DataWydaniaRW>", FileMode.Append);
                    sw.WriteLine();
                    sw.Write("            </tns:RWWartosc>", FileMode.Append);                    
                    sw.Flush();
                    liczbaRW_RWWartosc++;
                }
                sw.WriteLine();
            }

            foreach (DataRow row in dt.Rows)
            {
                object[] array = row.ItemArray;

                if (array[2].ToString().Contains("RW") || array[2].ToString().Contains("Rw") || array[2].ToString().Contains("rW") || array[2].ToString().Contains("rw"))
                {                    
                    sw.Write("            <tns:RWWiersz>", FileMode.Append);
                    sw.WriteLine();
                    sw.Write("                  <tns:Numer2RW>" + (array[2].ToString()).Substring(3) + "</tns:Numer2RW>", FileMode.Append);
                    sw.WriteLine();
                    sw.Write("                  <tns:KodTowaruRW>" + array[3].ToString() + "</tns:KodTowaruRW>", FileMode.Append);
                    sw.WriteLine();
                    sw.Write("                  <tns:NazwaTowaruRW>" + array[4].ToString() + "</tns:NazwaTowaruRW>", FileMode.Append);
                    sw.WriteLine();
                    sw.Write("                  <tns:IloscWydanaRW>" + array[7].ToString() + "</tns:IloscWydanaRW>", FileMode.Append);
                    sw.WriteLine();
                    if (array[5].ToString() == "")
                    {
                        sw.Write("                  <tns:JednostkaMiaryRW>" + "error" + "</tns:JednostkaMiaryRW>", FileMode.Append);
                        sw.WriteLine();
                    }
                    else if (array[5].ToString().Substring(array[5].ToString().Length - 1, 1).ToString() == "." || array[5].ToString().Substring(array[5].ToString().Length - 1, 1).ToString() == ",")
                    {
                        sw.Write("                  <tns:JednostkaMiaryRW>" + (array[5].ToString()).Substring(0, (array[5].ToString()).IndexOf(".")).ToUpper() + "</tns:JednostkaMiaryRW>", FileMode.Append);
                        sw.WriteLine();
                    }
                    else
                    {
                        sw.Write("                  <tns:JednostkaMiaryRW>" + (array[5].ToString()).ToUpper() + "</tns:JednostkaMiaryRW>", FileMode.Append);
                        sw.WriteLine();
                    }
                    sw.Write("                  <tns:CenaJednRW>" + (array[6].ToString()).Replace(",", ".") + "</tns:CenaJednRW>", FileMode.Append);
                    sw.WriteLine();
                    string WartoscPozycjiRW_bufor = "" + Double.Parse(array[7].ToString()) * Double.Parse(array[6].ToString()) + "";
                    sw.Write("                  <tns:WartoscPozycjiRW>" + WartoscPozycjiRW_bufor.Replace(",", ".") + "</tns:WartoscPozycjiRW>", FileMode.Append);
                    sw.WriteLine();
                    sw.Write("            </tns:RWWiersz>", FileMode.Append);
                    sw.Flush();
                    liczabRW_RWWiersz++;
                }
                sw.WriteLine();
            }
            string SumaRW_ciag = "";
            foreach (DataRow row in dt.Rows)
            {
                object[] array = row.ItemArray;
                
                if (array[2].ToString().Contains("RW") || array[2].ToString().Contains("Rw") || array[2].ToString().Contains("rW") || array[2].ToString().Contains("rw"))
                {
                    SumaRW_ciag = "" + Double.Parse(array[7].ToString()) * Convert.ToDouble(array[6].ToString()) + "";
                    sumaRW = sumaRW + (Double.Parse(array[7].ToString()) * Convert.ToDouble(array[6].ToString()));

                    liczbaRW_RWWartosc++;
                }

                

                //else if (liczabRW_RWCtrl > 0 && (array[2].ToString().Contains("RW") || array[2].ToString().Contains("Rw") || array[2].ToString().Contains("rW") || array[2].ToString().Contains("rw")))
                //{

                //    sw.Write("            <tns:RWCtrl>", FileMode.Append);
                //    sw.WriteLine();
                //    sw.Write("                  <tns:LiczbaRW>" + "1" + "</tns:LiczbaRW>", FileMode.Append);
                //    sw.WriteLine();
                //    string SumaRW_bufor = "" + Double.Parse(array[7].ToString()) * Convert.ToDouble(array[6].ToString()) + "";
                //    sw.Write("                  <tns:SumaRW>" + SumaRW_bufor.Replace(",", ".") + "</tns:SumaRW>", FileMode.Append);
                //    sw.WriteLine();
                //    sw.Write("            </tns:RWCtrl>", FileMode.Append);
                //    sw.Flush();
                //    liczbaRW_RWWartosc++;
                //}

            }
            sw.Write("            <tns:RWCtrl>", FileMode.Append);
            sw.WriteLine();
            sw.Write("                  <tns:LiczbaRW>" + liczabRW_RWWiersz + "</tns:LiczbaRW>", FileMode.Append);
            sw.WriteLine();
            string SumaRW_bufor = "" + sumaRW + "";
            sw.Write("                  <tns:SumaRW>" + SumaRW_bufor.Replace(",", ".") + "</tns:SumaRW>", FileMode.Append);
            sw.WriteLine();
            sw.Write("            </tns:RWCtrl>", FileMode.Append);
            sw.Flush();
            sw.WriteLine();

            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine("Przycisk ENTER kończy działanie programu...");
            Console.ReadLine();
        }
    }
}
