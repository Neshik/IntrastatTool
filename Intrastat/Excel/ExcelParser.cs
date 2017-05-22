using System;
using System.Windows;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.Serialization;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;

namespace Aqualight.ExcelParser
{
    static class ExcelParser
    {
        /// <summary>
        /// Holt die Bestellliste ohne Datei. Die Datei muss vorher aber schon einmal gecached worden sein.
        /// </summary>
        /// <returns></returns>
        public static List<BestellungElement> getExcelParser()
        {
            List<BestellungElement> list = new List<BestellungElement>();
            try
            {             
                IFormatter formatter = new BinaryFormatter();
                Stream stream = new FileStream("excel.bin", FileMode.Open, FileAccess.Read, FileShare.Read);
                list = (List<BestellungElement>)formatter.Deserialize(stream);
                return list;
            }
            catch(Exception e)
            {
                System.Console.WriteLine("Die Excelliste wurde noch nicht eingelesen." + e.ToString());
                return null;                   
            }
            
        }

        public static List<BestellungElement> getExcelParser(String excelFile)
        {
            List<BestellungElement> liste = new List<BestellungElement>();

            Excel.Application xlApp = new Excel.Application();
            try {
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(excelFile, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel._Worksheet worksheet = (Excel._Worksheet)xlWorkbook.Sheets.get_Item(2);

                Excel.Range excelRange = worksheet.UsedRange;

                //get an object array of all of the cells in the worksheet (their values)
                object[,] valueArray = (object[,])excelRange.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);


                for (int row = 1; row <= worksheet.UsedRange.Rows.Count; ++row)
                {
                    BestellungElement elem = new BestellungElement();
                    for (int col = 1; col <= worksheet.UsedRange.Columns.Count; ++col)
                    {
                        if (valueArray[row, col] != null)
                        {
                            switch (col){
                             
                                case 2: elem.Modelbezeichnung = valueArray[row, col].ToString(); break;
                                case 3: elem.Modelbezeichnung +=" "+(valueArray[row, col].ToString()); break;
                                case 4: elem.Modelbezeichnung += " " + valueArray[row, col].ToString(); break;
                                case 5: elem.Modelbezeichnung += " " + valueArray[row, col].ToString(); break;
                                case 6: elem.Modelbezeichnung += " " + valueArray[row, col].ToString(); break; 
                                case 7: elem.Rabatt = valueArray[row, col].ToString(); break;
                                case 8: try { elem.RabattStueckzahl = int.Parse(valueArray[row, col].ToString()); } catch { elem.RabattStueckzahl = 0; } break;
                                case 9: try { elem.Preis = Double.Parse(valueArray[row, col].ToString()); } catch { elem.Preis = 0.0; } break;
                                case 11: try { elem.PreisEndkunden = Double.Parse(valueArray[row, col].ToString()); } catch { elem.PreisEndkunden = 0.0; } break;
                                case 12: elem.Artikelnummer = valueArray[row, col].ToString(); break;
                                case 13: elem.EAN = valueArray[row, col].ToString(); break;
                                case 14: elem.Status = valueArray[row, col].ToString(); break;
                                case 15: try { elem.Gewicht = Double.Parse(valueArray[row, col].ToString()); } catch { elem.Gewicht = 0.0; } break;
                                case 16: elem.Verpackungsgroesse = valueArray[row, col].ToString(); break;
                                case 17: elem.Zollnummer = valueArray[row, col].ToString(); break;
                                case 18: elem.Infos = valueArray[row, col].ToString(); break;
                                case 19: elem.Infos += " " + valueArray[row, col].ToString(); break;
                            }
                        }                               
                    }
                    if (elem.Artikelnummer != null)
                    {
                        liste.Add(elem);
                        Console.WriteLine(elem.ToString());
                    }
                }         
            }
            catch (Exception e)
            {
                MessageBoxButton button = MessageBoxButton.OK;
                MessageBox.Show("Excel Datei nicht gefunden.\n"+e, excelFile, button);
            }
            return liste;


        }
        //
        private static int parseColumnRange(String range)
        {
            char firstLetter = range[0];
            return Char.ToUpper(firstLetter) - 64;

        }

        private static int parseRowRange(String range)
        {
            String endLetter = null;
            try {
                endLetter = range.Substring(1);
            }catch(ArgumentOutOfRangeException arg)
            {
                Console.WriteLine(arg.ToString() + " - Fehler beim Parsing der range in parseRowRange");
                return 0;
            }
            try
            {
                return int.Parse(endLetter);
            }
            catch(Exception e)
            {
                Console.WriteLine(e.ToString()+" - Fehler beim Parsing der Nummer in parseRowRange");
                return 0;
            }
        }

    }
}
