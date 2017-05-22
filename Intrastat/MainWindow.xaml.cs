using Aqualight;
using Aqualight.Artikelstamm;
using Aqualight.ExcelParser;
using Aqualight.Intrastat;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace AquaLight{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        bool intrastatSemaphor = false;
        HashSet<Artikel> hash;
     

        public MainWindow()
        {
           

        }

       
        private void ParseExcelBtn_Click(object sender, RoutedEventArgs e)
        {
            if (!intrastatSemaphor)
            {
                intrastatSemaphor = true;
                String file = @Exceldatei.Text;

                outputBox.AppendText("Lese Excel Datei: ");
                List<BestellungElement> list = ExcelParser.getExcelParser(file);
                IFormatter formatter = new BinaryFormatter();
                Stream stream = new FileStream("excel.bin",
                                         FileMode.Create,
                                         FileAccess.Write, FileShare.None);
                formatter.Serialize(stream, list);
                stream.Close();
                

                intrastatSemaphor = false;
            }
        }

        private void IntrastatBtn_Click(object sender, RoutedEventArgs e)
        {
            
            if (!intrastatSemaphor)
            {
                List<BestellungElement> list = new List<BestellungElement>();
                intrastatSemaphor = true;
                try
                {
                    outputBox.AppendText("Lese Excelliste aus dem Cache");
                    IFormatter formatter = new BinaryFormatter();
                    Stream stream = new FileStream("excel.bin", FileMode.Open, FileAccess.Read, FileShare.Read);                   
                    list = (List<BestellungElement>)formatter.Deserialize(stream);
                    outputBox.AppendText("erfolgreich.\n");
                }
                catch(Exception except)
                {                    
                    outputBox.AppendText("Fehlgeschlagen. Exceldatei muss erst eingelesen werden.");
                    Console.WriteLine(except.ToString());                   
                }

                if (list.Count > 0)
                {
                    
                    outputBox.AppendText("Durchsuche Datenbank nach Aufträgen: ");
                    AuftragspositionenListe auftragspositionenliste = new AuftragspositionenListe();
                    AquaLight.IntrastatBericht.IntrastatBericht intratstat = new AquaLight.IntrastatBericht.IntrastatBericht(((int)IntrastatYear.SelectedItem).ToString(), (String)IntrastatMonth.SelectedItem, list,odbc.Text, Ausgangspfad.Text, auftragspositionenliste);
                    if (intratstat != null)
                    {
                        outputBox.AppendText("erfolgreich.");
                        outputBox.AppendText("Die Datei für den Monat wurde erstellt und befindet sich im Ordner Intrastatbericht.");
                    }
                    else
                    {
                        outputBox.AppendText("Die Erstellung des Intrastatberichts ist fehlgeschlagen. Entweder die ODBC Datenquelle ist nicht richtig oder der Ordner ist nicht schreibbar für das Programm.");
                    }
                }
                else { outputBox.AppendText("fehlgeschlagen. Die Exceldatei konnte nicht gelesen werden. Entweder der Pfad ist falsch oder die Datei ist deformiert. \n"); }


                intrastatSemaphor = false;
            }

        }

        private void TabItem_Loaded(object sender, RoutedEventArgs e)
        {
            var months = System.Globalization.DateTimeFormatInfo.InvariantInfo.MonthNames;
            IntrastatMonth.ItemsSource = months;
            IntrastatMonth.SelectedItem = months[0];
            for (int i = 0; i < 5; i++) {
                IntrastatYear.Items.Add(DateTime.Now.Year - i);
            }
            IntrastatYear.SelectedIndex = 0;
            
     
        }

        private void ClearOldExcelData_Btn_Click(object sender, RoutedEventArgs e)
        {
            

        }

        private void ListBoxItem_Selected(object sender, RoutedEventArgs e)
        {

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            List<String> list = new List<string>();
            Artikelstamm artikelstamm = new Artikelstamm();
            hash = artikelstamm.Hash;
            Artikelliste.ItemsSource = list;
            
            foreach(Artikel artikel in hash)
            {
                list.Add(artikel.Bezeichnung);
            }
            


        }
    }
}
