using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Windows;
using EindOpdrachtLeandro.dao;
using EindOpdrachtLeandro.model;
using EindOpdrachtLeandro.Repositories;
using Microsoft.Office.Interop.Word;
using Section = Microsoft.Office.Interop.Word.Section;
using Window = System.Windows.Window;
using Word = Microsoft.Office.Interop.Word;

namespace EindOpdrachtLeandro
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();


        }

        SqlConnection connection;
        Iadresgegeven adresgegeven;
        Ibenodigheden benodigheden;
        Ihobby hobby;
        Ipersoon persoon;
        Isport sport;
        Isportclub sportclub;
        List<string> Uitlijning = new List<string>();
        Word.Application wrdApp;
        IList<Adresgegeven> adreslijst = new List<Adresgegeven>();
        IList<Adresgegeven> adreslijstkeuze = new List<Adresgegeven>();
        IList<Benodigheden> benodighedenlijst = new List<Benodigheden>();
        IList<Benodigheden> benodighedenlijstkeuze = new List<Benodigheden>();
        IList<Persoon> persoonlijst = new List<Persoon>();
        IList<Persoon> persoonlijstkeuze = new List<Persoon>();
        IList<Sport> sportlijst = new List<Sport>();
        IList<Sport> sportlijstkeuze = new List<Sport>();
        IList<Sportclub> sportclublijst = new List<Sportclub>();
        IList<Sportclub> sportclubKeuze = new List<Sportclub>();
        IList<Hobby> Hobbylijst = new List<Hobby>();
        IList<Hobby> Hobbylijstkeuze = new List<Hobby>();
        IEnumerable<string> sportclubnames = new List<string>();
        double numericValue;
        String FontValue;


        private void Window_Loaded(object sender, RoutedEventArgs e)
        { //Connectie aanmaken 
            connection = Singleton.Instance.GetDBConnection(); ;
            DataFromDB();
            InitGui();

            try //Indien de file niet aanwezig is moet je deze niet verwijderen.
            {
                File.Delete(Directory.GetParent(Environment.CurrentDirectory).Parent.FullName + "/Template1Bewerkt.docx");
            }
            catch (IOException ie)
            {
                Console.WriteLine("Dit bestand bestond niet");
                Console.WriteLine(ie.Message);
            }
            //Hier wordt een copy van de originele word file gemaakt zodat deze nooit veranderd.
            File.Copy(Directory.GetParent(Environment.CurrentDirectory).Parent.FullName + "/Template1Original.docx", Directory.GetParent(Environment.CurrentDirectory).Parent.FullName + "/Template1Bewerkt.docx");

        }
        public void DataFromDB()
        {
            //Alle gegevens inladen in de objecten van database
            adresgegeven = new daoAdresgegeven(connection);
            benodigheden = new daoBenodigheden(connection);
            persoon = new daoPersoon(connection);
            sport = new daoSport(connection);
            sportclub = new daoSportclub(connection);
            hobby = new daoHobby(connection);
            adreslijst = adresgegeven.GetAll();
            benodighedenlijst = benodigheden.GetAll();
            persoonlijst = persoon.GetAll();
            sportlijst = sport.GetAll();
            sportclublijst = sportclub.GetAll();
            Hobbylijst = hobby.GetAll();


        }
        public void InitGui()
        {
            //Lijst voor comboox declareren
            Uitlijning.Add("Links");
            Uitlijning.Add("Centreren");
            Uitlijning.Add("Rechts");

            //comboboxen vullen
            cmboAllignText.ItemsSource = Uitlijning;
            CmboAlignLogo.ItemsSource = Uitlijning;
            CmboAllignAanspreking.ItemsSource = Uitlijning;
            cmboAllignCompanyData.ItemsSource = Uitlijning;
            cmboAllignSlot.ItemsSource = Uitlijning;
            sportclubnames = sportclub.GetSportclubNaam().Distinct(); //Distinct gebruik ik om bijvoorbeeld Multiclub die er 2x inzit (2 verschillende sporten) deze maar 1x te tonen
            CmboSportclub.ItemsSource = sportclubnames;
        }


        private void Sportclub_Changed(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            //Hier ledig ik alle lijsten, indien ze switchen van club tijdens het runnen.
            sportclubKeuze.Clear();
            Hobbylijstkeuze.Clear();
            persoonlijstkeuze.Clear();
            benodighedenlijstkeuze.Clear();
            adreslijstkeuze.Clear();
            sportlijstkeuze.Clear();

            //Hier haalt methode de gewenste sportclub(s) uit
            Sportclubselecteren();

            //Hier haalt methode de Hobbies uit de lijst van sportclub
            HobbiesSelecteren(Sportclubselecteren());

            //Hier haalt methode de personen uit de lijst van de sportclub
            PersonenPersportclub();

            //Hier haalt methode  het adres(sen) uit de lijst van sportclub
            AdressenSelecteren(Sportclubselecteren());

            //Hier haalt methode de correcte sporten uit de lijst van sportclub
            SportenSelecteren(Sportclubselecteren());

            // Hier haalt methode de benodigdheden uit per sport.--> Hier moet ik nog een oplossing vinden...
            BenodigdhedenSelecteren(Sportclubselecteren());

        }

        private void Button_WordAanmaken(object sender, RoutedEventArgs e)
        {  //Controle of de persoon wel een sportclub heeft aangeduid.
            if (sportclubKeuze.Count == 0) { MessageBox.Show("Je moet een keuze maken voor welke sportclub je een mail wil sturen"); }
            else
            { //Nieuwe app en docu aanmaken + verwijzen welk document we willen toevoegen
                var app = new Word.Application();
                var doc = new Document();
                doc = app.Documents.Add(Directory.GetParent(Environment.CurrentDirectory).Parent.FullName + "/Template1Bewerkt.docx");

                //Functie die een generic tabel creert volgens de benodigheden van de sportclub -> Param is een Doc die je meegeeft zodat functie weet in welke doc je werkt.
                Tabelaanmaken(doc);
                //Opstarten van de Wordapplicatie
                app.Visible = true;
                //Hier maak ik een copy van de "Template" voor ik deze invul van gegevens.
                doc.Range(doc.Content.Start, doc.Content.End).Copy();

                //Hier is er een loop aan voor het aantal personen zodat de Template dit aantal keer gekopieerd wordt.
                for (int i = 0; i < persoonlijstkeuze.Count; i++)
                { // Hier is een loop gemaakt indien je meer dan 1 sportclub in lijst hebt (zoals 1club met meerdere sporten)
                    for (int k = 0; k < sportclubKeuze.Count; k++)
                    {
                        //Hier worden Alle fields in mijn document 1 per 1 doorlopen
                        foreach (Microsoft.Office.Interop.Word.Field Mergefields in doc.Fields)
                        {
                            //Bij logo moet er iets anders gebeuren dan bij al derest, daarom zit deze in de else.
                            if (Mergefields.Result.Text != "«Logo_sportclub»")
                            {
                                Mergefields.Select();
                                app.Selection.TypeText(MailmergeCase(Mergefields, k, i)); //Hier wordt gebruik gemaakt van een methode die de correcte field opvult.
                            }
                            else
                            { //Toevoegen van Logo waar het field = «Logo_sportclub»
                                Mergefields.Select();
                                //Hard gecodeerde path/deeltje uit database. 
                                var foto = app.Selection.InlineShapes.AddPicture(Directory.GetParent(Environment.CurrentDirectory).Parent.FullName +
                                    "/Logo/" + sportclubKeuze[0].Logo_sportclub + ".jpg");

                                //Hoogte en breedte zetten zodat iedere logo even groot is.
                                foto.Height = 60;
                                foto.Width = 120;
                            }
                        }
                    }
                    //Hier wordt gecontroleerd hoeveel pagina's nog moeten geplakt worden -1 Wordt geplaatst omdat je anders altijd 1 teveel plakt.
                    if (i != persoonlijstkeuze.Count - 1)
                    {
                        doc.Range(doc.Content.End - 1).Paste();
                    }
                }
                //Dit wordt gedaan omdat in  taakbeheer : details Word bleef runnen.
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            }

        }

        private void Button_Template(object sender, RoutedEventArgs e)
        {
            //Word App en Doc instellen +openen
            wrdApp = new Word.Application();
            wrdApp.Visible = true;
            wrdApp.Documents.Open(Directory.GetParent(Environment.CurrentDirectory).Parent.FullName + "/Template1Bewerkt.docx");

            //Dit wordt gedaan omdat in mijn taakbeheer mijn details Word bleef runnen.
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wrdApp);

        }
        private void WelcomeText_Click(object sender, RoutedEventArgs e)
        {
            //Word app + document aanmaken + instellen in welk document gewerkt word.
            var app = new Microsoft.Office.Interop.Word.Application();
            var doc = new Microsoft.Office.Interop.Word.Document();
            doc = app.Documents.Add(Directory.GetParent(Environment.CurrentDirectory).Parent.FullName + "/Template1Bewerkt.docx");
            //Hard gecodeerde Sectie van template aanspreken.
            Section s = doc.Sections[4];
            //Volledige "Afstand" van sectie bepalen.
            Range r = doc.Sections[4].Range;

            //Controleren of gebruiker zelf iets ingevuld heeft als text
            if (!TxtbxWelkomText.Text.Equals(""))
            {
                r.Text = TxtbxWelkomText.Text.Trim() + "\n\r";
            }
            else
            { //Voor geprogrammeerde text voor alle sportclubs.
                r.Text = "Welkom beste leden bij het begin van het nieuwe jaar, we zijn blij om u nog steeds als lid te zien." +
                    " We hebben natuurlijk nieuwe spullen nodig om het nieuw sportjaar goed in te zetten, dit vind u terug in de tabel hieronder.\n\r";
            }
            //Maken dat de range van de ingetypte woorden niet afkapt. dus zelfs groote instellen
            r.SetRange(0, r.Text.Length);
            //toevoegen als paragraaf
            s.Range.Paragraphs.Add(r);
            //opslaan &sluiten
            doc.SaveAs2(Directory.GetParent(Environment.CurrentDirectory).Parent.FullName + "/Template1Bewerkt.docx");
            doc.Close();

            //Dit om bepaalde achtergrond taken te stoppen.
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
        }
        private void Changed_Lettertype(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            var app = new Microsoft.Office.Interop.Word.Application();
            var doc = new Microsoft.Office.Interop.Word.Document();
            doc = app.Documents.Add(Directory.GetParent(Environment.CurrentDirectory).Parent.FullName + "/Template1Bewerkt.docx");
            //Colledige content Font aanpassen naar geselecteerde waarde.
            doc.Content.Font.Name = CmboFontFamily.SelectedItem.ToString();
            doc.SaveAs2(Directory.GetParent(Environment.CurrentDirectory).Parent.FullName + "/Template1Bewerkt.docx");
            doc.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);


        }
        private void CmboAlignLogo_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            var app = new Microsoft.Office.Interop.Word.Application();
            var doc = new Microsoft.Office.Interop.Word.Document();
            doc = app.Documents.Add(Directory.GetParent(Environment.CurrentDirectory).Parent.FullName + "/Template1Bewerkt.docx");
            //Hard gecodeerde sectie vanuit template
            Section s = doc.Sections[1];
            //Switch case om de uitlijning correct uit te voeren , Links staat er niet bij omdat dit "default" zo staat.
            switch (CmboAlignLogo.SelectedItem)
            {
                case "Rechts":
                    s.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    break;
                case "Centreren":
                    s.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    break;
                default:
                    s.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    break;

            }
            doc.SaveAs2(Directory.GetParent(Environment.CurrentDirectory).Parent.FullName + "/Template1Bewerkt.docx");
            doc.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
        }
        private void Change_Bedrijfgegevens(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {

            var app = new Microsoft.Office.Interop.Word.Application();
            var doc = new Microsoft.Office.Interop.Word.Document();
            doc = app.Documents.Add(Directory.GetParent(Environment.CurrentDirectory).Parent.FullName + "/Template1Bewerkt.docx");
            //Hard gecodeerde sectie vanuit template
            Section s = doc.Sections[2];
            //Switch case om de uitlijning correct uit te voeren , Links staat er niet bij omdat dit "default" zo staat.
            switch (cmboAllignCompanyData.SelectedItem)
            {
                case "Rechts":
                    s.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    break;
                case "Centreren":
                    s.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    break;
                default:
                    s.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    break;

            }
            doc.SaveAs2(Directory.GetParent(Environment.CurrentDirectory).Parent.FullName + "/Template1Bewerkt.docx");
            doc.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
        }
        private void Change_Aanspreking(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            var app = new Microsoft.Office.Interop.Word.Application();
            var doc = new Microsoft.Office.Interop.Word.Document();
            doc = app.Documents.Add(Directory.GetParent(Environment.CurrentDirectory).Parent.FullName + "/Template1Bewerkt.docx");
            //Hard gecodeerd sectie uit template
            Section s = doc.Sections[3];
            //Switch case om de uitlijning correct uit te voeren , Links staat er niet bij omdat dit "default" zo staat.
            switch (CmboAllignAanspreking.SelectedItem)
            {
                case "Rechts":
                    s.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    break;
                case "Centreren":
                    s.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    break;
                default:
                    s.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    break;

            }
            doc.SaveAs2(Directory.GetParent(Environment.CurrentDirectory).Parent.FullName + "/Template1Bewerkt.docx");
            doc.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
        }
        private void Change_Text(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            var app = new Microsoft.Office.Interop.Word.Application();
            var doc = new Microsoft.Office.Interop.Word.Document();
            doc = app.Documents.Add(Directory.GetParent(Environment.CurrentDirectory).Parent.FullName + "/Template1Bewerkt.docx");
            //Hard gecodeerd uit template
            Section s = doc.Sections[4];
            //Switch case om de uitlijning correct uit te voeren , Links staat er niet bij omdat dit "default" zo staat.
            switch (cmboAllignText.SelectedItem)
            {
                case "Rechts":
                    s.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    break;
                case "Centreren":
                    s.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    break;
                default:
                    s.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    break;

            }
            doc.SaveAs2(Directory.GetParent(Environment.CurrentDirectory).Parent.FullName + "/Template1Bewerkt.docx");
            doc.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
        }
        private void cmboAllignSlot_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            var app = new Microsoft.Office.Interop.Word.Application();
            var doc = new Microsoft.Office.Interop.Word.Document();
            doc = app.Documents.Add(Directory.GetParent(Environment.CurrentDirectory).Parent.FullName + "/Template1Bewerkt.docx");
            //Hard gecodeerd vanuit template
            Section s = doc.Sections[6];
            //Switch case om de uitlijning correct uit te voeren , Links staat er niet bij omdat dit "default" zo staat.
            switch (cmboAllignSlot.SelectedItem)
            {
                case "Rechts":
                    s.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    break;
                case "Centreren":
                    s.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    break;
                default:
                    s.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    break;

            }
            doc.SaveAs2(Directory.GetParent(Environment.CurrentDirectory).Parent.FullName + "/Template1Bewerkt.docx");
            doc.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
        }
        private void UpDown_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        { //Door dat numeric updown van syncfusion geen selectionchanged even heeft wordt dit gedaan door een mousleave event.
            //Om niet steeds de code te moeten uitvoeren wordt gecontroleerd of de waarde anders is dan voorheen.
            if (NmrLettergrootte.Value.Value != numericValue)
            {
                var app = new Microsoft.Office.Interop.Word.Application();
                var doc = new Microsoft.Office.Interop.Word.Document();
                doc = app.Documents.Add(Directory.GetParent(Environment.CurrentDirectory).Parent.FullName + "/Template1Bewerkt.docx");
                //Over geheel document wordt Size aangepast.
                doc.Content.Font.Size = Convert.ToInt32(NmrLettergrootte.Value.Value);
                doc.SaveAs2(Directory.GetParent(Environment.CurrentDirectory).Parent.FullName + "/Template1Bewerkt.docx");
                numericValue = NmrLettergrootte.Value.GetValueOrDefault();
                doc.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);

            }
        }

        //EIGEN METHODES !

        //Methode  om Tabel aan te maken , Parameter is Document dat moet meegegeven worden.
        private void Tabelaanmaken(Document doc)
        { //Het aantal velden doorlopen van het document via een loop
            for (int velden = 1; velden <= doc.Fields.Count; velden++)
            {
                if (doc.Fields[velden].Result.Text == "«Tabel_benodigheden»")
                { //hier wordt generieke tabel gemaakt, de count+1 is omdat er al een "Header" hard gecodeerd wordt. paramters = range van de sectie, lijstbenodigheden +,2 coloms,Autofit true , Autofit naar content
                    Word.Table table = doc.Fields[velden].Result.Tables.Add(doc.Sections[5].Range, benodighedenlijstkeuze.Count + 1, 2, DefaultTableBehavior: WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior: WdAutoFitBehavior.wdAutoFitContent);
                    //Header aanmaken 
                    table.Rows[1].Cells[1].Range.Text = "Beschrijving";
                    table.Rows[1].Cells[2].Range.Text = "Prijs";

                    for (int rijen = 2; rijen < benodighedenlijstkeuze.Count + 2; rijen++)
                    { //De tabel opvullen met de correcte gegevens.
                        table.Rows[rijen].Cells[1].Range.Text = benodighedenlijstkeuze[rijen - 2].Beschrijving;
                        table.Rows[rijen].Cells[2].Range.Text = "" + benodighedenlijstkeuze[rijen - 2].Prijs;
                    }
                }
            }


        }
        private string Datummeegeven()
        {
            //datum van vandaag in string plaatsen
            var DatumTijd = DateTime.Now;
            var tempDatum = DatumTijd.Day + "/" + DatumTijd.Month + "/" + DatumTijd.Year;
            string Datum = tempDatum.ToString();
            return Datum;
        }
        private IList<Sportclub> Sportclubselecteren()
        { //lijst doorlopen
            for (int i = 0; i < sportclublijst.Count; i++)
            { //Id's ophalen van sportclub met zelfde naam.
                if (sportclublijst[i].Naam_sportclub == CmboSportclub.SelectedItem.ToString())
                { //toevoegen
                    sportclubKeuze.Add(sportclublijst[i]);
                }
            }
            return sportclubKeuze;
        }
        //Hier wordt de hobby lijst aangemaakt die gelinked is aan de sportclub
        private void HobbiesSelecteren(IList<Sportclub> sportclubs)
        {
            //Aantal sportclubs doorlopen -> Dit is in het geval de Sportclub meer dan 1x voorkomt. Multiclub als voorbeeld.
            for (int k = 0; k < sportclubs.Count; k++)
            {
                //Hobbylijst doorlopen
                for (int j = 0; j < Hobbylijst.Count; j++)
                {   //Id sportclub = hobby.sportclub id --> toevoegen aan lijst
                    if (sportclubs[k].Id == Hobbylijst[j].Sportclub_id)
                    {   //Controleren of er geen dubbels worden toegevoegd. zoals bij Mulitclub (2 sporten voor 1 club)
                        if (!Hobbylijstkeuze.Contains(Hobbylijst[j]))
                        {

                            Hobbylijstkeuze.Add(Hobbylijst[j]);
                        }
                    }
                }
            }
        }
        private void PersonenPersportclub()
        { //Doorlopen van  hobbylijst die gelinkt is  met gekozen sportclub
            for (int i = 0; i < Hobbylijstkeuze.Count; i++)
            { //Doorlopen personen
                for (int k = 0; k < persoonlijst.Count; k++)
                {// als de persoon voorkomt in de gelinkte hobby lijst , voeg de persoon toe.
                    if (Hobbylijstkeuze[i].Persoon_id == persoonlijst[k].Id)
                    {
                        persoonlijstkeuze.Add(persoonlijst[k]);
                        Console.WriteLine(persoonlijstkeuze[i].Emailadres);

                    }
                }
            }
        }

        private void AdressenSelecteren(IList<Sportclub> gekozensportclubs)
        { //doorlopen Adressen
            for (int i = 0; i < adreslijst.Count; i++)
            { //Doorlopen sportclubs
                for (int j = 0; j < gekozensportclubs.Count; j++)
                    //Doorlopen personen lijst die gelinkt is aan gekozen sportclub
                    for (int l = 0; l < persoonlijstkeuze.Count; l++)
                    { //Hier zowel adressen van sportclubs als personen toevoegen.
                        if ((adreslijst[i].Id == gekozensportclubs[j].Adresgegeven_id) || (adreslijst[i].Id == persoonlijstkeuze[l].Adresgeeven_id))
                        { //Hier check ik of ik niet gewoon  dezelfde toevoeg -> geen duplicaten
                            if (!adreslijstkeuze.Contains(adreslijst[i]))
                            {
                                adreslijstkeuze.Add(adreslijst[i]);
                            }
                        }
                    }
            }
        }  
    private string AdresPerpersoonSelecteren(int persoonlijstloop)//Opgepast functie werkt alleen wanneer geplaatst in loop 
        {
        var TextMailmerge = "";
            //Doorlopen lijst personen gelinkt aan gekozen sportclub(s)
        for (int i = 0; i < persoonlijstkeuze.Count; i++)
        { //Doorlopen van adressenlijst die gelinkt is aan hobby, die alweer gelinkt is aan Sportclubs
            for (int j = 0; j < adreslijstkeuze.Count; j++)
            { //Indien de persoon adresid gevonden wordt in de  gelinkte adressenlijst tsla deze op als string. en return deze waarde
                if (persoonlijstkeuze[persoonlijstloop].Adresgeeven_id == adreslijstkeuze[j].Id)
                {

                    TextMailmerge = adreslijstkeuze[j].Adres + " " + adreslijstkeuze[j].Gemeente + " " + adreslijstkeuze[j].Postcode;
                }
            }
        }
        return TextMailmerge;
    }
    private String AdresPerSportclubSelecteren(int Sportclubloop) //Opgepast functie werkt alleen wanneer geplaatst in loop 
    {
        var TextMailmerge = "";
            //adressenlijst die gelinkt is aan hobby, die op zijn beurt gelinkt is aan de gekozen sportclub(s)
        for (int j = 0; j < adreslijstkeuze.Count; j++)
        { //Indien sportclub adres in de gelinkte lijst aanwezig , deze als String weergeven met alle data.
            if (sportclubKeuze[Sportclubloop].Adresgegeven_id == adreslijstkeuze[j].Id)
            {
                TextMailmerge = adreslijstkeuze[j].Adres + " " + adreslijstkeuze[j].Gemeente + " " + adreslijstkeuze[j].Postcode;
            }

        }
        return TextMailmerge;
    }
    private void SportenSelecteren(IList<Sportclub> gekozensportclubs)
    { //Doorlopen Sportlijst
        for (int i = 0; i < sportlijst.Count; i++)
        { //Doorlopen sportclubs --> deze parameter is een gelinkte list.
            for (int k = 0; k < gekozensportclubs.Count; k++)
            { //na controle toevoegen van Sport
                if (sportlijst[i].Id == gekozensportclubs[k].Sport_id)
                { //Hier check ik of ik niet gewoon dezelfde toevoeg
                    if (!sportlijstkeuze.Contains(sportlijst[i]))
                    {
                        sportlijstkeuze.Add(sportlijst[i]);
                    }
                }
            }
        }

    }
    private void BenodigdhedenSelecteren(IList<Sportclub> gekozensportclubs)
    { //Doorlopen benodigdheden
        for (int i = 0; i < benodighedenlijst.Count; i++)
        {//Doorlopen gelinkte lijst van sportclub(s)
            for (int j = 0; j < gekozensportclubs.Count; j++)
            { //als de sportid van benodigdheden  in de gekozen sportclubs gelijk zijn , voeg deze benodigheid toe
                if (benodighedenlijst[i].Sport_id == gekozensportclubs[j].Sport_id)
                {//Hier check ik of ik niet gewoon 2x dezelfde toevoeg
                    if (!benodighedenlijstkeuze.Contains(benodighedenlijst[i]))
                    {
                        benodighedenlijstkeuze.Add(benodighedenlijst[i]);
                    }
                }
            }
        }
    }
    private String MailmergeCase(Field Mergefields, int Sportclubloop, int persoonlijstloop) //Opgepast! Functie werkt alléén in geneste loops, ->  
    {
        var TextMailmerge = "";
            //Switch case aangemaakt om TextMailmerge  correct op te vullen naar de correcte Mergefield.
        switch (Mergefields.Result.Text)

        {
            case "«Naam_sportclub»":
                TextMailmerge = sportclubKeuze[0].Naam_sportclub; //deze mag op 0 blijven staan gebruiken naam om lijsten te laden.

                break;
            case "«Emailadres_sportclub»":
                if (sportclubKeuze[Sportclubloop].Emailadres != TextMailmerge) // --> duplicaten uithalen, maar indien 2 verschillende wordt deze ook weergegeven
                        //bv Multiclub heeft 2 sporten, iedere "dienst" kan zijn eigen email hebben.
                {
                    TextMailmerge = sportclubKeuze[Sportclubloop].Emailadres;
                }
                break;
            case "«Adres_sportclub»":
                TextMailmerge = AdresPerSportclubSelecteren(Sportclubloop);
                break;
            case "«Voornaam_persoon»":
                TextMailmerge = persoonlijstkeuze[persoonlijstloop].Voornaam;
                break;
            case "«Achternaam_persoon»":
                TextMailmerge = persoonlijstkeuze[persoonlijstloop].Achternaam;
                break;
            case "«Geboortedatum»":
                TextMailmerge = "" + persoonlijstkeuze[persoonlijstloop].Geboortedatum;
                break;
            case "«Emailadres_persoon»":
                if (persoonlijstkeuze[persoonlijstloop].Emailadres != TextMailmerge)// --> Duplicaten uithalen, als 1 persoon 2 sporten uitoefent bij zelfde sportclub 
                {
                    TextMailmerge = persoonlijstkeuze[persoonlijstloop].Emailadres;
                }
                break;
            case "«Adres_persoon»":
                TextMailmerge = AdresPerpersoonSelecteren(persoonlijstloop);
                break;
            case "«Datum_verstuurd»":
                TextMailmerge = Datummeegeven();
                break;
        }
        return TextMailmerge;
    }


}
}