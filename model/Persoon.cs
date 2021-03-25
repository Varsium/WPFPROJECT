using System;

namespace EindOpdrachtLeandro.model
{
    class Persoon
    {
        private int id;
        private string voornaam;
        private string achternaam;
        private DateTime geboortedatum;
        private string figuur;
        private String emailadres;
        private int adresgeeven_id;

        //GETTERS & setters aanmaken
        public int Id { get => id; set => id = value; }
        public string Voornaam { get => voornaam; set => voornaam = value; }
        public string Achternaam { get => achternaam; set => achternaam = value; }
        public DateTime Geboortedatum { get => geboortedatum; set => geboortedatum = value; }
        public string Figuur { get => figuur; set => figuur = value; }
        public string Emailadres { get => emailadres; set => emailadres = value; }
        public int Adresgeeven_id { get => adresgeeven_id; set => adresgeeven_id = value; }



    }
}
