using System;

namespace EindOpdrachtLeandro.model
{
    class Adresgegeven
    {
        private int id;
        private String adres;
        private String postcode;
        private String gemeente;

        //Getters  en setters aanmaken 
        public int Id { get => id; set => id = value; }
        public string Adres { get => adres; set => adres = value; }
        public string Postcode { get => postcode; set => postcode = value; }
        public string Gemeente { get => gemeente; set => gemeente = value; }

    }
}
