namespace EindOpdrachtLeandro.model
{
    class Benodigheden
    {
        private int id;
        private string beschrijving;
        private double prijs;
        private int sport_id;

        //Getters& setters van variabelen 
        public int Id { get => id; set => id = value; }
        public string Beschrijving { get => beschrijving; set => beschrijving = value; }
        public double Prijs { get => prijs; set => prijs = value; }
        public int Sport_id { get => sport_id; set => sport_id = value; }



    }
}
