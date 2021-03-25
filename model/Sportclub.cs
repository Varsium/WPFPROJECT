
namespace EindOpdrachtLeandro.model
{
    class Sportclub
    {
        private int id;
        private string naam_sportclub;
        private string emailadres;
        private string logo_sportclub;
        private int adresgegeven_id;
        private int sport_id;

        //Getters & setters aanmaken
        public int Id { get => id; set => id = value; }
        public string Naam_sportclub { get => naam_sportclub; set => naam_sportclub = value; }
        public string Emailadres { get => emailadres; set => emailadres = value; }
        public string Logo_sportclub { get => logo_sportclub; set => logo_sportclub = value; }
        public int Adresgegeven_id { get => adresgegeven_id; set => adresgegeven_id = value; }
        public int Sport_id { get => sport_id; set => sport_id = value; }



    }
}
