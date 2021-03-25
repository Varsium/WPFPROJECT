using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using EindOpdrachtLeandro.model;
using EindOpdrachtLeandro.Repositories;

namespace EindOpdrachtLeandro.dao
{
    class daoPersoon : Ipersoon
    {
        private SqlConnection connection;

        public daoPersoon(SqlConnection con)
        {
            this.connection = con;
        }

        public IList<Persoon> GetAll()
        {
            List<Persoon> personen = new List<Persoon>();
            var sql = new SqlCommand("Select * from Persoon", connection); //query
            SqlDataReader reader = sql.ExecuteReader();
            while (reader.Read())
            {//Objecten genereren en in lijst opslaan
                personen.Add(new Persoon
                {
                    Id = Convert.ToInt32(reader[0]),
                    Voornaam = Convert.ToString(reader[1]),
                    Achternaam = Convert.ToString(reader[2]),
                    Geboortedatum = Convert.ToDateTime(reader[3]),
                    Figuur = Convert.ToString(reader[4]),
                    Emailadres = Convert.ToString(reader[5]),
                    Adresgeeven_id = Convert.ToInt32(reader[6])

                });
            }
            reader.Close();
            return personen;
        }
    }
}
