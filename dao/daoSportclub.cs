using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using EindOpdrachtLeandro.model;
using EindOpdrachtLeandro.Repositories;


namespace EindOpdrachtLeandro.dao
{
    class daoSportclub : Isportclub
    {
        SqlConnection connection;

        public daoSportclub(SqlConnection con)
        {
            this.connection = con;
        }

        public IList<Sportclub> GetAll()
        {
            var sportclubs = new List<Sportclub>();
            var sql = new SqlCommand("Select * from Sportclub", connection); //Query
            SqlDataReader reader = sql.ExecuteReader();
            while (reader.Read())
            { //Lijst opvullen met objecten
                sportclubs.Add(new Sportclub
                {
                    Id = Convert.ToInt32(reader[0]),
                    Naam_sportclub = Convert.ToString(reader[1]),
                    Emailadres = Convert.ToString(reader[2]),
                    Logo_sportclub = Convert.ToString(reader[3]),
                    Adresgegeven_id = Convert.ToInt32(reader[4]),
                    Sport_id = Convert.ToInt32(reader[5]),

                });

            }
            reader.Close();

            return sportclubs;
        }
        public IList<string> GetSportclubNaam() //extra functie gemaakt om de namen te scheiden.
        {
            var sportclubnames = new List<String>();
            for (int i = 0; i < GetAll().Count; i++)
            {
                sportclubnames.Add(
                GetAll()[i].Naam_sportclub

                );
            }

            return sportclubnames;
        }
    }
}
