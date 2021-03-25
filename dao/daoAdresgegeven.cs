using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using EindOpdrachtLeandro.model;
using EindOpdrachtLeandro.Repositories;

namespace EindOpdrachtLeandro.dao
{
    class daoAdresgegeven : Iadresgegeven
    {
        private SqlConnection connection;

        public daoAdresgegeven(SqlConnection con)
        {
            this.connection = con;
        }

        public IList<Adresgegeven> GetAll()
        {
            var Adressen = new List<Adresgegeven>(); //Lijst aanmaken

            var sql = new SqlCommand("Select * from Adresgegeven", connection); //query opslaan/maken 
            SqlDataReader reader = sql.ExecuteReader();//uitvoeren query
            while (reader.Read())
            { // objecten genereren en opslaan in lijst
                Adressen.Add(new Adresgegeven
                {
                    Id = Convert.ToInt32(reader[0]),
                    Adres = Convert.ToString(reader[1]),
                    Postcode = Convert.ToString(reader[2]),
                    Gemeente = Convert.ToString(reader[3])
                });
            }
            reader.Close(); //Reader sluiten voor volgende methods

            return Adressen;
        }
    }
}
