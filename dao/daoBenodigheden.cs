using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using EindOpdrachtLeandro.model;
using EindOpdrachtLeandro.Repositories;

namespace EindOpdrachtLeandro.dao
{
    class daoBenodigheden : Ibenodigheden
    {
        public SqlConnection connection;

        public daoBenodigheden(SqlConnection con)
        {
            this.connection = con;
        }

        public IList<Benodigheden> GetAll()
        {
            var Lijstbenodigheden = new List<Benodigheden>();
            var sql = new SqlCommand("Select * from Benodigdheden", connection); //query aanmaken
            SqlDataReader Reader = sql.ExecuteReader();
            while (Reader.Read())
            { //objecten in lijst steken.
                Lijstbenodigheden.Add(
                new Benodigheden
                {
                    Id = Convert.ToInt32(Reader[0]),
                    Beschrijving = Convert.ToString(Reader[1]),
                    Prijs = Convert.ToDouble(Reader[2]),
                    Sport_id = Convert.ToInt32(Reader[3])
                }

                    );
            }
            Reader.Close();
            return Lijstbenodigheden;
        }

    }
}

