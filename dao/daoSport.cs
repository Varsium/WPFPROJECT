using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using EindOpdrachtLeandro.model;
using EindOpdrachtLeandro.Repositories;

namespace EindOpdrachtLeandro.dao
{
    class daoSport : Isport
    {
        private SqlConnection connection;

        public daoSport(SqlConnection con)
        {
            this.connection = con;
        }

        public IList<Sport> GetAll()
        {
            List<Sport> sports = new List<Sport>();
            var sql = new SqlCommand("select * from Sport", connection); //query
            SqlDataReader reader = sql.ExecuteReader();
            while (reader.Read())
            {//Object genereren en in lijst opslaan
                sports.Add(new Sport
                {
                    Id = Convert.ToInt32(reader[0]),
                    SporT = Convert.ToString(reader[1])
                });
            }
            reader.Close();
            return sports;
        }

    }
}
