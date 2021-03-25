using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using EindOpdrachtLeandro.model;
using EindOpdrachtLeandro.Repositories;

namespace EindOpdrachtLeandro.dao
{
    class daoHobby : Ihobby
    {
        private SqlConnection connection;

        public daoHobby(SqlConnection connection)
        {
            this.connection = connection;
        }
        public IList<Hobby> GetAll()
        {
            var hobbies = new List<Hobby>();
            var sql = new SqlCommand("Select * from hobby", connection); //query
            SqlDataReader reader = sql.ExecuteReader();
            while (reader.Read())
            { //objecten maken, en in lijst steken
                hobbies.Add(new Hobby
                {
                    Id = Convert.ToInt32(reader[0]),
                    Sportclub_id = Convert.ToInt32(reader[1]),
                    Persoon_id = Convert.ToInt32(reader[2])

                });
            }
            reader.Close();
            return hobbies;
        }

    }

}
