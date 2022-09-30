using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using Npgsql;

namespace UrnaDeMaria.Uteis
{
    public class DAL
    {
        private static string Server = "localhost";
        private static string Database = "UrnaEletronica";
        private static string User = "postgres";
        private static string Password = "123";
        private static string port = "5432";
        private static string ConnectionString = $"Server={Server};Port={port};User Id={User};Password={Password};Database={Database};";
        NpgsqlConnection Connection;
        public string connString = null;
        

        public DAL()
        {
            connString = String.Format("Server={0};Port={1};User Id={2};Password={3};Database={4};",
                                                        Server, port, User, Password, Database);

            Connection = new NpgsqlConnection(ConnectionString);
            Connection.Open();
        }


        public DataTable RetDataTable(string sql)
        {
            DataTable data = new DataTable();
            NpgsqlCommand Command = new NpgsqlCommand(sql, Connection);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(Command);
            da.Fill(data);
            return data;
        }

        //Espera um parâmetro do tipo string 
        //contendo um comando SQL do tipo INSERT, UPDATE, DELETE
        public void ExecutarComandoSQL(string sql)
        {
            NpgsqlCommand Command = new NpgsqlCommand(sql, Connection);
            Command.ExecuteNonQuery();
        }
         
    }
}
