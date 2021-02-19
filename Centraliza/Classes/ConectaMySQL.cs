using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;

namespace Centraliza.Classes
{
    class ConectaMySQL
    {
        private static FuncoesBanco func = new FuncoesBanco();
        public MySqlConnection AbreMySQL()
        {
            func.SelecionaBanco();
            if (func.Banco == "mysql")
            {
                MySqlConnection conexao = new MySqlConnection("SERVER=" + func.Servidor + ";PORT=3306;DATABASE=" + func.NomeDB + ";UID=" + func.UID + ";PASSWORD=" + func.Password +"");
                conexao.Open();
                return conexao;
            }
            else
            {
                return null;
            }
        }
        public MySqlConnection FechaMySQL()
        {
            MySqlConnection conexao = new MySqlConnection();
            conexao.Close();
            return conexao;
        }
        public void ExecutaComando(string SQL)
        {
            MySqlCommand mySqlCommand = new MySqlCommand();
            mySqlCommand.CommandType = System.Data.CommandType.Text;
            mySqlCommand.CommandText = SQL;

            mySqlCommand.Connection = AbreMySQL();

            mySqlCommand.ExecuteNonQuery();
        }
        public MySqlDataReader ExecutaConsulta(string SQL)
        {
            MySqlCommand mySqlCommand = new MySqlCommand();
            mySqlCommand.CommandType = System.Data.CommandType.Text;
            mySqlCommand.CommandText = SQL;

            mySqlCommand.Connection = AbreMySQL();

            return mySqlCommand.ExecuteReader();
        }

    }
}
