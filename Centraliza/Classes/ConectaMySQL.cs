using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace Centraliza.Classes
{
    class ConectaMySQL
    {
        private static FuncoesBanco func = new FuncoesBanco();
        public MySqlConnection AbreMySQL()
        {
            try
            {
                MySqlConnection conexao = new MySqlConnection("SERVER=" + "192.168.56.1" + ";PORT=3306;DATABASE=" + "mysolar" + ";UID=" + "MASTER" + ";PASSWORD=" + "303304" + "");
                conexao.Open();
                return conexao;
            }
            catch(Exception ex)
            {
                MessageBox.Show("Erro ao conectar-se com o Banco de Dados. " + Environment.NewLine + ex.Message);
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
