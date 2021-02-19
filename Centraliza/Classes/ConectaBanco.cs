using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Centraliza
{
    static class ConectaBanco
    {
        public static SqlConnection AbreBanco()
        {
                SqlConnection conexao = new SqlConnection();
                //conexao.ConnectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=D:\ProgramaSolar\Centraliza\Centraliza\BDsolar.mdf;Integrated Security=True;Connect Timeout=30";
                //conexao.ConnectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=\\ANDERSON-PC\Cia Solar\Centraliza\BDsolar.mdf;Integrated Security=True;Connect Timeout=30";
                conexao.ConnectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Centraliza\Centraliza\Banco\BDsolar.mdf;Integrated Security=True;Connect Timeout=30";
                conexao.Open();
                return conexao;
        }

        public static SqlDataReader ExecutaConsulta(string SQL)
        {
            SqlCommand comando = new SqlCommand();
            comando.CommandType = CommandType.Text;
            comando.CommandText = SQL;

            comando.Connection = AbreBanco();

            return comando.ExecuteReader();
        }
        public static void ExecutaComando(string SQL)
        {
            SqlCommand comando = new SqlCommand();
            comando.CommandType = CommandType.Text;
            comando.CommandText = SQL;
            comando.Connection = AbreBanco();

            comando.ExecuteNonQuery();
        }

        public static SqlConnection FechaBanco()
        {
            SqlConnection conexao = new SqlConnection();
            conexao.Close();
            return conexao;
        }
    }
}
