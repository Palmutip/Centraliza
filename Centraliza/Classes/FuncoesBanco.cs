using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;
using Centraliza.Classes;

namespace Centraliza
{
    class FuncoesBanco
    {
        //Variaveis
        ConectaMySQL conectaMySQL = new ConectaMySQL();
        public int Id { get; set; }
        public string id { get; set; }
        public string Login { get; set; }
        public string Senha { get; set; }
        public string Custo { get; set; }
        public string Media { get; set; }
        public string Janeiro { get; set; }
        public string Fevereiro { get; set; }
        public string Marco { get; set; }
        public string Abril { get; set; }
        public string Maio { get; set; }
        public string Junho { get; set; }
        public string Julho { get; set; }
        public string Agosto { get; set; }
        public string Setembro { get; set; }
        public string Outubro { get; set; }
        public string Novembro { get; set; }
        public string Dezembro { get; set; }
        public int Disponibilidade { get; set; }
        public int CliJan { get; set; }
        public int CliFev { get; set; }
        public int CliMar { get; set; }
        public int CliAbr { get; set; }
        public int CliMai { get; set; }
        public int CliJun { get; set; }
        public int CliJul { get; set; }
        public int CliAgo { get; set; }
        public int CliSet { get; set; }
        public int CliOut { get; set; }
        public int CliNov { get; set; }
        public int CliDez { get; set; }
        public string Nome { get; set; }
        public string CPF { get; set; }
        public string Endereco { get; set; }
        public string Cidade { get; set; }
        public string UF { get; set; }
        public string email { get; set; }
        public string Telefone { get; set; }
        public string Celular { get; set; }
        public string PotenciaInstalada { get; set; }
        public string QuantidadeModulos { get; set; }
        public string QuantidadeInversores { get; set; }
        public string ModeloInversor { get; set; }
        public string ModeloModulo { get; set; }
        public string Identificacao { get; set; }
        public string PotenciaInv { get; set; }
        public string PotenciaMod { get; set; }
        public string ComrpimentoMod { get; set; }
        public string LarguraMod { get; set; }
        public string GarantiaInv { get; set; }
        public string GarantiaMod { get; set; }
        public string MarcaInversor { get; set; }
        public string MarcaMod { get; set; }
        public string EficienciaInv { get; set; }
        public string Qtdmppt { get; set; }
        public string TemperaturaModulo { get; set; }
        public string Numero { get; set; }
        public string Complemento { get; set; }
        public string Bairro { get; set; }
        public string CEP { get; set; }
        public string Fases { get; set; }
        public string Tensao { get; set; }
        public string RegistroINMETRO { get; set; }
        public string Material { get; set; }
        public string Celulas { get; set; }
        public string RegistroInmetro { get; set; }
        public string MediaConsumo { get; set; }
        public int Proposta { get; set; }
        public string Contato { get; set; }
        public string Consumoanual { get;  set; }
        public string Perdas { get;  set; }
        public string Obs { get;  set; }
        public string Valorinv { get;  set; }
        public string Valorequip { get;  set; }
        public string Valorsist { get;  set; }
        public string Estu { get;  set; }
        public string Transformador { get;  set; }
        public string StringBox { get;  set; }
        public string Credito { get;  set; }
        public string Banco { get; set; }
        public string Caminho { get; set; }
        public string Servidor { get; set; }
        public string NomeDB { get; set; }
        public string UID { get; set; }
        public string Password { get; set; }
        public byte[] Foto { get; set; }
        public string NumeroCliente { get; set; }
        public string NumeroInstalacao { get; set; }
        public string Classe { get; set; }
        public string Latitude { get; set; }
        public string Longitude { get; set; }
        public string Disjuntor { get; set; }
        public string CargaInstalada { get; set; }
        public string Arranjo { get;  set; }

        //Comandos gerais
        public void LimpaTambem()
        {

            Id = 0;
            id = "";
            Login = "";
            Senha = "";
            Custo = "";
            Media = "";
            Janeiro = "";
            Fevereiro = "";
            Marco = "";
            Abril = "";
            Maio = "";
            Junho = "";
            Julho = "";
            Agosto = "";
            Setembro = "";
            Outubro = "";
            Novembro = "";
            Dezembro = "";
            Disponibilidade = 0;
            CliJan = 0;
            CliFev = 0;
            CliMar = 0;
            CliAbr = 0;
            CliMai = 0;
            CliJun = 0;
            CliJul = 0;
            CliAgo = 0;
            CliSet = 0;
            CliOut = 0;
            CliNov = 0;
            CliDez = 0;
            Nome = "";
            CPF = "";
            Endereco = "";
            Cidade = "";
            UF = "";
            email = "";
            Telefone = "";
            Celular = "";
            PotenciaInstalada = "";
            QuantidadeModulos = "";
            QuantidadeInversores = "";
            ModeloInversor = "";
            ModeloModulo = "";
            Identificacao = "";
            PotenciaInv = "";
            PotenciaMod = "";
            ComrpimentoMod = "";
            LarguraMod = "";
            GarantiaInv = "";
            GarantiaMod = "";
            MarcaInversor = "";
            MarcaMod = "";
            EficienciaInv = "";
            Qtdmppt = "";
            TemperaturaModulo = "";
            Numero = "";
            Complemento = "";
            Bairro = "";
            CEP = "";
            Fases = "";
            Tensao = "";
            RegistroINMETRO = "";
            Material = "";
            Celulas = "";
            RegistroInmetro = "";
            MediaConsumo = "";
            Proposta = 0;
            Contato = "";
            Consumoanual = "";
            Perdas = "";
            Obs = "";
            Valorinv = "";
            Valorequip = "";
            Valorsist = "";
            Estu = "";
            Banco = "";
            Caminho = "";
            Servidor = "";
            NomeDB = "";
            UID = "";
            Password = "";
            Transformador = "";
            StringBox = "";
            Credito = "";
            Arranjo = "";
            NumeroCliente = "";
            NumeroInstalacao = "";
            Classe = "";
            Latitude = "";
            Longitude = "";
            Disjuntor = "";
            CargaInstalada = "";
    }

        //Tabela Clientes
        public void CriaTableClientes()
        {
            string SQL;
            //SQL = "DROP TABLE Clientes;";

            //MySqlDataReader dados = conectaMySQL.ExecutaConsulta(SQL);
            //SqlDataReader dados = ConectaBanco.ExecutaConsulta(SQL);
            //ConectaBanco.FechaBanco();
            //conectaMySQL.FechaMySQL();
            //string SQL;
            /*SQL = "CREATE TABLE Clientes (Id INT NOT NULL AUTO_INCREMENT, Nome VARCHAR(255) NOT NULL, " +
                  "CPF_CNPJ NVARCHAR(255) NOT NULL, CEP NVARCHAR(255) NOT NULL, Endereco NVARCHAR(255) NOT NULL, Numero NVARCHAR(255) NOT NULL," +
                  "Complemento  NVARCHAR(255) NOT NULL, Bairro NVARCHAR(255) NOT NULL, Cidade NVARCHAR(255) NOT NULL, " +
                  "UF  NVARCHAR(255) NOT NULL, email NVARCHAR(255) NOT NULL, Telefone  NVARCHAR(255) NOT NULL," +
                  "Celular NVARCHAR(255) NOT NULL, Quantidade_Inversores " +
                  "NVARCHAR(255) NOT NULL, Marca_Inversor NVARCHAR(255) NOT NULL, Modelo_Inversor NVARCHAR(255) NOT NULL, Quantidade_Modulos " +
                  "NVARCHAR(255) NOT NULL, Marca_Modulo NVARCHAR(255) NOT NULL, Modelo_Modulo NVARCHAR(255) NOT NULL, Consumo_Medio NVARCHAR(255) NOT NULL," +
                  "Identificacao NVARCHAR(255) NOT NULL, PRIMARY KEY (Id));";*/
            SQL = "CREATE TABLE Clientes (Id INT IDENTITY (1, 1) NOT NULL, Nome NVARCHAR(255) NOT NULL, " +
            "CPF_CNPJ NVARCHAR(255) NOT NULL, CEP NVARCHAR(255) NOT NULL, Endereco NVARCHAR(255) NOT NULL, Numero NVARCHAR(255) NOT NULL," +
            "Complemento  NVARCHAR(255) NOT NULL, Bairro NVARCHAR(255) NOT NULL, Cidade NVARCHAR(255) NOT NULL, " +
            "UF  NVARCHAR(255) NOT NULL, email NVARCHAR(255) NOT NULL, Telefone  NVARCHAR(255) NOT NULL," +
            "Celular NVARCHAR(255) NOT NULL, Quantidade_Inversores " +
            "NVARCHAR(255) NOT NULL, Marca_Inversor NVARCHAR(255) NOT NULL, Modelo_Inversor NVARCHAR(255) NOT NULL, Quantidade_Modulos " +
            "NVARCHAR(255) NOT NULL, Marca_Modulo NVARCHAR(255) NOT NULL, Modelo_Modulo NVARCHAR(255) NOT NULL, Consumo_Medio NVARCHAR(255) NOT NULL," +
            "Identificacao NVARCHAR(255) NOT NULL, PRIMARY KEY CLUSTERED ([Id] ASC));";

            //MySqlDataReader dados = conectaMySQL.ExecutaConsulta(SQL);
            SqlDataReader dados = ConectaBanco.ExecutaConsulta(SQL);
            ConectaBanco.FechaBanco();
            //conectaMySQL.FechaMySQL();
        }
        public bool InserirCliente(string BD)
        {
            if (BD == "local")
            {
                string SQL;

                SQL = "Insert into Clientes (Nome, CPF_CNPJ, CEP, Endereco, Numero, Complemento, Bairro, Cidade, UF, email, Telefone, Celular, Quantidade_Inversores, Marca_Inversor, Modelo_Inversor, Quantidade_Modulos, Marca_Modulo, Modelo_Modulo , Consumo_Medio, Identificacao) values('" + Nome + "','" + CPF + "','" + CEP + "','" + Endereco + "','" + Numero + "','" + Complemento + "','" + Bairro + "','" + Cidade + "','" + UF + "','" + email + "','" + Telefone + "','" + Celular + "','" + QuantidadeInversores + "','" + MarcaInversor + "','" + ModeloInversor + "','" + QuantidadeModulos + "','" + MarcaMod + "','" + ModeloModulo + "','" + MediaConsumo + "','" + Identificacao + "')";

                ConectaBanco.ExecutaComando(SQL);
                ConectaBanco.FechaBanco();
            }
            else if (BD == "mysql")
            {
                string SQL;

                SQL = "Insert into Clientes (Nome, CPF_CNPJ, CEP, Endereco, Numero, Complemento, Bairro, Cidade, UF, email, Telefone, Celular, Quantidade_Inversores, Marca_Inversor, Modelo_Inversor, Quantidade_Modulos, Marca_Modulo, Modelo_Modulo , Consumo_Medio, Identificacao) values('" + Nome + "','" + CPF + "','" + CEP + "','" + Endereco + "','" + Numero + "','" + Complemento + "','" + Bairro + "','" + Cidade + "','" + UF + "','" + email + "','" + Telefone + "','" + Celular + "','" + QuantidadeInversores + "','" + MarcaInversor + "','" + ModeloInversor + "','" + QuantidadeModulos + "','" + MarcaMod + "','" + ModeloModulo + "','" + MediaConsumo + "','" + Identificacao + "')";

                conectaMySQL.ExecutaComando(SQL);
                conectaMySQL.FechaMySQL();
            }
            return true;
        }
        public bool AlterarCliente(string BD)
        {
            if (BD == "local")
            {
                string SQL;

                SQL = "Update Clientes set Nome='" + Nome + "', CPF_CNPJ='" + CPF + "', CEP='" + CEP + "', Endereco='" + Endereco + "', Numero='" + Numero + "', Complemento='" + Complemento + "'," +
                    " Bairro='" + Bairro + "', Cidade='" + Cidade + "', UF='" + UF + "'," +
                    " email='" + email + "', Telefone='" + Telefone + "', Celular='" + Celular + "'," +
                    "Quantidade_Modulos='" + QuantidadeModulos + "', Marca_Modulo ='" + MarcaMod + "', Modelo_Modulo ='" + ModeloModulo + "', Quantidade_Inversores ='" + QuantidadeInversores + "', Marca_Inversor ='" + MarcaInversor + "', Modelo_Inversor ='" + ModeloInversor + "', Consumo_Medio ='" + MediaConsumo + "', " +
                    "Identificacao='" + Identificacao + "' where Id =" + id;

                ConectaBanco.ExecutaComando(SQL);
                ConectaBanco.FechaBanco();
            }
            else if (BD == "mysql")
            {
                string SQL;

                SQL = "Update Clientes set Nome='" + Nome + "', CPF_CNPJ='" + CPF + "', CEP='" + CEP + "', Endereco='" + Endereco + "', Numero='" + Numero + "', Complemento='" + Complemento + "'," +
                    " Bairro='" + Bairro + "', Cidade='" + Cidade + "', UF='" + UF + "'," +
                    " email='" + email + "', Telefone='" + Telefone + "', Celular='" + Celular + "'," +
                    "Quantidade_Modulos='" + QuantidadeModulos + "', Marca_Modulo ='" + MarcaMod + "', Modelo_Modulo ='" + ModeloModulo + "', Quantidade_Inversores ='" + QuantidadeInversores + "', Marca_Inversor ='" + MarcaInversor + "', Modelo_Inversor ='" + ModeloInversor + "', Consumo_Medio ='" + MediaConsumo + "', " +
                    "Identificacao='" + Identificacao + "' where Id =" + id;

                conectaMySQL.ExecutaComando(SQL);
                conectaMySQL.FechaMySQL();
            }
            return true;
        }
        public void SelecionaCliente(string BD)
        {
            if (BD == "local")
            {
                string SQL;
                SQL = "Select * from Clientes where ID=" + id;

                SqlDataReader dados = ConectaBanco.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    Nome = dados["Nome"].ToString();
                    CPF = dados["CPF_CNPJ"].ToString();
                    Endereco = dados["Endereco"].ToString();
                    Numero = dados["Numero"].ToString();
                    Complemento = dados["Complemento"].ToString();
                    Bairro = dados["Bairro"].ToString();
                    CEP = dados["CEP"].ToString();
                    Cidade = dados["Cidade"].ToString();
                    UF = dados["UF"].ToString();
                    email = dados["email"].ToString();
                    Telefone = dados["Telefone"].ToString();
                    Celular = dados["Celular"].ToString();
                    QuantidadeModulos = dados["Quantidade_Modulos"].ToString();
                    QuantidadeInversores = dados["Quantidade_Inversores"].ToString();
                    ModeloInversor = dados["Modelo_Inversor"].ToString();
                    ModeloModulo = dados["Modelo_Modulo"].ToString();
                    MarcaInversor = dados["Marca_Inversor"].ToString();
                    MarcaMod = dados["Marca_Modulo"].ToString();
                    MediaConsumo = dados["Consumo_Medio"].ToString();
                    Identificacao = dados["Identificacao"].ToString();
                }
                ConectaBanco.FechaBanco();
            }
            else if (BD == "mysql")
            {
                string SQL;
                SQL = "Select * from Clientes where ID=" + id;

                MySqlDataReader dados = conectaMySQL.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    Nome = dados["Nome"].ToString();
                    CPF = dados["CPF_CNPJ"].ToString();
                    Endereco = dados["Endereco"].ToString();
                    Numero = dados["Numero"].ToString();
                    Complemento = dados["Complemento"].ToString();
                    Bairro = dados["Bairro"].ToString();
                    CEP = dados["CEP"].ToString();
                    Cidade = dados["Cidade"].ToString();
                    UF = dados["UF"].ToString();
                    email = dados["email"].ToString();
                    Telefone = dados["Telefone"].ToString();
                    Celular = dados["Celular"].ToString();
                    QuantidadeModulos = dados["Quantidade_Modulos"].ToString();
                    QuantidadeInversores = dados["Quantidade_Inversores"].ToString();
                    ModeloInversor = dados["Modelo_Inversor"].ToString();
                    ModeloModulo = dados["Modelo_Modulo"].ToString();
                    MarcaInversor = dados["Marca_Inversor"].ToString();
                    MarcaMod = dados["Marca_Modulo"].ToString();
                    MediaConsumo = dados["Consumo_Medio"].ToString();
                    Identificacao = dados["Identificacao"].ToString();
                }
                conectaMySQL.FechaMySQL();
            }
        }
        public DataTable AtualizaClientes(string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Clientes";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Clientes";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }
        }
        public DataTable ClienteNome(string TextoPesquisa, string coluna, string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Clientes where " + coluna + " like '%" + TextoPesquisa + "%'";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Clientes where " + coluna + " like '%" + TextoPesquisa + "%'";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }
        }
        public void ExcluiClientes(string BD)
        {
            if (BD == "local")
            {
                ConectaBanco.ExecutaComando("Delete from Clientes where ID=" + id);
                ConectaBanco.FechaBanco();
            }
            else if (BD == "mysql")
            {
                conectaMySQL.ExecutaComando("Delete from Clientes where ID=" + id);
                conectaMySQL.FechaMySQL();
            }
        }
        public bool Reseed(string BD)
        {
            if (BD == "local")
            {
                string SQL;

                SQL = "DBCC CHECKIDENT ('Clientes', RESEED, 0)";

                ConectaBanco.ExecutaComando(SQL);
                ConectaBanco.FechaBanco();

                /*SQL = "DBCC CHECKIDENT ('Paineis', RESEED, 0)";

                ConectaBanco.ExecutaComando(SQL);
                ConectaBanco.FechaBanco();

                SQL = "DBCC CHECKIDENT ('Inversores', RESEED, 0)";

                ConectaBanco.ExecutaComando(SQL);
                ConectaBanco.FechaBanco();*/
            }
            else if (BD == "mysql")
            {
                string SQL;

                SQL = "DBCC CHECKIDENT ('Clientes', RESEED, 1);";

                conectaMySQL.ExecutaComando(SQL);
                conectaMySQL.FechaMySQL();
            }
            return true;
        }

        //Tabela Inversores
        public void CriaTableInversores()
        {
            string SQL;
            SQL = "CREATE TABLE Inversores (Id INT NOT NULL AUTO_INCREMENT, Marca VARCHAR(255) NOT NULL, " +
                  "Modelo NVARCHAR(255) NOT NULL, Potencia NVARCHAR(255) NOT NULL, Fases NVARCHAR(255) NOT NULL," +
                  "Tensao  NVARCHAR(255) NOT NULL, Eficiencia NVARCHAR(255) NOT NULL, Garantia NVARCHAR(255) NOT NULL," +
                  "Quantidade_MPPT  NVARCHAR(255) NOT NULL, RegistroINMETRO NVARCHAR(255) NOT NULL, PRIMARY KEY (Id));";

            MySqlDataReader dados = conectaMySQL.ExecutaConsulta(SQL);

            conectaMySQL.FechaMySQL();
        }
        public bool InserirInversor(string BD)
        {
            if (BD == "local")
            {
                string SQL;

                SQL = "Insert into Inversores (Marca, Modelo, Potencia, Fases, Tensao,Eficiencia, Garantia, Quantidade_MPPT, RegistroINMETRO) values('" + MarcaInversor + "','" + ModeloInversor + "','" + PotenciaInv + "','" + Fases + "','" + Tensao + "','" + EficienciaInv + "','" + GarantiaInv + "','" + Qtdmppt + "','" + RegistroINMETRO + "')";

                ConectaBanco.ExecutaComando(SQL);
                ConectaBanco.FechaBanco();
            }
            else if (BD == "mysql")
            {
                string SQL;

                SQL = "Insert into Inversores (Marca, Modelo, Potencia, Fases, Tensao,Eficiencia, Garantia, Quantidade_MPPT, RegistroINMETRO) values('" + MarcaInversor + "','" + ModeloInversor + "','" + PotenciaInv + "','" + Fases + "','" + Tensao + "','" + EficienciaInv + "','" + GarantiaInv + "','" + Qtdmppt + "','" + RegistroINMETRO + "');";

                conectaMySQL.ExecutaComando(SQL);
                conectaMySQL.FechaMySQL();
            }
            return true;
        }
        public bool AlterarInversor(string BD)
        {
            if (BD == "local")
            {
                string SQL;

                SQL = "Update Inversores set Marca='" + MarcaInversor + "', Modelo='" + ModeloInversor + "', Potencia='" + PotenciaInv + "', Fases ='" + Fases + "',Tensao='" + Tensao + "'," +
                    "Eficiencia='" + EficienciaInv + "', Garantia='" + GarantiaInv + "'," +
                    " Quantidade_MPPT='" + Qtdmppt + "', RegistroINMETRO='" + RegistroINMETRO + "' where Id =" + id;

                ConectaBanco.ExecutaComando(SQL);
                ConectaBanco.FechaBanco();
            }
            else if (BD == "mysql")
            {
                string SQL;

                SQL = "Update Inversores set Marca='" + MarcaInversor + "', Modelo='" + ModeloInversor + "', Potencia='" + PotenciaInv + "', Fases ='" + Fases + "',Tensao='" + Tensao + "'," +
                    "Eficiencia='" + EficienciaInv + "', Garantia='" + GarantiaInv + "'," +
                    " Quantidade_MPPT='" + Qtdmppt + "', RegistroINMETRO='" + RegistroINMETRO + "' where Id =" + id;

                conectaMySQL.ExecutaComando(SQL);
                conectaMySQL.FechaMySQL();
            }
            return true;
        }
        public void SelecionaInversor(string BD, string modelo)
        {
            if (BD == "local")
            {
                string SQL;
                SQL = "Select * from Inversores where Modelo ='" + modelo + "'";

                SqlDataReader dados = ConectaBanco.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    MarcaInversor = dados["Marca"].ToString();
                    ModeloInversor = dados["Modelo"].ToString();
                    PotenciaInv = dados["Potencia"].ToString();
                    Fases = dados["Fases"].ToString();
                    Tensao = dados["Tensao"].ToString();
                    EficienciaInv = dados["Eficiencia"].ToString();
                    GarantiaInv = dados["Garantia"].ToString();
                    Qtdmppt = dados["Quantidade_MPPT"].ToString();
                    RegistroINMETRO = dados["RegistroINMETRO"].ToString();
                }
                ConectaBanco.FechaBanco();
            }
            else if (BD == "mysql")
            {
                string SQL;
                SQL = "Select * from Inversores where Modelo ='" + modelo + "'";

                MySqlDataReader dados = conectaMySQL.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    MarcaInversor = dados["Marca"].ToString();
                    ModeloInversor = dados["Modelo"].ToString();
                    PotenciaInv = dados["Potencia"].ToString();
                    Fases = dados["Fases"].ToString();
                    Tensao = dados["Tensao"].ToString();
                    EficienciaInv = dados["Eficiencia"].ToString();
                    GarantiaInv = dados["Garantia"].ToString();
                    Qtdmppt = dados["Quantidade_MPPT"].ToString();
                    RegistroINMETRO = dados["RegistroINMETRO"].ToString();
                }
                conectaMySQL.FechaMySQL();
            }
        }
        public void SelecionaInversorModelo(string BD, string modelo)
        {
            if (BD == "local")
            {
                string SQL;
                SQL = "Select * from Inversores where Modelo='" + modelo + "'";

                SqlDataReader dados = ConectaBanco.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    MarcaInversor = dados["Marca"].ToString();
                    ModeloInversor = dados["Modelo"].ToString();
                    PotenciaInv = dados["Potencia"].ToString();
                    Fases = dados["Fases"].ToString();
                    Tensao = dados["Tensao"].ToString();
                    EficienciaInv = dados["Eficiencia"].ToString();
                    GarantiaInv = dados["Garantia"].ToString();
                    Qtdmppt = dados["Quantidade_MPPT"].ToString();
                    RegistroINMETRO = dados["RegistroINMETRO"].ToString();
                }
                ConectaBanco.FechaBanco();
            }
            else if (BD == "mysql")
            {
                string SQL;
                SQL = "Select * from Inversores where Modelo='" + modelo + "'";

                MySqlDataReader dados = conectaMySQL.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    MarcaInversor = dados["Marca"].ToString();
                    ModeloInversor = dados["Modelo"].ToString();
                    PotenciaInv = dados["Potencia"].ToString();
                    Fases = dados["Fases"].ToString();
                    Tensao = dados["Tensao"].ToString();
                    EficienciaInv = dados["Eficiencia"].ToString();
                    GarantiaInv = dados["Garantia"].ToString();
                    Qtdmppt = dados["Quantidade_MPPT"].ToString();
                    RegistroINMETRO = dados["RegistroINMETRO"].ToString();
                }
                conectaMySQL.FechaMySQL();
            }
        }
        public void PesquisaModInv(string aux, string BD)
        {
            if (BD == "local")
            {
                string SQL;
                SQL = "Select * from Inversores where Modelo = '" + aux + "'";

                SqlDataReader dados = ConectaBanco.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    MarcaInversor = dados["Marca"].ToString();
                    ModeloInversor = dados["Modelo"].ToString();
                    PotenciaInv = dados["Potencia"].ToString();
                    Fases = dados["Fases"].ToString();
                    Tensao = dados["Tensao"].ToString();
                    EficienciaInv = dados["Eficiencia"].ToString();
                    GarantiaInv = dados["Garantia"].ToString();
                    Qtdmppt = dados["Quantidade_MPPT"].ToString();
                    RegistroINMETRO = dados["RegistroINMETRO"].ToString();
                }
                ConectaBanco.FechaBanco();
            }
            else if (BD == "mysql")
            {
                string SQL;
                SQL = "Select * from Inversores where Modelo = '" + aux + "'";

                MySqlDataReader dados = conectaMySQL.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    MarcaInversor = dados["Marca"].ToString();
                    ModeloInversor = dados["Modelo"].ToString();
                    PotenciaInv = dados["Potencia"].ToString();
                    Fases = dados["Fases"].ToString();
                    Tensao = dados["Tensao"].ToString();
                    EficienciaInv = dados["Eficiencia"].ToString();
                    GarantiaInv = dados["Garantia"].ToString();
                    Qtdmppt = dados["Quantidade_MPPT"].ToString();
                    RegistroINMETRO = dados["RegistroINMETRO"].ToString();
                }
                conectaMySQL.FechaMySQL();
            }
        }
        public void PesquisaMarcaInv(string aux, string BD)
        {
            if (BD == "local")
            {
                string SQL;
                SQL = "Select Marca from Inversores where Modelo = '" + aux + "'";

                SqlDataReader dados = ConectaBanco.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    MarcaInversor = dados["Marca"].ToString();
                }
                ConectaBanco.FechaBanco();
            }
            else if (BD == "mysql")
            {
                string SQL;
                SQL = "Select Marca from Inversores where Modelo = '" + aux + "'";

                MySqlDataReader dados = conectaMySQL.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    MarcaInversor = dados["Marca"].ToString();
                }
                conectaMySQL.FechaMySQL();
            }
        }
        public void PesquisaPotInv(string aux, string BD)
        {
            if (BD == "local")
            {
                string SQL;
                SQL = "Select * from Inversores where Modelo = '" + aux + "'";

                SqlDataReader dados = ConectaBanco.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    PotenciaInv = dados["Potencia"].ToString();
                    GarantiaInv = dados["Garantia"].ToString();
                }
                ConectaBanco.FechaBanco();
            }
            else if (BD == "mysql")
            {
                string SQL;
                SQL = "Select * from Inversores where Modelo = '" + aux + "'";

                MySqlDataReader dados = conectaMySQL.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    PotenciaInv = dados["Potencia"].ToString();
                    GarantiaInv = dados["Garantia"].ToString();
                }
                conectaMySQL.FechaMySQL();
            }
        }
        public DataTable AtualizaInversor(string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Inversores";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Inversores";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }
        }
        public void ExcluiInversor(string BD, string modelo)
        {
            if (BD == "local")
            {
                ConectaBanco.ExecutaComando("Delete from Inversores where Modelo='" + modelo + "'");
                ConectaBanco.FechaBanco();
            }
            else if (BD == "mysql")
            {
                conectaMySQL.ExecutaComando("Delete from Inversores where Modelo='" + modelo + "'");
                conectaMySQL.FechaMySQL();
            }
        }
        public DataTable MarcaInv1(string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Marca from Inversores Group By Marca";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Marca from Inversores Group By Marca";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }

        }
        public DataTable TensaoInv1(string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Tensao from Inversores Group By Tensao";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Tensao from Inversores Group By Tensao";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }

        }
        public DataTable ModeloInv1(string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Modelo from Inversores Group By Modelo";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Modelo from Inversores Group By Modelo";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }

        }
        public DataTable MarcaInv(string TextoPesquisa, string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Inversores where Marca like '%" + TextoPesquisa + "%' Order By Modelo";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Inversores where Marca like '%" + TextoPesquisa + "%' Order By Modelo";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }
            
        }
        public DataTable ModeloInv(string TextoPesquisa, string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Inversores where Modelo like '%" + TextoPesquisa + "%' Order By Modelo";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Inversores where Modelo like '%" + TextoPesquisa + "%' Order By Modelo";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }
        }
        public DataTable FiltroInv(string TextoPesquisa, string filtro, string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Inversores where " + filtro + " like '%" + TextoPesquisa + "%' Order By Marca";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Inversores where " + filtro + " like '%" + TextoPesquisa + "%' Order By Marca";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }
        }
        public DataTable EficienciaInversor(string TextoPesquisa, string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Inversores where Eficiencia like '%" + TextoPesquisa + "%' Order By Modelo";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Inversores where Eficiencia like '%" + TextoPesquisa + "%' Order By Modelo";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }
        }
        public DataTable FasesInversor(string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Fases from Inversores Group By Fases";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Fases from Inversores Group By Fases";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }
        }
        public DataTable GarantiaInversor(string TextoPesquisa, string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Inversores where Garantia like '%" + TextoPesquisa + "%' Order By Modelo";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Inversores where Garantia like '%" + TextoPesquisa + "%' Order By Modelo";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }
        }
        public DataTable qtdmpptInv(string TextoPesquisa, string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Inversores where Quantidade_MPPT like '%" + TextoPesquisa + "%' Order By Modelo";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Inversores where Quantidade_MPPT like '%" + TextoPesquisa + "%' Order By Modelo";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }
        }
        public DataTable TodosInv(string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Marca from Inversores Group By Marca";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Marca from Inversores Group By Marca";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }
            
        }

        //Tabela Credenciais
        public void CriaTableCredenciais()
        {
            string SQL;
            SQL = "DROP TABLE Credenciais;";

            MySqlDataReader dados = conectaMySQL.ExecutaConsulta(SQL);

            conectaMySQL.FechaMySQL();

            //string SQL;
            SQL = "CREATE TABLE Credenciais (Id INT NOT NULL AUTO_INCREMENT, Login VARCHAR(50) NULL, " +
                  "Senha NVARCHAR(50) NULL, Nome_Completo NVARCHAR(250) NULL, email NVARCHAR(50) NULL, Foto LONGBLOB NULL, PRIMARY KEY (Id));";

            /*MySqlDataReader*/dados = conectaMySQL.ExecutaConsulta(SQL);

            conectaMySQL.FechaMySQL();
        }
        public bool InserirCredenciais(string BD)
        {
            if (BD == "local")
            {
                string SQL;

                SQL = "Insert into Credenciais (Login, Senha, Nome_Completo, email, Foto) values('" + Login + "','" + Senha + "','" + Nome + "','" + email + "','" + Foto + "')";

                ConectaBanco.ExecutaComando(SQL);
                ConectaBanco.FechaBanco();
            }
            else if (BD == "mysql")
            {
                string SQL;

                SQL = "Insert into Credenciais (Login, Senha, Nome_Completo, email, Foto) values('" + Login + "','" + Senha + "','" + Nome + "','" + email + "','" + Foto + "')";

                conectaMySQL.ExecutaComando(SQL);
                conectaMySQL.FechaMySQL();
            }
            return true;
        }
        public bool AlterarCredenciais(string BD)
        {
            if (BD == "local")
            {
                string SQL;

                SQL = "Update Credenciais set Login='" + Login + "', Senha='" + Senha + "' where Id =" + Id;

                ConectaBanco.ExecutaComando(SQL);
                ConectaBanco.FechaBanco();
            }
            else if (BD == "mysql")
            {
                string SQL;

                SQL = "Update Credenciais set Login='" + Login + "', Senha='" + Senha + "' where Id =" + Id;

                conectaMySQL.ExecutaComando(SQL);
                conectaMySQL.FechaMySQL();
            }
            return true;
        }
        public bool PesquisaLogin(string aux, string BD)
        {
            if (BD == "local")
            {
                string SQL;
                SQL = "Select * from Credenciais where Login = '" + aux + "'";

                SqlDataReader dados = ConectaBanco.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    Id = Int32.Parse(dados["Id"].ToString());
                    Login = dados["Login"].ToString();
                    Senha = dados["Senha"].ToString();
                    Nome = dados["Nome_Completo"].ToString();
                    email = dados["email"].ToString();
                    ConectaBanco.FechaBanco();
                    return true;
                }
                else
                {
                    ConectaBanco.FechaBanco();
                    return false;
                }
                /*while (dados.Read())
                {
                    Foto = (byte[])(dados["Foto"]);
                }*/
            }
            else if (BD == "mysql")
            {
                string SQL;
                SQL = "Select * from Credenciais where Login = '" + aux + "'";

                MySqlDataReader dados = conectaMySQL.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    Id = Int32.Parse(dados["Id"].ToString());
                    Login = dados["Login"].ToString();
                    Senha = dados["Senha"].ToString();
                    Nome = dados["Nome_Completo"].ToString();
                    email = dados["email"].ToString();
                    conectaMySQL.FechaMySQL();
                    return true;
                }
                else
                {
                    conectaMySQL.FechaMySQL();
                    return false;
                }
                /*
                while (dados.Read())
                {
                    Foto = (byte[])(dados["Foto"]);
                }*/
            }
            else
            {
                return false;
            }
        }
        public DataTable PesquisaCredenciais(string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Login, Senha from Credenciais";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Login, Senha from Credenciais;";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }
            
        }
        public void Exclui(string BD)
        {
            if (BD == "local")
            {
                ConectaBanco.ExecutaComando("Delete from Credenciais where ID=" + Id);
                ConectaBanco.FechaBanco();
            }
            else if (BD == "mysql")
            {
                conectaMySQL.ExecutaComando("Delete from Credenciais where ID=" + Id);
                conectaMySQL.FechaMySQL();
            }
        }
        public void MostaFoto(string aux,string BD)
        {
            if (BD == "local")
            {
                string SQL;
                SQL = "Select Foto from Credenciais where Login = '" + aux + "'";

                SqlDataReader dados = ConectaBanco.ExecutaConsulta(SQL);
                ConectaBanco.FechaBanco();
            }
            else if (BD == "mysql")
            {
                string SQL;
                SQL = "Select Foto from Credenciais where Login = '" + aux + "'";

                MySqlDataReader dados = conectaMySQL.ExecutaConsulta(SQL);

                while (dados.Read())
                {
                    Foto = (byte[])(dados["Foto"]);
                }
                conectaMySQL.FechaMySQL();
            }
        }

        //Tabela Cidades
        /*public void ClimaBrasil(string aux)
        {
            string SQL;
            SQL = "Select * from Cidades where Cidade = '" + aux + "'";

            SqlDataReader dados = ConectaBanco.ExecutaConsulta(SQL);

            if (dados.Read())
            {
                CliJan = dados["Janeiro"].ToString();
                CliFev = dados["Fevereiro"].ToString();
                CliMar = dados["Marco"].ToString();
                CliAbr = dados["Abril"].ToString();
                CliMai = dados["Maio"].ToString();
                CliJun = dados["Junho"].ToString();
                CliJul = dados["Julho"].ToString();
                CliAgo = dados["Agosto"].ToString();
                CliSet = dados["Setembro"].ToString();
                CliOut = dados["Outubro"].ToString();
                CliNov = dados["Novembro"].ToString();
                CliDez = dados["Dezembro"].ToString();
            }
        }*/
        public void PesquisaMediaTemp(string aux, string BD)
        {

            string SQL;
            SQL = "Select Media from Cidades where Cidade = '" + aux + "'";

            SqlDataReader dados = ConectaBanco.ExecutaConsulta(SQL);

            if (dados.Read())
            {
                Media = dados["Media"].ToString();
            }
            ConectaBanco.FechaBanco();
        }

        //Tabela Paineis
        public void CriaTablePainel()
        {
            string SQL;
            SQL = "CREATE TABLE Paineis (Id INT NOT NULL AUTO_INCREMENT, Marca VARCHAR(255) NOT NULL, " +
                  "Modelo NVARCHAR(255) NOT NULL, Potencia DOUBLE NOT NULL, Material NVARCHAR(255) NOT NULL," +
                  "Celulas  NVARCHAR(255) NOT NULL, Coeficiente_Temperatura DOUBLE NOT NULL, Comprimento DOUBLE NOT NULL," +
                  "Largura DOUBLE NOT NULL, Garantia INT NOT NULL, RegistroInmetro  NVARCHAR(255) NOT NULL, PRIMARY KEY (Id));";

            MySqlDataReader dados = conectaMySQL.ExecutaConsulta(SQL);

            conectaMySQL.FechaMySQL();
        }
        public bool InserirPainel(string BD)
        {
            if (BD == "local")
            {
                string SQL;

                SQL = "Insert into Paineis (Marca, Modelo, Potencia, Material, Celulas, Coeficiente_Temperatura, Comprimento, Largura, Garantia, RegistroInmetro) values('" + MarcaMod + "','" + ModeloModulo + "'," +
                    "'" + Convert.ToDouble(PotenciaMod) + "','" + Material + "','" + Celulas + "','" + Convert.ToDouble(TemperaturaModulo) + "','" + Convert.ToDouble(ComrpimentoMod) + "','" + Convert.ToDouble(LarguraMod) + "', '" + Convert.ToInt32(GarantiaMod) + "','" + RegistroInmetro + "')";

                ConectaBanco.ExecutaComando(SQL);
                ConectaBanco.FechaBanco();
            }
            else if (BD == "mysql")
            {
                string SQL;

                SQL = "Insert into Paineis (Marca, Modelo, Potencia, Material, Celulas, Coeficiente_Temperatura, Comprimento, Largura, Garantia, RegistroInmetro) values('" + MarcaMod + "','" + ModeloModulo + "'," +
                    "'" + Convert.ToDouble(PotenciaMod) + "','" + Material + "','" + Celulas + "','" + Convert.ToDouble(TemperaturaModulo) + "','" + Convert.ToDouble(ComrpimentoMod) + "','" + Convert.ToDouble(LarguraMod) + "', '" + Convert.ToInt32(GarantiaMod) + "','" + RegistroInmetro + "');";

                conectaMySQL.ExecutaComando(SQL);
                conectaMySQL.FechaMySQL();
            }
            return true;
        }
        public void PesquisaPotMod(string aux, string BD)
        {
            if (BD == "local")
            {
                string SQL;
                SQL = "Select Potencia from Paineis where Modelo = '" + aux + "'";

                SqlDataReader dados = ConectaBanco.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    PotenciaMod = dados["Potencia"].ToString();
                }
                ConectaBanco.FechaBanco();
            }
            else if (BD == "mysql")
            {
                string SQL;
                SQL = "Select Potencia from Paineis where Modelo = '" + aux + "'";

                MySqlDataReader dados = conectaMySQL.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    PotenciaMod = dados["Potencia"].ToString();
                }
                conectaMySQL.FechaMySQL();
            }
        }
        public void PesquisaModMod(string aux, string BD)
        {
            if (BD == "local")
            {
                string SQL;
                SQL = "Select * from Paineis where Modelo = '" + aux + "'";

                SqlDataReader dados = ConectaBanco.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    MarcaMod = dados["Marca"].ToString();
                    ModeloModulo = dados["Modelo"].ToString();
                    PotenciaMod = dados["Potencia"].ToString();
                    Material = dados["Material"].ToString();
                    Celulas = dados["Celulas"].ToString();
                    TemperaturaModulo = dados["Coeficiente_Temperatura"].ToString();
                    ComrpimentoMod = dados["Comprimento"].ToString();
                    LarguraMod = dados["Largura"].ToString();
                    GarantiaMod = dados["Garantia"].ToString();
                    RegistroInmetro = dados["RegistroInmetro"].ToString();
                }
                ConectaBanco.FechaBanco();
            }
            else if (BD == "mysql")
            {
                string SQL;
                SQL = "Select * from Paineis where Modelo = '" + aux + "'";

                MySqlDataReader dados = conectaMySQL.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    MarcaMod = dados["Marca"].ToString();
                    ModeloModulo = dados["Modelo"].ToString();
                    PotenciaMod = dados["Potencia"].ToString();
                    Material = dados["Material"].ToString();
                    Celulas = dados["Celulas"].ToString();
                    TemperaturaModulo = dados["Coeficiente_Temperatura"].ToString();
                    ComrpimentoMod = dados["Comprimento"].ToString();
                    LarguraMod = dados["Largura"].ToString();
                    GarantiaMod = dados["Garantia"].ToString();
                    RegistroInmetro = dados["RegistroInmetro"].ToString();
                }
                conectaMySQL.FechaMySQL();
            }
        }
        public void PesquisaMarcaMod(string aux, string BD)
        {
            if (BD == "local")
            {
                string SQL;
                SQL = "Select Marca from Paineis where Modelo = '" + aux + "'";

                SqlDataReader dados = ConectaBanco.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    MarcaMod = dados["Marca"].ToString();
                }
                ConectaBanco.FechaBanco();
            }
            else if (BD == "mysql")
            {
                string SQL;
                SQL = "Select Marca from Paineis where Modelo = '" + aux + "'";

                MySqlDataReader dados = conectaMySQL.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    MarcaMod = dados["Marca"].ToString();
                }
                conectaMySQL.FechaMySQL();
            }           
        }
        public void PesquisaDimensoesMod(string aux, string BD)
        {
            if (BD == "local")
            {
                string SQL;
                SQL = "Select * from Paineis where Modelo = '" + aux + "'";

                SqlDataReader dados = ConectaBanco.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    ComrpimentoMod = dados["Comprimento"].ToString();
                    LarguraMod = dados["Largura"].ToString();
                    GarantiaMod = dados["Garantia"].ToString();
                }
                ConectaBanco.FechaBanco();
            }
            else if (BD == "mysql")
            {
                string SQL;
                SQL = "Select * from Paineis where Modelo = '" + aux + "'";

                MySqlDataReader dados = conectaMySQL.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    ComrpimentoMod = dados["Comprimento"].ToString();
                    LarguraMod = dados["Largura"].ToString();
                    GarantiaMod = dados["Garantia"].ToString();
                }
                conectaMySQL.FechaMySQL();
            }
        }
        public bool AlterarPaineis(string BD)
        {
            if (BD == "local")
            {
                string SQL;

                SQL = "Update Paineis set Marca='" + MarcaMod + "', Modelo='" + ModeloModulo + "', Potencia='" + PotenciaMod + "', Material='" + Material + "', Celulas='" + Celulas + "', Coeficiente_Temperatura='" + TemperaturaModulo + "'," +
                    " Comprimento='" + ComrpimentoMod + "', Largura='" + LarguraMod + "', Garantia='" + GarantiaMod + "', RegistroInmetro='" + RegistroInmetro + "' where Id =" + id;

                ConectaBanco.ExecutaComando(SQL);
                ConectaBanco.FechaBanco();
            }
            else if (BD == "mysql")
            {
                string SQL;

                SQL = "Update Paineis set Marca='" + MarcaMod + "', Modelo='" + ModeloModulo + "', Potencia='" + PotenciaMod + "', Material='" + Material + "', Celulas='" + Celulas + "', Coeficiente_Temperatura='" + TemperaturaModulo + "'," +
                    " Comprimento='" + ComrpimentoMod + "', Largura='" + LarguraMod + "', Garantia='" + GarantiaMod + "', RegistroInmetro='" + RegistroInmetro + "' where Id =" + id;

                conectaMySQL.ExecutaComando(SQL);
                conectaMySQL.FechaMySQL();
            }
            return true;
        }
        public void ExcluiPainel(string BD, string modelo)
        {
            if (BD == "local")
            {
                ConectaBanco.ExecutaComando("Delete from Paineis where Modelo ='" + modelo + "'");
                ConectaBanco.FechaBanco();
            }
            else if (BD == "mysql")
            {
                conectaMySQL.ExecutaComando("Delete from Paineis where Modelo ='" + modelo + "'");
                conectaMySQL.FechaMySQL();
            }            
        }
        public void SelecionaPainel(string BD, string modelo)
        {
            if (BD == "local")
            {
                string SQL;
                SQL = "Select * from Paineis where Modelo ='" + modelo + "'";

                SqlDataReader dados = ConectaBanco.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    MarcaMod = dados["Marca"].ToString();
                    ModeloModulo = dados["Modelo"].ToString();
                    PotenciaMod = dados["Potencia"].ToString();
                    Material = dados["Material"].ToString();
                    Celulas = dados["Celulas"].ToString();
                    TemperaturaModulo = dados["Coeficiente_Temperatura"].ToString();
                    ComrpimentoMod = dados["Comprimento"].ToString();
                    LarguraMod = dados["Largura"].ToString();
                    GarantiaMod = dados["Garantia"].ToString();
                    RegistroInmetro = dados["RegistroInmetro"].ToString();
                }
                ConectaBanco.FechaBanco();
            }
            else if (BD == "mysql")
            {
                string SQL;
                SQL = "Select * from Paineis where Modelo ='" + modelo + "'";

                MySqlDataReader dados = conectaMySQL.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    MarcaMod = dados["Marca"].ToString();
                    ModeloModulo = dados["Modelo"].ToString();
                    PotenciaMod = dados["Potencia"].ToString();
                    Material = dados["Material"].ToString();
                    Celulas = dados["Celulas"].ToString();
                    TemperaturaModulo = dados["Coeficiente_Temperatura"].ToString();
                    ComrpimentoMod = dados["Comprimento"].ToString();
                    LarguraMod = dados["Largura"].ToString();
                    GarantiaMod = dados["Garantia"].ToString();
                    RegistroInmetro = dados["RegistroInmetro"].ToString();
                }
                conectaMySQL.FechaMySQL();
            }
        }
        public void SelecionaPainelModelo(string BD, string modelo)
        {
            if (BD == "local")
            {
                string SQL;
                SQL = "Select * from Paineis where Modelo= '" + modelo + "'";

                SqlDataReader dados = ConectaBanco.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    MarcaMod = dados["Marca"].ToString();
                    ModeloModulo = dados["Modelo"].ToString();
                    PotenciaMod = dados["Potencia"].ToString();
                    Material = dados["Material"].ToString();
                    Celulas = dados["Celulas"].ToString();
                    TemperaturaModulo = dados["Coeficiente_Temperatura"].ToString();
                    ComrpimentoMod = dados["Comprimento"].ToString();
                    LarguraMod = dados["Largura"].ToString();
                    GarantiaMod = dados["Garantia"].ToString();
                    RegistroInmetro = dados["RegistroInmetro"].ToString();
                }
                ConectaBanco.FechaBanco();
            }
            else if (BD == "mysql")
            {
                string SQL;
                SQL = "Select * from Paineis where Modelo='" + modelo + "'";

                MySqlDataReader dados = conectaMySQL.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    MarcaMod = dados["Marca"].ToString();
                    ModeloModulo = dados["Modelo"].ToString();
                    PotenciaMod = dados["Potencia"].ToString();
                    Material = dados["Material"].ToString();
                    Celulas = dados["Celulas"].ToString();
                    TemperaturaModulo = dados["Coeficiente_Temperatura"].ToString();
                    ComrpimentoMod = dados["Comprimento"].ToString();
                    LarguraMod = dados["Largura"].ToString();
                    GarantiaMod = dados["Garantia"].ToString();
                    RegistroInmetro = dados["RegistroInmetro"].ToString();
                }
                conectaMySQL.FechaMySQL();
            }
        }
        public DataTable AtualizaPaineis(string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Paineis";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Paineis";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }
        }
        public DataTable FiltroMod(string TextoPesquisa, string filtro, string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Paineis where " + filtro + " like '%" + TextoPesquisa + "%' Order By Marca";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Paineis where " + filtro + " like '%" + TextoPesquisa + "%' Order By Marca";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }
            
        }
        public DataTable MarcaModuloOrc(string TextoPesquisa, string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Paineis where Marca like '%" + TextoPesquisa + "%' Order By Modelo";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Paineis where Marca like '%" + TextoPesquisa + "%' Order By Modelo";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }

        }
        public DataTable MarcaModulo(string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Marca from Paineis Group By Marca";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Marca from Paineis Group By Marca";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }
            
        }
        public DataTable ModeloMod(string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Modelo from Paineis Group By Modelo";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Modelo from Paineis Group By Modelo";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }

        }
        public DataTable CelulasModulo(string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Celulas from Paineis Group By Celulas";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Celulas from Paineis Group By Celulas";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }

        }
        public DataTable PotenciaModulo(string TextoPesquisa, string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Paineis where Potencia like '%" + TextoPesquisa + "%'";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Paineis where Potencia like '%" + TextoPesquisa + "%'";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }
            
        }
        public DataTable TemperaturaMod(string TextoPesquisa, string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Paineis where Coeficiente_Temperatura like '%" + TextoPesquisa + "%'";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();

                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Paineis where Coeficiente_Temperatura like '%" + TextoPesquisa + "%'";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();

                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }
            
        }
        public DataTable ComprimentoMod(string TextoPesquisa, string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Paineis where Comprimento like '%" + TextoPesquisa + "%'";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Paineis where Comprimento like '%" + TextoPesquisa + "%'";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }
            
        }
        public DataTable MaterialMod(string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Material from Paineis Group By Material";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Material from Paineis Group By Material";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }

        }
        public DataTable GarantiaModulo(string TextoPesquisa, string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Paineis where Garantia like '%" + TextoPesquisa + "%'";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Paineis where Garantia like '%" + TextoPesquisa + "%'";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }
        }
        public DataTable LarguraModulo(string TextoPesquisa, string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Paineis where Largura like '%" + TextoPesquisa + "%'";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Paineis where Largura like '%" + TextoPesquisa + "%'";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }
        }
        public DataTable TodosMod(string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Marca from Paineis Group By Marca";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Marca from Paineis Group By Marca";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }
            
        }

        //Tabela Padroes
        public void PesquisaCustoDisp(string aux, string BD)
        {

            string SQL;
            SQL = "Select Custo from Padroes where Fase = '" + aux + "'";

            SqlDataReader dados = ConectaBanco.ExecutaConsulta(SQL);

            if (dados.Read())
            {
                Custo = dados["Custo"].ToString();
            }
            ConectaBanco.FechaBanco();
        }

        //Tabela Temperatura
        public void PesquisaJaneiro(string aux, string BD)
        {

            string SQL;
            SQL = "Select * from Temperatura where Cidade = '" + aux + "'";

            SqlDataReader dados = ConectaBanco.ExecutaConsulta(SQL);

            if (dados.Read())
            {
                Janeiro = dados["Janeiro"].ToString();
                Fevereiro = dados["Fevereiro"].ToString();
                Marco = dados["Marco"].ToString();
                Abril = dados["Abril"].ToString();
                Maio = dados["Maio"].ToString();
                Junho = dados["Junho"].ToString();
                Julho = dados["Julho"].ToString();
                Agosto = dados["Agosto"].ToString();
                Setembro = dados["Setembro"].ToString();
                Outubro = dados["Outubro"].ToString();
                Novembro = dados["Novembro"].ToString();
                Dezembro = dados["Dezembro"].ToString();
            }
            ConectaBanco.FechaBanco();
        }

        //Orçamento
        public void CriaTableOrcamento()
        {
            string SQL;
            /*SQL = "CREATE TABLE Orcamento (Id INT NOT NULL AUTO_INCREMENT, Nome VARCHAR(255) NOT NULL, " +
                  "Contato NVARCHAR(255) NOT NULL, Cep NVARCHAR(255) NOT NULL, Endereco NVARCHAR(255) NOT NULL," +
                  "Numero  NVARCHAR(255) NOT NULL, Bairro NVARCHAR(255) NOT NULL, Cidade NVARCHAR(255) NOT NULL," +
                  "Tarifa  NVARCHAR(255) NOT NULL, Consumo_Anual NVARCHAR(255) NOT NULL, Estrutura  NVARCHAR(255) NOT NULL," +
                  "Quantidade_Paineis NVARCHAR(255) NOT NULL, Marca_Paineis NVARCHAR(255) NOT NULL, Modelo_Paineis " +
                  "NVARCHAR(255) NOT NULL, Quantidade_Inversores NVARCHAR(255) NOT NULL, Marca_Inversores " +
                  "NVARCHAR(255) NOT NULL, Modelo_Inversores NVARCHAR(255) NOT NULL, Valor_Orcamento NVARCHAR(255) NOT NULL," +
                  "Valor_Equipamentos NVARCHAR(255) NOT NULL, Valor_Inversor NVARCHAR(255) NOT NULL, Perdas  NVARCHAR(255) NOT NULL," +
                  "Obs TEXT NOT NULL, tjan INT NOT NULL, tfev INT NOT NULL, tmar INT NOT NULL, tabr INT NOT NULL," +
                  "tmai INT NOT NULL, tjun INT NOT NULL, tjul INT NOT NULL, tago INT NOT NULL, tset INT NOT NULL," +
                  "tout INT NOT NULL, tnov INT NOT NULL, tdez INT NOT NULL, somadisp INT NOT NULL, PRIMARY KEY (Id));";*/
            SQL = "CREATE TABLE Orcamento (Id INT IDENTITY (1, 1) NOT NULL, Nome VARCHAR(255) NOT NULL, " +
            "Contato NVARCHAR(255) NOT NULL, Cep NVARCHAR(255) NOT NULL, Endereco NVARCHAR(255) NOT NULL," +
            "Numero  NVARCHAR(255) NOT NULL, Bairro NVARCHAR(255) NOT NULL, Cidade NVARCHAR(255) NOT NULL," +
            "Tarifa  NVARCHAR(255) NOT NULL, Consumo_Anual NVARCHAR(255) NOT NULL, Estrutura  NVARCHAR(255) NOT NULL," +
            "Quantidade_Paineis NVARCHAR(255) NOT NULL, Marca_Paineis NVARCHAR(255) NOT NULL, Modelo_Paineis " +
            "NVARCHAR(255) NOT NULL, Quantidade_Inversores NVARCHAR(255) NOT NULL, Marca_Inversores " +
            "NVARCHAR(255) NOT NULL, Modelo_Inversores NVARCHAR(255) NOT NULL, Valor_Orcamento NVARCHAR(255) NOT NULL," +
            "Valor_Equipamentos NVARCHAR(255) NOT NULL, Valor_Inversor NVARCHAR(255) NOT NULL, Perdas  NVARCHAR(255) NOT NULL," +
            "Obs TEXT NOT NULL, tjan INT NOT NULL, tfev INT NOT NULL, tmar INT NOT NULL, tabr INT NOT NULL," +
            "tmai INT NOT NULL, tjun INT NOT NULL, tjul INT NOT NULL, tago INT NOT NULL, tset INT NOT NULL," +
            "tout INT NOT NULL, tnov INT NOT NULL, tdez INT NOT NULL, somadisp INT NOT NULL, PRIMARY KEY CLUSTERED ([Id] ASC));";

            //MySqlDataReader dados = conectaMySQL.ExecutaConsulta(SQL);
            SqlDataReader dados = ConectaBanco.ExecutaConsulta(SQL);
            ConectaBanco.FechaBanco();
            //conectaMySQL.FechaMySQL();
        }
        public bool Salvaorc(string BD)
        {
            if(BD == "local")
            {
                string SQL;

                SQL = "Insert into Orcamento (nome, contato, cep, endereco, numero, bairro, cidade, Tarifa, Consumo_Anual, " +
                    "estrutura, Quantidade_Paineis, Marca_Paineis, Modelo_Paineis, Quantidade_Inversores, Marca_Inversores, Modelo_Inversores, "+
                    " Valor_Orcamento, Valor_Equipamentos, Valor_Inversor, Perdas, Obs, tjan, tfev, tmar, tabr, tmai, tjun, tjul, tago, tset, tout, tnov, tdez, somadisp) " +
                    "values('" + Nome + "','" + Contato + "','" + CEP + "','" + Endereco + "','" + Numero + "', " +
                    "'" + Bairro + "','" + Cidade + "','" + Custo + "','" + Consumoanual + "','" + Estu + "','" + QuantidadeModulos + "',"+
                    "'" + MarcaMod + "','" + ModeloModulo + "','" + QuantidadeInversores + "','" + MarcaInversor + "',"+
                    "'" + ModeloInversor + "','" + Valorsist + "','" + Valorequip + "','" + Valorinv + "','" + Perdas + "','" + Obs + "','" + CliJan + "', " +
                    "'" + CliFev + "','" + CliMar + "','" + CliAbr + "','" + CliMai + "','" + CliJun + "','" + CliJul + "','" + CliAgo + "','" + CliSet + "', " +
                    "'" + CliOut + "','" + CliNov + "','" + CliDez + "','" + Disponibilidade + "')";
            
                ConectaBanco.ExecutaComando(SQL);
                ConectaBanco.FechaBanco();

            }
            else if (BD == "mysql")
            {
                string SQL;

                SQL = "Insert into Orcamento (nome, contato, cep, endereco, numero, bairro, cidade, Tarifa, Consumo_Anual, " +
                "estrutura, Quantidade_Paineis, Marca_Paineis, Modelo_Paineis, Quantidade_Inversores, Marca_Inversores, Modelo_Inversores, " +
                " Valor_Orcamento, Valor_Equipamentos, Valor_Inversor, Perdas, Obs, tjan, tfev, tmar, tabr, tmai, tjun, tjul, tago, tset, tout, tnov, tdez, somadisp) " +
                "values('" + Nome + "','" + Contato + "','" + CEP + "','" + Endereco + "','" + Numero + "', " +
                "'" + Bairro + "','" + Cidade + "','" + Custo + "','" + Consumoanual + "','" + Estu + "','" + QuantidadeModulos + "'," +
                "'" + MarcaMod + "','" + ModeloModulo + "','" + QuantidadeInversores + "','" + MarcaInversor + "'," +
                "'" + ModeloInversor + "','" + Valorsist + "','" + Valorequip + "','" + Valorinv + "','" + Perdas + "','" + Obs + "','" + CliJan + "', " +
                "'" + CliFev + "','" + CliMar + "','" + CliAbr + "','" + CliMai + "','" + CliJun + "','" + CliJul + "','" + CliAgo + "','" + CliSet + "', " +
                "'" + CliOut + "','" + CliNov + "','" + CliDez + "','" + Disponibilidade + "');";

                conectaMySQL.ExecutaComando(SQL);
                conectaMySQL.FechaMySQL();
            }
            return true;
        }
        public bool Alterarorc(string BD)
        {
            if (BD == "local")
            {
                string SQL;

                SQL = "Update Orcamento set Nome ='" + Nome + "', Contato ='" + Contato + "', Cep ='" + CEP + "', Endereco ='" + Endereco + "', Numero ='" + Numero + "'," +
                    " Bairro ='" + Bairro + "', Cidade ='" + Cidade + "', Cusokwh ='" + Custo + "', Consumoanual ='" + Consumoanual + "'," +
                    " Estrutura ='" + Estu + "', Quantidade_Paineis ='" + QuantidadeModulos + "', Marca_Paineis ='" + MarcaMod + "', Modelo_Paineis ='" + ModeloModulo + "', " +
                    " Quantidade_Inversores ='" + QuantidadeInversores + "', Marca_Inversores ='" + MarcaInversor + "', Modelo_Inversores ='" + ModeloInversor + "', " +
                    " Valor_Orcamento ='" + Valorsist + "', Valor_Equipamentos ='" + Valorequip + "', Valor_Inversor='" + Valorinv + "', Perdas = '" + Perdas + "', Obs = '" + Obs + "', " +
                    " tjan = '" + CliJan + "', tfev ='" + CliFev + "', tmar ='" + CliMar+ "', tabr ='" + CliAbr + "', tmai ='" + CliMai + "', tjun ='" + CliJun + "', "+
                    " tjul ='" + CliJul + "', tago ='" + CliAgo + "', tset ='" + CliSet + "', tout ='" + CliOut + "', tnov ='" + CliNov + "', tdez ='" + CliDez + "' where Id =" + Id;

                ConectaBanco.ExecutaComando(SQL);
                ConectaBanco.FechaBanco();
            }
            else if (BD == "mysql")
            {
                string SQL;

                SQL = "Update Orcamento set Nome ='" + Nome + "', Contato ='" + Contato + "', Cep ='" + CEP + "', Endereco ='" + Endereco + "', Numero ='" + Numero + "'," +
                " Bairro ='" + Bairro + "', Cidade ='" + Cidade + "', Cusokwh ='" + Custo + "', Consumoanual ='" + Consumoanual + "'," +
                " Estrutura ='" + Estu + "', Quantidade_Paineis ='" + QuantidadeModulos + "', Marca_Paineis ='" + MarcaMod + "', Modelo_Paineis ='" + ModeloModulo + "', " +
                " Quantidade_Inversores ='" + QuantidadeInversores + "', Marca_Inversores ='" + MarcaInversor + "', Modelo_Inversores ='" + ModeloInversor + "', " +
                " Valor_Orcamento ='" + Valorsist + "', Valor_Equipamentos ='" + Valorequip + "', Valor_Inversor='" + Valorinv + "', Perdas = '" + Perdas + "', Obs = '" + Obs + "', " +
                " tjan = '" + CliJan + "', tfev ='" + CliFev + "', tmar ='" + CliMar + "', tabr ='" + CliAbr + "', tmai ='" + CliMai + "', tjun ='" + CliJun + "', " +
                " tjul ='" + CliJul + "', tago ='" + CliAgo + "', tset ='" + CliSet + "', tout ='" + CliOut + "', tnov ='" + CliNov + "', tdez ='" + CliDez + "' where Id =" + Id + ";";

                conectaMySQL.ExecutaComando(SQL);
                conectaMySQL.FechaMySQL();
            }
            
            return true;
        }    
        public void SelecionaProposta(string BD)
        {
            if (BD == "local")
            {
                string SQL;
                SQL = "Select Id from Orcamento";

                SqlDataReader dados = ConectaBanco.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    Id = Int16.Parse(dados["Id"].ToString());
                }
                ConectaBanco.FechaBanco();
            }
            else if (BD == "mysql")
            {
                string SQL;
                SQL = "Select Id from Orcamento";

                MySqlDataReader dados = conectaMySQL.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    Id = Int16.Parse(dados["Id"].ToString());
                }
                conectaMySQL.FechaMySQL();
            }
            
        }
        public void SelecionaOrcamento(string BD)
        {
            if (BD == "local")
            {
                string SQL;
                SQL = "Select * from Orcamento where Id=" + Id;

                SqlDataReader dados = ConectaBanco.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    Nome = dados["Nome"].ToString();
                    Contato = dados["Contato"].ToString();
                    CEP = dados["CEP"].ToString();
                    Endereco = dados["Endereco"].ToString();
                    Numero = dados["Numero"].ToString();
                    Bairro = dados["Bairro"].ToString();
                    Cidade = dados["Cidade"].ToString();
                    Custo = dados["Tarifa"].ToString();
                    Consumoanual = dados["Consumo_Anual"].ToString();
                    Estu = dados["Estrutura"].ToString();
                    QuantidadeModulos = dados["Quantidade_Paineis"].ToString();
                    MarcaMod = dados["Marca_Paineis"].ToString();
                    ModeloModulo = dados["Modelo_Paineis"].ToString();
                    QuantidadeInversores = dados["Quantidade_Inversores"].ToString();
                    MarcaInversor  = dados["Marca_Inversores"].ToString();
                    ModeloInversor = dados["Modelo_Inversores"].ToString();
                    Valorsist = dados["Valor_Orcamento"].ToString();
                    Valorequip = dados["Valor_Equipamentos"].ToString();
                    Valorinv = dados["Valor_Inversor"].ToString();
                    Perdas = dados["Perdas"].ToString();
                    Obs = dados["Obs"].ToString();

                    CliJan = Int32.Parse(dados["tjan"].ToString());
                    CliFev = Int32.Parse(dados["tfev"].ToString());
                    CliMar = Int32.Parse(dados["tmar"].ToString());
                    CliAbr = Int32.Parse(dados["tabr"].ToString());
                    CliMai = Int32.Parse(dados["tmai"].ToString());
                    CliJun = Int32.Parse(dados["tjun"].ToString());
                    CliJul = Int32.Parse(dados["tjul"].ToString());
                    CliAgo = Int32.Parse(dados["tago"].ToString());
                    CliSet = Int32.Parse(dados["tset"].ToString());
                    CliOut = Int32.Parse(dados["tout"].ToString());
                    CliNov = Int32.Parse(dados["tnov"].ToString());
                    CliDez = Int32.Parse(dados["tdez"].ToString());
                    Disponibilidade = Int32.Parse(dados["somadisp"].ToString());
                }
                ConectaBanco.FechaBanco();
            }
            else if (BD == "mysql")
            {
                string SQL;
                SQL = "Select * from Orcamento where Id=" + Id;

                MySqlDataReader dados = conectaMySQL.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    Nome = dados["Nome"].ToString();
                    Contato = dados["Contato"].ToString();
                    CEP = dados["CEP"].ToString();
                    Endereco = dados["Endereco"].ToString();
                    Numero = dados["Numero"].ToString();
                    Bairro = dados["Bairro"].ToString();
                    Cidade = dados["Cidade"].ToString();
                    Custo = dados["Tarifa"].ToString();
                    Consumoanual = dados["Consumo_Anual"].ToString();
                    Estu = dados["Estrutura"].ToString();
                    QuantidadeModulos = dados["Quantidade_Paineis"].ToString();
                    MarcaMod = dados["Marca_Paineis"].ToString();
                    ModeloModulo = dados["Modelo_Paineis"].ToString();
                    QuantidadeInversores = dados["Quantidade_Inversores"].ToString();
                    MarcaInversor = dados["Marca_Inversores"].ToString();
                    ModeloInversor = dados["Modelo_Inversores"].ToString();
                    Valorsist = dados["Valor_Orcamento"].ToString();
                    Valorequip = dados["Valor_Equipamentos"].ToString();
                    Valorinv = dados["Valor_Inversor"].ToString();
                    Perdas = dados["Perdas"].ToString();
                    Obs = dados["Obs"].ToString();

                    CliJan = Int32.Parse(dados["tjan"].ToString());
                    CliFev = Int32.Parse(dados["tfev"].ToString());
                    CliMar = Int32.Parse(dados["tmar"].ToString());
                    CliAbr = Int32.Parse(dados["tabr"].ToString());
                    CliMai = Int32.Parse(dados["tmai"].ToString());
                    CliJun = Int32.Parse(dados["tjun"].ToString());
                    CliJul = Int32.Parse(dados["tjul"].ToString());
                    CliAgo = Int32.Parse(dados["tago"].ToString());
                    CliSet = Int32.Parse(dados["tset"].ToString());
                    CliOut = Int32.Parse(dados["tout"].ToString());
                    CliNov = Int32.Parse(dados["tnov"].ToString());
                    CliDez = Int32.Parse(dados["tdez"].ToString());
                    Disponibilidade = Int32.Parse(dados["somadisp"].ToString());
                }
                conectaMySQL.FechaMySQL();
            }
            
        }
        public DataTable CarregaOrc(string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Id, Nome, Contato, CEP, Endereco, Numero, Bairro, Cidade, Tarifa, Consumo_Anual, Estrutura, " +
                    " Quantidade_Paineis, Marca_Paineis, Modelo_Paineis, Quantidade_Inversores, Marca_Inversores, Modelo_Inversores, " +
                    " Valor_Orcamento, Valor_Equipamentos, Valor_Inversor, Perdas, Obs from Orcamento";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();
                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Id, Nome, Contato, CEP, Endereco, Numero, Bairro, Cidade, Tarifa, Consumo_Anual, Estrutura, " +
                    " Quantidade_Paineis, Marca_Paineis, Modelo_Paineis, Quantidade_Inversores, Marca_Inversores, Modelo_Inversores, " +
                    " Valor_Orcamento, Valor_Equipamentos, Valor_Inversor, Perdas, Obs from Orcamento";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }   
        }
        public DataTable PreencheMarcaModOrc(string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Marca_Paineis from Orcamento where Id =" + Id;

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Marca_Paineis from Orcamento where Id =" + Id;

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }
            
        }
        public DataTable PreencheModModOrc(string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Modelo_Paineis from Orcamento where Id =" + Id;

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Modelo_Paineis from Orcamento where Id =" + Id;

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;

            }
            else
            {
                return null;
            }

        }
        public DataTable PreencheMarcaInvOrc(string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Marca_Inversores from Orcamento where Id =" + Id;

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Marca_Inversores from Orcamento where Id =" + Id;

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }

        }
        public DataTable PreencheModIUnvOrc(string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Modelo_Inversores from Orcamento where Id =" + Id;

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Modelo_Inversores from Orcamento where Id =" + Id;

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }

        }
        public DataTable PesquisaOrc(string TextoPesquisa, string coluna, string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Orcamento where " + coluna + " like '%" + TextoPesquisa + "%'";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Orcamento where " + coluna + " like '%" + TextoPesquisa + "%'";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }

        }

        //
        //
        //MySQL
        //
        //
        public void DropTable()
        {
            string SQL;
            SQL = "DROP TABLE Orcamento;";

            MySqlDataReader dados = conectaMySQL.ExecutaConsulta(SQL);

            conectaMySQL.FechaMySQL();
        }
        public void Excl()
        {
            string SQL;

            SQL = "Delete from Orcamento where Id=" + Id + ";";

            conectaMySQL.ExecutaComando(SQL);
            conectaMySQL.FechaMySQL();
        }
        public DataTable Verifica()
        {
            MySqlCommand cmd = new MySqlCommand();
            cmd.Connection = conectaMySQL.AbreMySQL();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "Select * From Orcamento";

            MySqlDataAdapter adaptador = new MySqlDataAdapter();
            adaptador.SelectCommand = cmd;

            DataTable dt = new DataTable();
            adaptador.Fill(dt);
            conectaMySQL.FechaMySQL();

            return dt;
        }

        //Escolher o Banco
        public void SelecionaBanco()
        {
            string SQL;
            SQL = "Select * from Banco";

            SqlDataReader dados = ConectaBanco.ExecutaConsulta(SQL);

            if (dados.Read())
            {
                Banco = dados["Banco"].ToString();
                Servidor = dados["Servidor"].ToString();
                NomeDB = dados["NomeBD"].ToString();
                UID = dados["UID"].ToString();
                Password = dados["Senha"].ToString();
    }
        }
        public void AlterarBanco(string Banco)
        {
            string SQL;
            SQL = "Update Banco set Banco ='" + Banco + "'";
            ConectaBanco.ExecutaComando(SQL);
        }
        public void EscolheBanco()
        {
            string SQL;
            SQL = "Select Banco from Banco";

            SqlDataReader dados = ConectaBanco.ExecutaConsulta(SQL);

            if (dados.Read())
            {
                    Banco = dados["Banco"].ToString();
            }
        }
        public void AlterarBancoCaminho(string caminho)
        {
            string SQL;
            SQL = "Update Banco set Caminho ='" + caminho + "'";
            ConectaBanco.ExecutaComando(SQL);
        }

        //Proposta
        public void SelecionaProposta1(string BD)
        {
            if (BD == "local")
            {
                string SQL;
                SQL = "Select Proposta from Proposta where Id=" + 1;

                SqlDataReader dados = ConectaBanco.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    Proposta = Int16.Parse(dados["Proposta"].ToString());
                }
                ConectaBanco.FechaBanco();
            }
            else if(BD == "mysql")
            {
                string SQL;
                SQL = "Select Proposta from Proposta where Id=" + 1;

                MySqlDataReader dados = conectaMySQL.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    Proposta = Int16.Parse(dados["Proposta"].ToString());
                }
                conectaMySQL.FechaMySQL();
            }
        }
        public bool Maisumaproposta1(string BD)
        {
            if (BD == "local")
            {
                string SQL;

                SQL = "Update Proposta set Proposta='" + Proposta + "' where Id =" + 1;

                ConectaBanco.ExecutaComando(SQL);
                ConectaBanco.FechaBanco();
            }
            else if(BD == "mysql")
            {
                string SQL;

                SQL = "Update Proposta set Proposta='" + Proposta + "' where Id =" + 1;

                conectaMySQL.ExecutaComando(SQL);
                conectaMySQL.FechaMySQL();
            }
            return true;
        }

        //Projeto
        public void CriaTableProjetos()
        {
            string SQL;
            //SQL = "DROP TABLE Clientes;";

            //MySqlDataReader dados = conectaMySQL.ExecutaConsulta(SQL);
            //SqlDataReader dados = ConectaBanco.ExecutaConsulta(SQL);
            //ConectaBanco.FechaBanco();
            //conectaMySQL.FechaMySQL();
            //string SQL;
            /*SQL = "CREATE TABLE Clientes (Id INT NOT NULL AUTO_INCREMENT, Nome VARCHAR(255) NOT NULL, " +
                  "CPF_CNPJ NVARCHAR(255) NOT NULL, CEP NVARCHAR(255) NOT NULL, Endereco NVARCHAR(255) NOT NULL, Numero NVARCHAR(255) NOT NULL," +
                  "Complemento  NVARCHAR(255) NOT NULL, Bairro NVARCHAR(255) NOT NULL, Cidade NVARCHAR(255) NOT NULL, " +
                  "UF  NVARCHAR(255) NOT NULL, email NVARCHAR(255) NOT NULL, Telefone  NVARCHAR(255) NOT NULL," +
                  "Celular NVARCHAR(255) NOT NULL, Quantidade_Inversores " +
                  "NVARCHAR(255) NOT NULL, Marca_Inversor NVARCHAR(255) NOT NULL, Modelo_Inversor NVARCHAR(255) NOT NULL, Quantidade_Modulos " +
                  "NVARCHAR(255) NOT NULL, Marca_Modulo NVARCHAR(255) NOT NULL, Modelo_Modulo NVARCHAR(255) NOT NULL, Consumo_Medio NVARCHAR(255) NOT NULL," +
                  "Identificacao NVARCHAR(255) NOT NULL, PRIMARY KEY (Id));";*/
            SQL = "CREATE TABLE Projetos (Id INT IDENTITY (1, 1) NOT NULL, Nome NVARCHAR(255) NOT NULL, " +
            "CPF_CNPJ NVARCHAR(255) NOT NULL, CEP NVARCHAR(255) NOT NULL, Endereco NVARCHAR(255) NOT NULL, Numero NVARCHAR(255) NOT NULL," +
            "Complemento  NVARCHAR(255) NOT NULL, Bairro NVARCHAR(255) NOT NULL, Cidade NVARCHAR(255) NOT NULL, " +
            "UF  NVARCHAR(255) NOT NULL, email NVARCHAR(255) NOT NULL, Telefone  NVARCHAR(255) NOT NULL," +
            "Celular NVARCHAR(255) NOT NULL, Quantidade_Inversores " +
            "NVARCHAR(255) NOT NULL, Marca_Inversor NVARCHAR(255) NOT NULL, Modelo_Inversor NVARCHAR(255) NOT NULL, Quantidade_Modulos " +
            "NVARCHAR(255) NOT NULL, Marca_Modulo NVARCHAR(255) NOT NULL, Modelo_Modulo NVARCHAR(255) NOT NULL, Consumo_Medio NVARCHAR(255) NOT NULL," +
            "Identificacao NVARCHAR(255) NOT NULL, PRIMARY KEY CLUSTERED ([Id] ASC));";

            //MySqlDataReader dados = conectaMySQL.ExecutaConsulta(SQL);
            SqlDataReader dados = ConectaBanco.ExecutaConsulta(SQL);
            ConectaBanco.FechaBanco();
            //conectaMySQL.FechaMySQL();
        }
        public void SelecionaProjeto(string BD)
        {
            if (BD == "local")
            {
                string SQL;
                SQL = "Select * from Projetos where Id=" + id;

                SqlDataReader dados = ConectaBanco.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    NumeroCliente = dados["Numero_do_Cliente"].ToString();
                    NumeroInstalacao = dados["Numero_da_Instalacao"].ToString();
                    Classe = dados["Classe"].ToString();
                    Latitude = dados["Latitude"].ToString();
                    Longitude = dados["Longitude"].ToString();
                    Disjuntor = dados["Disjuntor"].ToString();
                    CargaInstalada = dados["Carga_instalada"].ToString();
                    Nome = dados["Titular"].ToString();
                    CPF = dados["CPF_CNPJ"].ToString();
                    Endereco = dados["Endereco"].ToString();
                    Numero = dados["Numero"].ToString();
                    Complemento = dados["Complemento"].ToString();
                    Bairro = dados["Bairro"].ToString();
                    CEP = dados["CEP"].ToString();
                    Cidade = dados["Municipio"].ToString();
                    UF = dados["Estado"].ToString();
                    Telefone = dados["Telefone"].ToString();
                    Celular = dados["Celular"].ToString();
                    email = dados["Email"].ToString();
                    QuantidadeModulos = dados["Qtd_Modulo"].ToString();
                    MarcaMod = dados["Marca_Modulos"].ToString();
                    ModeloModulo = dados["Modelo_Modulos"].ToString();
                    QuantidadeInversores = dados["Qtd_Inversor"].ToString();
                    MarcaInversor = dados["Marca_Inversor"].ToString();
                    ModeloInversor = dados["Modelo_Inversor"].ToString();
                    Tensao = dados["Tensao_de_Atendimento"].ToString();
                    Estu = dados["Estrutura"].ToString();
                    Transformador = dados["Transformador"].ToString();
                    StringBox = dados["String_Box"].ToString();
                    Credito = dados["Qtd_Credito"].ToString();
                    Arranjo = dados["Arranjo"].ToString();
                }
                ConectaBanco.FechaBanco();
            }
            else if (BD == "mysql")
            {
                string SQL;
                SQL = "Select * from Orcamento where Id=" + id;

                MySqlDataReader dados = conectaMySQL.ExecutaConsulta(SQL);

                if (dados.Read())
                {
                    NumeroCliente = dados["Numero_do_Cliente"].ToString();
                    NumeroInstalacao = dados["Numero_da_Instalacao"].ToString();
                    Classe = dados["Classe"].ToString();
                    Latitude = dados["Latitude"].ToString();
                    Longitude = dados["Longitude"].ToString();
                    Disjuntor = dados["Disjuntor"].ToString();
                    CargaInstalada = dados["Carga_instalada"].ToString();
                    Nome = dados["Titular"].ToString();
                    CPF = dados["CPF_CNPJ"].ToString();
                    Endereco = dados["Endereco"].ToString();
                    Numero = dados["Numero"].ToString();
                    Complemento = dados["Complemento"].ToString();
                    Bairro = dados["Bairro"].ToString();
                    CEP = dados["CEP"].ToString();
                    Cidade = dados["Municipio"].ToString();
                    UF = dados["Estado"].ToString();
                    Telefone = dados["Telefone"].ToString();
                    Celular = dados["Celular"].ToString();
                    email = dados["Email"].ToString();
                    QuantidadeModulos = dados["Qtd_Modulo"].ToString();
                    MarcaMod = dados["Marca_Modulos"].ToString();
                    ModeloModulo = dados["Modelo_Modulos"].ToString();
                    QuantidadeInversores = dados["Qtd_Inversor"].ToString();
                    MarcaInversor = dados["Marca_Inversor"].ToString();
                    ModeloInversor = dados["Modelo_Inversor"].ToString();
                    Tensao = dados["Tensao_de_Atendimento"].ToString();
                    Estu = dados["Estrutura"].ToString();
                    Transformador = dados["Transformador"].ToString();
                    StringBox = dados["String_Box"].ToString();
                    Credito = dados["Qtd_Credito"].ToString();
                    Arranjo = dados["Arranjo"].ToString();
                }
                conectaMySQL.FechaMySQL();
            }

        }
        public bool SalvaProj(string BD)
        {
            if (BD == "local")
            {
                string SQL;

                SQL = "Insert into Projetos (Numero_do_Cliente, Numero_da_Instalacao, Classe, Titular, CPF_CNPJ, Endereco, Numero," +
                    "Complemento, Bairro, CEP, Municipio, Estado, Telefone, Celular, Email, Qtd_Modulo, Marca_Modulos, Modelo_Modulos," +
                    "Qtd_Inversor, Marca_Inversor, Modelo_Inversor, Latitude, Longitude, Disjuntor, Carga_Instalada, Tensao_de_Atendimento, " +
                    "Estrutura, Transformador, String_Box, Qtd_Credito, Arranjo) " +
                    "values('" + NumeroCliente + "','" + NumeroInstalacao + "','" + Classe + "','" + Nome + "','" + CPF + "', " +
                    "'" + Endereco + "','" + Numero + "','" + Complemento + "','" + Bairro + "','" + CEP + "','" + Cidade + "'," +
                    "'" + UF + "','" + Telefone + "','" + Celular + "','" + email + "'," +
                    "'" + QuantidadeModulos + "','" + MarcaMod + "','" + ModeloModulo + "','" + QuantidadeInversores + "','" + MarcaInversor + "','" + ModeloInversor + "','" + Latitude + "', " +
                    "'" + Longitude + "','" + Disjuntor + "','" + CargaInstalada + "','" + Tensao + "','" + Estu + "','" + Transformador + "','" + StringBox + "','" + Credito + "', " +
                    "'" + Arranjo + "')";

                ConectaBanco.ExecutaComando(SQL);
                ConectaBanco.FechaBanco();

            }
            else if (BD == "mysql")
            {
                string SQL;

                SQL = "Insert into Projetos (Numero_do_Cliente, Numero_da_Instalacao, Classe, Titular, CPF_CNPJ, Endereco, Numero," +
                    "Complemento, Bairro, CEP, Municipio, Estado, Telefone, Celular, Email, Qtd_Modulo, Marca_Modulos, Modelo_Modulos" +
                    "Qtd_Inversor, Marca_Inversor, Modelo_Inversor, Latitude, Longitude, Disjuntor, Carga_Instalada, Tensao_de_Atendimento, " +
                    "Estrutura, Transformador, String_Box, Qtd_Credito, Arranjo) " +
                    "values('" + NumeroCliente + "','" + NumeroInstalacao + "','" + Classe + "','" + Nome + "','" + CPF + "', " +
                    "'" + Endereco + "','" + Numero + "','" + Complemento + "','" + Bairro + "','" + CEP + "','" + Cidade + "'," +
                    "'" + UF + "','" + Telefone + "','" + Celular + "','" + email + "'," +
                    "'" + QuantidadeModulos + "','" + MarcaMod + "','" + ModeloModulo + "','" + QuantidadeInversores + "','" + MarcaInversor + "','" + ModeloInversor + "','" + Latitude + "', " +
                    "'" + Longitude + "','" + Disjuntor + "','" + CargaInstalada + "','" + Tensao + "','" + Estu + "','" + Transformador + "','" + StringBox + "','" + Credito + "', " +
                    "'" + Arranjo + "');";

                conectaMySQL.ExecutaComando(SQL);
                conectaMySQL.FechaMySQL();
            }
            return true;
        }
        public bool AlteraProj(string BD)
        {
            if (BD == "local")
            {
                string SQL;

                SQL = "Update Projetos set Numero_do_Cliente ='" + NumeroCliente + "', Numero_da_Instalacao ='" + NumeroInstalacao + "', Classe ='" + Classe + "', Titular ='" + Nome + "', CPF_CNPJ ='" + CPF + "'," +
                    " Endereco ='" + Endereco + "', Numero ='" + Numero + "', Complemento ='" + Complemento + "', Bairro ='" + Bairro + "'," +
                    " CEP ='" + CEP + "', Municipio ='" + Cidade + "', Estado ='" + UF + "', Telefone ='" + Telefone + "', " +
                    " Celular ='" + Celular + "', Email ='" + email + "', Qtd_Modulo ='" + QuantidadeModulos + "', " +
                    " Marca_Modulos ='" + MarcaMod + "', Modelo_Modulos ='" + ModeloModulo + "', Qtd_Inversor='" + QuantidadeInversores + "', Marca_Inversor = '" + MarcaInversor + "', Modelo_Inversor = '" + ModeloInversor + "', " +
                    " Latitude = '" + Latitude + "', Longitude ='" + Longitude + "', Disjuntor ='" + Disjuntor + "', Carga_Instalada ='" + CargaInstalada + "', Tensao_de_Atendimento ='" + Tensao + "', Estrutura ='" + Estu + "', " +
                    " Transformador ='" + Transformador + "', String_Box ='" + StringBox + "', Qtd_Credito ='" + Credito + "', Arranjo ='" + Arranjo + "' where Id =" + id;

                ConectaBanco.ExecutaComando(SQL);
                ConectaBanco.FechaBanco();
            }
            else if (BD == "mysql")
            {
                string SQL;

                SQL = "Update Projetos set Numero_do_Cliente ='" + NumeroCliente + "', Numero_da_Instalacao ='" + NumeroInstalacao + "', Classe ='" + Classe + "', Titular ='" + Nome + "', CPF_CNPJ ='" + CPF + "'," +
                    " Endereco ='" + Endereco + "', Numero ='" + Numero + "', Complemento ='" + Complemento + "', Bairro ='" + Bairro + "'," +
                    " CEP ='" + CEP + "', Municipio ='" + Cidade + "', Estado ='" + UF + "', Telefone ='" + Telefone + "', " +
                    " Celular ='" + Celular + "', Email ='" + email + "', Qtd_Modulo ='" + QuantidadeModulos + "', " +
                    " Marca_Modulos ='" + MarcaMod + "', Modelo_Modulos ='" + ModeloModulo + "', Qtd_Inversor='" + QuantidadeInversores + "', Marca_Inversor = '" + MarcaInversor + "', Modelo_Inversor = '" + ModeloInversor + "', " +
                    " Latitude = '" + Latitude + "', Longitude ='" + Longitude + "', Disjuntor ='" + Disjuntor + "', Carga_Instalada ='" + CargaInstalada + "', Tensao_de_Atendimento ='" + Tensao + "', Estrutura ='" + Estu + "', " +
                    " Transformador ='" + Transformador + "', String_Box ='" + StringBox + "', Qtd_Credito ='" + Credito + "', Arranjo ='" + Arranjo + "' where Id =" + id;

                conectaMySQL.ExecutaComando(SQL);
                conectaMySQL.FechaMySQL();
            }

            return true;
        }
        public DataTable PesquisaProjeto(string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Projetos";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Projetos";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }

        }
        public DataTable PesquisaProjFiltro(string TextoPesquisa, string coluna, string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Projetos where " + coluna + " like '%" + TextoPesquisa + "%'";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select * from Projetos where " + coluna + " like '%" + TextoPesquisa + "%'";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }

        }

        //Fornecedores
        public DataTable BuscaFornecedor(string BD)
        {
            if (BD == "local")
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = ConectaBanco.AbreBanco();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Fornecedor from Fornecedor Order By Fornecedor";

                SqlDataAdapter adaptador = new SqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                ConectaBanco.FechaBanco();

                return dt;
            }
            else if (BD == "mysql")
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conectaMySQL.AbreMySQL();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Fornecedor from Fornecedor Order By Fornecedor";

                MySqlDataAdapter adaptador = new MySqlDataAdapter();
                adaptador.SelectCommand = cmd;

                DataTable dt = new DataTable();
                adaptador.Fill(dt);
                conectaMySQL.FechaMySQL();

                return dt;
            }
            else
            {
                return null;
            }

        }
    }
}
