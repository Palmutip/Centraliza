using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using CefSharp;
using CefSharp.WinForms;
using FireSharp.Config;
using FireSharp.Interfaces;
using FireSharp.Response;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Diagnostics;
using System.Drawing.Imaging;

namespace Centraliza
{
    public partial class TelaInicial : Form
    {
        //Variaveis e objetos
        int qtdorc = 0, i=4,j=0;
        int tjan, tfev, tmar, tabr, tmai, tjun, tjul, tago, tout, tnov, tset, tdez, somadisp;
        double potenciagerada, gerano, payback, valorsisjuros, total = 0, resultado = 0, percas = 1, simpot = 0;
        string payback1, potger1, gerano1, germes1, contot, conmes1, dimensao1, caixaacumulado, foto;
        string valorsist, valsisjur, valparcela, valorequip, valorrestante1, Banco, pottotalmodprojeto, pottotalinvprojeto;
        String imageLocation = "";
        bool editando = false, validado = false, duplo=false;

        Login login = new Login();
        FuncoesBanco func = new FuncoesBanco();
        Orcamento orcamento = new Orcamento();

        //Construtores
        public TelaInicial()
        {
            InitializeComponent();
            pgbstatusorca.Maximum = 26;
            pgbstatusorca.Value = 0;
            pgbstatusproj.Maximum = 26;
            pgbstatusproj.Value = 0;
            foto = "";
        }
        public TelaInicial(string Login)
        {
            InitializeComponent();
            pgbstatusorca.Maximum = 26;
            pgbstatusorca.Value = 0;
            pgbstatusproj.Maximum = 26;
            pgbstatusproj.Value = 0;
            func.Login = Login;
        }
        private void TelaInicial_Load(object sender, EventArgs e)
        {
            //func.SelecionaBanco();
            //if (func.Banco == "local" || func.Banco == "mysql")
            //{
            //    Banco = func.Banco;
            //}
            Banco = "mysql";
            login.Owner = this;
            login.ShowDialog();
            if (!validado)
            {
                Application.Exit();
            }
            else
            {
                CarregaFoto();
                if (func.Banco == "local")
                {
                    rbtnlocal.Checked = true;
                }
                else if (func.Banco == "mysql")
                {
                    rbtnmysql.Checked = true;
                }
                CarregaDataGrid();
                PaineisPrincipais(pnlinicio);
                PictureBoxRedondo();
                CarregaCombobox();
            }
        }

        //Funções
        private void ClicaInicio()
        {
            this.btnclientes.BackColor = Color.FromArgb(255, 255, 255);
            this.btnequipamentos.BackColor = Color.FromArgb(255, 255, 255);
            this.btnorcamento.BackColor = Color.FromArgb(255, 255, 255);
            this.btnprojeto.BackColor = Color.FromArgb(255, 255, 255);
            this.btninicio.BackColor = Color.FromArgb(255,83, 19);

            this.btnconfiguracoes.BackColor = Color.FromArgb(255, 255, 255);
            this.btnsair.BackColor = Color.FromArgb(255, 255, 255);

            this.btnclientes.Image = Properties.Resources.user_cinza;
            this.btnequipamentos.Image = Properties.Resources.painel_cinza;
            this.btnorcamento.Image = Properties.Resources.moeda_cinza;
            this.btnprojeto.Image = Properties.Resources.edit_cinza;
            this.btninicio.Image = Properties.Resources.home_branca;

            this.btnconfiguracoes.Image = Properties.Resources.conf_cinza;
            this.btnsair.Image = Properties.Resources.exit_cinza;

            this.btnclientes.ForeColor = Color.FromArgb(120, 120, 120);
            this.btnequipamentos.ForeColor = Color.FromArgb(120, 120, 120);
            this.btnorcamento.ForeColor = Color.FromArgb(120, 120, 120);
            this.btnprojeto.ForeColor = Color.FromArgb(120, 120, 120);
            this.btninicio.ForeColor = Color.FromArgb(255, 255, 255);

            this.btnconfiguracoes.ForeColor = Color.FromArgb(120, 120, 120);
            this.btnsair.ForeColor = Color.FromArgb(120, 120, 120);

            editando = false;

            PaineisPrincipais(pnlinicio);
        }
        private void ClicaClientes()
        {
            this.btnclientes.BackColor = Color.FromArgb(255, 83, 19);
            this.btnequipamentos.BackColor = Color.FromArgb(255, 255, 255);
            this.btnorcamento.BackColor = Color.FromArgb(255, 255, 255);
            this.btnprojeto.BackColor = Color.FromArgb(255, 255, 255);
            this.btninicio.BackColor = Color.FromArgb(255, 255, 255);

            this.btnconfiguracoes.BackColor = Color.FromArgb(255, 255, 255);
            this.btnsair.BackColor = Color.FromArgb(255, 255, 255);

            this.btnclientes.Image = Properties.Resources.user_branca;
            this.btnequipamentos.Image = Properties.Resources.painel_cinza;
            this.btnorcamento.Image = Properties.Resources.moeda_cinza;
            this.btnprojeto.Image = Properties.Resources.edit_cinza;
            this.btninicio.Image = Properties.Resources.home_cinza;

            this.btnconfiguracoes.Image = Properties.Resources.conf_cinza;
            this.btnsair.Image = Properties.Resources.exit_cinza;

            this.btnclientes.ForeColor = Color.FromArgb(255, 255, 255);
            this.btnequipamentos.ForeColor = Color.FromArgb(120, 120, 120);
            this.btnorcamento.ForeColor = Color.FromArgb(120, 120, 120);
            this.btnprojeto.ForeColor = Color.FromArgb(120, 120, 120);
            this.btninicio.ForeColor = Color.FromArgb(120, 120, 120);

            this.btnconfiguracoes.ForeColor = Color.FromArgb(120, 120, 120);
            this.btnsair.ForeColor = Color.FromArgb(120, 120, 120);

            editando = false;

            PaineisPrincipais(pnlclientes);

        }
        private void ClicaEquipamentos()
        {
            this.btnclientes.BackColor = Color.FromArgb(255, 255, 255);
            this.btnequipamentos.BackColor = Color.FromArgb(255, 83, 19);
            this.btnorcamento.BackColor = Color.FromArgb(255, 255, 255);
            this.btnprojeto.BackColor = Color.FromArgb(255, 255, 255);
            this.btninicio.BackColor = Color.FromArgb(255, 255, 255);

            this.btnconfiguracoes.BackColor = Color.FromArgb(255, 255, 255);
            this.btnsair.BackColor = Color.FromArgb(255, 255, 255);

            this.btnclientes.Image = Properties.Resources.user_cinza;
            this.btnequipamentos.Image = Properties.Resources.painel_branca;
            this.btnorcamento.Image = Properties.Resources.moeda_cinza;
            this.btnprojeto.Image = Properties.Resources.edit_cinza;
            this.btninicio.Image = Properties.Resources.home_cinza;

            this.btnconfiguracoes.Image = Properties.Resources.conf_cinza;
            this.btnsair.Image = Properties.Resources.exit_cinza;

            this.btnclientes.ForeColor = Color.FromArgb(120, 120, 120);
            this.btnequipamentos.ForeColor = Color.FromArgb(255, 255, 255);
            this.btnorcamento.ForeColor = Color.FromArgb(120, 120, 120);
            this.btnprojeto.ForeColor = Color.FromArgb(120, 120, 120);
            this.btninicio.ForeColor = Color.FromArgb(120, 120, 120);

            this.btnconfiguracoes.ForeColor = Color.FromArgb(120, 120, 120);
            this.btnsair.ForeColor = Color.FromArgb(120, 120, 120);

            editando = false;

        }
        private void ClicaProjeto()
        {
            this.btnclientes.BackColor = Color.FromArgb(255, 255, 255);
            this.btnequipamentos.BackColor = Color.FromArgb(255, 255, 255);
            this.btnorcamento.BackColor = Color.FromArgb(255, 255, 255);
            this.btnprojeto.BackColor = Color.FromArgb(255, 83, 19);
            this.btninicio.BackColor = Color.FromArgb(255, 255, 255);

            this.btnconfiguracoes.BackColor = Color.FromArgb(255, 255, 255);
            this.btnsair.BackColor = Color.FromArgb(255, 255, 255);

            this.btnclientes.Image = Properties.Resources.user_cinza;
            this.btnequipamentos.Image = Properties.Resources.painel_cinza;
            this.btnorcamento.Image = Properties.Resources.moeda_cinza;
            this.btnprojeto.Image = Properties.Resources.edit_branca;
            this.btninicio.Image = Properties.Resources.home_cinza;

            this.btnconfiguracoes.Image = Properties.Resources.conf_cinza;
            this.btnsair.Image = Properties.Resources.exit_cinza;

            this.btnclientes.ForeColor = Color.FromArgb(120, 120, 120);
            this.btnequipamentos.ForeColor = Color.FromArgb(120, 120, 120);
            this.btnorcamento.ForeColor = Color.FromArgb(120, 120, 120);
            this.btnprojeto.ForeColor = Color.FromArgb(255, 255, 255);
            this.btninicio.ForeColor = Color.FromArgb(120, 120, 120);

            this.btnconfiguracoes.ForeColor = Color.FromArgb(120, 120, 120);
            this.btnsair.ForeColor = Color.FromArgb(120, 120, 120);

            editando = false;

        }
        private void ClicaOrcamento()
        {
            this.btnclientes.BackColor = Color.FromArgb(255, 255, 255);
            this.btnequipamentos.BackColor = Color.FromArgb(255, 255, 255);
            this.btnorcamento.BackColor = Color.FromArgb(255, 83, 19);
            this.btnprojeto.BackColor = Color.FromArgb(255, 255, 255);
            this.btninicio.BackColor = Color.FromArgb(255, 255, 255);

            this.btnconfiguracoes.BackColor = Color.FromArgb(255, 255, 255);
            this.btnsair.BackColor = Color.FromArgb(255, 255, 255);

            this.btnclientes.Image = Properties.Resources.user_cinza;
            this.btnequipamentos.Image = Properties.Resources.painel_cinza;
            this.btnorcamento.Image = Properties.Resources.moeda_branca;
            this.btnprojeto.Image = Properties.Resources.edit_cinza;
            this.btninicio.Image = Properties.Resources.home_cinza;

            this.btnconfiguracoes.Image = Properties.Resources.conf_cinza;
            this.btnsair.Image = Properties.Resources.exit_cinza;

            this.btnclientes.ForeColor = Color.FromArgb(120, 120, 120);
            this.btnequipamentos.ForeColor = Color.FromArgb(120, 120, 120);
            this.btnorcamento.ForeColor = Color.FromArgb(255, 255, 255);
            this.btnprojeto.ForeColor = Color.FromArgb(120, 120, 120);
            this.btninicio.ForeColor = Color.FromArgb(120, 120, 120);

            this.btnconfiguracoes.ForeColor = Color.FromArgb(120, 120, 120);
            this.btnsair.ForeColor = Color.FromArgb(120, 120, 120);

            editando = false;

            PaineisPrincipais(pnlorcasalvos);
        }
        private void ClicaConf()
        {
            this.btnclientes.BackColor = Color.FromArgb(255, 255, 255);
            this.btnequipamentos.BackColor = Color.FromArgb(255, 255, 255);
            this.btnorcamento.BackColor = Color.FromArgb(255, 255, 255);
            this.btnprojeto.BackColor = Color.FromArgb(255, 255, 255);
            this.btninicio.BackColor = Color.FromArgb(255, 255, 255);

            this.btnconfiguracoes.BackColor = Color.FromArgb(255, 83, 19);
            this.btnsair.BackColor = Color.FromArgb(255, 255, 255);

            this.btnclientes.Image = Properties.Resources.user_cinza;
            this.btnequipamentos.Image = Properties.Resources.painel_cinza;
            this.btnorcamento.Image = Properties.Resources.moeda_cinza;
            this.btnprojeto.Image = Properties.Resources.edit_cinza;
            this.btninicio.Image = Properties.Resources.home_cinza;

            this.btnconfiguracoes.Image = Properties.Resources.conf_branca;
            this.btnsair.Image = Properties.Resources.exit_cinza;

            this.btnclientes.ForeColor = Color.FromArgb(120, 120, 120);
            this.btnequipamentos.ForeColor = Color.FromArgb(120, 120, 120);
            this.btnorcamento.ForeColor = Color.FromArgb(120, 120, 120);
            this.btnprojeto.ForeColor = Color.FromArgb(120, 120, 120);
            this.btninicio.ForeColor = Color.FromArgb(120, 120, 120);

            this.btnconfiguracoes.ForeColor = Color.FromArgb(255, 255, 255);
            this.btnsair.ForeColor = Color.FromArgb(120, 120, 120);

            editando = false;

        }
        private void ClicaSair()
        {
            this.btnclientes.BackColor = Color.FromArgb(255, 255, 255);
            this.btnequipamentos.BackColor = Color.FromArgb(255, 255, 255);
            this.btnorcamento.BackColor = Color.FromArgb(255, 255, 255);
            this.btnprojeto.BackColor = Color.FromArgb(255, 255, 255);
            this.btninicio.BackColor = Color.FromArgb(255, 255, 255);

            this.btnconfiguracoes.BackColor = Color.FromArgb(255, 255, 255);
            this.btnsair.BackColor = Color.FromArgb(255, 83, 19);

            this.btnclientes.Image = Properties.Resources.user_cinza;
            this.btnequipamentos.Image = Properties.Resources.painel_cinza;
            this.btnorcamento.Image = Properties.Resources.moeda_cinza;
            this.btnprojeto.Image = Properties.Resources.edit_cinza;
            this.btninicio.Image = Properties.Resources.home_cinza;

            this.btnconfiguracoes.Image = Properties.Resources.conf_cinza;
            this.btnsair.Image = Properties.Resources.exit_branca;

            this.btnclientes.ForeColor = Color.FromArgb(120, 120, 120);
            this.btnequipamentos.ForeColor = Color.FromArgb(120, 120, 120);
            this.btnorcamento.ForeColor = Color.FromArgb(120, 120, 120);
            this.btnprojeto.ForeColor = Color.FromArgb(120, 120, 120);
            this.btninicio.ForeColor = Color.FromArgb(120, 120, 120);

            this.btnconfiguracoes.ForeColor = Color.FromArgb(120, 120, 120);
            this.btnsair.ForeColor = Color.FromArgb(255, 255, 255);

            editando = false;

        }
        private void ConsumoUnidades()
        {
            PaineisPrincipais(pnlorcamento1);
        }
        private void recolhe()
        {
            if (pnlprincipal.Visible)
            {
                /*pnlinicio.Location = new Point(129, 110);
                pnlinicio.Size = new Size(1110, 523);
                pnlinicio.Anchor = (AnchorStyles.Bottom | AnchorStyles.Right | AnchorStyles.Top | AnchorStyles.Left);
                pnlinicio.AutoSize = true;*/
                pnlprincipal.Visible = false;
                pnlrec.Visible = true;
                pnlmenurec.Visible = true;
                pnlconf.Visible = false;
            }
            else
            {
                /*pnlinicio.Location = new Point(309, 110);
                pnlinicio.Size = new Size(922, 523);
                pnlinicio.Anchor = (AnchorStyles.Bottom | AnchorStyles.Right | AnchorStyles.Top | AnchorStyles.Left);
                pnlinicio.AutoSize = true;*/
                pnlprincipal.Visible = true;
                pnlrec.Visible = false;
                pnlmenurec.Visible = false;
                pnlconf.Visible = true;
            }
        }
        private void verificacor()
        {
            if (this.btnsair.BackColor == Color.FromArgb(255, 83, 19))
            {
                this.btnclientesrec.BackColor = Color.FromArgb(255, 255, 255);
                this.btnequipamentosrec.BackColor = Color.FromArgb(255, 255, 255);
                this.btnorcamentorec.BackColor = Color.FromArgb(255, 255, 255);
                this.btnprojetorec.BackColor = Color.FromArgb(255, 255, 255);
                this.btniniciorec.BackColor = Color.FromArgb(255, 255, 255);

                this.btnconfiguracoesrec.BackColor = Color.FromArgb(255, 255, 255);
                this.btnsairrec.BackColor = Color.FromArgb(255, 83, 19);

                this.btnclientesrec.Image = Properties.Resources.user_cinza;
                this.btnequipamentosrec.Image = Properties.Resources.painel_cinza;
                this.btnorcamentorec.Image = Properties.Resources.moeda_cinza;
                this.btnprojetorec.Image = Properties.Resources.edit_cinza;
                this.btniniciorec.Image = Properties.Resources.home_cinza;

                this.btnconfiguracoesrec.Image = Properties.Resources.conf_cinza;
                this.btnsairrec.Image = Properties.Resources.exit_branca;
            }
            else if (this.btnconfiguracoes.BackColor == Color.FromArgb(255, 83, 19))
            {
                this.btnclientesrec.BackColor = Color.FromArgb(255, 255, 255);
                this.btnequipamentosrec.BackColor = Color.FromArgb(255, 255, 255);
                this.btnorcamentorec.BackColor = Color.FromArgb(255, 255, 255);
                this.btnprojetorec.BackColor = Color.FromArgb(255, 255, 255);
                this.btniniciorec.BackColor = Color.FromArgb(255, 255, 255);

                this.btnconfiguracoesrec.BackColor = Color.FromArgb(255, 83, 19);
                this.btnsairrec.BackColor = Color.FromArgb(255, 255, 255);

                this.btnclientesrec.Image = Properties.Resources.user_cinza;
                this.btnequipamentosrec.Image = Properties.Resources.painel_cinza;
                this.btnorcamentorec.Image = Properties.Resources.moeda_cinza;
                this.btnprojetorec.Image = Properties.Resources.edit_cinza;
                this.btniniciorec.Image = Properties.Resources.home_cinza;

                this.btnconfiguracoesrec.Image = Properties.Resources.conf_branca;
                this.btnsairrec.Image = Properties.Resources.exit_cinza;
            }
            else if (this.btnprojeto.BackColor == Color.FromArgb(255, 83, 19))
            {
                this.btnclientesrec.BackColor = Color.FromArgb(255, 255, 255);
                this.btnequipamentosrec.BackColor = Color.FromArgb(255, 255, 255);
                this.btnorcamentorec.BackColor = Color.FromArgb(255, 255, 255);
                this.btnprojetorec.BackColor = Color.FromArgb(255, 83, 19);
                this.btniniciorec.BackColor = Color.FromArgb(255, 255, 255);

                this.btnconfiguracoesrec.BackColor = Color.FromArgb(255, 255, 255);
                this.btnsairrec.BackColor = Color.FromArgb(255, 255, 255);

                this.btnclientesrec.Image = Properties.Resources.user_cinza;
                this.btnequipamentosrec.Image = Properties.Resources.painel_cinza;
                this.btnorcamentorec.Image = Properties.Resources.moeda_cinza;
                this.btnprojetorec.Image = Properties.Resources.edit_branca;
                this.btniniciorec.Image = Properties.Resources.home_cinza;

                this.btnconfiguracoesrec.Image = Properties.Resources.conf_cinza;
                this.btnsairrec.Image = Properties.Resources.exit_cinza;
            }
            else if (this.btninicio.BackColor == Color.FromArgb(255, 83, 19))
            {
                this.btnclientesrec.BackColor = Color.FromArgb(255, 255, 255);
                this.btnequipamentosrec.BackColor = Color.FromArgb(255, 255, 255);
                this.btnorcamentorec.BackColor = Color.FromArgb(255, 255, 255);
                this.btnprojetorec.BackColor = Color.FromArgb(255, 255, 255);
                this.btniniciorec.BackColor = Color.FromArgb(255, 83, 19);

                this.btnconfiguracoesrec.BackColor = Color.FromArgb(255, 255, 255);
                this.btnsairrec.BackColor = Color.FromArgb(255, 255, 255);

                this.btnclientesrec.Image = Properties.Resources.user_cinza;
                this.btnequipamentosrec.Image = Properties.Resources.painel_cinza;
                this.btnorcamentorec.Image = Properties.Resources.moeda_cinza;
                this.btnprojetorec.Image = Properties.Resources.edit_cinza;
                this.btniniciorec.Image = Properties.Resources.home_branca;

                this.btnconfiguracoesrec.Image = Properties.Resources.conf_cinza;
                this.btnsairrec.Image = Properties.Resources.exit_cinza;
            }
            else if (this.btnequipamentos.BackColor == Color.FromArgb(255, 83, 19))
            {
                this.btnclientesrec.BackColor = Color.FromArgb(255, 255, 255);
                this.btnequipamentosrec.BackColor = Color.FromArgb(255, 83, 19);
                this.btnorcamentorec.BackColor = Color.FromArgb(255, 255, 255);
                this.btnprojetorec.BackColor = Color.FromArgb(255, 255, 255);
                this.btniniciorec.BackColor = Color.FromArgb(255, 255, 255);

                this.btnconfiguracoesrec.BackColor = Color.FromArgb(255, 255, 255);
                this.btnsairrec.BackColor = Color.FromArgb(255, 255, 255);

                this.btnclientesrec.Image = Properties.Resources.user_cinza;
                this.btnequipamentosrec.Image = Properties.Resources.painel_branca;
                this.btnorcamentorec.Image = Properties.Resources.moeda_cinza;
                this.btnprojetorec.Image = Properties.Resources.edit_cinza;
                this.btniniciorec.Image = Properties.Resources.home_cinza;

                this.btnconfiguracoesrec.Image = Properties.Resources.conf_cinza;
                this.btnsairrec.Image = Properties.Resources.exit_cinza;
            }
            else if (this.btnorcamento.BackColor == Color.FromArgb(255, 83, 19))
            {
                this.btnclientesrec.BackColor = Color.FromArgb(255, 255, 255);
                this.btnequipamentosrec.BackColor = Color.FromArgb(255, 255, 255);
                this.btnorcamentorec.BackColor = Color.FromArgb(255, 83, 19);
                this.btnprojetorec.BackColor = Color.FromArgb(255, 255, 255);
                this.btniniciorec.BackColor = Color.FromArgb(255, 255, 255);

                this.btnconfiguracoesrec.BackColor = Color.FromArgb(255, 255, 255);
                this.btnsairrec.BackColor = Color.FromArgb(255, 255, 255);

                this.btnclientesrec.Image = Properties.Resources.user_cinza;
                this.btnequipamentosrec.Image = Properties.Resources.painel_cinza;
                this.btnorcamentorec.Image = Properties.Resources.moeda_branca;
                this.btnprojetorec.Image = Properties.Resources.edit_cinza;
                this.btniniciorec.Image = Properties.Resources.home_cinza;

                this.btnconfiguracoesrec.Image = Properties.Resources.conf_cinza;
                this.btnsairrec.Image = Properties.Resources.exit_cinza;
            }
            else if (this.btnclientes.BackColor == Color.FromArgb(255, 83, 19))
            {
                this.btnclientesrec.BackColor = Color.FromArgb(255, 83, 19);
                this.btnequipamentosrec.BackColor = Color.FromArgb(255, 255, 255);
                this.btnorcamentorec.BackColor = Color.FromArgb(255, 255, 255);
                this.btnprojetorec.BackColor = Color.FromArgb(255, 255, 255);
                this.btniniciorec.BackColor = Color.FromArgb(255, 255, 255);

                this.btnconfiguracoesrec.BackColor = Color.FromArgb(255, 255, 255);
                this.btnsairrec.BackColor = Color.FromArgb(255, 255, 255);

                this.btnclientesrec.Image = Properties.Resources.user_branca;
                this.btnequipamentosrec.Image = Properties.Resources.painel_cinza;
                this.btnorcamentorec.Image = Properties.Resources.moeda_cinza;
                this.btnprojetorec.Image = Properties.Resources.edit_cinza;
                this.btniniciorec.Image = Properties.Resources.home_cinza;

                this.btnconfiguracoesrec.Image = Properties.Resources.conf_cinza;
                this.btnsairrec.Image = Properties.Resources.exit_cinza;
            }

        }
        private void Salvaorcamento()
        {
            func.Nome = txtnome.Text;
            func.Contato = txtcontato.Text;
            func.CEP = txtcep.Text;
            func.Endereco = txtendereco.Text;
            func.Numero = txtnumero.Text;
            func.Bairro = txtbairro.Text;
            func.Cidade = txtcidade.Text;
            func.Custo = txtkwh.Text;
            func.Consumoanual = txtcontot.Text;
            func.Estu = cbxestrutura.Text;
            func.QuantidadeModulos = txtqtdpaineis.Text;
            func.MarcaMod = cbxmarcamod.Text;
            func.ModeloModulo = cbxmodpaineis.Text;
            func.QuantidadeInversores = txtqtdinv.Text;
            func.MarcaInversor = cbxmarcainv.Text;
            func.ModeloInversor = cbxmodinv.Text;
            func.Valorsist = txtvalorsist.Text;
            func.Valorequip = txtvalorequip.Text;
            func.Valorinv = txtcustoinversor.Text;
            func.Obs = txtobs.Text;
            func.CliJan = tjan;
            func.CliFev = tfev;
            func.CliMar = tmar;
            func.CliAbr = tabr;
            func.CliMai = tmai;
            func.CliJun = tjun;
            func.CliJul = tjul;
            func.CliAgo = tago;
            func.CliSet = tset;
            func.CliOut = tout;
            func.CliNov = tnov;
            func.CliDez = tdez;
            func.Disponibilidade = somadisp;
            if (rbtn0.Checked)
            {
                func.Perdas = "0";
            }
            else if (rbtn5.Checked)
            {
                func.Perdas = "5";
            }
            else if (rbtn7.Checked)
            {
                func.Perdas = "7";
            }
            else if (rbtn10.Checked)
            {
                func.Perdas = "10";
            }
            else if (rbtn12.Checked)
            {
                func.Perdas = "12";
            }
            else if (rbtn15.Checked)
            {
                func.Perdas = "15";
            }
            else if (rbtn20.Checked)
            {
                func.Perdas = "20";
            }
            else if (rbtn25.Checked)
            {
                func.Perdas = "25";
            }
            else if (rbtn30.Checked)
            {
                func.Perdas = "30";
            }
            else if (rbtn35.Checked)
            {
                func.Perdas = "35";
            }
            else if (rbtn40.Checked)
            {
                func.Perdas = "40";
            }

            if (!editando)
            {
                func.Salvaorc(Banco);
            }
            else
            {
                func.Alterarorc(Banco);
            }
            
            dgvorcamentos.DataSource = func.CarregaOrc(Banco);
        }
        private void EditaOrcamento()
        {
            func.Id = Convert.ToInt32(dgvorcamentos.CurrentRow.Cells[0].Value);
            func.SelecionaOrcamento(Banco);

            txtnome.Text = func.Nome;
            txtcontato.Text = func.Contato;
            txtcep.Text = func.CEP;
            txtendereco.Text = func.Endereco;
            txtnumero.Text = func.Numero;
            txtbairro.Text = func.Bairro;
            txtcidade.Text = func.Cidade;
            txtkwh.Text = func.Custo;
            txtcontot.Text = func.Consumoanual;
            cbxestrutura.Text = func.Estu;
            txtqtdpaineis.Text = func.QuantidadeModulos;
            /*cbxmarcamod.Text = func.MarcaMod;
            cbxmodpaineis.Text = func.ModeloModulo;*/
            //var Dados = func.PreencheMarcaModOrc(Banco);
            //cbxmarcamod.DataSource = Dados;
            //cbxmarcamod.ValueMember = "Marca_Paineis";
            //cbxmarcamod.DisplayMember = "Marca_Paineis";
            cbxmarcamod.Text = func.MarcaMod;
            //var Dados2 = func.PreencheModModOrc(Banco);
            //cbxmodpaineis.DataSource = Dados2;
            //cbxmodpaineis.ValueMember = "Modelo_Paineis";
            //cbxmodpaineis.DisplayMember = "Modelo_Paineis";
            cbxmodpaineis.Text = func.ModeloModulo;
            txtqtdinv.Text = func.QuantidadeInversores;
            cbxmarcainv.Text = func.MarcaInversor;
            cbxmodinv.Text = func.ModeloInversor;
            //var Dados3 = func.PreencheMarcaInvOrc(Banco);
            //cbxmarcainv.DataSource = Dados3;
            //cbxmarcainv.ValueMember = "Marca_Inversores";
            //cbxmarcainv.DisplayMember = "Marca_Inversores";
            //var Dados4 = func.PreencheModIUnvOrc(Banco);
            //cbxmodinv.DataSource = Dados4;
            //cbxmodinv.ValueMember = "Modelo_Inversores";
            //cbxmodinv.DisplayMember = "Modelo_Inversores";
            txtvalorsist.Text = func.Valorsist;
            txtvalorequip.Text = func.Valorequip;
            txtcustoinversor.Text = func.Valorinv;
            switch (Int32.Parse(func.Perdas))
            {
                case 0:
                    rbtn0.Checked = true;
                    break;
                case 5:
                    rbtn5.Checked = true;
                    break;
                case 7:
                    rbtn7.Checked = true;
                    break;
                case 10:
                    rbtn10.Checked = true;
                    break;
                case 12:
                    rbtn12.Checked = true;
                    break;
                case 15:
                    rbtn15.Checked = true;
                    break;
                case 20:
                    rbtn20.Checked = true;
                    break;
                case 25:
                    rbtn25.Checked = true;
                    break;
                case 30:
                    rbtn30.Checked = true;
                    break;
                case 35:
                    rbtn35.Checked = true;
                    break;
                case 40:
                    rbtn40.Checked = true;
                    break;
                default:
                    rbtn0.Checked = true;
                    break;

            }
            txtobs.Text = func.Obs;

            txtjan.Text = func.CliJan.ToString();
            txtfev.Text = func.CliFev.ToString();
            txtmar.Text = func.CliMar.ToString();
            txtabr.Text = func.CliAbr.ToString();
            txtmai.Text = func.CliMai.ToString();
            txtjun.Text = func.CliJun.ToString();
            txtjul.Text = func.CliJul.ToString();
            txtago.Text = func.CliAgo.ToString();
            txtset.Text = func.CliSet.ToString();
            txtout.Text = func.CliOut.ToString();
            txtnov.Text = func.CliNov.ToString();
            txtdez.Text = func.CliDez.ToString();
            somadisp = func.Disponibilidade;
        }
        private void Calculos()
        {
            //Calculo e formatação de texto das variaveis
            func.PesquisaPotInv(cbxmodinv.Text, Banco);
            func.SelecionaProposta(Banco);
            int prop = func.Id;
            func.SelecionaProposta1(Banco);
            int prop1 = func.Proposta;
            prop1++;
            func.Proposta = prop1;
            func.Maisumaproposta1(Banco);
            double conmes = double.Parse(txtcontot.Text) / 12;
            conmes1 = conmes.ToString();
            conmes1 = string.Format("{0:0,0}", conmes);
            double valorrestante = double.Parse(txtvalorsist.Text) - double.Parse(txtvalorequip.Text);
            valorrestante1 = valorrestante.ToString();
            valorrestante1 = string.Format("{0:0,0}", valorrestante);
            func.PesquisaPotMod(cbxmodpaineis.Text, Banco);
            double germes = (((double.Parse(func.PotenciaMod) * double.Parse(txtqtdpaineis.Text)) * 0.83) / 1000) * 30 * 4.87;
            if (rbtn5.Checked)
            {
                germes *= 0.95;
            }
            else if (rbtn7.Checked)
            {
                germes *= 0.93;
            }
            else if (rbtn10.Checked)
            {
                germes *= 0.9;
            }
            else if (rbtn12.Checked)
            {
                germes *= 0.88;
            }
            else if (rbtn15.Checked)
            {
                germes *= 0.85;
            }
            else if (rbtn20.Checked)
            {
                germes *= 0.80;
            }
            else if (rbtn25.Checked)
            {
                germes *= 0.75;
            }
            else if (rbtn30.Checked)
            {
                germes *= 1.1;
            }
            else if (rbtn35.Checked)
            {
                germes *= 1.075;
            }
            else if (rbtn40.Checked)
            {
                germes *= 1.05;
            }
            germes1 = germes.ToString();
            germes1 = string.Format("{0:0,0}", germes);
            gerano = (((double.Parse(func.PotenciaMod) * double.Parse(txtqtdpaineis.Text)) * 0.83) / 1000) * 365 * 4.87;
            if (rbtn5.Checked)
            {
                gerano *= 0.95;
            }
            else if (rbtn7.Checked)
            {
                gerano *= 0.93;
            }
            else if (rbtn10.Checked)
            {
                gerano *= 0.9;
            }
            else if (rbtn12.Checked)
            {
                gerano *= 0.88;
            }
            else if (rbtn15.Checked)
            {
                gerano *= 0.85;
            }
            else if (rbtn20.Checked)
            {
                gerano *= 0.80;
            }
            else if (rbtn25.Checked)
            {
                gerano *= 0.75;
            }
            else if (rbtn30.Checked)
            {
                gerano *= 1.1;
            }
            else if (rbtn35.Checked)
            {
                gerano *= 1.075;
            }
            else if (rbtn40.Checked)
            {
                gerano *= 1.05;
            }
            gerano1 = gerano.ToString();
            gerano1 = string.Format("{0:0,0}", gerano);
            func.PesquisaDimensoesMod(cbxmodpaineis.Text, Banco);
            double dimensao = 2 * double.Parse(txtqtdpaineis.Text);
            dimensao1 = dimensao.ToString();
            dimensao1 = string.Format("{0:0,0.00}", dimensao);
            valorequip = txtvalorequip.Text;
            valorequip = string.Format("{0:0,0}", Int64.Parse(txtvalorequip.Text));
            valorsist = txtvalorsist.Text;
            valorsist = string.Format("{0:0,0}", Int64.Parse(txtvalorsist.Text));
            valorsisjuros = double.Parse(txtvalorsist.Text);
            valorsisjuros = (valorsisjuros * 1.18) / 12;
            valparcela = valorsisjuros.ToString();
            valparcela = string.Format("{0:0,0}", valorsisjuros);
            valsisjur = string.Format("{0:0,0}", (double.Parse(valparcela) * 12));
            contot = txtcontot.Text;
            contot = string.Format("{0:0,0}", Int64.Parse(txtcontot.Text));
            potenciagerada = double.Parse(func.PotenciaMod) * double.Parse(txtqtdpaineis.Text) / 1000;
            potger1 = potenciagerada.ToString();
            potger1 = string.Format("{0:0,0.00}", potenciagerada);
            if (txtcustoinversor.Text == "" || txtcustoinversor.Enabled == false || txtcustoinversor.Text == string.Empty)
            {
                payback = (((double.Parse(txtvalorsist.Text) + 1) * double.Parse(txtkwh.Text)) / gerano);
                payback1 = payback.ToString();
                payback1 = string.Format("{0:0.0}", payback);
            }
            else
            {
                payback = (((double.Parse(txtvalorsist.Text) + double.Parse(txtcustoinversor.Text)) * double.Parse(txtkwh.Text)) / gerano);
                payback1 = payback.ToString();
                payback1 = string.Format("{0:0.0}", payback);
            }

            //calculo retorno financeiro
            double auxiliar = double.Parse(txtkwh.Text);
            double auxiliarano = (gerano * 0.99) - (gerano * 0.985);
            double auxiliarano1 = gerano;
            double geracaoanual = double.Parse(txtvalorsist.Text);
            double caixaanual = (auxiliar * gerano);
            double caixaacumlado1 = double.Parse(txtvalorsist.Text) * (-1);
            if (txtkwh.Text != "")
            {
                for (int j = 0; j < 9; j++)
                {
                    auxiliar *= 1.1;
                    auxiliarano1 -= auxiliarano;
                    caixaacumlado1 += caixaanual;
                    caixaanual = (auxiliar * auxiliarano1);
                }
            }
            caixaacumlado1 += caixaanual - double.Parse(txtvalorsist.Text);

            caixaacumulado = caixaacumlado1.ToString();
            caixaacumulado = string.Format("{0:0,0.00}", caixaacumlado1);
        }
        private void Qtdorca()
        {
            switch (qtdorc)
            {
                case 0:
                    PaineisPrincipais(pnlorcamento0);

                    gbxuc1.Visible = false;
                    gbxuc1.Enabled = false;
                    gbxuc2.Visible = false;
                    gbxuc2.Enabled = false;
                    gbxuc3.Visible = false;
                    gbxuc3.Enabled = false;
                    gbxuc4.Visible = false;
                    gbxuc4.Enabled = false;

                    btnrecolhe1.Visible = false;
                    pnlpreencheuc.Visible = false;
                    pnl2uc.Visible = false;
                    pnl3uc.Visible = false;
                    pnl4uc.Visible = false;
                    pnlexp1.Visible = false;
                    pnlexp2.Visible = false;
                    pnlexp3.Visible = false;
                    pnlexp4.Visible = false;

                    break;
                case 1:
                    gbxuc1.Visible = true;
                    gbxuc1.Enabled = true;
                    gbxuc2.Visible = false;
                    gbxuc2.Enabled = false;
                    gbxuc3.Visible = false;
                    gbxuc3.Enabled = false;
                    gbxuc4.Visible = false;
                    gbxuc4.Enabled = false;

                    btnrecolhe1.Visible = false;
                    pnlpreencheuc.Visible = false;
                    pnl2uc.Visible = false;
                    pnl3uc.Visible = false;
                    pnl4uc.Visible = false;
                    pnlexp1.Visible = false;
                    pnlexp2.Visible = false;
                    pnlexp3.Visible = false;
                    pnlexp4.Visible = false;
                    break;
                case 2:
                    gbxuc1.Visible = true;
                    gbxuc1.Enabled = true;
                    gbxuc2.Visible = true;
                    gbxuc2.Enabled = true;
                    gbxuc3.Visible = false;
                    gbxuc3.Enabled = false;
                    gbxuc4.Visible = false;
                    gbxuc4.Enabled = false;

                    btnrecolhe1.Visible = true;
                    pnlpreencheuc.Visible = true;
                    pnl2uc.Visible = false;
                    pnl3uc.Visible = false;
                    pnl4uc.Visible = false;
                    pnlexp1.Visible = true;
                    pnlexp2.Visible = true;
                    pnlexp3.Visible = false;
                    pnlexp4.Visible = false;
                    break;
                case 3:
                    gbxuc1.Visible = true;
                    gbxuc1.Enabled = true;
                    gbxuc2.Visible = true;
                    gbxuc2.Enabled = true;
                    gbxuc3.Visible = true;
                    gbxuc3.Enabled = true;
                    gbxuc4.Visible = false;
                    gbxuc4.Enabled = false;

                    btnrecolhe1.Visible = true;
                    pnlpreencheuc.Visible = true;
                    pnl2uc.Visible = false;
                    pnl3uc.Visible = false;
                    pnl4uc.Visible = false;
                    pnlexp1.Visible = true;
                    pnlexp2.Visible = true;
                    pnlexp3.Visible = true;
                    pnlexp4.Visible = false;
                    break;
                case 4:
                    gbxuc1.Visible = true;
                    gbxuc1.Enabled = true;
                    gbxuc2.Visible = true;
                    gbxuc2.Enabled = true;
                    gbxuc3.Visible = true;
                    gbxuc3.Enabled = true;
                    gbxuc4.Visible = true;
                    gbxuc4.Enabled = true;

                    btnrecolhe1.Visible = true;
                    pnlpreencheuc.Visible = true;
                    pnl2uc.Visible = false;
                    pnl3uc.Visible = false;
                    pnl4uc.Visible = false;
                    pnlexp1.Visible = true;
                    pnlexp2.Visible = true;
                    pnlexp3.Visible = true;
                    pnlexp4.Visible = true;
                    break;
                default:
                    PaineisPrincipais(pnlorcamento0);
                    gbxuc1.Visible = false;
                    gbxuc1.Enabled = false;
                    gbxuc2.Visible = false;
                    gbxuc2.Enabled = false;
                    gbxuc3.Visible = false;
                    gbxuc3.Enabled = false;
                    gbxuc4.Visible = false;
                    gbxuc4.Enabled = false;

                    btnrecolhe1.Visible = false;
                    pnlpreencheuc.Visible = false;
                    pnl2uc.Visible = false;
                    pnl3uc.Visible = false;
                    pnl4uc.Visible = false;
                    pnlexp1.Visible = false;
                    pnlexp2.Visible = false;
                    pnlexp3.Visible = false;
                    pnlexp4.Visible = false;
                    break;
            }
        }
        private void CamposObrigatorios()
        {
            switch (qtdorc)
            {
                case 0:
                    break;
                case 1:
                    if (cbxclasse.Text != "" && cbxpadrao.Text != "" && txttarifa.Text != "")
                    {
                        PaineisPrincipais(pnlorcamento2);
                    }
                    else
                    {
                        MessageBox.Show("Preencha todos os campos obrigatórios");
                    }
                    break;
                case 2:
                    if (cbxclasse.Text != "" && cbxpadrao.Text != "" && txttarifa.Text != "" && cbxclasse2.Text != "" && cbxpadrao2.Text != "" &&
                        txttarifa2.Text != "" && txtidentificacaouc1.Text != "" && txtidentificacaouc2.Text != "")
                    {
                        PaineisPrincipais(pnlorcamento2);
                    }
                    else
                    {
                        MessageBox.Show("Preencha todos os campos obrigatórios");
                    }
                    break;
                case 3:
                    if (cbxclasse.Text != "" && cbxpadrao.Text != "" && txttarifa.Text != "" && cbxclasse2.Text != "" && cbxpadrao2.Text != "" &&
                        txttarifa2.Text != "" && txtidentificacaouc1.Text != "" && txtidentificacaouc2.Text != "" && cbxclasse3.Text != "" && cbxpadrao3.Text != "" &&
                        txttarifa3.Text != "" && txtidentificacaouc3.Text != "")
                    {
                        PaineisPrincipais(pnlorcamento2);
                    }
                    else
                    {
                        MessageBox.Show("Preencha todos os campos obrigatórios");
                    }
                    break;
                case 4:
                    if (cbxclasse.Text != "" && cbxpadrao.Text != "" && txttarifa.Text != "" && cbxclasse2.Text != "" && cbxpadrao2.Text != "" &&
                        txttarifa2.Text != "" && txtidentificacaouc1.Text != "" && txtidentificacaouc2.Text != "" && cbxclasse3.Text != "" && cbxpadrao3.Text != "" &&
                        txttarifa3.Text != "" && txtidentificacaouc3.Text != "" && cbxclasse4.Text != "" && cbxpadrao4.Text != "" &&
                        txttarifa4.Text != "" && txtidentificacaouc4.Text != "")
                    {
                        PaineisPrincipais(pnlorcamento2);
                    }
                    else
                    {
                        MessageBox.Show("Preencha todos os campos obrigatórios");
                    }
                    break;
                default:
                    break;
                }
            }
        private void Limpacampos()
        {
            chartgeracao.Series.Clear();
            chartretornofin.Series.Clear();
            txtjan.Text = "";
            txtfev.Text = "";
            txtmar.Text = "";
            txtabr.Text = "";
            txtmai.Text = "";
            txtjun.Text = "";
            txtjul.Text = "";
            txtago.Text = "";
            txtset.Text = "";
            txtout.Text = "";
            txtnov.Text = "";
            txtdez.Text = "";

            txtjan2.Text = "";
            txtfev2.Text = "";
            txtmar2.Text = "";
            txtabr2.Text = "";
            txtmai2.Text = "";
            txtjun2.Text = "";
            txtjul2.Text = "";
            txtago2.Text = "";
            txtset2.Text = "";
            txtout2.Text = "";
            txtnov2.Text = "";
            txtdez2.Text = "";

            txtjan3.Text = "";
            txtfev3.Text = "";
            txtmar3.Text = "";
            txtabr3.Text = "";
            txtmai3.Text = "";
            txtjun3.Text = "";
            txtjul3.Text = "";
            txtago3.Text = "";
            txtset3.Text = "";
            txtout3.Text = "";
            txtnov3.Text = "";
            txtdez3.Text = "";

            txtjan4.Text = "";
            txtfev4.Text = "";
            txtmar4.Text = "";
            txtabr4.Text = "";
            txtmai4.Text = "";
            txtjun4.Text = "";
            txtjul4.Text = "";
            txtago4.Text = "";
            txtset4.Text = "";
            txtout4.Text = "";
            txtnov4.Text = "";
            txtdez4.Text = "";

            txtidentificacaouc1.Text = "";
            cbxclasse.SelectedIndex = -1;
            cbxpadrao.SelectedIndex = -1;
            txttarifa.Text = "";
            txtidentificacaouc2.Text = "";
            cbxclasse2.SelectedIndex = -1;
            cbxpadrao2.SelectedIndex = -1;
            txttarifa2.Text = "";
            txtidentificacaouc3.Text = "";
            cbxclasse3.SelectedIndex = -1;
            cbxpadrao3.SelectedIndex = -1;
            txttarifa3.Text = "";
            txtidentificacaouc4.Text = "";
            cbxclasse4.SelectedIndex = -1;
            cbxpadrao4.SelectedIndex = -1;
            txttarifa4.Text = "";

            txtmedia.Text = "";
            txtpotenciasug.Text = "";
            txtanual.Text = "";

            txtnome.Text = "";
            txtcontato.Text = "";
            txtcontato.Mask = "(00) 0000-0000";
            txtcep.Text = "";
            txtendereco.Text = "";
            txtnumero.Text = "";
            txtbairro.Text = "";
            txtcidade.Text = "";
            txtkwh.Text = "";
            txtcontot.Text = "";
            txtqtdinv.Text = "";
            cbxestrutura.SelectedIndex = -1;
            txtqtdpaineis.Text = "";
            txtobs.Text = "";
            txtvalorequip.Text = "";
            txtvalorsist.Text = "";
            txtcustoinversor.Text = "";

            txtnomecliprojeto.Text = "";
            mtxtcpjcliprojeto.Text = string.Empty;
            mtxtcpjcliprojeto.Mask = "000.000.000-00";
            txtnumcliproj.Text = "";
            txtnuminstproj.Text = "";
            txtcargainstproj.Text = "";
            cbxfiltroprojeto.Text = "";
            cbxclasseproj.SelectedIndex = -1;
            cbxpadraoproj.SelectedIndex = -1;
            cbxdisjproj.SelectedIndex = -1;
            cbxtensoesatenproj.SelectedIndex = -1;
            cbxestruturaproj.SelectedIndex = -1;
            mtxtlatitudeproj.Text = "";
            mtxtlongitudeproj.Text = "";
            cbxstringboxproj.SelectedIndex = -1;
            txtarranjoproj.Text = "";
            txtqtdinstproj.Value = 1;

            txtnomecliente.Text = string.Empty;
            mtxtcpfcnpjcliente.Text = string.Empty;
            mtxtcpfcnpjcliente.Mask = "000.000.000-00";
            txtenderecocliente.Text = string.Empty;
            txtnumerocliente.Text = string.Empty;
            txtcomplementocliente.Text = string.Empty;
            txtbairrocliente.Text = string.Empty;
            mtxtcepcliente.Text = string.Empty;
            txtcidadecliente.Text = string.Empty;
            cbxufcliente.Text = string.Empty;
            txtemailcliente.Text = string.Empty;
            mtxttelefone.Text = string.Empty;
            mtxtcelularcliente.Text = string.Empty;
            txtqtdinvcliente.Value = 1;
            txtqtdmodmodcliente.Value = 1;
            txtconsmedcliente.Text = string.Empty;
            txtidentificacaocliente.Text = string.Empty;

            lblnomecliente.Text = "";
            lblidentificacaocliente.Text = "";
            lblcpfcliente.Text = "";
            lblenderecocliente.Text = "";
            lblcepcliente.Text = "";
            lblcidadeufcliente.Text = "";
            lbltelefonecliente.Text = "";
            lblcelularcliente.Text = "";
            lblemailcliente.Text = "";
            lblqtdinvcliente.Text = "";
            lblqtdmodcliente.Text = "";

            cbxmarcamodequipamentos.Text = "";
            cbxmodelomodequipamentos.Text = "";
            txtpotenciamodequipamentos.Text = "";
            txtcoefmodequipamentos.Text = "";
            cbxmaterialmodequipamentos.Text = "";
            cbxcelulasmodequipamentos.Text = "";
            txtcompmodequip.Text = "";
            txtlargmodequip.Text = "";
            txtgarantiamodequip.Text = "";
            txtreginmmodequip.Text = "";
            cbxmaterialmodequipamentos.SelectedIndex = -1;

            cbxmarcainvequip.Text = "";
            cbxmodinvequip.Text = "";
            txtpotenciainvequip.Text = "";
            cbxfasesinvequip.Text = "";
            cbxfasesinvequip.SelectedIndex = -1;
            cbxtensaoinvequip.Text = "";
            txteficienciainvequip.Text = "";
            txtgarantiainvequip.Text = "";
            txtreginminvequip.Text = "";
            txtqtdmpptinvequip.Value = 1;

            btnsalvarprojeto.Text = "Salvar";
            btnsalva.Text = "Salvar";
            btncadastrarinversor.Text = "Cadastrar";
            btncadastrarmodulo.Text = "Cadastrar";
            btnlimpainvequip.Text = "Limpar Campos";
            btnlimpamodequip.Text = "Limpar Campos";
            btnlimpaouexclui.Text = "Limpar Campos";

            cbxfiltromodulo.SelectedIndex = -1;
            cbxfiltrocliente.SelectedIndex = -1;
            cbxfiltroorcamento.SelectedIndex = -1;
            cbxfiltroprojeto.SelectedIndex = -1;
            cbxfiltroinv.SelectedIndex = -1;

            txtprocuracliente.Text = "O que você procura?";
            txtprocuracliente.ForeColor = Color.Silver;
            txtbusca.Text = "O que você procura?";
            txtbusca.ForeColor = Color.Silver;
            txtprocuramodulo.Text = "O que você procura?";
            txtprocuramodulo.ForeColor = Color.Silver;
            txtprocuraprojeto.Text = "O que você procura?";
            txtprocuraprojeto.ForeColor = Color.Silver;
            txtprocurainversor.Text = "O que você procura?";
            txtprocurainversor.ForeColor = Color.Silver;

            pnladicionacliente.Visible = false;
            pnlvisualizacli.Visible = false;
            pnlnovoproj1.Visible = false;
            pnlnovoproj2.Visible = false;
            pnlnovoproj3.Visible = false;
            pnlfinalizaprojeto.Visible = false;
            panel20.Visible = false;
            panel18.Visible = false;

            pgbstatusproj.Value = 0;
            pgbstatusorca.Value = 0;

            cbxtransformadorproj.SelectedIndex = 0;

            editando = false;

            func.LimpaTambem();

            rbtn0.Checked = true;

            txteqcomissao.Text = "";
            txteqinstalacao.Text = "";
            txteqoutros.Text = "";
            txteqprojeto.Text = "";
            txteqservico.Text = "";
            txteqtotal.Text = "";

            pnlsimulaorc.Visible = false;
            rbtnsim0.Checked = true;

            txtsimtot1.Text = "";
            txtsimserv1.Text = "";
            txtsimoutros1.Text = "";
            txtsiminsta1.Text = "";
            txtsimcom1.Text = "";
            txtsimproj1.Text = "";
            txtsimvalkit1.Text = "";
            txtsimger1.Text = "";
            txtsimqtdplaca1.Value = 0;
            txtsimqtdinv1.Value = 0;
            cbxsimmodmod1.SelectedIndex = -1;
            cbxsimmodinv1.SelectedIndex = -1;
            cbxsimparceiro1.SelectedIndex = -1;
            txtsimpotger1.Text = "";
            chksimop1.Checked = false;

            txtsimtot5.Text = "";
            txtsimserv5.Text = "";
            txtsimoutros5.Text = "";
            txtsiminsta5.Text = "";
            txtsimcom5.Text = "";
            txtsimproj5.Text = "";
            txtsimvalkit5.Text = "";
            txtsimger5.Text = "";
            txtsimqtdplaca5.Value = 0;
            txtsimqtdinv5.Value = 0;
            cbxsimmodmod5.SelectedIndex = -1;
            cbxsimmodinv5.SelectedIndex = -1;
            cbxsimparceiro5.SelectedIndex = -1;
            txtsimpotger5.Text = "";
            chksimop5.Checked = false;

            txtsimtot2.Text = "";
            txtsimserv2.Text = "";
            txtsimoutros2.Text = "";
            txtsiminsta2.Text = "";
            txtsimcom2.Text = "";
            txtsimproj2.Text = "";
            txtsimvalkit2.Text = "";
            txtsimger2.Text = "";
            txtsimqtdplaca2.Value = 0;
            txtsimqtdinv2.Value = 0;
            cbxsimmodmod2.SelectedIndex = -1;
            cbxsimmodinv2.SelectedIndex = -1;
            cbxsimparceiro2.SelectedIndex = -1;
            txtsimpotger2.Text = "";
            chksimop2.Checked = false;

            txtsimtot3.Text = "";
            txtsimserv3.Text = "";
            txtsimoutros3.Text = "";
            txtsiminsta3.Text = "";
            txtsimcom3.Text = "";
            txtsimproj3.Text = "";
            txtsimvalkit3.Text = "";
            txtsimger3.Text = "";
            txtsimqtdplaca3.Value = 0;
            txtsimqtdinv3.Value = 0;
            cbxsimmodmod3.SelectedIndex = -1;
            cbxsimmodinv3.SelectedIndex = -1;
            cbxsimparceiro3.SelectedIndex = -1;
            txtsimpotger3.Text = "";
            chksimop3.Checked = false;

            txtsimtot4.Text = "";
            txtsimserv4.Text = "";
            txtsimoutros4.Text = "";
            txtsiminsta4.Text = "";
            txtsimcom4.Text = "";
            txtsimproj4.Text = "";
            txtsimvalkit4.Text = "";
            txtsimger4.Text = "";
            txtsimqtdplaca4.Value = 0;
            txtsimqtdinv4.Value = 0;
            cbxsimmodmod4.SelectedIndex = -1;
            cbxsimmodinv4.SelectedIndex = -1;
            cbxsimparceiro4.SelectedIndex = -1;
            txtsimpotger4.Text = "";
            chksimop4.Checked = false;

            percas = 1;

            foto = "";
        }
        private void Preencheconsumo()
        {
            pnlpreencheuc.Visible = true;
            pnl2uc.Visible = false;
            pnl3uc.Visible = false;
            pnl4uc.Visible = false;
        }
        private void ControlaStrings()
        {
            switch (qtdorc)
            {
                case 0:
                    lblstringuc1.Text = "";
                    lblstringuc2.Text = "";
                    lblstringuc3.Text = "";
                    lblstringuc4.Text = "";
                    break;
                case 1:
                    lblstringuc1.Text = cbxclasse.Text + " " + cbxpadrao.Text;
                    lblstringuc2.Text = "";
                    lblstringuc3.Text = "";
                    lblstringuc4.Text = "";
                    break;
                case 2:
                    lblstringuc1.Text = txtidentificacaouc1.Text + " " + cbxclasse.Text + " " + cbxpadrao.Text;
                    lblstringuc2.Text = txtidentificacaouc2.Text + " " + cbxclasse2.Text + " " + cbxpadrao2.Text; 
                    lblstringuc3.Text = "";
                    lblstringuc4.Text = "";
                    break;
                case 3:
                    lblstringuc1.Text = txtidentificacaouc1.Text + " " + cbxclasse.Text + " " + cbxpadrao.Text;
                    lblstringuc2.Text = txtidentificacaouc2.Text + " " + cbxclasse2.Text + " " + cbxpadrao2.Text;
                    lblstringuc3.Text = txtidentificacaouc3.Text + " " + cbxclasse3.Text + " " + cbxpadrao3.Text;
                    lblstringuc4.Text = "";
                    break;
                case 4:
                    lblstringuc1.Text = txtidentificacaouc1.Text + " " + cbxclasse.Text + " " + cbxpadrao.Text;
                    lblstringuc2.Text = txtidentificacaouc2.Text + " " + cbxclasse2.Text + " " + cbxpadrao2.Text;
                    lblstringuc3.Text = txtidentificacaouc3.Text + " " + cbxclasse3.Text + " " + cbxpadrao3.Text;
                    lblstringuc4.Text = txtidentificacaouc4.Text + " " + cbxclasse4.Text + " " + cbxpadrao4.Text;
                    break;
                default:
                    lblstringuc1.Text = "";
                    lblstringuc2.Text = "";
                    lblstringuc3.Text = "";
                    lblstringuc4.Text = "";
                    break;
            }
        }
        private bool Dimensionamento()
        {
            switch (qtdorc)
            {
                case 1:
                    if ((txtjan.Text != "" && txtago.Text != "" && txtfev.Text != "" && txtmar.Text != "" && txtabr.Text != "" && txtmai.Text != "" && txtjun.Text != "" 
                        && txtjul.Text != "" && txtset.Text != "" && txtout.Text != "" && txtnov.Text != "" && txtdez.Text != "") || editando == true)
                    {
                        double mensal = 0;

                        //Consumo
                        double total;
                        total = orcamento.CalculaConsumoTotal(double.Parse(txtjan.Text), double.Parse(txtfev.Text), double.Parse(txtmar.Text), double.Parse(txtabr.Text),
                            double.Parse(txtmai.Text), double.Parse(txtjun.Text), double.Parse(txtjul.Text), double.Parse(txtago.Text), double.Parse(txtset.Text), double.Parse(txtout.Text),
                            double.Parse(txtnov.Text), double.Parse(txtdez.Text));

                        //Consumo com Tarifa
                        double resultado = orcamento.CalculaConsumoComTarifa(double.Parse(txtjan.Text), double.Parse(txtfev.Text), double.Parse(txtmar.Text), double.Parse(txtabr.Text),
                            double.Parse(txtmai.Text), double.Parse(txtjun.Text), double.Parse(txtjul.Text), double.Parse(txtago.Text), double.Parse(txtset.Text), double.Parse(txtout.Text),
                            double.Parse(txtnov.Text), double.Parse(txtdez.Text), double.Parse(txtkwh.Text));

                        int disp = orcamento.UmaUC(cbxpadrao.Text);

                        double potencia = Math.Ceiling((((((total) / 365) / 4.87) * 1000) / 0.84) / 335) * 335 / 1000;
                        mensal = total / 12;

                        string textototal = string.Format("{0:0,0.00}", total);
                        string textomensal = string.Format("{0:0,0.00}", mensal);
                        string textopotencia = string.Format("{0:0,0.0}", potencia);

                        double tarifex = resultado / total;
                        txtkwh.Text = tarifex.ToString("0.000000000");
                        txtanual.Text = textototal;
                        string texte = string.Format("{0:0}", double.Parse(txtanual.Text));
                        txtcontot.Text = texte;
                        txtmedia.Text = textomensal;
                        txtpotenciasug.Text = textopotencia;
                        return true;
                    }
                    else
                    {
                        MessageBox.Show("Preencha todos os meses!");
                        return false;
                    }
                    break;
                case 2:
                    if (txtjan.Text != "" && txtago.Text != "" && txtfev.Text != "" && txtmar.Text != "" && txtabr.Text != "" && txtmai.Text != "" && txtjun.Text != "" 
                        && txtjul.Text != "" && txtset.Text != "" && txtout.Text != "" && txtnov.Text != "" && txtdez.Text != "" && txtjan2.Text != "" && txtago2.Text != "" 
                        && txtfev2.Text != "" && txtmar2.Text != "" && txtabr2.Text != "" && txtmai2.Text != "" && txtjun2.Text != "" && txtjul2.Text != "" && txtset2.Text != "" 
                        && txtout2.Text != "" && txtnov2.Text != "" && txtdez2.Text != "")
                    {
                        double mensal = 0;

                        //Consumo
                        double total1;
                        total1 = orcamento.CalculaConsumoTotal(double.Parse(txtjan.Text), double.Parse(txtfev.Text), double.Parse(txtmar.Text), double.Parse(txtabr.Text),
                            double.Parse(txtmai.Text), double.Parse(txtjun.Text), double.Parse(txtjul.Text), double.Parse(txtago.Text), double.Parse(txtset.Text), double.Parse(txtout.Text),
                            double.Parse(txtnov.Text), double.Parse(txtdez.Text));
                        double total2;
                        total2 = orcamento.CalculaConsumoTotal(double.Parse(txtjan2.Text), double.Parse(txtfev2.Text), double.Parse(txtmar2.Text), double.Parse(txtabr2.Text),
                            double.Parse(txtmai2.Text), double.Parse(txtjun2.Text), double.Parse(txtjul2.Text), double.Parse(txtago2.Text), double.Parse(txtset2.Text), double.Parse(txtout2.Text),
                            double.Parse(txtnov2.Text), double.Parse(txtdez2.Text));
                        total = total1 + total2;

                        //Consumo com Tarifa
                        double resultado1 = orcamento.CalculaConsumoComTarifa(double.Parse(txtjan.Text), double.Parse(txtfev.Text), double.Parse(txtmar.Text), double.Parse(txtabr.Text),
                            double.Parse(txtmai.Text), double.Parse(txtjun.Text), double.Parse(txtjul.Text), double.Parse(txtago.Text), double.Parse(txtset.Text), double.Parse(txtout.Text),
                            double.Parse(txtnov.Text), double.Parse(txtdez.Text), double.Parse(txttarifa.Text));
                        double resultado2 = orcamento.CalculaConsumoComTarifa(double.Parse(txtjan2.Text), double.Parse(txtfev2.Text), double.Parse(txtmar2.Text), double.Parse(txtabr2.Text),
                            double.Parse(txtmai2.Text), double.Parse(txtjun2.Text), double.Parse(txtjul2.Text), double.Parse(txtago2.Text), double.Parse(txtset2.Text), double.Parse(txtout2.Text),
                            double.Parse(txtnov2.Text), double.Parse(txtdez2.Text), double.Parse(txttarifa2.Text));
                        resultado = resultado1 + resultado2;

                        int disp = orcamento.DuasUC(cbxpadrao.Text, cbxpadrao2.Text);

                        double potencia = Math.Ceiling((((((total) / 365) / 4.87) * 1000) / 0.84) / 335) * 335 / 1000;
                        mensal = total / 12;

                        string textototal = string.Format("{0:0,0.00}", total);
                        string textomensal = string.Format("{0:0,0.00}", mensal);
                        string textopotencia = string.Format("{0:0,0.0}", potencia);

                        double tarifex = resultado / total;
                        txtkwh.Text = tarifex.ToString("0.000000000");
                        txtanual.Text = textototal;
                        string texte = string.Format("{0:0}", double.Parse(txtanual.Text));
                        txtcontot.Text = texte;
                        txtmedia.Text = textomensal;
                        txtpotenciasug.Text = textopotencia;
                        return true;
                    }
                    else
                    {
                        MessageBox.Show("Preencha todos os meses!");
                        return false;
                    }
                    break;
                case 3:
                    if (txtjan.Text != "" && txtago.Text != "" && txtfev.Text != "" && txtmar.Text != "" && txtabr.Text != "" && txtmai.Text != "" && txtjun.Text != ""
                        && txtjul.Text != "" && txtset.Text != "" && txtout.Text != "" && txtnov.Text != "" && txtdez.Text != "" && txtjan2.Text != "" && txtago2.Text != ""
                        && txtfev2.Text != "" && txtmar2.Text != "" && txtabr2.Text != "" && txtmai2.Text != "" && txtjun2.Text != "" && txtjul2.Text != "" && txtset2.Text != ""
                        && txtout2.Text != "" && txtnov2.Text != "" && txtdez2.Text != "" && txtjan3.Text != "" && txtago3.Text != "" && txtfev3.Text != "" && txtmar3.Text != "" 
                        && txtabr3.Text != "" && txtmai3.Text != "" && txtjun3.Text != "" && txtjul3.Text != "" && txtset3.Text != "" && txtout3.Text != "" && txtnov3.Text != "" 
                        && txtdez3.Text != "")
                    {
                        double mensal = 0;

                        //Consumo
                        double total1;
                        total1 = orcamento.CalculaConsumoTotal(double.Parse(txtjan.Text), double.Parse(txtfev.Text), double.Parse(txtmar.Text), double.Parse(txtabr.Text),
                            double.Parse(txtmai.Text), double.Parse(txtjun.Text), double.Parse(txtjul.Text), double.Parse(txtago.Text), double.Parse(txtset.Text), double.Parse(txtout.Text),
                            double.Parse(txtnov.Text), double.Parse(txtdez.Text));
                        double total2;
                        total2 = orcamento.CalculaConsumoTotal(double.Parse(txtjan2.Text), double.Parse(txtfev2.Text), double.Parse(txtmar2.Text), double.Parse(txtabr2.Text),
                            double.Parse(txtmai2.Text), double.Parse(txtjun2.Text), double.Parse(txtjul2.Text), double.Parse(txtago2.Text), double.Parse(txtset2.Text), double.Parse(txtout2.Text),
                            double.Parse(txtnov2.Text), double.Parse(txtdez2.Text));
                        double total3;
                        total3 = orcamento.CalculaConsumoTotal(double.Parse(txtjan3.Text), double.Parse(txtfev3.Text), double.Parse(txtmar3.Text), double.Parse(txtabr3.Text),
                            double.Parse(txtmai3.Text), double.Parse(txtjun3.Text), double.Parse(txtjul3.Text), double.Parse(txtago3.Text), double.Parse(txtset3.Text), double.Parse(txtout3.Text),
                            double.Parse(txtnov3.Text), double.Parse(txtdez3.Text));
                        total = total1 + total2 + total3;

                        //Consumo com Tarifa
                        double resultado1 = orcamento.CalculaConsumoComTarifa(double.Parse(txtjan.Text), double.Parse(txtfev.Text), double.Parse(txtmar.Text), double.Parse(txtabr.Text),
                            double.Parse(txtmai.Text), double.Parse(txtjun.Text), double.Parse(txtjul.Text), double.Parse(txtago.Text), double.Parse(txtset.Text), double.Parse(txtout.Text),
                            double.Parse(txtnov.Text), double.Parse(txtdez.Text), double.Parse(txttarifa.Text));
                        double resultado2 = orcamento.CalculaConsumoComTarifa(double.Parse(txtjan2.Text), double.Parse(txtfev2.Text), double.Parse(txtmar2.Text), double.Parse(txtabr2.Text),
                            double.Parse(txtmai2.Text), double.Parse(txtjun2.Text), double.Parse(txtjul2.Text), double.Parse(txtago2.Text), double.Parse(txtset2.Text), double.Parse(txtout2.Text),
                            double.Parse(txtnov2.Text), double.Parse(txtdez2.Text), double.Parse(txttarifa2.Text));
                        double resultado3 = orcamento.CalculaConsumoComTarifa(double.Parse(txtjan3.Text), double.Parse(txtfev3.Text), double.Parse(txtmar3.Text), double.Parse(txtabr3.Text),
                            double.Parse(txtmai3.Text), double.Parse(txtjun3.Text), double.Parse(txtjul3.Text), double.Parse(txtago3.Text), double.Parse(txtset3.Text), double.Parse(txtout3.Text),
                            double.Parse(txtnov3.Text), double.Parse(txtdez3.Text), double.Parse(txttarifa3.Text));
                        resultado = resultado1 + resultado2 + resultado3;

                        int disp = orcamento.TresUC(cbxpadrao.Text, cbxpadrao2.Text, cbxpadrao3.Text);

                        double potencia = Math.Ceiling((((((total) / 365) / 4.87) * 1000) / 0.84) / 335) * 335 / 1000;
                        mensal = total / 12;

                        string textototal = string.Format("{0:0,0.00}", total);
                        string textomensal = string.Format("{0:0,0.00}", mensal);
                        string textopotencia = string.Format("{0:0,0.0}", potencia);

                        double tarifex = resultado / total;
                        txtkwh.Text = tarifex.ToString("0.000000000");
                        txtanual.Text = textototal;
                        string texte = string.Format("{0:0}", double.Parse(txtanual.Text));
                        txtcontot.Text = texte;
                        txtmedia.Text = textomensal;
                        txtpotenciasug.Text = textopotencia;
                        return true;
                    }
                    else
                    {
                        MessageBox.Show("Preencha todos os meses!");
                        return false;
                    }
                    break;
                case 4:
                    if (txtjan.Text != "" && txtago.Text != "" && txtfev.Text != "" && txtmar.Text != "" && txtabr.Text != "" && txtmai.Text != "" && txtjun.Text != ""
                        && txtjul.Text != "" && txtset.Text != "" && txtout.Text != "" && txtnov.Text != "" && txtdez.Text != "" && txtjan2.Text != "" && txtago2.Text != ""
                        && txtfev2.Text != "" && txtmar2.Text != "" && txtabr2.Text != "" && txtmai2.Text != "" && txtjun2.Text != "" && txtjul2.Text != "" && txtset2.Text != ""
                        && txtout2.Text != "" && txtnov2.Text != "" && txtdez2.Text != "" && txtjan3.Text != "" && txtago3.Text != "" && txtfev3.Text != "" && txtmar3.Text != ""
                        && txtabr3.Text != "" && txtmai3.Text != "" && txtjun3.Text != "" && txtjul3.Text != "" && txtset3.Text != "" && txtout3.Text != "" && txtnov3.Text != ""
                        && txtdez3.Text != "" && txtjan4.Text != "" && txtago4.Text != "" && txtfev4.Text != "" && txtmar4.Text != "" && txtabr4.Text != "" && txtmai4.Text != "" 
                        && txtjun4.Text != "" && txtjul4.Text != "" && txtset4.Text != "" && txtout4.Text != "" && txtnov4.Text != "" && txtdez4.Text != "")
                    {
                        double mensal = 0;

                        //Consumo
                        double total1;
                        total1 = orcamento.CalculaConsumoTotal(double.Parse(txtjan.Text), double.Parse(txtfev.Text), double.Parse(txtmar.Text), double.Parse(txtabr.Text),
                            double.Parse(txtmai.Text), double.Parse(txtjun.Text), double.Parse(txtjul.Text), double.Parse(txtago.Text), double.Parse(txtset.Text), double.Parse(txtout.Text),
                            double.Parse(txtnov.Text), double.Parse(txtdez.Text));
                        double total2;
                        total2 = orcamento.CalculaConsumoTotal(double.Parse(txtjan2.Text), double.Parse(txtfev2.Text), double.Parse(txtmar2.Text), double.Parse(txtabr2.Text),
                            double.Parse(txtmai2.Text), double.Parse(txtjun2.Text), double.Parse(txtjul2.Text), double.Parse(txtago2.Text), double.Parse(txtset2.Text), double.Parse(txtout2.Text),
                            double.Parse(txtnov2.Text), double.Parse(txtdez2.Text));
                        double total3;
                        total3 = orcamento.CalculaConsumoTotal(double.Parse(txtjan3.Text), double.Parse(txtfev3.Text), double.Parse(txtmar3.Text), double.Parse(txtabr3.Text),
                            double.Parse(txtmai3.Text), double.Parse(txtjun3.Text), double.Parse(txtjul3.Text), double.Parse(txtago3.Text), double.Parse(txtset3.Text), double.Parse(txtout3.Text),
                            double.Parse(txtnov3.Text), double.Parse(txtdez3.Text));
                        double total4;
                        total4 = orcamento.CalculaConsumoTotal(double.Parse(txtjan4.Text), double.Parse(txtfev4.Text), double.Parse(txtmar4.Text), double.Parse(txtabr4.Text),
                            double.Parse(txtmai4.Text), double.Parse(txtjun4.Text), double.Parse(txtjul4.Text), double.Parse(txtago4.Text), double.Parse(txtset4.Text), double.Parse(txtout4.Text),
                            double.Parse(txtnov4.Text), double.Parse(txtdez4.Text));
                        total = total1 + total2 + total3 + total4;

                        //Consumo com Tarifa
                        double resultado1 = orcamento.CalculaConsumoComTarifa(double.Parse(txtjan.Text), double.Parse(txtfev.Text), double.Parse(txtmar.Text), double.Parse(txtabr.Text),
                            double.Parse(txtmai.Text), double.Parse(txtjun.Text), double.Parse(txtjul.Text), double.Parse(txtago.Text), double.Parse(txtset.Text), double.Parse(txtout.Text),
                            double.Parse(txtnov.Text), double.Parse(txtdez.Text), double.Parse(txttarifa.Text));
                        double resultado2 = orcamento.CalculaConsumoComTarifa(double.Parse(txtjan2.Text), double.Parse(txtfev2.Text), double.Parse(txtmar2.Text), double.Parse(txtabr2.Text),
                            double.Parse(txtmai2.Text), double.Parse(txtjun2.Text), double.Parse(txtjul2.Text), double.Parse(txtago2.Text), double.Parse(txtset2.Text), double.Parse(txtout2.Text),
                            double.Parse(txtnov2.Text), double.Parse(txtdez2.Text), double.Parse(txttarifa2.Text));
                        double resultado3 = orcamento.CalculaConsumoComTarifa(double.Parse(txtjan3.Text), double.Parse(txtfev3.Text), double.Parse(txtmar3.Text), double.Parse(txtabr3.Text),
                            double.Parse(txtmai3.Text), double.Parse(txtjun3.Text), double.Parse(txtjul3.Text), double.Parse(txtago3.Text), double.Parse(txtset3.Text), double.Parse(txtout3.Text),
                            double.Parse(txtnov3.Text), double.Parse(txtdez3.Text), double.Parse(txttarifa3.Text));
                        double resultado4 = orcamento.CalculaConsumoComTarifa(double.Parse(txtjan4.Text), double.Parse(txtfev4.Text), double.Parse(txtmar4.Text), double.Parse(txtabr4.Text),
                            double.Parse(txtmai4.Text), double.Parse(txtjun4.Text), double.Parse(txtjul4.Text), double.Parse(txtago4.Text), double.Parse(txtset4.Text), double.Parse(txtout4.Text),
                            double.Parse(txtnov4.Text), double.Parse(txtdez4.Text), double.Parse(txttarifa4.Text));
                        resultado = resultado1 + resultado2 + resultado3 + resultado4;

                        int disp = orcamento.QuatroUC(cbxpadrao.Text, cbxpadrao2.Text, cbxpadrao3.Text, cbxpadrao4.Text);

                        double potencia = Math.Ceiling((((((total) / 365) / 4.87) * 1000) / 0.84) / 335) * 335 / 1000;
                        mensal = total / 12;

                        string textototal = string.Format("{0:0,0.00}", total);
                        string textomensal = string.Format("{0:0,0.00}", mensal);
                        string textopotencia = string.Format("{0:0,0.0}", potencia);

                        double tarifex = resultado / total;
                        txtkwh.Text = tarifex.ToString("0.000000000");
                        txtanual.Text = textototal;
                        string texte = string.Format("{0:0}", double.Parse(txtanual.Text));
                        txtcontot.Text = texte;
                        txtmedia.Text = textomensal;
                        txtpotenciasug.Text = textopotencia;
                        return true;
                    }
                    else
                    {
                        MessageBox.Show("Preencha todos os meses!");
                        return false;
                    }
                    break;
                default:
                    MessageBox.Show("Verifique se todos os campos foram preenchidos corretamente!");
                    return false;
                    break;
            }
            
        }
        private void SomaConsumo()
        {
            tjan = 0; tfev = 0; tmar = 0; tabr = 0; tmai = 0; tjun = 0; somadisp=0;
            tjul = 0; tago = 0; tset = 0; tout = 0; tnov = 0; tdez = 0;
            switch (qtdorc)
            { 
                case 1:
                    somadisp = orcamento.UmaUC(cbxpadrao.Text);
                    tjan = Int32.Parse(txtjan.Text) - somadisp;
                    tfev = Int32.Parse(txtfev.Text) - somadisp;
                    tmar = Int32.Parse(txtmar.Text) - somadisp;
                    tabr = Int32.Parse(txtabr.Text) - somadisp;
                    tmai = Int32.Parse(txtmai.Text) - somadisp;
                    tjun = Int32.Parse(txtjun.Text) - somadisp;
                    tjul = Int32.Parse(txtjul.Text) - somadisp;
                    tago = Int32.Parse(txtago.Text) - somadisp;
                    tset = Int32.Parse(txtset.Text) - somadisp;
                    tout = Int32.Parse(txtout.Text) - somadisp;
                    tnov = Int32.Parse(txtnov.Text) - somadisp;
                    tdez = Int32.Parse(txtdez.Text) - somadisp;

                    break;
                case 2:
                    somadisp = orcamento.UmaUC(cbxpadrao.Text);
                    tjan = Int32.Parse(txtjan.Text) - somadisp;
                    tfev = Int32.Parse(txtfev.Text) - somadisp;
                    tmar = Int32.Parse(txtmar.Text) - somadisp;
                    tabr = Int32.Parse(txtabr.Text) - somadisp;
                    tmai = Int32.Parse(txtmai.Text) - somadisp;
                    tjun = Int32.Parse(txtjun.Text) - somadisp;
                    tjul = Int32.Parse(txtjul.Text) - somadisp;
                    tago = Int32.Parse(txtago.Text) - somadisp;
                    tset = Int32.Parse(txtset.Text) - somadisp;
                    tout = Int32.Parse(txtout.Text) - somadisp;
                    tnov = Int32.Parse(txtnov.Text) - somadisp;
                    tdez = Int32.Parse(txtdez.Text) - somadisp;

                    somadisp = orcamento.UmaUC(cbxpadrao2.Text);
                    tjan += Int32.Parse(txtjan2.Text) - somadisp;
                    tfev += Int32.Parse(txtfev2.Text) - somadisp;
                    tmar += Int32.Parse(txtmar2.Text) - somadisp;
                    tabr += Int32.Parse(txtabr2.Text) - somadisp;
                    tmai += Int32.Parse(txtmai2.Text) - somadisp;
                    tjun += Int32.Parse(txtjun2.Text) - somadisp;
                    tjul += Int32.Parse(txtjul2.Text) - somadisp;
                    tago += Int32.Parse(txtago2.Text) - somadisp;
                    tset += Int32.Parse(txtset2.Text) - somadisp;
                    tout += Int32.Parse(txtout2.Text) - somadisp;
                    tnov += Int32.Parse(txtnov2.Text) - somadisp;
                    tdez += Int32.Parse(txtdez2.Text) - somadisp;
                    somadisp = orcamento.DuasUC(cbxpadrao.Text, cbxpadrao2.Text);
                    break;
                case 3:
                    somadisp = orcamento.UmaUC(cbxpadrao.Text);
                    tjan = Int32.Parse(txtjan.Text) - somadisp;
                    tfev = Int32.Parse(txtfev.Text) - somadisp;
                    tmar = Int32.Parse(txtmar.Text) - somadisp;
                    tabr = Int32.Parse(txtabr.Text) - somadisp;
                    tmai = Int32.Parse(txtmai.Text) - somadisp;
                    tjun = Int32.Parse(txtjun.Text) - somadisp;
                    tjul = Int32.Parse(txtjul.Text) - somadisp;
                    tago = Int32.Parse(txtago.Text) - somadisp;
                    tset = Int32.Parse(txtset.Text) - somadisp;
                    tout = Int32.Parse(txtout.Text) - somadisp;
                    tnov = Int32.Parse(txtnov.Text) - somadisp;
                    tdez = Int32.Parse(txtdez.Text) - somadisp;

                    somadisp = orcamento.UmaUC(cbxpadrao2.Text);
                    tjan += Int32.Parse(txtjan2.Text) - somadisp;
                    tfev += Int32.Parse(txtfev2.Text) - somadisp;
                    tmar += Int32.Parse(txtmar2.Text) - somadisp;
                    tabr += Int32.Parse(txtabr2.Text) - somadisp;
                    tmai += Int32.Parse(txtmai2.Text) - somadisp;
                    tjun += Int32.Parse(txtjun2.Text) - somadisp;
                    tjul += Int32.Parse(txtjul2.Text) - somadisp;
                    tago += Int32.Parse(txtago2.Text) - somadisp;
                    tset += Int32.Parse(txtset2.Text) - somadisp;
                    tout += Int32.Parse(txtout2.Text) - somadisp;
                    tnov += Int32.Parse(txtnov2.Text) - somadisp;
                    tdez += Int32.Parse(txtdez2.Text) - somadisp;

                    somadisp = orcamento.UmaUC(cbxpadrao3.Text);
                    tjan += Int32.Parse(txtjan3.Text) - somadisp;
                    tfev += Int32.Parse(txtfev3.Text) - somadisp;
                    tmar += Int32.Parse(txtmar3.Text) - somadisp;
                    tabr += Int32.Parse(txtabr3.Text) - somadisp;
                    tmai += Int32.Parse(txtmai3.Text) - somadisp;
                    tjun += Int32.Parse(txtjun3.Text) - somadisp;
                    tjul += Int32.Parse(txtjul3.Text) - somadisp;
                    tago += Int32.Parse(txtago3.Text) - somadisp;
                    tset += Int32.Parse(txtset3.Text) - somadisp;
                    tout += Int32.Parse(txtout3.Text) - somadisp;
                    tnov += Int32.Parse(txtnov3.Text) - somadisp;
                    tdez += Int32.Parse(txtdez3.Text) - somadisp;
                    somadisp = orcamento.TresUC(cbxpadrao.Text, cbxpadrao2.Text, cbxpadrao3.Text);
                    break;
                case 4:
                    somadisp = orcamento.UmaUC(cbxpadrao.Text);
                    tjan = Int32.Parse(txtjan.Text) - somadisp;
                    tfev = Int32.Parse(txtfev.Text) - somadisp;
                    tmar = Int32.Parse(txtmar.Text) - somadisp;
                    tabr = Int32.Parse(txtabr.Text) - somadisp;
                    tmai = Int32.Parse(txtmai.Text) - somadisp;
                    tjun = Int32.Parse(txtjun.Text) - somadisp;
                    tjul = Int32.Parse(txtjul.Text) - somadisp;
                    tago = Int32.Parse(txtago.Text) - somadisp;
                    tset = Int32.Parse(txtset.Text) - somadisp;
                    tout = Int32.Parse(txtout.Text) - somadisp;
                    tnov = Int32.Parse(txtnov.Text) - somadisp;
                    tdez = Int32.Parse(txtdez.Text) - somadisp;

                    somadisp = orcamento.UmaUC(cbxpadrao2.Text);
                    tjan += Int32.Parse(txtjan2.Text) - somadisp;
                    tfev += Int32.Parse(txtfev2.Text) - somadisp;
                    tmar += Int32.Parse(txtmar2.Text) - somadisp;
                    tabr += Int32.Parse(txtabr2.Text) - somadisp;
                    tmai += Int32.Parse(txtmai2.Text) - somadisp;
                    tjun += Int32.Parse(txtjun2.Text) - somadisp;
                    tjul += Int32.Parse(txtjul2.Text) - somadisp;
                    tago += Int32.Parse(txtago2.Text) - somadisp;
                    tset += Int32.Parse(txtset2.Text) - somadisp;
                    tout += Int32.Parse(txtout2.Text) - somadisp;
                    tnov += Int32.Parse(txtnov2.Text) - somadisp;
                    tdez += Int32.Parse(txtdez2.Text) - somadisp;

                    somadisp = orcamento.UmaUC(cbxpadrao3.Text);
                    tjan += Int32.Parse(txtjan3.Text) - somadisp;
                    tfev += Int32.Parse(txtfev3.Text) - somadisp;
                    tmar += Int32.Parse(txtmar3.Text) - somadisp;
                    tabr += Int32.Parse(txtabr3.Text) - somadisp;
                    tmai += Int32.Parse(txtmai3.Text) - somadisp;
                    tjun += Int32.Parse(txtjun3.Text) - somadisp;
                    tjul += Int32.Parse(txtjul3.Text) - somadisp;
                    tago += Int32.Parse(txtago3.Text) - somadisp;
                    tset += Int32.Parse(txtset3.Text) - somadisp;
                    tout += Int32.Parse(txtout3.Text) - somadisp;
                    tnov += Int32.Parse(txtnov3.Text) - somadisp;
                    tdez += Int32.Parse(txtdez3.Text) - somadisp;

                    somadisp = orcamento.UmaUC(cbxpadrao4.Text);
                    tjan += Int32.Parse(txtjan4.Text) - somadisp;
                    tfev += Int32.Parse(txtfev4.Text) - somadisp;
                    tmar += Int32.Parse(txtmar4.Text) - somadisp;
                    tabr += Int32.Parse(txtabr4.Text) - somadisp;
                    tmai += Int32.Parse(txtmai4.Text) - somadisp;
                    tjun += Int32.Parse(txtjun4.Text) - somadisp;
                    tjul += Int32.Parse(txtjul4.Text) - somadisp;
                    tago += Int32.Parse(txtago4.Text) - somadisp;
                    tset += Int32.Parse(txtset4.Text) - somadisp;
                    tout += Int32.Parse(txtout4.Text) - somadisp;
                    tnov += Int32.Parse(txtnov4.Text) - somadisp;
                    tdez += Int32.Parse(txtdez4.Text) - somadisp;
                    somadisp = orcamento.QuatroUC(cbxpadrao.Text, cbxpadrao2.Text, cbxpadrao3.Text, cbxpadrao4.Text);
                    break;
                default:
                    tjan = 0; tfev = 0; tmar = 0; tabr = 0; tmai = 0; tjun = 0; somadisp = 0; 
                    tjul = 0; tago = 0; tset = 0; tout = 0; tnov = 0; tdez = 0;
                    break;
            }
        }
        private void Graficos()
        {
            var retornofin = chartretornofin.ChartAreas[0];
            retornofin.AxisX.IntervalType = System.Windows.Forms.DataVisualization.Charting.DateTimeIntervalType.Number;

            var gera = chartgeracao.ChartAreas[0];
            gera.AxisX.IntervalType = System.Windows.Forms.DataVisualization.Charting.DateTimeIntervalType.Number;

            retornofin.AxisX.LabelStyle.Format = "";
            retornofin.AxisY.LabelStyle.Format = "";
            retornofin.AxisY.LabelStyle.IsEndLabelVisible = true;

            gera.AxisX.LabelStyle.Format = "";
            gera.AxisY.LabelStyle.Format = "";
            gera.AxisY.LabelStyle.IsEndLabelVisible = true;

            retornofin.AxisX.Minimum = 1;
            retornofin.AxisX.Maximum = 10;

            gera.AxisX.Minimum = 0;
            gera.AxisX.Maximum = 13;

            chartgeracao.Series.Add("Consumo");
            chartgeracao.Series.Add("Geração");

            chartgeracao.Series["Consumo"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
            chartgeracao.Series["Geração"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
            chartgeracao.Series["Consumo"].Color = Color.DarkSeaGreen;
            chartgeracao.Series["Geração"].Color = Color.Blue;
            chartgeracao.Series[0].IsVisibleInLegend = false;

            chartgeracao.Series["Consumo"].Points.AddXY("Jan", double.Parse(txtjan.Text));
            chartgeracao.Series["Consumo"].Points.AddXY("Fev", double.Parse(txtjan.Text));
            chartgeracao.Series["Consumo"].Points.AddXY("Mar", double.Parse(txtjan.Text));
            chartgeracao.Series["Consumo"].Points.AddXY("Abr", double.Parse(txtjan.Text));
            chartgeracao.Series["Consumo"].Points.AddXY("Mai", double.Parse(txtjan.Text));
            chartgeracao.Series["Consumo"].Points.AddXY("Jun", double.Parse(txtjan.Text));
            chartgeracao.Series["Consumo"].Points.AddXY("Jul", double.Parse(txtjan.Text));
            chartgeracao.Series["Consumo"].Points.AddXY("Ago", double.Parse(txtjan.Text));
            chartgeracao.Series["Consumo"].Points.AddXY("Set", double.Parse(txtjan.Text));
            chartgeracao.Series["Consumo"].Points.AddXY("Out", double.Parse(txtjan.Text));
            chartgeracao.Series["Consumo"].Points.AddXY("Nov", double.Parse(txtjan.Text));
            chartgeracao.Series["Consumo"].Points.AddXY("Dez", double.Parse(txtjan.Text));


            if (rbtn5.Checked)
            {
                chartgeracao.Series["Geração"].Points.AddXY("Janeiro", potenciagerada * 5.48 * 31 * 0.83 * 0.95);
                chartgeracao.Series["Geração"].Points.AddXY("Fevereiro", potenciagerada * 5.7 * 29 * 0.83 * 0.95);
                chartgeracao.Series["Geração"].Points.AddXY("Março", potenciagerada * 4.85 * 31 * 0.83 * 0.95);
                chartgeracao.Series["Geração"].Points.AddXY("Abril", potenciagerada * 4.59 * 30 * 0.83 * 0.95);
                chartgeracao.Series["Geração"].Points.AddXY("Maio", potenciagerada * 3.95 * 31 * 0.83 * 0.95);
                chartgeracao.Series["Geração"].Points.AddXY("Junho", potenciagerada * 3.76 * 30 * 0.83 * 0.95);
                chartgeracao.Series["Geração"].Points.AddXY("Julho", potenciagerada * 4.01 * 31 * 0.83 * 0.95);
                chartgeracao.Series["Geração"].Points.AddXY("Agosto", potenciagerada * 4.86 * 31 * 0.83 * 0.95);
                chartgeracao.Series["Geração"].Points.AddXY("Setembro", potenciagerada * 5.08 * 30 * 0.83 * 0.95);
                chartgeracao.Series["Geração"].Points.AddXY("Outubro", potenciagerada * 5.37 * 31 * 0.83 * 0.95);
                chartgeracao.Series["Geração"].Points.AddXY("Novembro", potenciagerada * 5.22 * 30 * 0.83 * 0.95);
                chartgeracao.Series["Geração"].Points.AddXY("Dezembro", potenciagerada * 5.59 * 31 * 0.83 * 0.95);
            }
            else if (rbtn7.Checked)
            {
                chartgeracao.Series["Geração"].Points.AddXY("Janeiro", potenciagerada * 5.48 * 31 * 0.83 * 0.93);
                chartgeracao.Series["Geração"].Points.AddXY("Fevereiro", potenciagerada * 5.7 * 29 * 0.83 * 0.93);
                chartgeracao.Series["Geração"].Points.AddXY("Março", potenciagerada * 4.85 * 31 * 0.83 * 0.93);
                chartgeracao.Series["Geração"].Points.AddXY("Abril", potenciagerada * 4.59 * 30 * 0.83 * 0.93);
                chartgeracao.Series["Geração"].Points.AddXY("Maio", potenciagerada * 3.95 * 31 * 0.83 * 0.93);
                chartgeracao.Series["Geração"].Points.AddXY("Junho", potenciagerada * 3.76 * 30 * 0.83 * 0.93);
                chartgeracao.Series["Geração"].Points.AddXY("Julho", potenciagerada * 4.01 * 31 * 0.83 * 0.93);
                chartgeracao.Series["Geração"].Points.AddXY("Agosto", potenciagerada * 4.86 * 31 * 0.83 * 0.93);
                chartgeracao.Series["Geração"].Points.AddXY("Setembro", potenciagerada * 5.08 * 30 * 0.83 * 0.93);
                chartgeracao.Series["Geração"].Points.AddXY("Outubro", potenciagerada * 5.37 * 31 * 0.83 * 0.93);
                chartgeracao.Series["Geração"].Points.AddXY("Novembro", potenciagerada * 5.22 * 30 * 0.83 * 0.93);
                chartgeracao.Series["Geração"].Points.AddXY("Dezembro", potenciagerada * 5.59 * 31 * 0.83 * 0.93);
            }
            else if (rbtn10.Checked)
            {
                chartgeracao.Series["Geração"].Points.AddXY("Janeiro", potenciagerada * 5.48 * 31 * 0.83 * 0.9);
                chartgeracao.Series["Geração"].Points.AddXY("Fevereiro", potenciagerada * 5.7 * 29 * 0.83 * 0.9);
                chartgeracao.Series["Geração"].Points.AddXY("Março", potenciagerada * 4.85 * 31 * 0.83 * 0.9);
                chartgeracao.Series["Geração"].Points.AddXY("Abril", potenciagerada * 4.59 * 30 * 0.83 * 0.9);
                chartgeracao.Series["Geração"].Points.AddXY("Maio", potenciagerada * 3.95 * 31 * 0.83 * 0.9);
                chartgeracao.Series["Geração"].Points.AddXY("Junho", potenciagerada * 3.76 * 30 * 0.83 * 0.9);
                chartgeracao.Series["Geração"].Points.AddXY("Julho", potenciagerada * 4.01 * 31 * 0.83 * 0.9);
                chartgeracao.Series["Geração"].Points.AddXY("Agosto", potenciagerada * 4.86 * 31 * 0.83 * 0.9);
                chartgeracao.Series["Geração"].Points.AddXY("Setembro", potenciagerada * 5.08 * 30 * 0.83 * 0.9);
                chartgeracao.Series["Geração"].Points.AddXY("Outubro", potenciagerada * 5.37 * 31 * 0.83 * 0.9);
                chartgeracao.Series["Geração"].Points.AddXY("Novembro", potenciagerada * 5.22 * 30 * 0.83 * 0.9);
                chartgeracao.Series["Geração"].Points.AddXY("Dezembro", potenciagerada * 5.59 * 31 * 0.83 * 0.9);
            }
            else if (rbtn12.Checked)
            {
                chartgeracao.Series["Geração"].Points.AddXY("Janeiro", potenciagerada * 5.48 * 31 * 0.83 * 0.88);
                chartgeracao.Series["Geração"].Points.AddXY("Fevereiro", potenciagerada * 5.7 * 29 * 0.83 * 0.88);
                chartgeracao.Series["Geração"].Points.AddXY("Março", potenciagerada * 4.85 * 31 * 0.83 * 0.88);
                chartgeracao.Series["Geração"].Points.AddXY("Abril", potenciagerada * 4.59 * 30 * 0.83 * 0.88);
                chartgeracao.Series["Geração"].Points.AddXY("Maio", potenciagerada * 3.95 * 31 * 0.83 * 0.88);
                chartgeracao.Series["Geração"].Points.AddXY("Junho", potenciagerada * 3.76 * 30 * 0.83 * 0.88);
                chartgeracao.Series["Geração"].Points.AddXY("Julho", potenciagerada * 4.01 * 31 * 0.83 * 0.88);
                chartgeracao.Series["Geração"].Points.AddXY("Agosto", potenciagerada * 4.86 * 31 * 0.83 * 0.88);
                chartgeracao.Series["Geração"].Points.AddXY("Setembro", potenciagerada * 5.08 * 30 * 0.83 * 0.88);
                chartgeracao.Series["Geração"].Points.AddXY("Outubro", potenciagerada * 5.37 * 31 * 0.83 * 0.88);
                chartgeracao.Series["Geração"].Points.AddXY("Novembro", potenciagerada * 5.22 * 30 * 0.83 * 0.88);
                chartgeracao.Series["Geração"].Points.AddXY("Dezembro", potenciagerada * 5.59 * 31 * 0.83 * 0.88);
            }
            else if (rbtn15.Checked)
            {
                chartgeracao.Series["Geração"].Points.AddXY("Janeiro", potenciagerada * 5.48 * 31 * 0.83 * 0.85);
                chartgeracao.Series["Geração"].Points.AddXY("Fevereiro", potenciagerada * 5.7 * 29 * 0.83 * 0.85);
                chartgeracao.Series["Geração"].Points.AddXY("Março", potenciagerada * 4.85 * 31 * 0.83 * 0.85);
                chartgeracao.Series["Geração"].Points.AddXY("Abril", potenciagerada * 4.59 * 30 * 0.83 * 0.85);
                chartgeracao.Series["Geração"].Points.AddXY("Maio", potenciagerada * 3.95 * 31 * 0.83 * 0.85);
                chartgeracao.Series["Geração"].Points.AddXY("Junho", potenciagerada * 3.76 * 30 * 0.83 * 0.85);
                chartgeracao.Series["Geração"].Points.AddXY("Julho", potenciagerada * 4.01 * 31 * 0.83 * 0.85);
                chartgeracao.Series["Geração"].Points.AddXY("Agosto", potenciagerada * 4.86 * 31 * 0.83 * 0.85);
                chartgeracao.Series["Geração"].Points.AddXY("Setembro", potenciagerada * 5.08 * 30 * 0.83 * 0.85);
                chartgeracao.Series["Geração"].Points.AddXY("Outubro", potenciagerada * 5.37 * 31 * 0.83 * 0.85);
                chartgeracao.Series["Geração"].Points.AddXY("Novembro", potenciagerada * 5.22 * 30 * 0.83 * 0.85);
                chartgeracao.Series["Geração"].Points.AddXY("Dezembro", potenciagerada * 5.59 * 31 * 0.83 * 0.85);
            }
            else if (rbtn20.Checked)
            {
                chartgeracao.Series["Geração"].Points.AddXY("Janeiro", potenciagerada * 5.48 * 31 * 0.83 * 0.8);
                chartgeracao.Series["Geração"].Points.AddXY("Fevereiro", potenciagerada * 5.7 * 29 * 0.83 * 0.8);
                chartgeracao.Series["Geração"].Points.AddXY("Março", potenciagerada * 4.85 * 31 * 0.83 * 0.8);
                chartgeracao.Series["Geração"].Points.AddXY("Abril", potenciagerada * 4.59 * 30 * 0.83 * 0.8);
                chartgeracao.Series["Geração"].Points.AddXY("Maio", potenciagerada * 3.95 * 31 * 0.83 * 0.8);
                chartgeracao.Series["Geração"].Points.AddXY("Junho", potenciagerada * 3.76 * 30 * 0.83 * 0.8);
                chartgeracao.Series["Geração"].Points.AddXY("Julho", potenciagerada * 4.01 * 31 * 0.83 * 0.8);
                chartgeracao.Series["Geração"].Points.AddXY("Agosto", potenciagerada * 4.86 * 31 * 0.83 * 0.8);
                chartgeracao.Series["Geração"].Points.AddXY("Setembro", potenciagerada * 5.08 * 30 * 0.83 * 0.8);
                chartgeracao.Series["Geração"].Points.AddXY("Outubro", potenciagerada * 5.37 * 31 * 0.83 * 0.8);
                chartgeracao.Series["Geração"].Points.AddXY("Novembro", potenciagerada * 5.22 * 30 * 0.83 * 0.8);
                chartgeracao.Series["Geração"].Points.AddXY("Dezembro", potenciagerada * 5.59 * 31 * 0.83 * 0.8);
            }
            else if (rbtn25.Checked)
            {
                chartgeracao.Series["Geração"].Points.AddXY("Janeiro", potenciagerada * 5.48 * 31 * 0.83 * 0.75);
                chartgeracao.Series["Geração"].Points.AddXY("Fevereiro", potenciagerada * 5.7 * 29 * 0.83 * 0.75);
                chartgeracao.Series["Geração"].Points.AddXY("Março", potenciagerada * 4.85 * 31 * 0.83 * 0.75);
                chartgeracao.Series["Geração"].Points.AddXY("Abril", potenciagerada * 4.59 * 30 * 0.83 * 0.75);
                chartgeracao.Series["Geração"].Points.AddXY("Maio", potenciagerada * 3.95 * 31 * 0.83 * 0.75);
                chartgeracao.Series["Geração"].Points.AddXY("Junho", potenciagerada * 3.76 * 30 * 0.83 * 0.75);
                chartgeracao.Series["Geração"].Points.AddXY("Julho", potenciagerada * 4.01 * 31 * 0.83 * 0.75);
                chartgeracao.Series["Geração"].Points.AddXY("Agosto", potenciagerada * 4.86 * 31 * 0.83 * 0.75);
                chartgeracao.Series["Geração"].Points.AddXY("Setembro", potenciagerada * 5.08 * 30 * 0.83 * 0.75);
                chartgeracao.Series["Geração"].Points.AddXY("Outubro", potenciagerada * 5.37 * 31 * 0.83 * 0.75);
                chartgeracao.Series["Geração"].Points.AddXY("Novembro", potenciagerada * 5.22 * 30 * 0.83 * 0.75);
                chartgeracao.Series["Geração"].Points.AddXY("Dezembro", potenciagerada * 5.59 * 31 * 0.83 * 0.75);
            }
            else if (rbtn30.Checked)
            {
                chartgeracao.Series["Geração"].Points.AddXY("Janeiro", potenciagerada * 5.48 * 31 * 0.83 * 1.1);
                chartgeracao.Series["Geração"].Points.AddXY("Fevereiro", potenciagerada * 5.7 * 29 * 0.83 * 1.1);
                chartgeracao.Series["Geração"].Points.AddXY("Março", potenciagerada * 4.85 * 31 * 0.83 * 1.1);
                chartgeracao.Series["Geração"].Points.AddXY("Abril", potenciagerada * 4.59 * 30 * 0.83 * 1.1);
                chartgeracao.Series["Geração"].Points.AddXY("Maio", potenciagerada * 3.95 * 31 * 0.83 * 1.1);
                chartgeracao.Series["Geração"].Points.AddXY("Junho", potenciagerada * 3.76 * 30 * 0.83 * 1.1);
                chartgeracao.Series["Geração"].Points.AddXY("Julho", potenciagerada * 4.01 * 31 * 0.83 * 1.1);
                chartgeracao.Series["Geração"].Points.AddXY("Agosto", potenciagerada * 4.86 * 31 * 0.83 * 1.1);
                chartgeracao.Series["Geração"].Points.AddXY("Setembro", potenciagerada * 5.08 * 30 * 0.83 * 1.1);
                chartgeracao.Series["Geração"].Points.AddXY("Outubro", potenciagerada * 5.37 * 31 * 0.83 * 1.1);
                chartgeracao.Series["Geração"].Points.AddXY("Novembro", potenciagerada * 5.22 * 30 * 0.83 * 1.1);
                chartgeracao.Series["Geração"].Points.AddXY("Dezembro", potenciagerada * 5.59 * 31 * 0.83 * 1.1);
            }
            else if (rbtn35.Checked)
            {
                chartgeracao.Series["Geração"].Points.AddXY("Janeiro", potenciagerada * 5.48 * 31 * 0.83 * 1.075);
                chartgeracao.Series["Geração"].Points.AddXY("Fevereiro", potenciagerada * 5.7 * 29 * 0.83 * 1.075);
                chartgeracao.Series["Geração"].Points.AddXY("Março", potenciagerada * 4.85 * 31 * 0.83 * 1.075);
                chartgeracao.Series["Geração"].Points.AddXY("Abril", potenciagerada * 4.59 * 30 * 0.83 * 1.075);
                chartgeracao.Series["Geração"].Points.AddXY("Maio", potenciagerada * 3.95 * 31 * 0.83 * 1.075);
                chartgeracao.Series["Geração"].Points.AddXY("Junho", potenciagerada * 3.76 * 30 * 0.83 * 1.075);
                chartgeracao.Series["Geração"].Points.AddXY("Julho", potenciagerada * 4.01 * 31 * 0.83 * 1.075);
                chartgeracao.Series["Geração"].Points.AddXY("Agosto", potenciagerada * 4.86 * 31 * 0.83 * 1.075);
                chartgeracao.Series["Geração"].Points.AddXY("Setembro", potenciagerada * 5.08 * 30 * 0.83 * 1.075);
                chartgeracao.Series["Geração"].Points.AddXY("Outubro", potenciagerada * 5.37 * 31 * 0.83 * 1.075);
                chartgeracao.Series["Geração"].Points.AddXY("Novembro", potenciagerada * 5.22 * 30 * 0.83 * 1.075);
                chartgeracao.Series["Geração"].Points.AddXY("Dezembro", potenciagerada * 5.59 * 31 * 0.83 * 1.075);
            }
            else if (rbtn40.Checked)
            {
                chartgeracao.Series["Geração"].Points.AddXY("Janeiro", potenciagerada * 5.48 * 31 * 0.83 * 1.05);
                chartgeracao.Series["Geração"].Points.AddXY("Fevereiro", potenciagerada * 5.7 * 29 * 0.83 * 1.05);
                chartgeracao.Series["Geração"].Points.AddXY("Março", potenciagerada * 4.85 * 31 * 0.83 * 1.05);
                chartgeracao.Series["Geração"].Points.AddXY("Abril", potenciagerada * 4.59 * 30 * 0.83 * 1.05);
                chartgeracao.Series["Geração"].Points.AddXY("Maio", potenciagerada * 3.95 * 31 * 0.83 * 1.05);
                chartgeracao.Series["Geração"].Points.AddXY("Junho", potenciagerada * 3.76 * 30 * 0.83 * 1.05);
                chartgeracao.Series["Geração"].Points.AddXY("Julho", potenciagerada * 4.01 * 31 * 0.83 * 1.05);
                chartgeracao.Series["Geração"].Points.AddXY("Agosto", potenciagerada * 4.86 * 31 * 0.83 * 1.05);
                chartgeracao.Series["Geração"].Points.AddXY("Setembro", potenciagerada * 5.08 * 30 * 0.83 * 1.05);
                chartgeracao.Series["Geração"].Points.AddXY("Outubro", potenciagerada * 5.37 * 31 * 0.83 * 1.05);
                chartgeracao.Series["Geração"].Points.AddXY("Novembro", potenciagerada * 5.22 * 30 * 0.83 * 1.05);
                chartgeracao.Series["Geração"].Points.AddXY("Dezembro", potenciagerada * 5.59 * 31 * 0.83 * 1.05);
            }
            else if (rbtn0.Checked)
            {
                chartgeracao.Series["Geração"].Points.AddXY("Jan", potenciagerada * 5.48 * 31 * 0.83);
                chartgeracao.Series["Geração"].Points.AddXY("Fev", potenciagerada * 5.7 * 29 * 0.83);
                chartgeracao.Series["Geração"].Points.AddXY("Mar", potenciagerada * 4.85 * 31 * 0.83);
                chartgeracao.Series["Geração"].Points.AddXY("Abr", potenciagerada * 4.59 * 30 * 0.83);
                chartgeracao.Series["Geração"].Points.AddXY("Mai", potenciagerada * 3.95 * 31 * 0.83);
                chartgeracao.Series["Geração"].Points.AddXY("Jun", potenciagerada * 3.76 * 30 * 0.83);
                chartgeracao.Series["Geração"].Points.AddXY("Jul", potenciagerada * 4.01 * 31 * 0.83);
                chartgeracao.Series["Geração"].Points.AddXY("Ago", potenciagerada * 4.86 * 31 * 0.83);
                chartgeracao.Series["Geração"].Points.AddXY("Set", potenciagerada * 5.08 * 30 * 0.83);
                chartgeracao.Series["Geração"].Points.AddXY("Out", potenciagerada * 5.37 * 31 * 0.83);
                chartgeracao.Series["Geração"].Points.AddXY("Nov", potenciagerada * 5.22 * 30 * 0.83);
                chartgeracao.Series["Geração"].Points.AddXY("Dez", potenciagerada * 5.59 * 31 * 0.83);
            }

            chartretornofin.Series.Add("Caixa Acumulado");
            chartretornofin.Series.Add("Custo");
            chartretornofin.Series["Caixa Acumulado"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            chartretornofin.Series["Custo"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            chartretornofin.Series["Caixa Acumulado"].Color = Color.DarkSeaGreen;
            chartretornofin.Series["Custo"].Color = Color.Blue;
            chartretornofin.Series[0].IsVisibleInLegend = false;

            for (int i = 1; i < 11; i++)
            {
                chartretornofin.Series["Custo"].Points.AddXY(i, double.Parse(txtcustoinversor.Text));
            }

            double resultado = (double.Parse(txtvalorsist.Text) - (gerano * double.Parse(txtkwh.Text))) * (-1);
            double aux = gerano;
            double tari = double.Parse(txtkwh.Text);
            double dif = (double.Parse(txtvalorsist.Text) * 0.99) - (double.Parse(txtvalorsist.Text) * 0.985);
            //
            chartretornofin.Series["Caixa Acumulado"].Points.AddXY(1, (double.Parse(txtvalorsist.Text) - (gerano * double.Parse(txtkwh.Text))) * (-1));
            //
            aux = aux * 0.99;
            tari = (double.Parse(txtkwh.Text) * 1.1);
            resultado = resultado + (aux * tari);
            //
            chartretornofin.Series["Caixa Acumulado"].Points.AddXY(2, resultado);
            //
            tari *= 1.1;
            aux = aux * 0.985;
            resultado = resultado + (aux * tari);
            //
            chartretornofin.Series["Caixa Acumulado"].Points.AddXY(3, resultado);
            //
            for (int i = 4; i < 11; i++)
            {
                aux = aux - ((i - 3) * dif);
                tari = tari * ((1 + (i - 3) * 0.1));
                resultado = resultado + (aux * tari);
                chartretornofin.Series["Caixa Acumulado"].Points.AddXY(i, resultado);
            }
        }
        private void PaineisPrincipais(Panel panel)
        {
            pnlorcamento0.Visible = false;
            pnlorcamento1.Visible = false;
            pnlorcamento2.Visible = false;
            pnlorcamento3.Visible = false;
            pnlorcamento4.Visible = false;
            pnlorcamento5.Visible = false;
            pnlinicio.Visible = false;
            pnlorcasalvos.Visible = false;
            pnlconfiguracao.Visible = false;
            pnlclientes.Visible = false;
            pnlprojeto.Visible = false;
            pnlfinalizaprojeto.Visible = false;
            pnlequipamentos.Visible = false;
            pnlmod1.Visible = false;
            pnlinv1.Visible = false;

            if (!pnlnovoproj0.Visible)
            {
                pnlnovoproj0.Visible = true;
                pnlnovoproj1.Visible = false;
                pnlnovoproj2.Visible = false;
                pnlnovoproj3.Visible = false;
            }
            
            panel.Visible = true;
        }
        private void CarregaCombobox()
        {
            //cbxmodelomodcliente.DataSource = Dados;
            //cbxmodelomodcliente.ValueMember = "Modelo";
            //cbxmodelomodcliente.DisplayMember = "Modelo";

            var Dados1 = func.TodosMod(Banco);
            cbxmarcamod.DataSource = Dados1;
            cbxmarcamoccliente.DataSource = Dados1;
            cbxmarcamoccliente.ValueMember = "Marca";
            cbxmarcamoccliente.DisplayMember = "Marca";

            var Dados = func.ModeloMod(Banco);
            cbxmodpaineis.DataSource = Dados;
            cbxmodpaineis.ValueMember = "Modelo";
            cbxmodpaineis.DisplayMember = "Modelo";

            var Dados3 = func.TodosInv(Banco);
            cbxmarcainv.DataSource = Dados3;
            cbxmarcainvcliente.DataSource = Dados3;
            cbxmarcainvcliente.ValueMember = "Marca";
            cbxmarcainvcliente.DisplayMember = "Marca";

            var Dados2_5 = func.ModeloInv1(Banco);
            cbxmodinv.DataSource = Dados2_5;
            cbxmodinv.ValueMember = "Modelo";
            cbxmodinv.DisplayMember = "Modelo";

            var Dados2 = func.MarcaInv(cbxmarcainv.Text, Banco);
            cbxmodeloinvcliente.DataSource = Dados2;
            cbxmodeloinvcliente.ValueMember = "Modelo";
            cbxmodeloinvcliente.DisplayMember = "Modelo";

            var Dados4 = func.MaterialMod(Banco);
            cbxmaterialmodequipamentos.DataSource = Dados4;
            cbxmaterialmodequipamentos.ValueMember = "Material";
            cbxmaterialmodequipamentos.DisplayMember = "Material";

            var Dados5 = func.FasesInversor(Banco);
            cbxfasesinvequip.DataSource = Dados5;
            cbxfasesinvequip.ValueMember = "Fases";
            cbxfasesinvequip.DisplayMember = "Fases";

            Dados5 = func.MarcaModulo(Banco);
            cbxmarcamodequipamentos.DataSource = Dados5;
            cbxmarcamodequipamentos.ValueMember = "Marca";
            cbxmarcamodequipamentos.DisplayMember = "Marca";

            Dados5 = func.ModeloMod(Banco);
            cbxmodelomodequipamentos.DataSource = Dados5;
            cbxmodelomodequipamentos.ValueMember = "Modelo";
            cbxmodelomodequipamentos.DisplayMember = "Modelo";

            cbxmodelomodcliente.DataSource = Dados5;
            cbxmodelomodcliente.ValueMember = "Modelo";
            cbxmodelomodcliente.DisplayMember = "Modelo";

            Dados5 = func.CelulasModulo(Banco);
            cbxcelulasmodequipamentos.DataSource = Dados5;
            cbxcelulasmodequipamentos.ValueMember = "Celulas";
            cbxcelulasmodequipamentos.DisplayMember = "Celulas";

            Dados5 = func.MarcaInv1(Banco);
            cbxmarcainvequip.DataSource = Dados5;
            cbxmarcainvequip.ValueMember = "Marca";
            cbxmarcainvequip.DisplayMember = "Marca";

            Dados5 = func.ModeloInv1(Banco);
            cbxmodinvequip.DataSource = Dados5;
            cbxmodinvequip.ValueMember = "Modelo";
            cbxmodinvequip.DisplayMember = "Modelo";

            Dados5 = func.TensaoInv1(Banco);
            cbxtensaoinvequip.DataSource = Dados5;
            cbxtensaoinvequip.ValueMember = "Tensao";
            cbxtensaoinvequip.DisplayMember = "Tensao";

            Dados5 = func.BuscaFornecedor(Banco);
            cbxsimparceiro1.DataSource = Dados5;
            cbxsimparceiro1.ValueMember = "Fornecedor";
            cbxsimparceiro1.DisplayMember = "Fornecedor";

            var Dados6 = func.BuscaFornecedor(Banco);
            cbxsimparceiro2.DataSource = Dados6;
            cbxsimparceiro2.ValueMember = "Fornecedor";
            cbxsimparceiro2.DisplayMember = "Fornecedor";

            var Dados7 = func.BuscaFornecedor(Banco);
            cbxsimparceiro3.DataSource = Dados7;
            cbxsimparceiro3.ValueMember = "Fornecedor";
            cbxsimparceiro3.DisplayMember = "Fornecedor";

            var Dados8 = func.BuscaFornecedor(Banco);
            cbxsimparceiro4.DataSource = Dados8;
            cbxsimparceiro4.ValueMember = "Fornecedor";
            cbxsimparceiro4.DisplayMember = "Fornecedor";

            var Dados9 = func.BuscaFornecedor(Banco);
            cbxsimparceiro5.DataSource = Dados9;
            cbxsimparceiro5.ValueMember = "Fornecedor";
            cbxsimparceiro5.DisplayMember = "Fornecedor";

            Dados5 = func.ModeloMod(Banco);
            cbxsimmodmod1.DataSource = Dados5;
            cbxsimmodmod1.ValueMember = "Modelo";
            cbxsimmodmod1.DisplayMember = "Modelo";

            Dados6 = func.ModeloMod(Banco);
            cbxsimmodmod2.DataSource = Dados6;
            cbxsimmodmod2.ValueMember = "Modelo";
            cbxsimmodmod2.DisplayMember = "Modelo";

            Dados7 = func.ModeloMod(Banco);
            cbxsimmodmod3.DataSource = Dados7;
            cbxsimmodmod3.ValueMember = "Modelo";
            cbxsimmodmod3.DisplayMember = "Modelo";

            Dados8 = func.ModeloMod(Banco);
            cbxsimmodmod4.DataSource = Dados8;
            cbxsimmodmod4.ValueMember = "Modelo";
            cbxsimmodmod4.DisplayMember = "Modelo";

            Dados9 = func.ModeloMod(Banco);
            cbxsimmodmod5.DataSource = Dados9;
            cbxsimmodmod5.ValueMember = "Modelo";
            cbxsimmodmod5.DisplayMember = "Modelo";

            Dados5 = func.ModeloInv1(Banco);
            cbxsimmodinv1.DataSource = Dados5;
            cbxsimmodinv1.ValueMember = "Modelo";
            cbxsimmodinv1.DisplayMember = "Modelo";

            Dados6 = func.ModeloInv1(Banco);
            cbxsimmodinv2.DataSource = Dados6;
            cbxsimmodinv2.ValueMember = "Modelo";
            cbxsimmodinv2.DisplayMember = "Modelo";

            Dados7 = func.ModeloInv1(Banco);
            cbxsimmodinv3.DataSource = Dados7;
            cbxsimmodinv3.ValueMember = "Modelo";
            cbxsimmodinv3.DisplayMember = "Modelo";

            Dados8 = func.ModeloInv1(Banco);
            cbxsimmodinv4.DataSource = Dados8;
            cbxsimmodinv4.ValueMember = "Modelo";
            cbxsimmodinv4.DisplayMember = "Modelo";

            Dados9 = func.ModeloInv1(Banco);
            cbxsimmodinv5.DataSource = Dados9;
            cbxsimmodinv5.ValueMember = "Modelo";
            cbxsimmodinv5.DisplayMember = "Modelo";
        }
        private void PictureBoxRedondo()
        {
            System.Drawing.Drawing2D.GraphicsPath gp = new System.Drawing.Drawing2D.GraphicsPath();
            gp.AddEllipse(0, 0, pbxUsuario.Width - 3, pbxUsuario.Height - 3);
            Region rg = new Region(gp);
            pbxUsuario.Region = rg;

            System.Drawing.Drawing2D.GraphicsPath gp1 = new System.Drawing.Drawing2D.GraphicsPath();
            gp1.AddEllipse(0, 0, pbxuploadfoto.Width - 3, pbxuploadfoto.Height - 3);
            Region rg1 = new Region(gp1);
            pbxuploadfoto.Region = rg1;
        }
        private void CarregaFoto()
        {
            
            txtcpass.Text = func.Password;
            txtcuser.Text = func.UID;
            txtcnomebd.Text = func.NomeDB;
            txtcip.Text = func.Servidor;

            func.PesquisaLogin(func.Login, Banco);
            if (func.Foto == null)
            {
                pbxuploadfoto.Image = Properties.Resources._void;
                pbxUsuario.Image = Properties.Resources.user_cinza_circulobranco;
            }
            else
            {
                MemoryStream ms = new MemoryStream(func.Foto);
                //pbxuploadfoto.Image = Image.FromStream(ms);
                //pbxUsuario.Image = Image.FromStream(ms);
            }
            label56.Text = func.Nome;
            txtcnomecompleto.Text = func.Nome;
            txtcusuario.Text = func.Login;
            txtcsenha.Text = func.Senha;
            txtcemail.Text = func.email;
            pbxUsuario.Image = Base64ToImage(func.FotoUsuario);
        }
        public bool ValidaLogin(string login, string senha)
        {
            func.PesquisaLogin(login, Banco);
            if (login == func.Login && senha == func.Senha)
            {
                validado = true;
                return true;
            }
            else
            {
                return false;
            }
        }
        public void CarregaDataGrid()
        {
            dgvcredenciais.DataSource = func.PesquisaCredenciais(Banco);
            dgvorcamentos.DataSource = func.CarregaOrc(Banco);
            dgvclientes.DataSource = func.AtualizaClientes(Banco);
            dgvencontraclienteprojeto.DataSource = func.AtualizaClientes(Banco);
            dgvprojetos.DataSource = func.PesquisaProjeto(Banco);
            dgvmodulos.DataSource = func.AtualizaPaineis(Banco);
            dgvinversores.DataSource = func.AtualizaInversor(Banco);
        }
        public void Projeto()
        {
            string pasta = @"C:\Centraliza\Projeto\";
            CriaMemorial(@"C:\Centraliza\Centraliza\MemorialDescritivo.docx", @"C:\Centraliza\Projeto\Memorial Descritivo Solar " + txtnomecliprojeto.Text + ".docx");
            pgbstatusproj.Value++;
            potenciagerada = double.Parse(func.PotenciaMod) * double.Parse(func.QuantidadeModulos) / 1000;
            if (potenciagerada < 10)
            {
                pgbstatusproj.Value++;
                CriaFormulario(@"C:\Centraliza\Centraliza\FORMULARIO_GD10kW.docx", @"C:\Centraliza\Projeto\FORMULARIO_GD10kW " + txtnome.Text + ".docx");
                pgbstatusproj.Value++;
            }
            else if (potenciagerada > 10 && potenciagerada <= 75)
            {
                pgbstatusproj.Value++;
                CriaFormulario(@"C:\Centraliza\Centraliza\FORMULARIO_GD_Maior_10kW.docx", @"C:\Centraliza\Projeto\FORMULARIO_GD_Maior_10kW " + txtnome.Text + ".docx");
                pgbstatusproj.Value++;
            }
            else
            {
                pgbstatusproj.Value++;
                CriaFormulario(@"C:\Centraliza\Centraliza\FORMULARIO_GD.docx", @"C:\Centraliza\Projeto\FORMULARIO_GD " + txtnome.Text + ".docx");
                pgbstatusproj.Value++;
            }
            func.SelecionaInversor(Banco, func.ModeloInversor);
            pgbstatusproj.Value++;
            if (func.RegistroINMETRO != "")
            {
                string fileName;
                string alvo;
                string Certificado = func.RegistroINMETRO;
                string pri = Certificado.Substring(0, 6);
                string ano = Certificado.Substring(7, 4);
                string certs = @"C:\Centraliza\Certificados\" + pri + @"\" + ano;
                string[] files = System.IO.Directory.GetFiles(certs);
                foreach (string s in files)
                {
                    fileName = System.IO.Path.GetFileName(s);
                    string destFile = System.IO.Path.Combine(pasta, fileName);
                    alvo = System.IO.Path.Combine(certs, fileName);
                    System.IO.File.Copy(s, destFile, true);
                }
                pgbstatusproj.Value++;
            }
            else
            {
                MessageBox.Show("O Inversor não possui certificado do Inmetro!", "Atenção");
                pgbstatusproj.Value++;
            }
            func.SelecionaPainel(Banco, func.ModeloModulo);
            if (func.RegistroInmetro != "")
            {
                string fileName;
                string alvo;
                string Certificado = func.RegistroInmetro;
                string pri = Certificado.Substring(0, 6);
                string ano = Certificado.Substring(7, 4);
                string certs = @"C:\Centraliza\Certificados\" + pri + @"\" + ano;
                string[] files = System.IO.Directory.GetFiles(certs);
                foreach (string s in files)
                {
                    fileName = System.IO.Path.GetFileName(s);
                    string destFile = System.IO.Path.Combine(pasta, fileName);
                    alvo = System.IO.Path.Combine(certs, fileName);
                    System.IO.File.Copy(s, destFile, true);
                }
                pgbstatusproj.Value++;
            }
            else
            {
                MessageBox.Show("O painel não possui certificado do Inmetro!", "Atenção");
                pgbstatusproj.Value++;
            }
            Process.Start("explorer.exe", pasta);
        }
        public bool VerificaArquivos()
        {
            //Arquivo dos Inversores
            string fileName = func.ModeloInversor + ".txt";
            string sourcePath = @"C:\Centraliza\Dados Equipamentos\Inversores\";
            string targetPath = @"C:\Centraliza\Projeto\";
            string sourceFile = System.IO.Path.Combine(sourcePath, fileName);
            string destFile = System.IO.Path.Combine(targetPath, fileName);
            //Arquivo dos Paineis
            string fileName2 = func.ModeloModulo + ".txt";
            string sourcePath2 = @"C:\Centraliza\Dados Equipamentos\Modulos\";
            string sourceFile2 = System.IO.Path.Combine(sourcePath2, fileName2);
            string destFile2 = System.IO.Path.Combine(targetPath, fileName2);
            if (File.Exists(sourceFile))
            {
                if (File.Exists(sourceFile2))
                {
                    return true;
                }
                else
                {
                    MessageBox.Show("Verifique o arquivo de dados do módulo", "Atenção");
                    return false;
                }
            }
            else if (!File.Exists(sourceFile))
            {
                MessageBox.Show("Verifique o arquivo de dados do inversor", "Atenção");
                return false;
            }
            else
            {
                MessageBox.Show("Verifique os arquivos de dados dos equipamentos do cliente", "Atenção");
                return false;
            }
        }
        private string ConverteBase64(string caminho)
        {
            using (Image image = Image.FromFile(caminho))
            {
                using (MemoryStream m = new MemoryStream())
                {
                    image.Save(m, image.RawFormat);
                    byte[] imageBytes = m.ToArray();

                    // Convert byte[] to Base64 String
                    string base64String = Convert.ToBase64String(imageBytes);
                    return base64String;
                }
            }
        }
        private Image Base64ToImage(string base64String)
        {
            // Convert base 64 string to byte[]
            byte[] imageBytes = Convert.FromBase64String(base64String);
            // Convert byte[] to Image
            using (var ms = new MemoryStream(imageBytes, 0, imageBytes.Length))
            {
                Image image = Image.FromStream(ms, true);
                return image;
            }
        }

        //Geração de documentos
        private void CreateWordDoc(object filename, object SaveAs)
        {
            Word.Application wordApp = new Word.Application();
            object missing = Missing.Value;
            Word.Document mywordDoc = null;

            if (File.Exists((string)filename))
            {
                object readOnly = false;
                object isVisible = false;
                wordApp.Visible = false;

                mywordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing);
                mywordDoc.Activate();
                pgbstatusorca.Value++;

                Calculos();

                //Achar e substituir
                this.orcamento.AcharESubstituir(wordApp, "<proposta>", func.Proposta.ToString());
                //this.orcamento.AcharESubstituir(wordApp, "<obs>", txtobs.Text);
                this.orcamento.AcharESubstituir(wordApp, "<obs>", " ");
                this.orcamento.AcharESubstituir(wordApp, "<cidade>", txtcidade.Text);
                this.orcamento.AcharESubstituir(wordApp, "<nome>", txtnome.Text);
                if (txtnumero.Text == "" && txtbairro.Text == "")
                {
                    this.orcamento.AcharESubstituir(wordApp, "<endereco>", txtendereco.Text + " ");
                }
                else if (txtnumero.Text == "" && txtbairro.Text != "")
                {
                    this.orcamento.AcharESubstituir(wordApp, "<endereco>", txtendereco.Text + ", " + txtbairro.Text + " ");
                }
                else if (txtnumero.Text != "" && txtbairro.Text == "")
                {
                    this.orcamento.AcharESubstituir(wordApp, "<endereco>", txtendereco.Text + ", " + txtnumero.Text + " ");
                }
                else
                {
                    this.orcamento.AcharESubstituir(wordApp, "<endereco>", txtendereco.Text + ", " + txtnumero.Text + ", " + txtbairro.Text);
                }

                this.orcamento.AcharESubstituir(wordApp, "<cep>", "CEP " + txtcep.Text);
                this.orcamento.AcharESubstituir(wordApp, "<contato>", txtcontato.Text);
                this.orcamento.AcharESubstituir(wordApp, "<potger>", potger1);
                this.orcamento.AcharESubstituir(wordApp, "<gerano>", gerano1);
                this.orcamento.AcharESubstituir(wordApp, "<germes>", germes1);
                this.orcamento.AcharESubstituir(wordApp, "<conano>", contot);
                this.orcamento.AcharESubstituir(wordApp, "<conmes>", conmes1);
                this.orcamento.AcharESubstituir(wordApp, "<sumdimmod>", dimensao1);
                this.orcamento.AcharESubstituir(wordApp, "<qtdmod>", txtqtdpaineis.Text);
                pgbstatusorca.Value++;
                //
                //Nova Configuração de Equipamentos
                //
                func.PesquisaModMod(cbxmodpaineis.Text, Banco);
                this.orcamento.AcharESubstituir(wordApp, "<modelomod>", "Módulos Fotovoltaicos " + func.MarcaMod + " " + cbxmodpaineis.Text + " " + func.Material + " " + func.Celulas + " " + func.PotenciaMod);
                this.orcamento.AcharESubstituir(wordApp, "<qtdinv>", txtqtdinv.Text);
                func.PesquisaModInv(cbxmodinv.Text, Banco);
                pgbstatusorca.Value++;
                //Plural Inversores
                if (Int32.Parse(txtqtdinv.Text) > 1)
                {
                    if (func.MarcaInversor == "AP System" || func.ModeloInversor == "Reno560" || func.ModeloInversor == "Reno560-LV")
                    {
                        this.orcamento.AcharESubstituir(wordApp, "<modeloinv>", "Microinversores " + func.MarcaInversor + " " + cbxmodinv.Text);
                        this.orcamento.AcharESubstituir(wordApp, "<plusinginv>", "Inversores");
                    }
                    else
                    {
                        this.orcamento.AcharESubstituir(wordApp, "<modeloinv>", "Inversores " + func.MarcaInversor + " " + cbxmodinv.Text);
                        this.orcamento.AcharESubstituir(wordApp, "<plusinginv>", "Inversores");
                    }
                }
                else
                {
                    if (func.MarcaInversor == "AP System" || func.ModeloInversor == "Reno560" || func.ModeloInversor == "Reno560-LV")
                    {
                        this.orcamento.AcharESubstituir(wordApp, "<modeloinv>", "Microinversor " + func.MarcaInversor + " " + cbxmodinv.Text);
                        this.orcamento.AcharESubstituir(wordApp, "<plusinginv>", "Inversor");
                    }
                    else
                    {
                        this.orcamento.AcharESubstituir(wordApp, "<modeloinv>", "Inversor " + func.MarcaInversor + " " + cbxmodinv.Text);
                        this.orcamento.AcharESubstituir(wordApp, "<plusinginv>", "Inversor");
                    }
                }
                pgbstatusorca.Value++;
                //Formas de pagamento
                this.orcamento.AcharESubstituir(wordApp, "<formspag>", "");
                this.orcamento.AcharESubstituir(wordApp, "<opcpag1>", "");
                this.orcamento.AcharESubstituir(wordApp, "<opcpag2>", "");
                this.orcamento.AcharESubstituir(wordApp, "<opcpag3>", "");
                this.orcamento.AcharESubstituir(wordApp, "<opcpag4>", "");

                this.orcamento.AcharESubstituir(wordApp, "<estrutura>", cbxestrutura.Text);
                this.orcamento.AcharESubstituir(wordApp, "<economia>", caixaacumulado);
                this.orcamento.AcharESubstituir(wordApp, "<garantiamod>", func.GarantiaMod + " anos");
                this.orcamento.AcharESubstituir(wordApp, "<garantiainv>", func.GarantiaInv + " anos");
                this.orcamento.AcharESubstituir(wordApp, "<valorsistema>", valorsist);
                this.orcamento.AcharESubstituir(wordApp, "<valorsistemajuros>", valsisjur);
                this.orcamento.AcharESubstituir(wordApp, "<valorsistemajurosparcela>", valparcela);
                this.orcamento.AcharESubstituir(wordApp, "<extenso>", orcamento.EscreverExtenso(Int64.Parse(txtvalorsist.Text)));
                valsisjur = string.Format("{0:0}", (double.Parse(valparcela) * 12));
                this.orcamento.AcharESubstituir(wordApp, "<extenso2>", orcamento.EscreverExtenso(Int64.Parse((valsisjur))));
                this.orcamento.AcharESubstituir(wordApp, "<valorequipamentos>", valorequip);
                this.orcamento.AcharESubstituir(wordApp, "<valorremanescente>", valorrestante1);
                this.orcamento.AcharESubstituir(wordApp, "<data>", DateTime.Now.ToString("dd' de 'MMMM' de 'yyyy"));
                pgbstatusorca.Value++;

                //Payback
                double tarifao = double.Parse(txtkwh.Text);
                double pb;
                if (txtcustoinversor.Text == "" || txtcustoinversor.Enabled == false || txtcustoinversor.Text == string.Empty)
                {
                    double t = (((double.Parse(txtvalorsist.Text) + 1) * double.Parse(txtkwh.Text)) / gerano);
                    string t1 = t.ToString();
                    t1 = string.Format("{0:0.0}", t);
                    this.orcamento.AcharESubstituir(wordApp, "<payback>", t1);
                }
                else
                {
                    this.orcamento.AcharESubstituir(wordApp, "<payback>", payback1);
                }
                pgbstatusorca.Value++;

            }
            else
            {
                MessageBox.Show("Arquivo não encontrado!");
            }

            //Save as
            mywordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing);

            mywordDoc.Close();
            wordApp.Quit();
            pgbstatusorca.Value++;
        }
        public void GerarPlanilha()
        {
            Excel.Application xlapp = new Excel.Application();
            Excel.Workbook wb = default(Excel.Workbook);

            switch (qtdorc)
            {
                case 1:
                    if (!editando)
                    {
                        wb = xlapp.Workbooks.Open(@"C:\Centraliza\Centraliza\temp-geracao-1.xlsx");
                        pgbstatusorca.Value++;
                    }
                    else
                    {
                        wb = xlapp.Workbooks.Open(@"C:\Centraliza\Centraliza\temp-geracao-salvo.xlsx");
                        pgbstatusorca.Value++;
                    }
                    break;
                case 2:
                    wb = xlapp.Workbooks.Open(@"C:\Centraliza\Centraliza\temp-geracao-2.xlsx");
                    pgbstatusorca.Value++;
                    break;
                case 3:
                    wb = xlapp.Workbooks.Open(@"C:\Centraliza\Centraliza\temp-geracao-3.xlsx");
                    pgbstatusorca.Value++;
                    break;
                case 4:
                    wb = xlapp.Workbooks.Open(@"C:\Centraliza\Centraliza\temp-geracao-4.xlsx");
                    pgbstatusorca.Value++;
                    break;
                default:
                    break;
            }

            //Produzido
            if (rbtn5.Checked)
            {
                wb.Worksheets[1].Cells.Replace("<1>", (potenciagerada * 5.48 * 31 * 0.83 * 0.95).ToString());
                wb.Worksheets[1].Cells.Replace("<2>", (potenciagerada * 5.7 * 29 * 0.83 * 0.95).ToString());
                wb.Worksheets[1].Cells.Replace("<3>", (potenciagerada * 4.85 * 31 * 0.83 * 0.95).ToString());
                wb.Worksheets[1].Cells.Replace("<4>", (potenciagerada * 4.59 * 30 * 0.83 * 0.95).ToString());
                wb.Worksheets[1].Cells.Replace("<5>", (potenciagerada * 3.95 * 31 * 0.83 * 0.95).ToString());
                wb.Worksheets[1].Cells.Replace("<6>", (potenciagerada * 3.76 * 30 * 0.83 * 0.95).ToString());
                wb.Worksheets[1].Cells.Replace("<7>", (potenciagerada * 4.01 * 31 * 0.83 * 0.95).ToString());
                wb.Worksheets[1].Cells.Replace("<8>", (potenciagerada * 4.86 * 31 * 0.83 * 0.95).ToString());
                wb.Worksheets[1].Cells.Replace("<9>", (potenciagerada * 5.08 * 30 * 0.83 * 0.95).ToString());
                wb.Worksheets[1].Cells.Replace("<10>", (potenciagerada * 5.37 * 31 * 0.83 * 0.95).ToString());
                wb.Worksheets[1].Cells.Replace("<11>", (potenciagerada * 5.22 * 30 * 0.83 * 0.95).ToString());
                wb.Worksheets[1].Cells.Replace("<12>", (potenciagerada * 5.59 * 31 * 0.83 * 0.95).ToString());
            }
            else if (rbtn7.Checked)
            {
                wb.Worksheets[1].Cells.Replace("<1>", (potenciagerada * 5.48 * 31 * 0.83 * 0.93).ToString());
                wb.Worksheets[1].Cells.Replace("<2>", (potenciagerada * 5.7 * 29 * 0.83 * 0.93).ToString());
                wb.Worksheets[1].Cells.Replace("<3>", (potenciagerada * 4.85 * 31 * 0.83 * 0.93).ToString());
                wb.Worksheets[1].Cells.Replace("<4>", (potenciagerada * 4.59 * 30 * 0.83 * 0.93).ToString());
                wb.Worksheets[1].Cells.Replace("<5>", (potenciagerada * 3.95 * 31 * 0.83 * 0.93).ToString());
                wb.Worksheets[1].Cells.Replace("<6>", (potenciagerada * 3.76 * 30 * 0.83 * 0.93).ToString());
                wb.Worksheets[1].Cells.Replace("<7>", (potenciagerada * 4.01 * 31 * 0.83 * 0.93).ToString());
                wb.Worksheets[1].Cells.Replace("<8>", (potenciagerada * 4.86 * 31 * 0.83 * 0.93).ToString());
                wb.Worksheets[1].Cells.Replace("<9>", (potenciagerada * 5.08 * 30 * 0.83 * 0.93).ToString());
                wb.Worksheets[1].Cells.Replace("<10>", (potenciagerada * 5.37 * 31 * 0.83 * 0.93).ToString());
                wb.Worksheets[1].Cells.Replace("<11>", (potenciagerada * 5.22 * 30 * 0.83 * 0.93).ToString());
                wb.Worksheets[1].Cells.Replace("<12>", (potenciagerada * 5.59 * 31 * 0.83 * 0.93).ToString());
            }
            else if (rbtn10.Checked)
            {
                wb.Worksheets[1].Cells.Replace("<1>", (potenciagerada * 5.48 * 31 * 0.83 * 0.9).ToString());
                wb.Worksheets[1].Cells.Replace("<2>", (potenciagerada * 5.7 * 29 * 0.83 * 0.9).ToString());
                wb.Worksheets[1].Cells.Replace("<3>", (potenciagerada * 4.85 * 31 * 0.83 * 0.9).ToString());
                wb.Worksheets[1].Cells.Replace("<4>", (potenciagerada * 4.59 * 30 * 0.83 * 0.9).ToString());
                wb.Worksheets[1].Cells.Replace("<5>", (potenciagerada * 3.95 * 31 * 0.83 * 0.9).ToString());
                wb.Worksheets[1].Cells.Replace("<6>", (potenciagerada * 3.76 * 30 * 0.83 * 0.9).ToString());
                wb.Worksheets[1].Cells.Replace("<7>", (potenciagerada * 4.01 * 31 * 0.83 * 0.9).ToString());
                wb.Worksheets[1].Cells.Replace("<8>", (potenciagerada * 4.86 * 31 * 0.83 * 0.9).ToString());
                wb.Worksheets[1].Cells.Replace("<9>", (potenciagerada * 5.08 * 30 * 0.83 * 0.9).ToString());
                wb.Worksheets[1].Cells.Replace("<10>", (potenciagerada * 5.37 * 31 * 0.83 * 0.9).ToString());
                wb.Worksheets[1].Cells.Replace("<11>", (potenciagerada * 5.22 * 30 * 0.83 * 0.9).ToString());
                wb.Worksheets[1].Cells.Replace("<12>", (potenciagerada * 5.59 * 31 * 0.83 * 0.9).ToString());
            }
            else if (rbtn12.Checked)
            {
                wb.Worksheets[1].Cells.Replace("<1>", (potenciagerada * 5.48 * 31 * 0.83 * 0.88).ToString());
                wb.Worksheets[1].Cells.Replace("<2>", (potenciagerada * 5.7 * 29 * 0.83 * 0.88).ToString());
                wb.Worksheets[1].Cells.Replace("<3>", (potenciagerada * 4.85 * 31 * 0.83 * 0.88).ToString());
                wb.Worksheets[1].Cells.Replace("<4>", (potenciagerada * 4.59 * 30 * 0.83 * 0.88).ToString());
                wb.Worksheets[1].Cells.Replace("<5>", (potenciagerada * 3.95 * 31 * 0.83 * 0.88).ToString());
                wb.Worksheets[1].Cells.Replace("<6>", (potenciagerada * 3.76 * 30 * 0.83 * 0.88).ToString());
                wb.Worksheets[1].Cells.Replace("<7>", (potenciagerada * 4.01 * 31 * 0.83 * 0.88).ToString());
                wb.Worksheets[1].Cells.Replace("<8>", (potenciagerada * 4.86 * 31 * 0.83 * 0.88).ToString());
                wb.Worksheets[1].Cells.Replace("<9>", (potenciagerada * 5.08 * 30 * 0.83 * 0.88).ToString());
                wb.Worksheets[1].Cells.Replace("<10>", (potenciagerada * 5.37 * 31 * 0.83 * 0.88).ToString());
                wb.Worksheets[1].Cells.Replace("<11>", (potenciagerada * 5.22 * 30 * 0.83 * 0.88).ToString());
                wb.Worksheets[1].Cells.Replace("<12>", (potenciagerada * 5.59 * 31 * 0.83 * 0.88).ToString());
            }
            else if (rbtn15.Checked)
            {
                wb.Worksheets[1].Cells.Replace("<1>", (potenciagerada * 5.48 * 31 * 0.83 * 0.85).ToString());
                wb.Worksheets[1].Cells.Replace("<2>", (potenciagerada * 5.7 * 29 * 0.83 * 0.85).ToString());
                wb.Worksheets[1].Cells.Replace("<3>", (potenciagerada * 4.85 * 31 * 0.83 * 0.85).ToString());
                wb.Worksheets[1].Cells.Replace("<4>", (potenciagerada * 4.59 * 30 * 0.83 * 0.85).ToString());
                wb.Worksheets[1].Cells.Replace("<5>", (potenciagerada * 3.95 * 31 * 0.83 * 0.85).ToString());
                wb.Worksheets[1].Cells.Replace("<6>", (potenciagerada * 3.76 * 30 * 0.83 * 0.85).ToString());
                wb.Worksheets[1].Cells.Replace("<7>", (potenciagerada * 4.01 * 31 * 0.83 * 0.85).ToString());
                wb.Worksheets[1].Cells.Replace("<8>", (potenciagerada * 4.86 * 31 * 0.83 * 0.85).ToString());
                wb.Worksheets[1].Cells.Replace("<9>", (potenciagerada * 5.08 * 30 * 0.83 * 0.85).ToString());
                wb.Worksheets[1].Cells.Replace("<10>", (potenciagerada * 5.37 * 31 * 0.83 * 0.85).ToString());
                wb.Worksheets[1].Cells.Replace("<11>", (potenciagerada * 5.22 * 30 * 0.83 * 0.85).ToString());
                wb.Worksheets[1].Cells.Replace("<12>", (potenciagerada * 5.59 * 31 * 0.83 * 0.85).ToString());
            }
            else if (rbtn20.Checked)
            {
                wb.Worksheets[1].Cells.Replace("<1>", (potenciagerada * 5.48 * 31 * 0.83 * 0.80).ToString());
                wb.Worksheets[1].Cells.Replace("<2>", (potenciagerada * 5.7 * 29 * 0.83 * 0.80).ToString());
                wb.Worksheets[1].Cells.Replace("<3>", (potenciagerada * 4.85 * 31 * 0.83 * 0.80).ToString());
                wb.Worksheets[1].Cells.Replace("<4>", (potenciagerada * 4.59 * 30 * 0.83 * 0.80).ToString());
                wb.Worksheets[1].Cells.Replace("<5>", (potenciagerada * 3.95 * 31 * 0.83 * 0.80).ToString());
                wb.Worksheets[1].Cells.Replace("<6>", (potenciagerada * 3.76 * 30 * 0.83 * 0.80).ToString());
                wb.Worksheets[1].Cells.Replace("<7>", (potenciagerada * 4.01 * 31 * 0.83 * 0.80).ToString());
                wb.Worksheets[1].Cells.Replace("<8>", (potenciagerada * 4.86 * 31 * 0.83 * 0.80).ToString());
                wb.Worksheets[1].Cells.Replace("<9>", (potenciagerada * 5.08 * 30 * 0.83 * 0.80).ToString());
                wb.Worksheets[1].Cells.Replace("<10>", (potenciagerada * 5.37 * 31 * 0.83 * 0.80).ToString());
                wb.Worksheets[1].Cells.Replace("<11>", (potenciagerada * 5.22 * 30 * 0.83 * 0.80).ToString());
                wb.Worksheets[1].Cells.Replace("<12>", (potenciagerada * 5.59 * 31 * 0.83 * 0.80).ToString());
            }
            else if (rbtn25.Checked)
            {
                wb.Worksheets[1].Cells.Replace("<1>", (potenciagerada * 5.48 * 31 * 0.83 * 0.75).ToString());
                wb.Worksheets[1].Cells.Replace("<2>", (potenciagerada * 5.7 * 29 * 0.83 * 0.75).ToString());
                wb.Worksheets[1].Cells.Replace("<3>", (potenciagerada * 4.85 * 31 * 0.83 * 0.75).ToString());
                wb.Worksheets[1].Cells.Replace("<4>", (potenciagerada * 4.59 * 30 * 0.83 * 0.75).ToString());
                wb.Worksheets[1].Cells.Replace("<5>", (potenciagerada * 3.95 * 31 * 0.83 * 0.75).ToString());
                wb.Worksheets[1].Cells.Replace("<6>", (potenciagerada * 3.76 * 30 * 0.83 * 0.75).ToString());
                wb.Worksheets[1].Cells.Replace("<7>", (potenciagerada * 4.01 * 31 * 0.83 * 0.75).ToString());
                wb.Worksheets[1].Cells.Replace("<8>", (potenciagerada * 4.86 * 31 * 0.83 * 0.75).ToString());
                wb.Worksheets[1].Cells.Replace("<9>", (potenciagerada * 5.08 * 30 * 0.83 * 0.75).ToString());
                wb.Worksheets[1].Cells.Replace("<10>", (potenciagerada * 5.37 * 31 * 0.83 * 0.75).ToString());
                wb.Worksheets[1].Cells.Replace("<11>", (potenciagerada * 5.22 * 30 * 0.83 * 0.75).ToString());
                wb.Worksheets[1].Cells.Replace("<12>", (potenciagerada * 5.59 * 31 * 0.83 * 0.75).ToString());
            }
            else if (rbtn30.Checked)
            {
                wb.Worksheets[1].Cells.Replace("<1>", (potenciagerada * 5.48 * 31 * 0.83 * 1.10).ToString());
                wb.Worksheets[1].Cells.Replace("<2>", (potenciagerada * 5.7 * 29 * 0.83 * 1.10).ToString());
                wb.Worksheets[1].Cells.Replace("<3>", (potenciagerada * 4.85 * 31 * 0.83 * 1.10).ToString());
                wb.Worksheets[1].Cells.Replace("<4>", (potenciagerada * 4.59 * 30 * 0.83 * 1.10).ToString());
                wb.Worksheets[1].Cells.Replace("<5>", (potenciagerada * 3.95 * 31 * 0.83 * 1.10).ToString());
                wb.Worksheets[1].Cells.Replace("<6>", (potenciagerada * 3.76 * 30 * 0.83 * 1.10).ToString());
                wb.Worksheets[1].Cells.Replace("<7>", (potenciagerada * 4.01 * 31 * 0.83 * 1.10).ToString());
                wb.Worksheets[1].Cells.Replace("<8>", (potenciagerada * 4.86 * 31 * 0.83 * 1.10).ToString());
                wb.Worksheets[1].Cells.Replace("<9>", (potenciagerada * 5.08 * 30 * 0.83 * 1.10).ToString());
                wb.Worksheets[1].Cells.Replace("<10>", (potenciagerada * 5.37 * 31 * 0.83 * 1.10).ToString());
                wb.Worksheets[1].Cells.Replace("<11>", (potenciagerada * 5.22 * 30 * 0.83 * 1.10).ToString());
                wb.Worksheets[1].Cells.Replace("<12>", (potenciagerada * 5.59 * 31 * 0.83 * 1.10).ToString());
            }
            else if (rbtn35.Checked)
            {
                wb.Worksheets[1].Cells.Replace("<1>", (potenciagerada * 5.48 * 31 * 0.83 * 1.075).ToString());
                wb.Worksheets[1].Cells.Replace("<2>", (potenciagerada * 5.7 * 29 * 0.83 * 1.075).ToString());
                wb.Worksheets[1].Cells.Replace("<3>", (potenciagerada * 4.85 * 31 * 0.83 * 1.075).ToString());
                wb.Worksheets[1].Cells.Replace("<4>", (potenciagerada * 4.59 * 30 * 0.83 * 1.075).ToString());
                wb.Worksheets[1].Cells.Replace("<5>", (potenciagerada * 3.95 * 31 * 0.83 * 1.075).ToString());
                wb.Worksheets[1].Cells.Replace("<6>", (potenciagerada * 3.76 * 30 * 0.83 * 1.075).ToString());
                wb.Worksheets[1].Cells.Replace("<7>", (potenciagerada * 4.01 * 31 * 0.83 * 1.075).ToString());
                wb.Worksheets[1].Cells.Replace("<8>", (potenciagerada * 4.86 * 31 * 0.83 * 1.075).ToString());
                wb.Worksheets[1].Cells.Replace("<9>", (potenciagerada * 5.08 * 30 * 0.83 * 1.075).ToString());
                wb.Worksheets[1].Cells.Replace("<10>", (potenciagerada * 5.37 * 31 * 0.83 * 1.075).ToString());
                wb.Worksheets[1].Cells.Replace("<11>", (potenciagerada * 5.22 * 30 * 0.83 * 1.075).ToString());
                wb.Worksheets[1].Cells.Replace("<12>", (potenciagerada * 5.59 * 31 * 0.83 * 1.075).ToString());
            }
            else if (rbtn40.Checked)
            {
                wb.Worksheets[1].Cells.Replace("<1>", (potenciagerada * 5.48 * 31 * 0.83 * 1.05).ToString());
                wb.Worksheets[1].Cells.Replace("<2>", (potenciagerada * 5.7 * 29 * 0.83 * 1.05).ToString());
                wb.Worksheets[1].Cells.Replace("<3>", (potenciagerada * 4.85 * 31 * 0.83 * 1.05).ToString());
                wb.Worksheets[1].Cells.Replace("<4>", (potenciagerada * 4.59 * 30 * 0.83 * 1.05).ToString());
                wb.Worksheets[1].Cells.Replace("<5>", (potenciagerada * 3.95 * 31 * 0.83 * 1.05).ToString());
                wb.Worksheets[1].Cells.Replace("<6>", (potenciagerada * 3.76 * 30 * 0.83 * 1.05).ToString());
                wb.Worksheets[1].Cells.Replace("<7>", (potenciagerada * 4.01 * 31 * 0.83 * 1.05).ToString());
                wb.Worksheets[1].Cells.Replace("<8>", (potenciagerada * 4.86 * 31 * 0.83 * 1.05).ToString());
                wb.Worksheets[1].Cells.Replace("<9>", (potenciagerada * 5.08 * 30 * 0.83 * 1.05).ToString());
                wb.Worksheets[1].Cells.Replace("<10>", (potenciagerada * 5.37 * 31 * 0.83 * 1.05).ToString());
                wb.Worksheets[1].Cells.Replace("<11>", (potenciagerada * 5.22 * 30 * 0.83 * 1.05).ToString());
                wb.Worksheets[1].Cells.Replace("<12>", (potenciagerada * 5.59 * 31 * 0.83 * 1.05).ToString());
            }
            else if (rbtn0.Checked)
            {
                wb.Worksheets[1].Cells.Replace("<1>", (potenciagerada * 5.48 * 31 * 0.83).ToString());
                wb.Worksheets[1].Cells.Replace("<2>", (potenciagerada * 5.7 * 29 * 0.83).ToString());
                wb.Worksheets[1].Cells.Replace("<3>", (potenciagerada * 4.85 * 31 * 0.83).ToString());
                wb.Worksheets[1].Cells.Replace("<4>", (potenciagerada * 4.59 * 30 * 0.83).ToString());
                wb.Worksheets[1].Cells.Replace("<5>", (potenciagerada * 3.95 * 31 * 0.83).ToString());
                wb.Worksheets[1].Cells.Replace("<6>", (potenciagerada * 3.76 * 30 * 0.83).ToString());
                wb.Worksheets[1].Cells.Replace("<7>", (potenciagerada * 4.01 * 31 * 0.83).ToString());
                wb.Worksheets[1].Cells.Replace("<8>", (potenciagerada * 4.86 * 31 * 0.83).ToString());
                wb.Worksheets[1].Cells.Replace("<9>", (potenciagerada * 5.08 * 30 * 0.83).ToString());
                wb.Worksheets[1].Cells.Replace("<10>", (potenciagerada * 5.37 * 31 * 0.83).ToString());
                wb.Worksheets[1].Cells.Replace("<11>", (potenciagerada * 5.22 * 30 * 0.83).ToString());
                wb.Worksheets[1].Cells.Replace("<12>", (potenciagerada * 5.59 * 31 * 0.83).ToString());
            }
            pgbstatusorca.Value++;

            if (qtdorc == 1)
            {
                //Disponibilidade
                int disponibilidade = orcamento.UmaUC(cbxpadrao.Text);
                wb.Worksheets[1].Cells.Replace("<disp>", disponibilidade.ToString());
                pgbstatusorca.Value++;

                double jan = orcamento.Zeradisp(double.Parse(txtjan.Text), cbxpadrao.Text);
                double fev = orcamento.Zeradisp(double.Parse(txtfev.Text), cbxpadrao.Text);
                double mar = orcamento.Zeradisp(double.Parse(txtmar.Text), cbxpadrao.Text);
                double abr = orcamento.Zeradisp(double.Parse(txtabr.Text), cbxpadrao.Text);
                double mai = orcamento.Zeradisp(double.Parse(txtmai.Text), cbxpadrao.Text);
                double jun = orcamento.Zeradisp(double.Parse(txtjun.Text), cbxpadrao.Text);
                double jul = orcamento.Zeradisp(double.Parse(txtjul.Text), cbxpadrao.Text);
                double ago = orcamento.Zeradisp(double.Parse(txtago.Text), cbxpadrao.Text);
                double set = orcamento.Zeradisp(double.Parse(txtset.Text), cbxpadrao.Text);
                double outu = orcamento.Zeradisp(double.Parse(txtout.Text), cbxpadrao.Text);
                double nov = orcamento.Zeradisp(double.Parse(txtnov.Text), cbxpadrao.Text);
                double dez = orcamento.Zeradisp(double.Parse(txtdez.Text), cbxpadrao.Text);
                pgbstatusorca.Value++;

                wb.Worksheets[1].Cells.Replace("<janeiro>", jan.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<fevereiro>", fev.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<marco>", mar.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<abril>", abr.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<maio>", mai.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<junho>", jun.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<julho>", jul.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<agosto>", ago.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<setembro>", set.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<outubro>", outu.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<novembro>", nov.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<dezembro>", dez.ToString("0"));
                pgbstatusorca.Value++;

                wb.Worksheets[1].Cells.Replace("<unidadecons>", cbxclasse.Text + " " + cbxpadrao.Text);
            }
            else if(qtdorc == 2)
            {
                //Disponibilidade
                int disponibilidade = orcamento.DuasUC(cbxpadrao.Text, cbxpadrao2.Text);
                wb.Worksheets[1].Cells.Replace("<disp>", disponibilidade.ToString());

                double jan = orcamento.Zeradisp(double.Parse(txtjan.Text), cbxpadrao.Text);
                double fev = orcamento.Zeradisp(double.Parse(txtfev.Text), cbxpadrao.Text);
                double mar = orcamento.Zeradisp(double.Parse(txtmar.Text), cbxpadrao.Text);
                double abr = orcamento.Zeradisp(double.Parse(txtabr.Text), cbxpadrao.Text);
                double mai = orcamento.Zeradisp(double.Parse(txtmai.Text), cbxpadrao.Text);
                double jun = orcamento.Zeradisp(double.Parse(txtjun.Text), cbxpadrao.Text);
                double jul = orcamento.Zeradisp(double.Parse(txtjul.Text), cbxpadrao.Text);
                double ago = orcamento.Zeradisp(double.Parse(txtago.Text), cbxpadrao.Text);
                double set = orcamento.Zeradisp(double.Parse(txtset.Text), cbxpadrao.Text);
                double outu = orcamento.Zeradisp(double.Parse(txtout.Text), cbxpadrao.Text);
                double nov = orcamento.Zeradisp(double.Parse(txtnov.Text), cbxpadrao.Text);
                double dez = orcamento.Zeradisp(double.Parse(txtdez.Text), cbxpadrao.Text);

                double jan2 = orcamento.Zeradisp(double.Parse(txtjan2.Text), cbxpadrao2.Text);
                double fev2 = orcamento.Zeradisp(double.Parse(txtfev2.Text), cbxpadrao2.Text);
                double mar2 = orcamento.Zeradisp(double.Parse(txtmar2.Text), cbxpadrao2.Text);
                double abr2 = orcamento.Zeradisp(double.Parse(txtabr2.Text), cbxpadrao2.Text);
                double mai2 = orcamento.Zeradisp(double.Parse(txtmai2.Text), cbxpadrao2.Text);
                double jun2 = orcamento.Zeradisp(double.Parse(txtjun2.Text), cbxpadrao2.Text);
                double jul2 = orcamento.Zeradisp(double.Parse(txtjul2.Text), cbxpadrao2.Text);
                double ago2 = orcamento.Zeradisp(double.Parse(txtago2.Text), cbxpadrao2.Text);
                double set2 = orcamento.Zeradisp(double.Parse(txtset2.Text), cbxpadrao2.Text);
                double outu2 = orcamento.Zeradisp(double.Parse(txtout2.Text), cbxpadrao2.Text);
                double nov2 = orcamento.Zeradisp(double.Parse(txtnov2.Text), cbxpadrao2.Text);
                double dez2 = orcamento.Zeradisp(double.Parse(txtdez2.Text), cbxpadrao2.Text);

                wb.Worksheets[1].Cells.Replace("<janeiro>", jan.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<janeiro2>", jan2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<fevereiro>", fev.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<fevereiro2>", fev2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<marco>", mar.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<marco2>", mar2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<abril>", abr.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<abril2>", abr2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<maio>", mai.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<maio2>", mai2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<junho>", jun.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<junho2>", jun2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<julho>", jul.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<julho2>", jul2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<agosto>", ago.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<agosto2>", ago2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<setembro>", set.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<setembro2>", set2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<outubro>", outu.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<outubro2>", outu2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<novembro>", nov.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<novembro2>", nov2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<dezembro>", dez.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<dezembro2>", dez2.ToString("0"));

                wb.Worksheets[1].Cells.Replace("<unidadecons>", cbxclasse.Text + " " + cbxpadrao.Text);
                wb.Worksheets[1].Cells.Replace("<unidadecons2>", cbxclasse2.Text + " " + cbxpadrao2.Text);
            }
            else if(qtdorc == 3)
            {
                //Disponibilidade
                int disponibilidade = orcamento.TresUC(cbxpadrao.Text, cbxpadrao2.Text, cbxpadrao3.Text);
                wb.Worksheets[1].Cells.Replace("<disp>", disponibilidade.ToString());

                double jan = orcamento.Zeradisp(double.Parse(txtjan.Text), cbxpadrao.Text);
                double fev = orcamento.Zeradisp(double.Parse(txtfev.Text), cbxpadrao.Text);
                double mar = orcamento.Zeradisp(double.Parse(txtmar.Text), cbxpadrao.Text);
                double abr = orcamento.Zeradisp(double.Parse(txtabr.Text), cbxpadrao.Text);
                double mai = orcamento.Zeradisp(double.Parse(txtmai.Text), cbxpadrao.Text);
                double jun = orcamento.Zeradisp(double.Parse(txtjun.Text), cbxpadrao.Text);
                double jul = orcamento.Zeradisp(double.Parse(txtjul.Text), cbxpadrao.Text);
                double ago = orcamento.Zeradisp(double.Parse(txtago.Text), cbxpadrao.Text);
                double set = orcamento.Zeradisp(double.Parse(txtset.Text), cbxpadrao.Text);
                double outu = orcamento.Zeradisp(double.Parse(txtout.Text), cbxpadrao.Text);
                double nov = orcamento.Zeradisp(double.Parse(txtnov.Text), cbxpadrao.Text);
                double dez = orcamento.Zeradisp(double.Parse(txtdez.Text), cbxpadrao.Text);

                double jan2 = orcamento.Zeradisp(double.Parse(txtjan2.Text), cbxpadrao2.Text);
                double fev2 = orcamento.Zeradisp(double.Parse(txtfev2.Text), cbxpadrao2.Text);
                double mar2 = orcamento.Zeradisp(double.Parse(txtmar2.Text), cbxpadrao2.Text);
                double abr2 = orcamento.Zeradisp(double.Parse(txtabr2.Text), cbxpadrao2.Text);
                double mai2 = orcamento.Zeradisp(double.Parse(txtmai2.Text), cbxpadrao2.Text);
                double jun2 = orcamento.Zeradisp(double.Parse(txtjun2.Text), cbxpadrao2.Text);
                double jul2 = orcamento.Zeradisp(double.Parse(txtjul2.Text), cbxpadrao2.Text);
                double ago2 = orcamento.Zeradisp(double.Parse(txtago2.Text), cbxpadrao2.Text);
                double set2 = orcamento.Zeradisp(double.Parse(txtset2.Text), cbxpadrao2.Text);
                double outu2 = orcamento.Zeradisp(double.Parse(txtout2.Text), cbxpadrao2.Text);
                double nov2 = orcamento.Zeradisp(double.Parse(txtnov2.Text), cbxpadrao2.Text);
                double dez2 = orcamento.Zeradisp(double.Parse(txtdez2.Text), cbxpadrao2.Text);

                double jan3 = orcamento.Zeradisp(double.Parse(txtjan3.Text), cbxpadrao3.Text);
                double fev3 = orcamento.Zeradisp(double.Parse(txtfev3.Text), cbxpadrao3.Text);
                double mar3 = orcamento.Zeradisp(double.Parse(txtmar3.Text), cbxpadrao3.Text);
                double abr3 = orcamento.Zeradisp(double.Parse(txtabr3.Text), cbxpadrao3.Text);
                double mai3 = orcamento.Zeradisp(double.Parse(txtmai3.Text), cbxpadrao3.Text);
                double jun3 = orcamento.Zeradisp(double.Parse(txtjun3.Text), cbxpadrao3.Text);
                double jul3 = orcamento.Zeradisp(double.Parse(txtjul3.Text), cbxpadrao3.Text);
                double ago3 = orcamento.Zeradisp(double.Parse(txtago3.Text), cbxpadrao3.Text);
                double set3 = orcamento.Zeradisp(double.Parse(txtset3.Text), cbxpadrao3.Text);
                double outu3 = orcamento.Zeradisp(double.Parse(txtout3.Text), cbxpadrao3.Text);
                double nov3 = orcamento.Zeradisp(double.Parse(txtnov3.Text), cbxpadrao3.Text);
                double dez3 = orcamento.Zeradisp(double.Parse(txtdez3.Text), cbxpadrao3.Text);

                wb.Worksheets[1].Cells.Replace("<janeiro>", jan.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<janeiro2>", jan2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<janeiro3>", jan3.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<fevereiro>", fev.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<fevereiro2>", fev2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<fevereiro3>", fev3.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<marco>", mar.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<marco2>", mar2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<marco3>", mar3.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<abril>", abr.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<abril2>", abr2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<abril3>", abr3.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<maio>", mai.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<maio2>", mai2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<maio3>", mai3.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<junho>", jun.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<junho2>", jun2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<junho3>", jun3.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<julho>", jul.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<julho2>", jul2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<julho3>", jul3.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<agosto>", ago.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<agosto2>", ago2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<agosto3>", ago3.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<setembro>", set.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<setembro2>", set2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<setembro3>", set3.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<outubro>", outu.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<outubro2>", outu2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<outubro3>", outu3.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<novembro>", nov.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<novembro2>", nov2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<novembro3>", nov3.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<dezembro>", dez.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<dezembro2>", dez2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<dezembro3>", dez3.ToString("0"));

                wb.Worksheets[1].Cells.Replace("<unidadecons>", cbxclasse.Text + " " + cbxpadrao.Text);
                wb.Worksheets[1].Cells.Replace("<unidadecons2>", cbxclasse2.Text + " " + cbxpadrao2.Text);
                wb.Worksheets[1].Cells.Replace("<unidadecons3>", cbxclasse3.Text + " " + cbxpadrao3.Text);
            }
            else if (qtdorc == 4)
            {
                //Disponibilidade
                int disponibilidade = orcamento.QuatroUC(cbxpadrao.Text, cbxpadrao2.Text, cbxpadrao3.Text, cbxpadrao4.Text);
                wb.Worksheets[1].Cells.Replace("<disp>", disponibilidade.ToString());

                double jan = orcamento.Zeradisp(double.Parse(txtjan.Text), cbxpadrao.Text);
                double fev = orcamento.Zeradisp(double.Parse(txtfev.Text), cbxpadrao.Text);
                double mar = orcamento.Zeradisp(double.Parse(txtmar.Text), cbxpadrao.Text);
                double abr = orcamento.Zeradisp(double.Parse(txtabr.Text), cbxpadrao.Text);
                double mai = orcamento.Zeradisp(double.Parse(txtmai.Text), cbxpadrao.Text);
                double jun = orcamento.Zeradisp(double.Parse(txtjun.Text), cbxpadrao.Text);
                double jul = orcamento.Zeradisp(double.Parse(txtjul.Text), cbxpadrao.Text);
                double ago = orcamento.Zeradisp(double.Parse(txtago.Text), cbxpadrao.Text);
                double set = orcamento.Zeradisp(double.Parse(txtset.Text), cbxpadrao.Text);
                double outu = orcamento.Zeradisp(double.Parse(txtout.Text), cbxpadrao.Text);
                double nov = orcamento.Zeradisp(double.Parse(txtnov.Text), cbxpadrao.Text);
                double dez = orcamento.Zeradisp(double.Parse(txtdez.Text), cbxpadrao.Text);

                double jan2 = orcamento.Zeradisp(double.Parse(txtjan2.Text), cbxpadrao2.Text);
                double fev2 = orcamento.Zeradisp(double.Parse(txtfev2.Text), cbxpadrao2.Text);
                double mar2 = orcamento.Zeradisp(double.Parse(txtmar2.Text), cbxpadrao2.Text);
                double abr2 = orcamento.Zeradisp(double.Parse(txtabr2.Text), cbxpadrao2.Text);
                double mai2 = orcamento.Zeradisp(double.Parse(txtmai2.Text), cbxpadrao2.Text);
                double jun2 = orcamento.Zeradisp(double.Parse(txtjun2.Text), cbxpadrao2.Text);
                double jul2 = orcamento.Zeradisp(double.Parse(txtjul2.Text), cbxpadrao2.Text);
                double ago2 = orcamento.Zeradisp(double.Parse(txtago2.Text), cbxpadrao2.Text);
                double set2 = orcamento.Zeradisp(double.Parse(txtset2.Text), cbxpadrao2.Text);
                double outu2 = orcamento.Zeradisp(double.Parse(txtout2.Text), cbxpadrao2.Text);
                double nov2 = orcamento.Zeradisp(double.Parse(txtnov2.Text), cbxpadrao2.Text);
                double dez2 = orcamento.Zeradisp(double.Parse(txtdez2.Text), cbxpadrao2.Text);

                double jan3 = orcamento.Zeradisp(double.Parse(txtjan3.Text), cbxpadrao3.Text);
                double fev3 = orcamento.Zeradisp(double.Parse(txtfev3.Text), cbxpadrao3.Text);
                double mar3 = orcamento.Zeradisp(double.Parse(txtmar3.Text), cbxpadrao3.Text);
                double abr3 = orcamento.Zeradisp(double.Parse(txtabr3.Text), cbxpadrao3.Text);
                double mai3 = orcamento.Zeradisp(double.Parse(txtmai3.Text), cbxpadrao3.Text);
                double jun3 = orcamento.Zeradisp(double.Parse(txtjun3.Text), cbxpadrao3.Text);
                double jul3 = orcamento.Zeradisp(double.Parse(txtjul3.Text), cbxpadrao3.Text);
                double ago3 = orcamento.Zeradisp(double.Parse(txtago3.Text), cbxpadrao3.Text);
                double set3 = orcamento.Zeradisp(double.Parse(txtset3.Text), cbxpadrao3.Text);
                double outu3 = orcamento.Zeradisp(double.Parse(txtout3.Text), cbxpadrao3.Text);
                double nov3 = orcamento.Zeradisp(double.Parse(txtnov3.Text), cbxpadrao3.Text);
                double dez3 = orcamento.Zeradisp(double.Parse(txtdez3.Text), cbxpadrao3.Text);

                double jan4 = orcamento.Zeradisp(double.Parse(txtjan4.Text), cbxpadrao.Text);
                double fev4 = orcamento.Zeradisp(double.Parse(txtfev4.Text), cbxpadrao.Text);
                double mar4 = orcamento.Zeradisp(double.Parse(txtmar4.Text), cbxpadrao.Text);
                double abr4 = orcamento.Zeradisp(double.Parse(txtabr4.Text), cbxpadrao.Text);
                double mai4 = orcamento.Zeradisp(double.Parse(txtmai4.Text), cbxpadrao.Text);
                double jun4 = orcamento.Zeradisp(double.Parse(txtjun4.Text), cbxpadrao.Text);
                double jul4 = orcamento.Zeradisp(double.Parse(txtjul4.Text), cbxpadrao.Text);
                double ago4 = orcamento.Zeradisp(double.Parse(txtago4.Text), cbxpadrao.Text);
                double set4 = orcamento.Zeradisp(double.Parse(txtset4.Text), cbxpadrao.Text);
                double outu4 = orcamento.Zeradisp(double.Parse(txtout4.Text), cbxpadrao.Text);
                double nov4 = orcamento.Zeradisp(double.Parse(txtnov4.Text), cbxpadrao.Text);
                double dez4 = orcamento.Zeradisp(double.Parse(txtdez4.Text), cbxpadrao.Text);

                wb.Worksheets[1].Cells.Replace("<janeiro>", jan.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<fevereiro>", fev.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<marco>", mar.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<abril>", abr.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<maio>", mai.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<junho>", jun.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<julho>", jul.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<agosto>", ago.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<setembro>", set.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<outubro>", outu.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<novembro>", nov.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<dezembro>", dez.ToString("0"));

                wb.Worksheets[1].Cells.Replace("<janeiro2>", jan2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<fevereiro2>", fev2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<marco2>", mar2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<abril2>", abr2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<maio2>", mai2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<junho2>", jun2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<julho2>", jul2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<agosto2>", ago2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<setembro2>", set2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<outubro2>", outu2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<novembro2>", nov2.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<dezembro2>", dez2.ToString("0"));

                wb.Worksheets[1].Cells.Replace("<janeiro3>", jan3.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<fevereiro3>", fev3.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<marco3>", mar3.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<abril3>", abr3.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<maio3>", mai3.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<junho3>", jun3.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<julho3>", jul3.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<agosto3>", ago3.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<setembro3>", set3.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<outubro3>", outu3.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<novembro3>", nov3.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<dezembro3>", dez3.ToString("0"));

                wb.Worksheets[1].Cells.Replace("<janeiro4>", jan4.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<fevereiro4>", fev4.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<marco4>", mar4.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<abril4>", abr4.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<maio4>", mai4.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<junho4>", jun4.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<julho4>", jul4.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<agosto4>", ago4.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<setembro4>", set4.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<outubro4>", outu4.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<novembro4>", nov4.ToString("0"));
                wb.Worksheets[1].Cells.Replace("<dezembro4>", dez4.ToString("0"));

                wb.Worksheets[1].Cells.Replace("<unidadecons>", cbxclasse.Text + " " + cbxpadrao.Text);
                wb.Worksheets[1].Cells.Replace("<unidadecons2>", cbxclasse2.Text + " " + cbxpadrao2.Text);
                wb.Worksheets[1].Cells.Replace("<unidadecons3>", cbxclasse3.Text + " " + cbxpadrao3.Text);
                wb.Worksheets[1].Cells.Replace("<unidadecons4>", cbxclasse4.Text + " " + cbxpadrao4.Text);
            }

            
            wb.SaveAs(@"C:\Centraliza\Orçamentos\Geração " + txtnome.Text + ".xlsx");
            pgbstatusorca.Value++;

            wb.Close();
            xlapp.Quit();
            pgbstatusorca.Value++;
        }
        public void GerarPlanilhaFin()
        {
            Excel.Application xlapp = new Excel.Application();
            Excel.Workbook wb = default(Excel.Workbook);
            pgbstatusorca.Value++;

            wb = xlapp.Workbooks.Open(@"C:\Centraliza\Centraliza\temp-financeiro-2.xlsx");
            pgbstatusorca.Value++;

            //Calculo Tarifa
            double tarifao = double.Parse(txtkwh.Text);
            pgbstatusorca.Value++;

            //Produzido
            wb.Worksheets[1].Cells.Replace("<gerano>", gerano.ToString());
            wb.Worksheets[1].Cells.Replace("<tarifa>", tarifao.ToString("0.0000000000"));
            wb.Worksheets[1].Cells.Replace("<valorsistema>", txtvalorsist.Text);
            pgbstatusorca.Value++;

            //Payback
            double pb;
            if (txtcustoinversor.Text == "" || txtcustoinversor.Enabled == false)
            {
                wb.Worksheets[1].Cells.Replace("<custo>", "0");
                pb = orcamento.CalculaPayback(tarifao, 0, double.Parse(txtvalorsist.Text), gerano);
            }
            else
            {
                wb.Worksheets[1].Cells.Replace("<custo>", double.Parse(txtcustoinversor.Text) / 10);
                pb = orcamento.CalculaPayback(tarifao, double.Parse(txtcustoinversor.Text), double.Parse(txtvalorsist.Text), gerano);
            }
            pgbstatusorca.Value++;

            wb.SaveAs(@"C:\Centraliza\Orçamentos\Retorno Financeiro " + txtnome.Text + ".xlsx");
            pgbstatusorca.Value++;

            wb.Close();
            xlapp.Quit();
            pgbstatusorca.Value++;
        }
        private void CriaMemorial(object filename, object SaveAs)
        {
            Word.Application wordApp = new Word.Application();
            object missing = Missing.Value;
            Word.Document mywordDoc = null;

            if (File.Exists((string)filename))
            {
                object readOnly = false;
                object isVisible = false;
                wordApp.Visible = false;

                mywordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing);
                mywordDoc.Activate();

                func.PesquisaDimensoesMod(func.ModeloModulo, Banco);
                double dimensao = 2 * double.Parse(func.QuantidadeModulos);
                string dimensao1 = dimensao.ToString();
                dimensao1 = string.Format("{0:0,0.00}", dimensao);
                double germes = (((double.Parse(func.PotenciaMod) * double.Parse(func.QuantidadeModulos)) * 0.83) / 1000) * 30 * 4.87;
                string germes1 = germes.ToString();
                germes1 = string.Format("{0:0,0}", germes);

                pgbstatusproj.Value++;
                this.orcamento.AcharESubstituir(wordApp, "<data>", DateTime.Now.ToString("dd' de 'MMMM' de 'yyyy"));
                this.orcamento.AcharESubstituir(wordApp, "<nome>", lbltitularprojeto.Text);
                this.orcamento.AcharESubstituir(wordApp, "<endereco>", lblenderecoprojeto.Text);
                this.orcamento.AcharESubstituir(wordApp, "<num>", lblnumeroenderecoprojeto.Text);
                this.orcamento.AcharESubstituir(wordApp, "<bairro>", lblbairroprojeto.Text);
                this.orcamento.AcharESubstituir(wordApp, "<cidade>", lblcidadeprojeto.Text);
                this.orcamento.AcharESubstituir(wordApp, "<uf>", func.UF);
                this.orcamento.AcharESubstituir(wordApp, "<latitude>", mtxtlatitudeproj.Text);
                this.orcamento.AcharESubstituir(wordApp, "<longitude>", mtxtlongitudeproj.Text);
                this.orcamento.AcharESubstituir(wordApp, "<numeroins>", txtnuminstproj.Text);
                pgbstatusproj.Value++;
                if (lblclasseprojeto.Text == "Comercial Monofásico" || lblclasseprojeto.Text == "Residencial Monofásico" || lblclasseprojeto.Text == "Rural Monofásico" || lblclasseprojeto.Text == "Industrial Monofásico")
                {
                    this.orcamento.AcharESubstituir(wordApp, "<padrao>", "Monofásico");
                }
                else if (lblclasseprojeto.Text == "Comercial Bifásico" || lblclasseprojeto.Text == "Residencial Bifásico" || lblclasseprojeto.Text == "Rural Bifásico" || lblclasseprojeto.Text == "Industrial Bifásico")
                {
                    this.orcamento.AcharESubstituir(wordApp, "<padrao>", "Bifásico");
                }
                else
                {
                    this.orcamento.AcharESubstituir(wordApp, "<padrao>", "Trifásico");
                }
                string aux = new string(cbxdisjproj.Text.Where(char.IsDigit).ToArray());
                this.orcamento.AcharESubstituir(wordApp, "<disjamp>", aux);
                this.orcamento.AcharESubstituir(wordApp, "<classe>", cbxclasseproj.Text);
                this.orcamento.AcharESubstituir(wordApp, "<conmes>", func.MediaConsumo);
                pgbstatusproj.Value++;
                //Invesores e Paineis
                this.orcamento.AcharESubstituir(wordApp, "<qtdinv>", func.QuantidadeInversores);
                this.orcamento.AcharESubstituir(wordApp, "<qtdmod>", func.QuantidadeModulos);
                func.PesquisaModMod(func.ModeloModulo, Banco);
                this.orcamento.AcharESubstituir(wordApp, "<marcamod>", func.MarcaMod + " " + func.ModeloModulo + " " + func.Material + " " + func.Celulas + " " + func.PotenciaMod);
                func.PesquisaModInv(func.ModeloInversor,Banco);
                //Plural Inversores
                if (Int32.Parse(func.QuantidadeInversores) > 1)
                {
                    if (func.MarcaInversor == "AP System" || func.ModeloInversor == "Reno560" || func.ModeloInversor == "Reno560-LV")
                    {
                        this.orcamento.AcharESubstituir(wordApp, "<marcainv>", "Microinversores " + func.MarcaInversor + " " + func.ModeloInversor + " de " + func.PotenciaInv + "W " + func.Fases + " " + func.Tensao);
                    }
                    else
                    {
                        this.orcamento.AcharESubstituir(wordApp, "<marcainv>", "Inversores " + func.MarcaInversor + " " + func.ModeloInversor + " de " + func.PotenciaInv + "W " + func.Fases + " " + func.Tensao);
                    }
                }
                else
                {
                    if (func.MarcaInversor == "AP System" || func.ModeloInversor == "Reno560" || func.ModeloInversor == "Reno560-LV")
                    {
                        this.orcamento.AcharESubstituir(wordApp, "<marcainv>", "Microinversor " + func.MarcaInversor + " " + func.ModeloInversor + " de " + func.PotenciaInv + "W " + func.Fases + " " + func.Tensao);
                    }
                    else
                    {
                        this.orcamento.AcharESubstituir(wordApp, "<marcainv>", "Inversor " + func.MarcaInversor + " " + func.ModeloInversor + " de " + func.PotenciaInv + "W " + func.Fases + " " + func.Tensao);
                    }
                }
                pgbstatusproj.Value++;
                if (cbxstringboxproj.Text != "")
                {
                    this.orcamento.AcharESubstituir(wordApp, "<stringbox>", "String Box CC " + cbxstringboxproj.Text);
                }
                else
                {
                    this.orcamento.AcharESubstituir(wordApp, "<stringbox>", "");
                }
                if (cbxtransformadorproj.Text != "Nenhum")
                {
                    this.orcamento.AcharESubstituir(wordApp, "<transformador>", "Transformador de " + cbxtransformadorproj.Text);
                }
                else
                {
                    this.orcamento.AcharESubstituir(wordApp, "<transformador>", "");
                }
                pgbstatusproj.Value++;
                this.orcamento.AcharESubstituir(wordApp, "<estrutura>", cbxestruturaproj.Text);
                func.PesquisaModInv(func.ModeloInversor,Banco);
                if (func.MarcaInversor == "AP System" || func.ModeloInversor == "Reno560" || func.ModeloInversor == "Reno560-LV")
                {
                   this.orcamento.AcharESubstituir(wordApp, "<conectores>", "CONECTORES BAIQI Macho Femea MEVS próprio para microinversor.");
                }
                else
                {
                   this.orcamento.AcharESubstituir(wordApp, "<conectores>", "CONECTORES MC4  com proteção UV e resistência a amoníaco (conforme a DLG) 1500h 70C/70% RH, 750ppm.");
                }
                pgbstatusproj.Value++;
                this.orcamento.AcharESubstituir(wordApp, "<cabos>", "CABOS SOLARES 6MM ATE 1800V CC ABNT NBR.");
                this.orcamento.AcharESubstituir(wordApp, "<sumdimmod>", dimensao1);
                this.orcamento.AcharESubstituir(wordApp, "<arranjo>", txtarranjoproj.Text);
                this.orcamento.AcharESubstituir(wordApp, "<modeloinv>", func.MarcaInversor + " " + func.ModeloInversor);
                this.orcamento.AcharESubstituir(wordApp, "<modelomod>", func.MarcaMod + " " + func.ModeloModulo);
                this.orcamento.AcharESubstituir(wordApp, "<reginv>", func.RegistroINMETRO);
                this.orcamento.AcharESubstituir(wordApp, "<regmod>", func.RegistroInmetro);
                this.orcamento.AcharESubstituir(wordApp, "<germes>", germes1);
                pgbstatusproj.Value++;

                //Arquivo dos Inversores
                string fileName = func.ModeloInversor + ".txt";
                string sourcePath = @"C:\Centraliza\Dados Equipamentos\Inversores\";
                string targetPath = @"C:\Centraliza\Projeto\";
                string sourceFile = System.IO.Path.Combine(sourcePath, fileName);
                string destFile = System.IO.Path.Combine(targetPath, fileName);
                System.IO.File.Copy(sourceFile, destFile, true);
                pgbstatusproj.Value++;

                //Arquivo dos Paineis
                string fileName2 = func.ModeloModulo + ".txt";
                string sourcePath2 = @"C:\Centraliza\Dados Equipamentos\Modulos\";
                string sourceFile2 = System.IO.Path.Combine(sourcePath2, fileName2);
                string destFile2 = System.IO.Path.Combine(targetPath, fileName2);
                System.IO.File.Copy(sourceFile2, destFile2, true);

                //Registro do INMETRO
                func.SelecionaInversor(Banco, func.ModeloInversor);
                func.SelecionaPainel(Banco,func.ModeloModulo);
                if (func.RegistroINMETRO != "")
                {
                    string Certificado = func.RegistroINMETRO;
                    string pri = Certificado.Substring(0, 6);
                    string ano = Certificado.Substring(7, 4);
                    string pasta = @"C:\Centraliza\Certificados\" + pri + @"\" + ano;
                    //Process.Start("explorer.exe", pasta);
                }
                else
                {
                    MessageBox.Show("O equipamento selecionado não possui certificado atrelado!", "Atenção");
                }
                if (func.RegistroInmetro != "")
                {
                    string Certificado = func.RegistroInmetro;
                    string pri = Certificado.Substring(0, 6);
                    string ano = Certificado.Substring(7, 4);
                    string fileName12 = @"C:\Centraliza\Certificados\" + pri + @"\" + ano;
                    string sourcePath12 = @"C:\Centraliza\Certificados\" + pri + @"\" + ano;
                    string targetPath12 = @"C:\Centraliza\Orçamentos";
                    string sourceFile12 = System.IO.Path.Combine(sourcePath12, fileName12);
                    string destFile12 = System.IO.Path.Combine(targetPath12, fileName12);
                    if (System.IO.Directory.Exists(sourcePath12))
                    {
                        string[] files = System.IO.Directory.GetFiles(sourcePath12);

                        // Copy the files and overwrite destination files if they already exist.
                        foreach (string s in files)
                        {
                            // Use static Path methods to extract only the file name from the path.
                            fileName12 = System.IO.Path.GetFileName(s);
                            destFile12 = System.IO.Path.Combine(targetPath12, fileName12);
                            System.IO.File.Copy(s, destFile, true);
                        }
                    }

                    //Process.Start("explorer.exe", fileName12);
                }
                else
                {
                    MessageBox.Show("O equipamento selecionado não possui certificado atrelado!", "Atenção");
                }
                pgbstatusproj.Value++;
            }
            else
            {
                MessageBox.Show("Arquivo não encontrado!");
            }

            //Save as
            mywordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing);
            pgbstatusproj.Value++;
            mywordDoc.Close();
            wordApp.Quit();
        }
        private void CriaFormulario(object filename, object SaveAs)
        {
            Word.Application wordApp = new Word.Application();
            object missing = Missing.Value;
            Word.Document mywordDoc = null;

            if (File.Exists((string)filename))
            {
                object readOnly = false;
                object isVisible = false;
                wordApp.Visible = false;

                mywordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing);
                mywordDoc.Activate();

                pgbstatusproj.Value++;
                pgbstatusproj.Value++;
                func.PesquisaDimensoesMod(func.ModeloModulo, Banco);
                double dimensao = 2 * double.Parse(func.QuantidadeModulos);
                string dimensao1 = dimensao.ToString();
                dimensao1 = string.Format("{0:0,0.00}", dimensao);
                double germes = (((double.Parse(func.PotenciaMod) * double.Parse(func.QuantidadeModulos)) * 0.83) / 1000) * 30 * 4.87;
                string germes1 = germes.ToString();
                germes1 = string.Format("{0:0,0}", germes);
                potenciagerada = double.Parse(func.PotenciaMod) * double.Parse(func.QuantidadeModulos) / 1000;
                string potger1 = potenciagerada.ToString();
                potger1 = string.Format("{0:0.00}", potenciagerada);

                pgbstatusproj.Value++;
                this.orcamento.AcharESubstituir(wordApp, "<numerocli>", lblnumerocliprojeto.Text);
                this.orcamento.AcharESubstituir(wordApp, "<numeroins>", lblnumeroinstprojeto.Text);
                this.orcamento.AcharESubstituir(wordApp, "<data>", DateTime.Now.ToString("dd' de 'MMMM' de 'yyyy"));
                this.orcamento.AcharESubstituir(wordApp, "<nome>", lbltitularprojeto.Text);
                this.orcamento.AcharESubstituir(wordApp, "<endereco>", lblenderecoprojeto.Text);
                this.orcamento.AcharESubstituir(wordApp, "<num>", lblnumeroenderecoprojeto.Text);
                this.orcamento.AcharESubstituir(wordApp, "<bairro>", lblbairroprojeto.Text);
                this.orcamento.AcharESubstituir(wordApp, "<cidade>", lblcidadeprojeto.Text);
                pgbstatusproj.Value++;
                this.orcamento.AcharESubstituir(wordApp, "<cep>", lblcepprojeto.Text);
                this.orcamento.AcharESubstituir(wordApp, "<uf>", func.UF);
                this.orcamento.AcharESubstituir(wordApp, "<doc>", func.CPF);
                this.orcamento.AcharESubstituir(wordApp, "<comp>", func.Complemento);
                this.orcamento.AcharESubstituir(wordApp, "<telefone>", func.Telefone);
                this.orcamento.AcharESubstituir(wordApp, "<contato>", func.Celular);
                pgbstatusproj.Value++;
                this.orcamento.AcharESubstituir(wordApp, "<latitude>", mtxtlatitudeproj.Text);
                this.orcamento.AcharESubstituir(wordApp, "<longitude>", mtxtlongitudeproj.Text);
                this.orcamento.AcharESubstituir(wordApp, "<fuso>", "23");
                this.orcamento.AcharESubstituir(wordApp, "<kwhinst>", txtcargainstproj.Text);
                string aux = new string(cbxdisjproj.Text.Where(char.IsDigit).ToArray());
                this.orcamento.AcharESubstituir(wordApp, "<disjamp>", aux);
                pgbstatusproj.Value++;
                this.orcamento.AcharESubstituir(wordApp, "<unidadecons>", lblclasseprojeto.Text);
                this.orcamento.AcharESubstituir(wordApp, "<tensao>", cbxtensoesatenproj.Text);
                pgbstatusproj.Value++;
                func.PesquisaPotInv(func.ModeloInversor,Banco);
                double resultado = double.Parse(func.PotenciaInv) * double.Parse(func.QuantidadeInversores) / 1000;
                this.orcamento.AcharESubstituir(wordApp, "<potinst>", resultado.ToString("0.00"));
                resultado = potenciagerada < resultado ? resultado = potenciagerada : resultado; 
                this.orcamento.AcharESubstituir(wordApp, "<potinstmenor>", resultado.ToString("0.00"));
                func.PesquisaModInv(func.ModeloInversor,Banco);
                func.PesquisaModMod(func.ModeloModulo,Banco);
                this.orcamento.AcharESubstituir(wordApp, "<qtdinv>", func.QuantidadeInversores);
                this.orcamento.AcharESubstituir(wordApp, "<qtdmod>", func.QuantidadeModulos);
                this.orcamento.AcharESubstituir(wordApp, "<fabmod>", func.MarcaMod);
                pgbstatusproj.Value++;
                this.orcamento.AcharESubstituir(wordApp, "<modmod>", func.ModeloModulo);
                this.orcamento.AcharESubstituir(wordApp, "<potmod>", potger1);
                this.orcamento.AcharESubstituir(wordApp, "<fabinv>", func.MarcaInversor);
                this.orcamento.AcharESubstituir(wordApp, "<modinv>", func.ModeloInversor);
                pgbstatusproj.Value++;
                this.orcamento.AcharESubstituir(wordApp, "<sumdimmod>", dimensao1);
                this.orcamento.AcharESubstituir(wordApp, "<qtdins>", txtqtdinstproj.Value);
                pgbstatusproj.Value++;

            }
            else
            {
                MessageBox.Show("Arquivo não encontrado!");
            }

            //Save as
            mywordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing);

            mywordDoc.Close();
            wordApp.Quit();
        }

        //Animação
        private void btninicio_Click(object sender, EventArgs e)
        {
            ClicaInicio();
            Limpacampos();
        }
        private void btnconfiguracoes_Click(object sender, EventArgs e)
        {
            ClicaConf();
            Limpacampos();
            PaineisPrincipais(pnlconfiguracao);
            btnperfil.BackgroundImage = Properties.Resources.FundoButton;
            btnconfgeral.BackgroundImage = null;
            pnlbconfiguracao.Visible = false;
            pnlbperfil.Visible = true;
        }
        private void btnorcamento_Click(object sender, EventArgs e)
        {
            ClicaOrcamento();
            Limpacampos();
        }
        private void btnequipamentos_Click(object sender, EventArgs e)
        {
            ClicaEquipamentos();
            Limpacampos();
            PaineisPrincipais(pnlequipamentos);
        }
        private void btnclientes_Click(object sender, EventArgs e)
        {
            ClicaClientes();
            Limpacampos();
        }
        private void btnprojeto_Click(object sender, EventArgs e)
        {
            ClicaProjeto();
            Limpacampos();
            PaineisPrincipais(pnlprojeto);
        }
        private void btnsair_Click(object sender, EventArgs e)
        {
            ClicaSair();
            Limpacampos();
            var resultado = MessageBox.Show("Tem certeza que deseja sair?", "Atenção", MessageBoxButtons.YesNo);
            if (resultado == DialogResult.Yes)
            {
                Application.Exit();
            }
        }
        private void btnmenu_Click(object sender, EventArgs e)
        {
            recolhe();
            verificacor();
        }
        private void btnmenurec_Click(object sender, EventArgs e)
        {
            recolhe();
            verificacor();
        }
        private void btniniciorec_Click(object sender, EventArgs e)
        {
            recolhe();
            ClicaInicio();
            Limpacampos();
        }
        private void btnclientesrec_Click(object sender, EventArgs e)
        {
            recolhe();
            ClicaClientes();
            Limpacampos();
        }
        private void btnequipamentosrec_Click(object sender, EventArgs e)
        {
            recolhe();
            ClicaEquipamentos();
            Limpacampos();
        }
        private void btnorcamentorec_Click(object sender, EventArgs e)
        {
            recolhe();
            ClicaOrcamento();
            Limpacampos();
        }
        private void btnprojetorec_Click(object sender, EventArgs e)
        {
            recolhe();
            ClicaProjeto();
            Limpacampos();
        }
        private void btnconfiguracoesrec_Click(object sender, EventArgs e)
        {
            recolhe();
            ClicaConf();
            Limpacampos();
        }
        private void btnsairrec_Click(object sender, EventArgs e)
        {
            recolhe();
            ClicaSair();
            Limpacampos();
        }
        private void btnnovoorcamento_Click(object sender, EventArgs e)
        {
            PaineisPrincipais(pnlorcamento0);
            Limpacampos();
        }

        //Transição de paineis
        private void btn1_Click(object sender, EventArgs e)
        {
            ConsumoUnidades();
            qtdorc = 1;
            Qtdorca();
        }
        private void btn2uc_Click(object sender, EventArgs e)
        {
            ConsumoUnidades();
            qtdorc = 2;
            Qtdorca();
        }
        private void btn3uc_Click(object sender, EventArgs e)
        {
            ConsumoUnidades();
            qtdorc = 3;
            Qtdorca();
        }
        private void btn4uc_Click(object sender, EventArgs e)
        {
            ConsumoUnidades();
            qtdorc = 4;
            Qtdorca();
        }
        private void btnpasso2_Click(object sender, EventArgs e)
        {
            CamposObrigatorios();
            ControlaStrings();
        }
        private void btnvolta1_Click(object sender, EventArgs e)
        {
            PaineisPrincipais(pnlorcamento0);

            qtdorc = 0;
            Qtdorca();
            Limpacampos();
        }
        private void button10_Click(object sender, EventArgs e)
        {
            if (!editando)
            {
                PaineisPrincipais(pnlorcamento1);

                Qtdorca();
            }
            else
            {
                qtdorc = 0;
                editando = false;
                PaineisPrincipais(pnlorcasalvos);
                gbxuc1.Visible = false;
                gbxuc1.Enabled = false;
                gbxuc2.Visible = false;
                gbxuc2.Enabled = false;
                gbxuc3.Visible = false;
                gbxuc3.Enabled = false;
                gbxuc4.Visible = false;
                gbxuc4.Enabled = false;

                btnrecolhe1.Visible = false;
                pnlpreencheuc.Visible = false;
                pnl2uc.Visible = false;
                pnl3uc.Visible = false;
                pnl4uc.Visible = false;
                pnlexp1.Visible = false;
                pnlexp2.Visible = false;
                pnlexp3.Visible = false;
                pnlexp4.Visible = false;
            }

        }
        private void button11_Click(object sender, EventArgs e)
        {
            if (Dimensionamento())
            {
                PaineisPrincipais(pnlorcamento3);
            }
            
        }
        private void button17_Click(object sender, EventArgs e)
        {
            PaineisPrincipais(pnlorcamento2);
        }
        private void btnvolta4_Click(object sender, EventArgs e)
        {
            chartgeracao.Series.Clear();
            chartretornofin.Series.Clear();
            PaineisPrincipais(pnlorcamento3);
        }
        private void button28_Click(object sender, EventArgs e)
        {
            PaineisPrincipais(pnlorcasalvos);
        }
        private void btnnovoorca_Click(object sender, EventArgs e)
        {
            PaineisPrincipais(pnlorcamento0);
        }

        //Transição de paineis da parte de consumo por Unidade Consumidora
        private void pictureBox5_Click(object sender, EventArgs e)
        {
            if (txtjan4.Text != "" && txtago4.Text != "" && txtfev4.Text != "" && txtmar4.Text != "" && txtabr4.Text != "" && txtmai4.Text != "" && txtjun4.Text != ""
                        && txtjul4.Text != "" && txtset4.Text != "" && txtout4.Text != "" && txtnov4.Text != "" && txtdez4.Text != "")
            {
                pnlexp4.BackColor = Color.FromArgb(45, 185, 59);
            }
            else
            {
                pnlexp4.BackColor = Color.FromArgb(255, 83, 19);
            }
            Preencheconsumo();
        }
        private void btnrecolhe3_Click(object sender, EventArgs e)
        {
            if (txtjan3.Text != "" && txtago3.Text != "" && txtfev3.Text != "" && txtmar3.Text != "" && txtabr3.Text != "" && txtmai3.Text != "" && txtjun3.Text != ""
                        && txtjul3.Text != "" && txtset3.Text != "" && txtout3.Text != "" && txtnov3.Text != "" && txtdez3.Text != "")
            {
                pnlexp3.BackColor = Color.FromArgb(45, 185, 59);
            }
            else
            {
                pnlexp3.BackColor = Color.FromArgb(255, 83, 19);
            }
            Preencheconsumo();
        }
        private void btnrecolhe2_Click(object sender, EventArgs e)
        {
            if (txtjan2.Text != "" && txtago2.Text != "" && txtfev2.Text != "" && txtmar2.Text != "" && txtabr2.Text != "" && txtmai2.Text != "" && txtjun2.Text != ""
                        && txtjul2.Text != "" && txtset2.Text != "" && txtout2.Text != "" && txtnov2.Text != "" && txtdez2.Text != "")
            {
                pnlexp2.BackColor = Color.FromArgb(45, 185, 59);
            }
            else
            {
                pnlexp2.BackColor = Color.FromArgb(255, 83, 19);
            }
            Preencheconsumo();
        }
        private void btnexp1_Click(object sender, EventArgs e)
        {
            pnlpreencheuc.Visible = false;
            pnl2uc.Visible = false;
            pnl3uc.Visible = false;
            pnl4uc.Visible = false;
        }
        private void btnexp2_Click(object sender, EventArgs e)
        {
            pnlpreencheuc.Visible = false;
            pnl2uc.Visible = true;
            pnl3uc.Visible = false;
            pnl4uc.Visible = false;
        }
        private void btnexp3_Click(object sender, EventArgs e)
        {
            pnlpreencheuc.Visible = false;
            pnl2uc.Visible = false;
            pnl3uc.Visible = true;
            pnl4uc.Visible = false;
        }
        private void btnexp4_Click(object sender, EventArgs e)
        {
            pnlpreencheuc.Visible = false;
            pnl2uc.Visible = false;
            pnl3uc.Visible = false;
            pnl4uc.Visible = true;
        }
        private void btnrecolhe1_Click(object sender, EventArgs e)
        {
            if (txtjan.Text != "" && txtago.Text != "" && txtfev.Text != "" && txtmar.Text != "" && txtabr.Text != "" && txtmai.Text != "" && txtjun.Text != ""
                        && txtjul.Text != "" && txtset.Text != "" && txtout.Text != "" && txtnov.Text != "" && txtdez.Text != "")
            {
                pnlexp1.BackColor = Color.FromArgb(45, 185, 59);
            }
            else
            {
                pnlexp1.BackColor = Color.FromArgb(255, 83, 19);
            }
                Preencheconsumo();
        }

        //Controle de objetos
        private void txttarifa_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((Char.IsLetter(e.KeyChar)) || (Char.IsWhiteSpace(e.KeyChar)))
                e.Handled = true;
        }
        private void txtcontato_Click(object sender, EventArgs e)
        {
            txtcontato.SelectionStart = 0;
        }
        private void txtcontato_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (txtcontato.Text.Length < 14)
                txtcontato.Mask = "(00) 0000-0000";
            else
            {
                txtcontato.Mask = "(00) 00000-0000";
                txtcontato.SelectionStart = txtcontato.TextLength - 1;
            }
        }
        private void txtcep_Leave(object sender, EventArgs e)
        {
            using (var ws = new WSCorreios.AtendeClienteClient())
            {
                try
                {
                    var endereco = ws.consultaCEP(txtcep.Text.Trim());
                    txtbairro.Text = endereco.bairro;
                    txtcidade.Text = endereco.cidade;
                    txtendereco.Text = endereco.end;
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.Message);
                }
            }
        }
        private void txtcep_Click(object sender, EventArgs e)
        {
            txtcep.SelectionStart = 0;
        }
        private void txtvalorsist_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((Char.IsLetter(e.KeyChar)) || (Char.IsWhiteSpace(e.KeyChar)) || (Char.IsSymbol(e.KeyChar)) || (Char.IsPunctuation(e.KeyChar)))
                e.Handled = true;
        }
        private void txtvalorequip_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((Char.IsLetter(e.KeyChar)) || (Char.IsWhiteSpace(e.KeyChar)) || (Char.IsSymbol(e.KeyChar)) || (Char.IsPunctuation(e.KeyChar)))
                e.Handled = true;
        }
        private void txtcustoinversor_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((Char.IsLetter(e.KeyChar)) || (Char.IsWhiteSpace(e.KeyChar)) || (Char.IsSymbol(e.KeyChar)) || (Char.IsPunctuation(e.KeyChar)))
                e.Handled = true;
        }
        private void txttarifa_Leave(object sender, EventArgs e)
        {
            if (txttarifa.Text != "")
            {
                string texte = string.Format("{0:0.00000000}", double.Parse(txttarifa.Text));
                txtkwh.Text = texte;
            }

        }
        private void txtjan_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((Char.IsLetter(e.KeyChar)) || (Char.IsWhiteSpace(e.KeyChar)) || (Char.IsSymbol(e.KeyChar)) || (Char.IsPunctuation(e.KeyChar)))
                e.Handled = true;
        }
        private void txtfev_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((Char.IsLetter(e.KeyChar)) || (Char.IsWhiteSpace(e.KeyChar)) || (Char.IsSymbol(e.KeyChar)) || (Char.IsPunctuation(e.KeyChar)))
                e.Handled = true;
        }
        private void txtmar_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((Char.IsLetter(e.KeyChar)) || (Char.IsWhiteSpace(e.KeyChar)) || (Char.IsSymbol(e.KeyChar)) || (Char.IsPunctuation(e.KeyChar)))
                e.Handled = true;
        }
        private void txtabr_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((Char.IsLetter(e.KeyChar)) || (Char.IsWhiteSpace(e.KeyChar)) || (Char.IsSymbol(e.KeyChar)) || (Char.IsPunctuation(e.KeyChar)))
                e.Handled = true;
        }
        private void txtmai_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((Char.IsLetter(e.KeyChar)) || (Char.IsWhiteSpace(e.KeyChar)) || (Char.IsSymbol(e.KeyChar)) || (Char.IsPunctuation(e.KeyChar)))
                e.Handled = true;
        }
        private void txtjun_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((Char.IsLetter(e.KeyChar)) || (Char.IsWhiteSpace(e.KeyChar)) || (Char.IsSymbol(e.KeyChar)) || (Char.IsPunctuation(e.KeyChar)))
                e.Handled = true;
        }
        private void txtjul_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((Char.IsLetter(e.KeyChar)) || (Char.IsWhiteSpace(e.KeyChar)) || (Char.IsSymbol(e.KeyChar)) || (Char.IsPunctuation(e.KeyChar)))
                e.Handled = true;
        }
        private void txtago_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((Char.IsLetter(e.KeyChar)) || (Char.IsWhiteSpace(e.KeyChar)) || (Char.IsSymbol(e.KeyChar)) || (Char.IsPunctuation(e.KeyChar)))
                e.Handled = true;
        }
        private void txtset_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((Char.IsLetter(e.KeyChar)) || (Char.IsWhiteSpace(e.KeyChar)) || (Char.IsSymbol(e.KeyChar)) || (Char.IsPunctuation(e.KeyChar)))
                e.Handled = true;
        }
        private void txtout_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((Char.IsLetter(e.KeyChar)) || (Char.IsWhiteSpace(e.KeyChar)) || (Char.IsSymbol(e.KeyChar)) || (Char.IsPunctuation(e.KeyChar)))
                e.Handled = true;
        }
        private void txtnov_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((Char.IsLetter(e.KeyChar)) || (Char.IsWhiteSpace(e.KeyChar)) || (Char.IsSymbol(e.KeyChar)) || (Char.IsPunctuation(e.KeyChar)))
                e.Handled = true;
        }
        private void txtdez_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((Char.IsLetter(e.KeyChar)) || (Char.IsWhiteSpace(e.KeyChar)) || (Char.IsSymbol(e.KeyChar)) || (Char.IsPunctuation(e.KeyChar)))
                e.Handled = true;
        }
        private void textBox1_Enter(object sender, EventArgs e)
        {
            if (txtbusca.Text == "O que você procura?")
            {
                txtbusca.Text = "";
            }
            txtbusca.ForeColor = Color.Black;
        }
        private void mtxtcpfcnpjcliente_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (mtxtcpfcnpjcliente.Text.Length < 14)
            {
                mtxtcpfcnpjcliente.Mask = "000.000.000-00";
            }   
            else
            {
                mtxtcpfcnpjcliente.Mask = "00.000.000/0000-00";
                mtxtcpfcnpjcliente.SelectionStart = mtxtcpfcnpjcliente.TextLength - (i-j);
                i--;
                if (j != 1)
                {
                    j++;
                }
            }
        }
        private void mtxtcepcliente_Leave(object sender, EventArgs e)
        {
            using (var ws = new WSCorreios.AtendeClienteClient())
            {
                try
                {
                    var endereco = ws.consultaCEP(mtxtcepcliente.Text.Trim());
                    txtbairrocliente.Text = endereco.bairro;
                    txtcidadecliente.Text = endereco.cidade;
                    txtenderecocliente.Text = endereco.end;
                    cbxufcliente.Text = endereco.uf;
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.Message);
                }
            }
        }
        private void mtxtcepcliente_Click(object sender, EventArgs e)
        {
            mtxtcepcliente.SelectionStart = 0;
        }
        private void mtxtcpfcnpjcliente_Click(object sender, EventArgs e)
        {
            mtxtcpfcnpjcliente.SelectionStart = 0;
            i = 4;
            j = 0;
        }
        private void mtxtcpfcnpjcliente_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Back)
            {
                mtxtcpfcnpjcliente.Mask = "000.000.000-00";
                i = 4;
                j = 0;
            }
        }
        private void txtbusca_Leave(object sender, EventArgs e)
        {
            if (txtbusca.Text == "")
            {
                txtbusca.Text = "O que você procura?";
                dgvorcamentos.DataSource = func.CarregaOrc(Banco);
            }
            
            txtbusca.ForeColor = Color.Silver;
        }
        private void txtbusca_TextChanged(object sender, EventArgs e)
        {
            string texto = txtbusca.Text;
            if (texto == "")
            {
                dgvorcamentos.DataSource = func.PesquisaOrc(txtbusca.Text, cbxfiltroorcamento.Text, Banco);
            }
        }
        private void txtprocuracliente_Enter(object sender, EventArgs e)
        {
            if (txtprocuracliente.Text == "O que você procura?")
            {
                txtprocuracliente.Text = "";
            }
            txtprocuracliente.ForeColor = Color.Black;
        }
        private void txtprocuracliente_Leave(object sender, EventArgs e)
        {
            if (txtprocuracliente.Text == "")
            {
                txtprocuracliente.Text = "O que você procura?";
                dgvclientes.DataSource = func.AtualizaClientes(Banco);
            }
            txtprocuracliente.ForeColor = Color.Silver;
        }
        private void cbxmarcamod_SelectionChangeCommitted(object sender, EventArgs e)
        {
            var Dados = func.MarcaModuloOrc(cbxmarcamod.Text, Banco);
            cbxmodpaineis.Enabled = true;
            cbxmodpaineis.DataSource = Dados;
            cbxmodpaineis.ValueMember = "Modelo";
            cbxmodpaineis.DisplayMember = "Modelo";
        }
        private void cbxmarcainv_SelectionChangeCommitted(object sender, EventArgs e)
        {
            var Dados = func.MarcaInv(cbxmarcainv.Text, Banco);
            cbxmodinv.Enabled = true;
            cbxmodinv.DataSource = Dados;
            cbxmodinv.ValueMember = "Modelo";
            cbxmodinv.DisplayMember = "Modelo";
        }
        private void btnfechasimulacao_Click(object sender, EventArgs e)
        {
            pnlsimulaorc.Visible = false;
            //Limpacampos();
        }
        private void btnconfirmasimulacao_Click(object sender, EventArgs e)
        {
            if (!duplo)
            {
                if (rbtnsim0.Checked)
                {
                    rbtn0.Checked = true;
                }
                if (rbtnsim5.Checked)
                {
                    rbtn5.Checked = true;
                }
                if (rbtnsim7.Checked)
                {
                    rbtn7.Checked = true;
                }
                if (rbtnsim10.Checked)
                {
                    rbtn10.Checked = true;
                }
                if (rbtnsim12.Checked)
                {
                    rbtn12.Checked = true;
                }
                if (rbtnsim15.Checked)
                {
                    rbtn15.Checked = true;
                }
                if (rbtnsim20.Checked)
                {
                    rbtn20.Checked = true;
                }
                if (rbtnsim25.Checked)
                {
                    rbtn25.Checked = true;
                }
                if (rbtnsim30.Checked)
                {
                    rbtn30.Checked = true;
                }
                if (rbtnsim35.Checked)
                {
                    rbtn35.Checked = true;
                }
                if (rbtnsim40.Checked)
                {
                    rbtn40.Checked = true;
                }
                if (chksimop1.Checked)
                {
                    func.SelecionaPainel(Banco, cbxsimmodmod1.Text);
                    cbxmarcamod.Text = func.MarcaMod;
                    cbxmodpaineis.Text = func.ModeloModulo;
                    func.SelecionaInversor(Banco, cbxsimmodinv1.Text);
                    cbxmarcainv.Text = func.MarcaInversor;
                    cbxmodinv.Text = func.ModeloInversor;
                    txtqtdinv.Text = txtsimqtdinv1.Value.ToString();
                    txtqtdpaineis.Text = txtsimqtdplaca1.Value.ToString();
                    txtvalorsist.Text = txtsimtot1.Text;
                    txtvalorequip.Text = txtsimcom1.Text;
                }
                else if (chksimop2.Checked)
                {
                    func.SelecionaPainel(Banco, cbxsimmodmod2.Text);
                    cbxmarcamod.Text = func.MarcaMod;
                    cbxmodpaineis.Text = func.ModeloModulo;
                    func.SelecionaInversor(Banco, cbxsimmodinv2.Text);
                    cbxmarcainv.Text = func.MarcaInversor;
                    cbxmodinv.Text = func.ModeloInversor;
                    txtqtdinv.Text = txtsimqtdinv2.Value.ToString();
                    txtqtdpaineis.Text = txtsimqtdplaca2.Value.ToString();
                    txtvalorsist.Text = txtsimtot2.Text;
                    txtvalorequip.Text = txtsimcom2.Text;
                }
                else if (chksimop3.Checked)
                {
                    func.SelecionaPainel(Banco, cbxsimmodmod3.Text);
                    cbxmarcamod.Text = func.MarcaMod;
                    cbxmodpaineis.Text = func.ModeloModulo;
                    func.SelecionaInversor(Banco, cbxsimmodinv3.Text);
                    cbxmarcainv.Text = func.MarcaInversor;
                    cbxmodinv.Text = func.ModeloInversor;
                    txtqtdinv.Text = txtsimqtdinv3.Value.ToString();
                    txtqtdpaineis.Text = txtsimqtdplaca3.Value.ToString();
                    txtvalorsist.Text = txtsimtot3.Text;
                    txtvalorequip.Text = txtsimcom3.Text;
                }
                else if (chksimop4.Checked)
                {
                    func.SelecionaPainel(Banco, cbxsimmodmod4.Text);
                    cbxmarcamod.Text = func.MarcaMod;
                    cbxmodpaineis.Text = func.ModeloModulo;
                    func.SelecionaInversor(Banco, cbxsimmodinv4.Text);
                    cbxmarcainv.Text = func.MarcaInversor;
                    cbxmodinv.Text = func.ModeloInversor;
                    txtqtdinv.Text = txtsimqtdinv4.Value.ToString();
                    txtqtdpaineis.Text = txtsimqtdplaca4.Value.ToString();
                    txtvalorsist.Text = txtsimtot4.Text;
                    txtvalorequip.Text = txtsimcom4.Text;
                }
                else if (chksimop5.Checked)
                {
                    func.SelecionaPainel(Banco, cbxsimmodmod5.Text);
                    cbxmarcamod.Text = func.MarcaMod;
                    cbxmodpaineis.Text = func.ModeloModulo;
                    func.SelecionaInversor(Banco, cbxsimmodinv5.Text);
                    cbxmarcainv.Text = func.MarcaInversor;
                    cbxmodinv.Text = func.ModeloInversor;
                    txtqtdinv.Text = txtsimqtdinv5.Value.ToString();
                    txtqtdpaineis.Text = txtsimqtdplaca5.Value.ToString();
                    txtvalorsist.Text = txtsimtot5.Text;
                    txtvalorequip.Text = txtsimcom5.Text;
                }
                pnlsimulaorc.Visible = false;
            }
            else
            {
                MessageBox.Show("Mais de um orçamento selecionado. Verifique sua simulação e tente novamente!");
            }
        }
        private void chksimop1_CheckedChanged(object sender, EventArgs e)
        {
            if(chksimop1.Checked == true)
            {
                if(chksimop2.Checked == false && chksimop3.Checked == false && chksimop4.Checked == false && chksimop5.Checked == false)
                {
                    pnlsim1.BackColor = Color.FromArgb(45, 185, 59);
                    duplo = false;
                }
                else
                {
                    pnlsim1.BackColor = Color.FromArgb(255, 83, 19);
                    duplo = true;
                }
            }
            else
            {
                pnlsim1.BackColor = Color.White;
            }
        }
        private void chksimop2_CheckedChanged(object sender, EventArgs e)
        {
            if (chksimop2.Checked == true)
            {
                if (chksimop1.Checked == false && chksimop3.Checked == false && chksimop4.Checked == false && chksimop5.Checked == false)
                {
                    pnlsim2.BackColor = Color.FromArgb(45, 185, 59);
                    duplo = false;
                }
                else
                {
                    pnlsim2.BackColor = Color.FromArgb(255, 83, 19);
                    duplo = true;
                }
            }
            else
            {
                pnlsim2.BackColor = Color.White;
            }
        }
        private void chksimop3_CheckedChanged(object sender, EventArgs e)
        {
            if (chksimop3.Checked == true)
            {
                if (chksimop2.Checked == false && chksimop1.Checked == false && chksimop4.Checked == false && chksimop5.Checked == false)
                {
                    pnlsim3.BackColor = Color.FromArgb(45, 185, 59);
                    duplo = false;
                }
                else
                {
                    pnlsim3.BackColor = Color.FromArgb(255, 83, 19);
                    duplo = true;
                }
            }
            else
            {
                pnlsim3.BackColor = Color.White;
            }
        }
        private void chksimop4_CheckedChanged(object sender, EventArgs e)
        {
            if (chksimop4.Checked == true)
            {
                if (chksimop2.Checked == false && chksimop3.Checked == false && chksimop1.Checked == false && chksimop5.Checked == false)
                {
                    pnlsim4.BackColor = Color.FromArgb(45, 185, 59);
                    duplo = false;
                }
                else
                {
                    pnlsim4.BackColor = Color.FromArgb(255, 83, 19);
                    duplo = true;
                }
            }
            else
            {
                pnlsim4.BackColor = Color.White;
            }
        }
        private void chksimop5_CheckedChanged(object sender, EventArgs e)
        {
            if (chksimop5.Checked == true)
            {
                if (chksimop2.Checked == false && chksimop3.Checked == false && chksimop4.Checked == false && chksimop1.Checked == false)
                {
                    pnlsim5.BackColor = Color.FromArgb(45, 185, 59);
                    duplo = false;
                }
                else
                {
                    pnlsim5.BackColor = Color.FromArgb(255, 83, 19);
                    duplo = true;
                }
            }
            else
            {
                pnlsim5.BackColor = Color.White;
            }
        }
        private void rbtnsim0_CheckedChanged(object sender, EventArgs e)
        {
            percas = 1;
            if (txtsimpotger1.Text != "" && cbxsimmodmod1.Text != "" )
            {
                func.PesquisaPotMod(cbxsimmodmod1.Text, Banco);
                txtsimpotger1.Text = (double.Parse(txtsimqtdplaca1.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger1.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca1.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger2.Text != "" && cbxsimmodmod2.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod2.Text, Banco);
                txtsimpotger2.Text = (double.Parse(txtsimqtdplaca2.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger2.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca2.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger3.Text != "" && cbxsimmodmod3.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod2.Text, Banco);
                txtsimpotger3.Text = (double.Parse(txtsimqtdplaca3.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger3.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca3.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger4.Text != "" && cbxsimmodmod4.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod4.Text, Banco);
                txtsimpotger4.Text = (double.Parse(txtsimqtdplaca4.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger4.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca4.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger5.Text != "" && cbxsimmodmod5.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod5.Text, Banco);
                txtsimpotger5.Text = (double.Parse(txtsimqtdplaca5.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger5.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca5.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
        }
        private void rbtnsim5_CheckedChanged(object sender, EventArgs e)
        {
            percas = 0.95;
            if (txtsimpotger1.Text != "" && cbxsimmodmod1.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod1.Text, Banco);
                txtsimpotger1.Text = (double.Parse(txtsimqtdplaca1.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger1.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca1.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger2.Text != "" && cbxsimmodmod2.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod2.Text, Banco);
                txtsimpotger2.Text = (double.Parse(txtsimqtdplaca2.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger2.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca2.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger3.Text != "" && cbxsimmodmod3.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod2.Text, Banco);
                txtsimpotger3.Text = (double.Parse(txtsimqtdplaca3.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger3.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca3.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger4.Text != "" && cbxsimmodmod4.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod4.Text, Banco);
                txtsimpotger4.Text = (double.Parse(txtsimqtdplaca4.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger4.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca4.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger5.Text != "" && cbxsimmodmod5.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod5.Text, Banco);
                txtsimpotger5.Text = (double.Parse(txtsimqtdplaca5.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger5.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca5.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
        }
        private void rbtnsim7_CheckedChanged(object sender, EventArgs e)
        {
            percas = 0.93;
            if (txtsimpotger1.Text != "" && cbxsimmodmod1.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod1.Text, Banco);
                txtsimpotger1.Text = (double.Parse(txtsimqtdplaca1.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger1.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca1.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger2.Text != "" && cbxsimmodmod2.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod2.Text, Banco);
                txtsimpotger2.Text = (double.Parse(txtsimqtdplaca2.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger2.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca2.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger3.Text != "" && cbxsimmodmod3.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod2.Text, Banco);
                txtsimpotger3.Text = (double.Parse(txtsimqtdplaca3.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger3.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca3.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger4.Text != "" && cbxsimmodmod4.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod4.Text, Banco);
                txtsimpotger4.Text = (double.Parse(txtsimqtdplaca4.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger4.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca4.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger5.Text != "" && cbxsimmodmod5.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod5.Text, Banco);
                txtsimpotger5.Text = (double.Parse(txtsimqtdplaca5.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger5.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca5.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
        }
        private void rbtnsim10_CheckedChanged(object sender, EventArgs e)
        {
            percas = 0.9;
            if (txtsimpotger1.Text != "" && cbxsimmodmod1.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod1.Text, Banco);
                txtsimpotger1.Text = (double.Parse(txtsimqtdplaca1.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger1.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca1.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger2.Text != "" && cbxsimmodmod2.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod2.Text, Banco);
                txtsimpotger2.Text = (double.Parse(txtsimqtdplaca2.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger2.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca2.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger3.Text != "" && cbxsimmodmod3.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod2.Text, Banco);
                txtsimpotger3.Text = (double.Parse(txtsimqtdplaca3.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger3.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca3.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger4.Text != "" && cbxsimmodmod4.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod4.Text, Banco);
                txtsimpotger4.Text = (double.Parse(txtsimqtdplaca4.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger4.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca4.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger5.Text != "" && cbxsimmodmod5.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod5.Text, Banco);
                txtsimpotger5.Text = (double.Parse(txtsimqtdplaca5.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger5.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca5.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
        }
        private void rbtnsim12_CheckedChanged(object sender, EventArgs e)
        {
            percas = 0.88;
            if (txtsimpotger1.Text != "" && cbxsimmodmod1.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod1.Text, Banco);
                txtsimpotger1.Text = (double.Parse(txtsimqtdplaca1.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger1.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca1.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger2.Text != "" && cbxsimmodmod2.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod2.Text, Banco);
                txtsimpotger2.Text = (double.Parse(txtsimqtdplaca2.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger2.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca2.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger3.Text != "" && cbxsimmodmod3.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod2.Text, Banco);
                txtsimpotger3.Text = (double.Parse(txtsimqtdplaca3.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger3.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca3.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger4.Text != "" && cbxsimmodmod4.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod4.Text, Banco);
                txtsimpotger4.Text = (double.Parse(txtsimqtdplaca4.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger4.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca4.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger5.Text != "" && cbxsimmodmod5.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod5.Text, Banco);
                txtsimpotger5.Text = (double.Parse(txtsimqtdplaca5.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger5.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca5.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
        }
        private void rbtnsim15_CheckedChanged(object sender, EventArgs e)
        {
            percas = 0.85;
            if (txtsimpotger1.Text != "" && cbxsimmodmod1.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod1.Text, Banco);
                txtsimpotger1.Text = (double.Parse(txtsimqtdplaca1.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger1.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca1.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger2.Text != "" && cbxsimmodmod2.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod2.Text, Banco);
                txtsimpotger2.Text = (double.Parse(txtsimqtdplaca2.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger2.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca2.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger3.Text != "" && cbxsimmodmod3.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod2.Text, Banco);
                txtsimpotger3.Text = (double.Parse(txtsimqtdplaca3.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger3.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca3.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger4.Text != "" && cbxsimmodmod4.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod4.Text, Banco);
                txtsimpotger4.Text = (double.Parse(txtsimqtdplaca4.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger4.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca4.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger5.Text != "" && cbxsimmodmod5.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod5.Text, Banco);
                txtsimpotger5.Text = (double.Parse(txtsimqtdplaca5.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger5.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca5.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
        }
        private void rbtnsim20_CheckedChanged(object sender, EventArgs e)
        {
            percas = 0.8;
            if (txtsimpotger1.Text != "" && cbxsimmodmod1.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod1.Text, Banco);
                txtsimpotger1.Text = (double.Parse(txtsimqtdplaca1.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger1.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca1.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger2.Text != "" && cbxsimmodmod2.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod2.Text, Banco);
                txtsimpotger2.Text = (double.Parse(txtsimqtdplaca2.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger2.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca2.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger3.Text != "" && cbxsimmodmod3.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod2.Text, Banco);
                txtsimpotger3.Text = (double.Parse(txtsimqtdplaca3.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger3.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca3.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger4.Text != "" && cbxsimmodmod4.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod4.Text, Banco);
                txtsimpotger4.Text = (double.Parse(txtsimqtdplaca4.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger4.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca4.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger5.Text != "" && cbxsimmodmod5.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod5.Text, Banco);
                txtsimpotger5.Text = (double.Parse(txtsimqtdplaca5.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger5.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca5.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
        }
        private void rbtnsim25_CheckedChanged(object sender, EventArgs e)
        {
            percas = 0.75;
            if (txtsimpotger1.Text != "" && cbxsimmodmod1.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod1.Text, Banco);
                txtsimpotger1.Text = (double.Parse(txtsimqtdplaca1.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger1.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca1.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger2.Text != "" && cbxsimmodmod2.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod2.Text, Banco);
                txtsimpotger2.Text = (double.Parse(txtsimqtdplaca2.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger2.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca2.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger3.Text != "" && cbxsimmodmod3.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod2.Text, Banco);
                txtsimpotger3.Text = (double.Parse(txtsimqtdplaca3.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger3.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca3.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger4.Text != "" && cbxsimmodmod4.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod4.Text, Banco);
                txtsimpotger4.Text = (double.Parse(txtsimqtdplaca4.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger4.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca4.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger5.Text != "" && cbxsimmodmod5.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod5.Text, Banco);
                txtsimpotger5.Text = (double.Parse(txtsimqtdplaca5.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger5.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca5.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
        }
        private void rbtnsim30_CheckedChanged(object sender, EventArgs e)
        {
            percas = 1.1;
            if (txtsimpotger1.Text != "" && cbxsimmodmod1.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod1.Text, Banco);
                txtsimpotger1.Text = (double.Parse(txtsimqtdplaca1.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger1.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca1.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger2.Text != "" && cbxsimmodmod2.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod2.Text, Banco);
                txtsimpotger2.Text = (double.Parse(txtsimqtdplaca2.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger2.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca2.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger3.Text != "" && cbxsimmodmod3.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod2.Text, Banco);
                txtsimpotger3.Text = (double.Parse(txtsimqtdplaca3.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger3.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca3.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger4.Text != "" && cbxsimmodmod4.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod4.Text, Banco);
                txtsimpotger4.Text = (double.Parse(txtsimqtdplaca4.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger4.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca4.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger5.Text != "" && cbxsimmodmod5.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod5.Text, Banco);
                txtsimpotger5.Text = (double.Parse(txtsimqtdplaca5.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger5.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca5.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
        }
        private void rbtnsim35_CheckedChanged(object sender, EventArgs e)
        {
            percas = 1.075;
            if (txtsimpotger1.Text != "" && cbxsimmodmod1.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod1.Text, Banco);
                txtsimpotger1.Text = (double.Parse(txtsimqtdplaca1.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger1.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca1.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger2.Text != "" && cbxsimmodmod2.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod2.Text, Banco);
                txtsimpotger2.Text = (double.Parse(txtsimqtdplaca2.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger2.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca2.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger3.Text != "" && cbxsimmodmod3.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod2.Text, Banco);
                txtsimpotger3.Text = (double.Parse(txtsimqtdplaca3.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger3.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca3.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger4.Text != "" && cbxsimmodmod4.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod4.Text, Banco);
                txtsimpotger4.Text = (double.Parse(txtsimqtdplaca4.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger4.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca4.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger5.Text != "" && cbxsimmodmod5.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod5.Text, Banco);
                txtsimpotger5.Text = (double.Parse(txtsimqtdplaca5.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger5.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca5.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
        }
        private void rbtnsim40_CheckedChanged(object sender, EventArgs e)
        {
            percas = 1.05;
            if (txtsimpotger1.Text != "" && cbxsimmodmod1.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod1.Text, Banco);
                txtsimpotger1.Text = (double.Parse(txtsimqtdplaca1.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger1.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca1.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger2.Text != "" && cbxsimmodmod2.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod2.Text, Banco);
                txtsimpotger2.Text = (double.Parse(txtsimqtdplaca2.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger2.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca2.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger3.Text != "" && cbxsimmodmod3.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod2.Text, Banco);
                txtsimpotger3.Text = (double.Parse(txtsimqtdplaca3.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger3.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca3.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger4.Text != "" && cbxsimmodmod4.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod4.Text, Banco);
                txtsimpotger4.Text = (double.Parse(txtsimqtdplaca4.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger4.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca4.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
            if (txtsimpotger5.Text != "" && cbxsimmodmod5.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod5.Text, Banco);
                txtsimpotger5.Text = (double.Parse(txtsimqtdplaca5.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger5.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca5.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
        }
        private void txtsimqtdplaca1_ValueChanged(object sender, EventArgs e)
        {
            if (cbxsimmodmod1.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod1.Text, Banco);
                txtsimpotger1.Text = (double.Parse(txtsimqtdplaca1.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger1.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca1.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
        }
        private void txtsimqtdplaca2_ValueChanged(object sender, EventArgs e)
        {
            if (cbxsimmodmod2.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod2.Text, Banco);
                txtsimpotger2.Text = (double.Parse(txtsimqtdplaca2.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger2.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca2.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
        }
        private void txtsimqtdplaca3_ValueChanged(object sender, EventArgs e)
        {
            if (cbxsimmodmod3.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod3.Text, Banco);
                txtsimpotger3.Text = (double.Parse(txtsimqtdplaca3.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger3.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca3.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
        }
        private void txtsimqtdplaca4_ValueChanged(object sender, EventArgs e)
        {
            if (cbxsimmodmod4.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod4.Text, Banco);
                txtsimpotger4.Text = (double.Parse(txtsimqtdplaca4.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger4.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca4.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
        }
        private void txtsimqtdplaca5_ValueChanged(object sender, EventArgs e)
        {
            if (cbxsimmodmod5.Text != "")
            {
                func.PesquisaPotMod(cbxsimmodmod5.Text, Banco);
                txtsimpotger5.Text = (double.Parse(txtsimqtdplaca5.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger5.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca5.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
        }
        private void cbxsimmodmod1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbxsimmodmod1.Text != "" && cbxsimmodmod1.Text != "System.Data.DataRowView")
            {
                func.PesquisaPotMod(cbxsimmodmod1.Text, Banco);
                txtsimpotger1.Text = (double.Parse(txtsimqtdplaca1.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger1.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca1.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
        }
        private void cbxsimmodmod2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbxsimmodmod2.Text != "" && cbxsimmodmod2.Text != "System.Data.DataRowView")
            {
                func.PesquisaPotMod(cbxsimmodmod2.Text, Banco);
                txtsimpotger2.Text = (double.Parse(txtsimqtdplaca2.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger2.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca2.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
        }
        private void cbxsimmodmod3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbxsimmodmod3.Text != "" && cbxsimmodmod3.Text != "System.Data.DataRowView")
            {
                func.PesquisaPotMod(cbxsimmodmod3.Text, Banco);
                txtsimpotger3.Text = (double.Parse(txtsimqtdplaca3.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger3.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca3.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
        }
        private void cbxsimmodmod4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbxsimmodmod4.Text != "" && cbxsimmodmod4.Text != "System.Data.DataRowView")
            {
                func.PesquisaPotMod(cbxsimmodmod4.Text, Banco);
                txtsimpotger4.Text = (double.Parse(txtsimqtdplaca4.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger4.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca4.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
        }
        private void cbxsimmodmod5_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbxsimmodmod5.Text != "" && cbxsimmodmod5.Text != "System.Data.DataRowView")
            {
                func.PesquisaPotMod(cbxsimmodmod5.Text, Banco);
                txtsimpotger5.Text = (double.Parse(txtsimqtdplaca5.Value.ToString()) * double.Parse(func.PotenciaMod)).ToString();
                txtsimger5.Text = ((((double.Parse(func.PotenciaMod) * double.Parse(txtsimqtdplaca5.Value.ToString())) * 0.83) / 1000) * 30 * 4.87 * percas).ToString();
            }
        }
        private void cbxsimmodinv1_SelectedIndexChanged(object sender, EventArgs e)
        {
            /*func.SelecionaInversor(Banco, cbxsimmodinv1.Text);
            if(int.Parse(func.PotenciaInv) <= 2000)
            {
                txteqcomissao.Text = "9";
                txteqoutros.Text = "9";
            }*/
        }
        private void cbxsimmodinv2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        private void cbxsimmodinv3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        private void cbxsimmodinv4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        private void cbxsimmodinv5_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        private void btnsimulaorc_Click(object sender, EventArgs e)
        {
            pnlsimulaorc.Visible = true;
        }

        private void btn5ucoumais_Click(object sender, EventArgs e)
        {

        }

        //Tela do Datagrid de orçamento
        private void pictureBox4_Click(object sender, EventArgs e)
        {
            dgvorcamentos.DataSource = func.PesquisaOrc(txtbusca.Text, cbxfiltroorcamento.Text, Banco);
        }
        private void txtbusca_KeyUp(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                dgvorcamentos.DataSource = func.PesquisaOrc(txtbusca.Text, cbxfiltroorcamento.Text, Banco);
            }
        }
        private void btneditaorca_Click(object sender, EventArgs e)
        {
            Limpacampos();
            CarregaCombobox();
            qtdorc = 1;
            Qtdorca();
            editando = true;
            EditaOrcamento();
            if (Dimensionamento())
            {
                PaineisPrincipais(pnlorcamento3);
            }
            
        }

        //Ultima tela orçamento
        private void btnpasso4_Click(object sender, EventArgs e)
        {
            if (!editando)
            {
                Calculos();
                Graficos();
                btnsalva.Text = "Salvar";
                PaineisPrincipais(pnlorcamento4);
            }
            else
            {
                Calculos();
                Graficos();
                btnsalva.Text = "Atualizar";
                PaineisPrincipais(pnlorcamento4);
            }

        }
        private void btnsalva_Click(object sender, EventArgs e)
        {
            if (!editando)
            {
                if (txtnome.Text == string.Empty || txtkwh.Text == string.Empty || txtvalorequip.Text == string.Empty || txtqtdpaineis.Text == string.Empty ||
                    txtqtdinv.Text == string.Empty || cbxmodinv.Text == string.Empty || cbxmodpaineis.Text == string.Empty)
                {
                    MessageBox.Show("Preencha todos os campos obrigatórios", "Atenção");
                }
                else
                {
                    string pasta = @"C:\Centraliza\Orçamentos";
                    if (!File.Exists(pasta + "\\Orçamento " + txtnome.Text + ".docx"))
                    {
                        CreateWordDoc(@"C:\Centraliza\Centraliza\temp-1.docx", @"C:\Centraliza\Orçamentos\Orçamento " + txtnome.Text + ".docx");
                    }
                    pgbstatusorca.Value++;
                    if (!File.Exists(pasta + "\\Retorno Financeiro " + txtnome.Text + ".xlsx"))
                    {
                        GerarPlanilhaFin();
                    }
                    pgbstatusorca.Value++;
                    if (!File.Exists(pasta + "\\Geração " + txtnome.Text + ".xlsx"))
                    {
                        GerarPlanilha();
                    }
                    pgbstatusorca.Value++;
                    chartgeracao.Series.Clear();
                    chartretornofin.Series.Clear();
                    Process.Start("explorer.exe", pasta);
                    SomaConsumo();
                    Salvaorcamento();
                    Limpacampos();

                    PaineisPrincipais(pnlorcamento5);
                }
            }
            else
            {
                if (txtnome.Text == string.Empty || txtkwh.Text == string.Empty || txtvalorequip.Text == string.Empty || txtqtdpaineis.Text == string.Empty ||
                    txtqtdinv.Text == string.Empty || cbxmodinv.Text == string.Empty || cbxmodpaineis.Text == string.Empty)
                {
                    MessageBox.Show("Preencha todos os campos obrigatórios", "Atenção");
                }
                else
                {
                    string pasta = @"C:\Centraliza\Orçamentos";
                    if (!File.Exists(pasta + "\\Orçamento " + txtnome.Text + ".docx"))
                    {
                        CreateWordDoc(@"C:\Centraliza\Centraliza\temp-1.docx", @"C:\Centraliza\Orçamentos\Orçamento " + txtnome.Text + ".docx");
                    }
                    pgbstatusorca.Value++;
                    if (!File.Exists(pasta + "\\Retorno Financeiro " + txtnome.Text + ".xlsx"))
                    {
                        GerarPlanilhaFin();
                    }
                    pgbstatusorca.Value++;
                    if (!File.Exists(pasta + "\\Geração " + txtnome.Text + ".xlsx"))
                    {
                        GerarPlanilha();
                    }
                    pgbstatusorca.Value++;
                    chartgeracao.Series.Clear();
                    chartretornofin.Series.Clear();
                    Process.Start("explorer.exe", pasta);
                    //SomaConsumo();
                    Salvaorcamento();
                    Limpacampos();
                    pgbstatusorca.Value = 0;

                    PaineisPrincipais(pnlorcamento5);
                }
            }
        }
        private void btnabrepasta_Click(object sender, EventArgs e)
        {
            if (txtnome.Text == string.Empty || txtkwh.Text == string.Empty || txtvalorequip.Text == string.Empty || txtqtdpaineis.Text == string.Empty ||
                txtqtdinv.Text == string.Empty || cbxmodinv.Text == string.Empty || cbxmodpaineis.Text == string.Empty)
            {
                MessageBox.Show("Preencha todos os campos obrigatórios", "Atenção");
            }
            else
            {
                string pasta = @"C:\Centraliza\Orçamentos";
                if (!File.Exists(pasta + "\\Orçamento " + txtnome.Text + ".docx"))
                {
                    CreateWordDoc(@"C:\Centraliza\Centraliza\temp-1.docx", @"C:\Centraliza\Orçamentos\Orçamento " + txtnome.Text + ".docx");
                }
                pgbstatusorca.Value++;
                if (!File.Exists(pasta + "\\Retorno Financeiro " + txtnome.Text + ".xlsx"))
                {
                    GerarPlanilhaFin();
                }
                pgbstatusorca.Value++;
                if (!File.Exists(pasta + "\\Geração " + txtnome.Text + ".xlsx"))
                {
                    GerarPlanilha();
                }
                pgbstatusorca.Value++;
                Process.Start("explorer.exe", pasta);
                pgbstatusorca.Value = 0;
            }

        }
        private void btnexcluir_Click(object sender, EventArgs e)
        {
            if (!editando)
            {
                var resultado = MessageBox.Show("Tem certeza que deseja excluir o orçamento?", "Atenção", MessageBoxButtons.YesNo);
                if (resultado == DialogResult.Yes)
                {
                    Limpacampos();
                    pgbstatusorca.Value = 23;
                    CarregaCombobox();
                    PaineisPrincipais(pnlorcamento0);
                    pgbstatusorca.Value = 0;
                }
            }
            else
            {
                var resultado = MessageBox.Show("Tem certeza que deseja excluir o orçamento?", "Atenção", MessageBoxButtons.YesNo);
                if (resultado == DialogResult.Yes)
                {
                    func.Excl();
                    Limpacampos();
                    pgbstatusorca.Value = 23;
                    CarregaCombobox();
                    PaineisPrincipais(pnlorcamento0);
                    pgbstatusorca.Value = 0;
                }
            }
        }

        //Configurações e usuário
        private void pbxUsuario_Click(object sender, EventArgs e)
        {
            ClicaConf();
            PaineisPrincipais(pnlconfiguracao);
            btnperfil.BackgroundImage = Properties.Resources.FundoButton;
            btnconfgeral.BackgroundImage = null;
            pnlbconfiguracao.Visible = false;
            pnlbperfil.Visible = true;
        }
        private void btnperfil_Click(object sender, EventArgs e)
        {
            btnperfil.BackgroundImage = Properties.Resources.FundoButton;
            btnconfgeral.BackgroundImage = null;
            pnlbconfiguracao.Visible = false;
            pnlbperfil.Visible = true;
        }
        private void btnconfgeral_Click(object sender, EventArgs e)
        {
            btnconfgeral.BackgroundImage = Properties.Resources.FundoButton;
            btnperfil.BackgroundImage = null;
            pnlbconfiguracao.Visible = true;
            pnlbperfil.Visible = false;
        }
        private void btnsalvarconf_Click(object sender, EventArgs e)
        {
            var resultado = MessageBox.Show("Tem certeza que deseja alterar o banco de dados? A aplicação será reiniciada", "Atenção", MessageBoxButtons.YesNo);
            if (resultado == DialogResult.Yes)
            {
                if (rbtnlocal.Checked)
                {
                    func.AlterarBanco("local");
                    Application.Restart();
                }
                else if (rbtnmysql.Checked)
                {
                    
                    func.AlterarBanco("mysql");
                    Application.Restart();
                }
            }
            
        }
        private void pbxuploadfoto_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog dialog = new OpenFileDialog();
                //dialog.Filter = "jpg files(*.jpg)|*.jpg|png files(*.png)|*.png|All Files(*.*)|*.*";
                dialog.Filter = "Foto (*.jpg, *.jpeg, *.bmp, *.png, *.tif, *.tiff)|*.jpg; *.jpeg; *.bmp; *.png; *.tif; *.tiff|All Files(*.*)|*.*";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    imageLocation = dialog.FileName;
                    foto = dialog.FileName;
                    pbxuploadfoto.ImageLocation = imageLocation;
                    pbxUsuario.ImageLocation = imageLocation;
                    PictureBoxRedondo();
                }
            }
            catch(Exception)
            {
                MessageBox.Show("Erro", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void btnsalvaperfil_Click(object sender, EventArgs e)
        {
            func.Nome = txtcnomecompleto.Text;
            func.Senha = txtcsenha.Text;
            func.Login = txtcusuario.Text;
            func.email = txtcemail.Text;
            func.FotoUsuario = ConverteBase64(foto);

            try
            {
                //func.InserirCredenciais(Banco);
                func.AtualizaUsario(Banco, label56.Text);
                MessageBox.Show("Dados atualizados com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao salvar dados do usuário. " + ex.Message, "Erro");
            }

        }

        //Painel de Clientes
        private void pbxbuscaclientes_Click(object sender, EventArgs e)
        {
            dgvclientes.DataSource = func.ClienteNome(txtprocuracliente.Text, cbxfiltrocliente.Text, Banco);
        }
        private void txtprocuracliente_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                dgvclientes.DataSource = func.ClienteNome(txtprocuracliente.Text, cbxfiltrocliente.Text, Banco);
            }
        }
        private void btnvisualizarcliente_Click(object sender, EventArgs e)
        {
            if (dgvclientes.Rows.Count == 0)
            {
                MessageBox.Show("Nenhum cliente cadastrado");
            }
            else
            {
                func.id = dgvclientes.CurrentRow.Cells[0].Value.ToString();
                func.SelecionaCliente(Banco);
                lblnomecliente.Text = func.Nome;
                lblidentificacaocliente.Text = func.Identificacao;
                lblcpfcliente.Text = func.CPF;
                lblenderecocliente.Text = func.Endereco + ", " + func.Numero + " " + func.Complemento;
                lblcepcliente.Text = func.CEP;
                lblcidadeufcliente.Text = func.Cidade + " - " + func.UF;
                lbltelefonecliente.Text = func.Telefone;
                lblcelularcliente.Text = func.Celular;
                lblemailcliente.Text = func.email;
                string q;
                if (func.MarcaInversor == "AP System" || func.ModeloInversor == "Reno560" || func.ModeloInversor == "Reno560-LV")
                {
                    if (Int16.Parse(func.QuantidadeInversores) > 1)
                    {
                        q = " Micro Inversores ";
                    }
                    else
                    {
                        q = " Micro Inversor ";
                    }
                }
                else
                {
                    if (Int16.Parse(func.QuantidadeInversores) > 1)
                    {
                        q = " Inversores ";
                    }
                    else
                    {
                        q = " Inversor ";
                    }
                }
                func.PesquisaPotInv(func.ModeloInversor, Banco);
                lblqtdinvcliente.Text = func.QuantidadeInversores + q + func.MarcaInversor + " " + func.ModeloInversor + "  " + (double.Parse(func.PotenciaInv) / 1000).ToString("0.00") + "KW";
                func.PesquisaPotMod(func.ModeloModulo, Banco);
                lblqtdmodcliente.Text = func.QuantidadeModulos + " Painéis " + func.MarcaMod + " " + func.ModeloModulo + "  " + (((double.Parse(func.QuantidadeModulos)) * (double.Parse(func.PotenciaMod)))/1000).ToString("0.00") + "KW";
                pnlvisualizacli.Visible = true;
            }
        }
        private void button11_Click_1(object sender, EventArgs e)
        {
            if (pnlvisualizacli.Visible == true)
            {
                pnlvisualizacli.Visible = false;
            }
            btncadastraouatualiza.Text = "Cadastrar";
            pnladicionacliente.Visible = true;
        }
        private void btncadastraouatualiza_Click(object sender, EventArgs e)
        {
            if (btncadastraouatualiza.Text == "Cadastrar")
            {
                if (txtnomecliente.Text == "" && mtxtcpfcnpjcliente.Text == "" && txtenderecocliente.Text == "" && txtnumerocliente.Text == "" && txtbairrocliente.Text == "" && mtxtcepcliente.Text == "_____-__" &&
                    txtcidadecliente.Text == "" && cbxufcliente.Text == "" && (mtxtcelularcliente.Text == "(__)_____-____" || mtxttelefone.Text == "(__)____-____" || txtemailcliente.Text == "") &&
                    txtqtdinvcliente.Value <= 0 && txtqtdmodmodcliente.Value <= 0 && cbxmarcainvcliente.Text == "" && cbxmarcamoccliente.Text == "" && cbxmodeloinvcliente.Text == "" && cbxmodelomodcliente.Text == "" &&
                    txtconsmedcliente.Text == "" && txtidentificacaocliente.Text == "")
                {
                    MessageBox.Show("Favor preencher todos os campos!");
                }
                else
                {
                    func.Nome = txtnomecliente.Text;
                    func.CPF = mtxtcpfcnpjcliente.Text;
                    func.Endereco = txtenderecocliente.Text;
                    func.Numero = txtnumerocliente.Text;
                    func.Complemento = txtcomplementocliente.Text;
                    func.Bairro = txtbairrocliente.Text;
                    func.CEP = mtxtcepcliente.Text;
                    func.Cidade = txtcidadecliente.Text;
                    func.UF = cbxufcliente.Text;
                    func.email = txtemailcliente.Text;
                    func.Telefone = mtxttelefone.Text;
                    func.Celular = mtxtcelularcliente.Text;
                    func.QuantidadeInversores = txtqtdinvcliente.Value.ToString();
                    func.QuantidadeModulos = txtqtdmodmodcliente.Value.ToString();
                    func.MarcaInversor = cbxmarcainvcliente.Text;
                    func.MarcaMod = cbxmarcamoccliente.Text;
                    func.ModeloInversor = cbxmodeloinvcliente.Text;
                    func.ModeloModulo = cbxmodelomodcliente.Text;
                    func.MediaConsumo = txtconsmedcliente.Text;
                    func.Identificacao = txtidentificacaocliente.Text;
                    func.InserirCliente(Banco);
                    MessageBox.Show("Cliente Cadastrado com Sucesso!");
                    dgvencontraclienteprojeto.DataSource = func.AtualizaClientes(Banco);
                    txtnomecliente.Text = string.Empty;
                    mtxtcpfcnpjcliente.Text = string.Empty;
                    txtenderecocliente.Text = string.Empty;
                    txtnumerocliente.Text = string.Empty;
                    txtcomplementocliente.Text = string.Empty;
                    txtbairrocliente.Text = string.Empty;
                    mtxtcepcliente.Text = string.Empty;
                    txtcidadecliente.Text = string.Empty;
                    cbxufcliente.Text = string.Empty;
                    txtemailcliente.Text = string.Empty;
                    mtxttelefone.Text = string.Empty;
                    mtxtcelularcliente.Text = string.Empty;
                    txtqtdinvcliente.Value = 1;
                    txtqtdmodmodcliente.Value = 1;
                    txtconsmedcliente.Text = string.Empty;
                    txtidentificacaocliente.Text = string.Empty;

                    pnladicionacliente.Visible = false;
                }

            }
            else
            {
                if (txtnomecliente.Text == "" && mtxtcpfcnpjcliente.Text == "" && txtenderecocliente.Text == "" && txtnumerocliente.Text == "" && txtbairrocliente.Text == "" && mtxtcepcliente.Text == "_____-__" &&
                    txtcidadecliente.Text == "" && cbxufcliente.Text == "" && (mtxtcelularcliente.Text == "(__)_____-____" || mtxttelefone.Text == "(__)____-____" || txtemailcliente.Text == "") &&
                    txtqtdinvcliente.Value <= 0 && txtqtdmodmodcliente.Value <= 0 && cbxmarcainvcliente.Text == "" && cbxmarcamoccliente.Text == "" && cbxmodeloinvcliente.Text == "" && cbxmodelomodcliente.Text == "" &&
                    txtconsmedcliente.Text == "" && txtidentificacaocliente.Text == "")
                {
                    MessageBox.Show("Favor preencher todos os campos!");
                }
                else
                {
                    func.Nome = txtnomecliente.Text;
                    func.CPF = mtxtcpfcnpjcliente.Text;
                    func.Endereco = txtenderecocliente.Text;
                    func.Numero = txtnumerocliente.Text;
                    func.Complemento = txtcomplementocliente.Text;
                    func.Bairro = txtbairrocliente.Text;
                    func.CEP = mtxtcepcliente.Text;
                    func.Cidade = txtcidadecliente.Text;
                    func.UF = cbxufcliente.Text;
                    func.email = txtemailcliente.Text;
                    func.Telefone = mtxttelefone.Text;
                    func.Celular = mtxtcelularcliente.Text;
                    func.QuantidadeInversores = txtqtdinvcliente.Value.ToString();
                    func.QuantidadeModulos = txtqtdmodmodcliente.Value.ToString();
                    func.MarcaInversor = cbxmarcainvcliente.Text;
                    func.MarcaMod = cbxmarcamoccliente.Text;
                    func.ModeloInversor = cbxmodeloinvcliente.Text;
                    func.ModeloModulo = cbxmodelomodcliente.Text;
                    func.MediaConsumo = txtconsmedcliente.Text;
                    func.Identificacao = txtidentificacaocliente.Text;
                    func.AlterarCliente(Banco);
                    MessageBox.Show("Cliente Atualizado com Sucesso!");
                    txtnomecliente.Text = string.Empty;
                    mtxtcpfcnpjcliente.Text = string.Empty;
                    txtenderecocliente.Text = string.Empty;
                    txtnumerocliente.Text = string.Empty;
                    txtcomplementocliente.Text = string.Empty;
                    txtbairrocliente.Text = string.Empty;
                    mtxtcepcliente.Text = string.Empty;
                    txtcidadecliente.Text = string.Empty;
                    cbxufcliente.Text = string.Empty;
                    txtemailcliente.Text = string.Empty;
                    mtxttelefone.Text = string.Empty;
                    mtxtcelularcliente.Text = string.Empty;
                    txtqtdinvcliente.Value = 1;
                    txtqtdmodmodcliente.Value = 1;
                    txtconsmedcliente.Text = string.Empty;
                    txtidentificacaocliente.Text = string.Empty;

                    pnladicionacliente.Visible = false;
                }
            }
            dgvclientes.DataSource = func.AtualizaClientes(Banco);
        }
        private void btnvoltacadastracliente_Click(object sender, EventArgs e)
        {
            txtnomecliente.Text = string.Empty;
            mtxtcpfcnpjcliente.Text = string.Empty;
            txtenderecocliente.Text = string.Empty;
            txtnumerocliente.Text = string.Empty;
            txtcomplementocliente.Text = string.Empty;
            txtbairrocliente.Text = string.Empty;
            mtxtcepcliente.Text = string.Empty;
            txtcidadecliente.Text = string.Empty;
            cbxufcliente.Text = string.Empty;
            txtemailcliente.Text = string.Empty;
            mtxttelefone.Text = string.Empty;
            mtxtcelularcliente.Text = string.Empty;
            txtqtdinvcliente.Value = 1;
            txtqtdmodmodcliente.Value = 1;
            txtconsmedcliente.Text = string.Empty;
            txtidentificacaocliente.Text = string.Empty;

            pnladicionacliente.Visible = false;
        }
        private void cbxmarcainvcliente_SelectionChangeCommitted(object sender, EventArgs e)
        {
            var Dados = func.MarcaInv(cbxmarcainvcliente.Text, Banco);
            cbxmodeloinvcliente.DataSource = Dados;
            cbxmodeloinvcliente.ValueMember = "Modelo";
            cbxmodeloinvcliente.DisplayMember = "Modelo";
        }
        private void cbxmarcamoccliente_SelectionChangeCommitted(object sender, EventArgs e)
        {
            var Dados = func.MarcaModuloOrc(cbxmarcamoccliente.Text, Banco);
            cbxmodelomodcliente.DataSource = Dados;
            cbxmodelomodcliente.ValueMember = "Modelo";
            cbxmodelomodcliente.DisplayMember = "Modelo";
        }
        private void btneditarcliente_Click(object sender, EventArgs e)
        {
            btncadastraouatualiza.Text = "Atualizar";
            func.SelecionaCliente(Banco);
            txtnomecliente.Text = func.Nome;
            mtxtcpfcnpjcliente.Text = func.CPF;
            txtenderecocliente.Text = func.Endereco;
            txtnumerocliente.Text = func.Numero;
            txtcomplementocliente.Text = func.Complemento;
            txtbairrocliente.Text = func.Bairro;
            mtxtcepcliente.Text = func.CEP;
            txtcidadecliente.Text = func.Cidade;
            cbxufcliente.Text = func.UF;
            txtemailcliente.Text = func.email;
            mtxttelefone.Text = func.Telefone;
            mtxtcelularcliente.Text = func.Celular;
            txtqtdinvcliente.Value = Convert.ToDecimal(func.QuantidadeInversores);
            txtqtdmodmodcliente.Value = Convert.ToDecimal(func.QuantidadeModulos);
            txtconsmedcliente.Text = func.MediaConsumo;
            txtidentificacaocliente.Text = func.Identificacao;

            cbxmarcainvcliente.Text = func.MarcaInversor;
            cbxmarcamoccliente.Text = func.MarcaMod;
            cbxmodeloinvcliente.Text = func.ModeloInversor;
            cbxmodelomodcliente.Text = func.ModeloModulo;

            pnladicionacliente.Visible = true;
            pnlvisualizacli.Visible = false;

        }
        private void btnabrirpastacliente_Click(object sender, EventArgs e)
        {
            string pasta = @"\\ANDERSON-PC\Cia Solar\OneDrive\Instalados\" + func.Identificacao;
            try
            {
                Process.Start("explorer.exe", pasta);
            }
            catch
            {
                MessageBox.Show("Não foi possivel encontrar a pasta. Verifique se o computador do Anderson está ligado e possui uma conexão ativa com a internet" + e);
            }
        }
        private void btnexcluircliente_Click(object sender, EventArgs e)
        {
            var resultado = MessageBox.Show("Tem certeza que deseja excluir o cliente?", "Atenção", MessageBoxButtons.YesNo);
            if (resultado == DialogResult.Yes)
            {
                func.ExcluiClientes(Banco);
                dgvclientes.DataSource = func.AtualizaClientes(Banco);
                MessageBox.Show("Excluído com sucesso","Sucesso");
                pnlvisualizacli.Visible = false;
            }
        }
        private void btnvoltavisualizacli_Click(object sender, EventArgs e)
        {
            pnlvisualizacli.Visible = false;
        }
        private void btnlimpaouexclui_Click(object sender, EventArgs e)
        {
            Limpacampos();
            CarregaCombobox();
        }
        private void btnmonitoramentocliente_Click(object sender, EventArgs e)
        {
            if (func.ModeloInversor == "Reno560" || func.ModeloInversor == "Reno560-LV")
            {
                string url = "http://www.renovigiportal.com/";
                Process.Start(url);
            }
            else if (func.ModeloInversor.Contains("Primo") || func.ModeloInversor.Contains("Symo") || func.ModeloInversor.Contains("Eco"))
            {
                string url = "https://www.solarweb.com/";
                Process.Start(url);
            }
            else if (func.ModeloInversor.Contains("Sunny"))
            {
                string url = "https://www.sunnyportal.com/";
                Process.Start(url);
            }
            else if (func.ModeloInversor.Contains("REFUone"))
            {
                string url = "https://refu-log.com/";
                Process.Start(url);
            }
            else if (func.ModeloInversor.Contains("SE27.6k") || func.ModeloInversor.Contains("SE17K") || func.ModeloInversor.Contains("SE75k") || func.ModeloInversor.Contains("SE100k"))
            {
                string url = "https://monitoring.solaredge.com/solaredge-web/p/home";
                Process.Start(url);
            }
            else
            {
                string url = "http://www.renovigi.solar/cus/renovigi/index_po.html?1578427279140";
                Process.Start(url);
            }
        }

        //Painel Projeto
        private void btnvoltaproj1_Click(object sender, EventArgs e)
        {
            pnlnovoproj0.Visible = true;
            pnlnovoproj1.Visible = false;
            pnlnovoproj2.Visible = false;
            pnlnovoproj3.Visible = false;
            Limpacampos();
        }
        private void btncadastraclienteprojeto_Click(object sender, EventArgs e)
        {
            Limpacampos();
            ClicaClientes();
            PaineisPrincipais(pnlclientes);
            if (pnlvisualizacli.Visible == true)
            {
                pnlvisualizacli.Visible = false;
            }
            btncadastraouatualiza.Text = "Cadastrar";
            pnladicionacliente.Visible = true;
        }
        private void btniniciarprojeto_Click(object sender, EventArgs e)
        {
            if(dgvclientes.Rows.Count == 0)
            {
                MessageBox.Show("Nenhum cliente cadastrado", "Atenção");
            }
            else
            {
                func.id = dgvencontraclienteprojeto.CurrentRow.Cells[0].Value.ToString();
                func.SelecionaCliente(Banco);
                if(func.Celular.Contains("(35)") || func.Telefone.Contains("(35)"))
                {
                    txtnomecliprojeto.Text = func.Nome;
                    mtxtcpjcliprojeto.Text = func.CPF;
                    pnlnovoproj1.Visible = false;
                    pnlnovoproj2.Visible = true;
                    pnlnovoproj3.Visible = false;
                }
                else
                {
                    MessageBox.Show("O cliente deve possuir ao menos um telefone ou celular válido!", "Atenção");
                }
            }
            
        }
        private void pbxbuscaprojeto_Click(object sender, EventArgs e)
        {
            dgvprojetos.DataSource = func.PesquisaProjFiltro(txtprocuraprojeto.Text, cbxfiltroprojeto.Text, Banco);
        }
        private void txtprocuraprojeto_Enter(object sender, EventArgs e)
        {
            if (txtprocuraprojeto.Text == "O que você procura?")
            {
                txtprocuraprojeto.Text = "";
            }
            txtprocuraprojeto.ForeColor = Color.Black;
        }
        private void txtprocuraprojeto_Leave(object sender, EventArgs e)
        {
            if (txtprocuraprojeto.Text == "")
            {
                txtprocuraprojeto.Text = "O que você procura?";
                dgvinversores.DataSource = func.AtualizaInversor(Banco);
            }
            txtprocuraprojeto.ForeColor = Color.Silver;
        }
        private void txtprocuraprojeto_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                dgvprojetos.DataSource = func.PesquisaProjFiltro(txtprocuraprojeto.Text, cbxfiltroprojeto.Text, Banco);
            }
        }
        private void btnvoltaproj2_Click(object sender, EventArgs e)
        {
            Limpacampos();
            pnlnovoproj1.Visible = true;
            pnlnovoproj2.Visible = false;
            pnlnovoproj3.Visible = false;
        }
        private void btnnovoprojeto_Click(object sender, EventArgs e)
        {
            pnlnovoproj0.Visible = false;
            pnlnovoproj1.Visible = true;
            pnlnovoproj2.Visible = false;
            pnlnovoproj3.Visible = false;
        }
        private void brnproxproj3_Click(object sender, EventArgs e)
        {
            if(txtnomecliprojeto.Text != "" && mtxtcpjcliprojeto.Text != string.Empty && txtnumcliproj.Text != "" && txtnuminstproj.Text != "" && txtcargainstproj.Text != "" &&
                cbxclasseproj.Text != "" && cbxpadraoproj.Text != "" && cbxdisjproj.Text != "" && mtxtlatitudeproj.Text != "" && mtxtlongitudeproj.Text != "" && 
                cbxstringboxproj.Text != "" && txtarranjoproj.Text != "" && cbxtensoesatenproj.Text != "" && cbxestruturaproj.Text != "")
            {
                if(txtqtdinstproj.Value <= 0)
                {
                    MessageBox.Show("Numero de instalações a receber o crédito deve ser maior que zero", "Atenção");
                }
                else
                {
                    lblnumerocliprojeto.Text = txtnumcliproj.Text;
                    lblnumeroinstprojeto.Text = txtnuminstproj.Text;
                    lblclasseprojeto.Text = cbxclasseproj.Text + " " + cbxpadraoproj.Text;
                    lbltitularprojeto.Text = txtnomecliprojeto.Text;
                    lblcpjcnpjprojeto.Text = mtxtcpjcliprojeto.Text;
                    lblenderecoprojeto.Text = func.Endereco;
                    lblnumeroenderecoprojeto.Text = func.Numero;
                    lblcomplementoprojeto.Text = func.Complemento;
                    lblbairroprojeto.Text = func.Bairro;
                    lblcepprojeto.Text = func.CEP;
                    lblcidadeprojeto.Text = func.Cidade;
                    lblestadoprojeto.Text = func.UF;
                    lbltelefoneprojeto.Text = func.Telefone;
                    lblcelularprojeto.Text = func.Celular;
                    lblqtdinvprojeto.Text = func.QuantidadeInversores;
                    lblqtdmodprojeto.Text = func.QuantidadeModulos;
                    lblmarcainvprojeto.Text = func.MarcaInversor;
                    lblmarcamodprojeto.Text = func.MarcaMod;
                    lblmodinvprojeto.Text = func.ModeloInversor;
                    lblmodmodprojeto.Text = func.ModeloModulo;
                    lblarranjoprojeto.Text = (Int16.Parse(func.QuantidadeModulos) * 2).ToString();
                    lblqtducprojeto.Text = txtqtdinstproj.Value.ToString();
                    func.SelecionaInversorModelo(Banco,func.ModeloInversor);
                    func.SelecionaPainelModelo(Banco, func.ModeloModulo);
                    string aux = ((Double.Parse(func.QuantidadeModulos) * Double.Parse(func.PotenciaMod))/1000).ToString();
                    pottotalmodprojeto = aux;
                    pottotalmodprojeto = string.Format("{0:0,0.00}", aux);
                    aux = ((Double.Parse(func.QuantidadeInversores) * Double.Parse(func.PotenciaInv)) / 1000).ToString();
                    pottotalinvprojeto = aux;
                    pottotalinvprojeto = string.Format("{0:0,0.00}", aux);
                    lblpottotalmodprojeto.Text = pottotalmodprojeto;
                    lblpottotinvprojeto.Text = pottotalinvprojeto;
                    pnlnovoproj1.Visible = false;
                    pnlnovoproj2.Visible = false;
                    pnlnovoproj3.Visible = true;
                }
            }
            else
            {
                MessageBox.Show("Favor preencher todos os campos", "Atenção");
            }
        }
        private void button45_Click(object sender, EventArgs e)
        {
            pnlnovoproj1.Visible = false;
            pnlnovoproj2.Visible = true;
            pnlnovoproj3.Visible = false;
        }
        private void btngerarprojeto_Click(object sender, EventArgs e)
        {
            if (VerificaArquivos())
            {
                Projeto();
            }
        }
        private void btnsalvarprojeto_Click(object sender, EventArgs e)
        {
            if (btnsalvarprojeto.Text == "Salvar")
            {
                if (VerificaArquivos())
                {
                    Projeto();
                    func.NumeroCliente = txtnumcliproj.Text;
                    func.NumeroInstalacao = txtnuminstproj.Text;
                    func.Classe = lblclasseprojeto.Text;
                    func.Latitude = mtxtlatitudeproj.Text;
                    func.Longitude = mtxtlongitudeproj.Text;
                    func.Disjuntor = cbxdisjproj.Text;
                    func.CargaInstalada = txtcargainstproj.Text;
                    func.Tensao = cbxtensoesatenproj.Text;
                    func.Estu = cbxestruturaproj.Text;
                    func.Transformador = cbxtransformadorproj.Text;
                    func.StringBox = cbxstringboxproj.Text;
                    func.Credito = txtqtdinstproj.Value.ToString();
                    func.Arranjo = txtarranjoproj.Text;
                    func.SalvaProj(Banco);
                    pnlnovoproj1.Visible = false;
                    pnlnovoproj2.Visible = false;
                    pnlnovoproj3.Visible = false;
                    pnlprojeto.Visible = false;
                    pnlfinalizaprojeto.Visible = true;
                }
            }
            else
            {
                if (VerificaArquivos())
                {
                    Projeto();
                    func.NumeroCliente = txtnumcliproj.Text;
                    func.NumeroInstalacao = txtnuminstproj.Text;
                    func.Classe = lblclasseprojeto.Text;
                    func.Latitude = mtxtlatitudeproj.Text;
                    func.Longitude = mtxtlongitudeproj.Text;
                    func.Disjuntor = cbxdisjproj.Text;
                    func.CargaInstalada = txtcargainstproj.Text;
                    func.Tensao = cbxtensoesatenproj.Text;
                    func.Estu = cbxestruturaproj.Text;
                    func.Transformador = cbxtransformadorproj.Text;
                    func.StringBox = cbxstringboxproj.Text;
                    func.Credito = txtqtdinstproj.Value.ToString();
                    func.Arranjo = txtarranjoproj.Text;
                    func.AlteraProj(Banco);
                    pnlnovoproj1.Visible = false;
                    pnlnovoproj2.Visible = false;
                    pnlnovoproj3.Visible = false;
                    pnlprojeto.Visible = false;
                    pnlfinalizaprojeto.Visible = true;
                }
            }
        }
        private void btnexcluirprojeto_Click(object sender, EventArgs e)
        {
            var resultado = MessageBox.Show("Tem certeza que deseja excluir o projeto?", "Atenção", MessageBoxButtons.YesNo);
            if (resultado == DialogResult.Yes)
            {
                Limpacampos();
                CarregaCombobox();
                PaineisPrincipais(pnlprojeto);
            }
        }
        private void button43_Click(object sender, EventArgs e)
        {
            CarregaDataGrid();
            pnlnovoproj1.Visible = false;
            pnlnovoproj2.Visible = false;
            pnlnovoproj3.Visible = false;
            pnlfinalizaprojeto.Visible = false;
            Limpacampos();
            PaineisPrincipais(pnlprojeto);
        }
        private void button42_Click(object sender, EventArgs e)
        {
            CarregaDataGrid();
            Limpacampos();
            PaineisPrincipais(pnlprojeto);
            pnlnovoproj1.Visible = true;
            pnlnovoproj2.Visible = false;
            pnlnovoproj3.Visible = false;
            pnlfinalizaprojeto.Visible = false;
        }
        private void btnvisualizarprojeto_Click(object sender, EventArgs e)
        {
            btnsalvarprojeto.Text = "Atualizar";

            func.id = dgvprojetos.CurrentRow.Cells[0].Value.ToString();
            func.SelecionaProjeto(Banco);

            txtnomecliprojeto.Text = func.Nome;
            mtxtcpjcliprojeto.Text = func.CPF;

            txtnumcliproj.Text = func.NumeroCliente;
            txtnuminstproj.Text = func.NumeroInstalacao;
            mtxtlatitudeproj.Text = func.Latitude;
            mtxtlongitudeproj.Text = func.Longitude;
            cbxdisjproj.Text = func.Disjuntor;
            txtcargainstproj.Text = func.CargaInstalada;
            cbxtensoesatenproj.Text = func.Tensao;
            cbxestruturaproj.Text = func.Estu;
            cbxtransformadorproj.Text = func.Transformador;
            cbxstringboxproj.Text = func.StringBox;
            txtqtdinstproj.Value = Decimal.Round(decimal.Parse(func.Credito));
            txtarranjoproj.Text = func.Arranjo;
            if (func.Classe.Contains("Residencial"))
            {
                cbxclasseproj.Text = "Residencial";
            }
            if (func.Classe.Contains("Industrial"))
            {
                cbxclasseproj.Text = "Industrial";
            }
            if (func.Classe.Contains("Comercial"))
            {
                cbxclasseproj.Text = "Comercial";
            }
            if (func.Classe.Contains("Rural"))
            {
                cbxclasseproj.Text = "Rural";
            }
            if (func.Classe.Contains("Monofasico"))
            {
                cbxpadraoproj.Text = "Monofasico";
            }
            if (func.Classe.Contains("Bifasico"))
            {
                cbxpadraoproj.Text = "Bifasico";
            }
            if (func.Classe.Contains("Trifasico"))
            {
                cbxpadraoproj.Text = "Trifasico";
            }
            lblnumerocliprojeto.Text = func.NumeroCliente;
            lblnumeroinstprojeto.Text = func.NumeroInstalacao;
            lblclasseprojeto.Text = func.Classe;
            lbltitularprojeto.Text = func.Nome;
            lblcpjcnpjprojeto.Text = func.CPF;
            lblenderecoprojeto.Text = func.Endereco;
            lblnumeroenderecoprojeto.Text = func.Numero;
            lblcomplementoprojeto.Text = func.Complemento;
            lblbairroprojeto.Text = func.Bairro;
            lblcepprojeto.Text = func.CEP;
            lblcidadeprojeto.Text = func.Cidade;
            lblestadoprojeto.Text = func.UF;
            lbltelefoneprojeto.Text = func.Telefone;
            lblcelularprojeto.Text = func.Celular;
            lblqtdinvprojeto.Text = func.QuantidadeInversores;
            lblqtdmodprojeto.Text = func.QuantidadeModulos;
            lblmarcainvprojeto.Text = func.MarcaInversor;
            lblmarcamodprojeto.Text = func.MarcaMod;
            lblmodinvprojeto.Text = func.ModeloInversor;
            lblmodmodprojeto.Text = func.ModeloModulo;
            lblarranjoprojeto.Text = (Int16.Parse(func.QuantidadeModulos) * 2).ToString();
            lblqtducprojeto.Text = func.Credito;
            func.SelecionaInversorModelo(Banco, func.ModeloInversor);
            func.SelecionaPainelModelo(Banco, func.ModeloModulo);
            string aux = ((Double.Parse(func.QuantidadeModulos) * Double.Parse(func.PotenciaMod)) / 1000).ToString();
            pottotalmodprojeto = aux;
            pottotalmodprojeto = string.Format("{0:0,0.00}", aux);
            aux = ((Double.Parse(func.QuantidadeInversores) * Double.Parse(func.PotenciaInv)) / 1000).ToString();
            pottotalinvprojeto = aux;
            pottotalinvprojeto = string.Format("{0:0,0.00}", aux);
            lblpottotalmodprojeto.Text = pottotalmodprojeto;
            lblpottotinvprojeto.Text = pottotalinvprojeto;

            pnlnovoproj0.Visible = false;
            pnlnovoproj1.Visible = false;
            pnlnovoproj2.Visible = false;
            pnlnovoproj3.Visible = true;
        }

        //Painel Equipamentos
        private void lblmodulo_Click(object sender, EventArgs e)
        {
            PaineisPrincipais(pnlmod1);
        }
        private void pbxinversor_Click(object sender, EventArgs e)
        {
            PaineisPrincipais(pnlinv1);
        }
        private void pbxmodulo_Click(object sender, EventArgs e)
        {
            PaineisPrincipais(pnlmod1);
        }
        private void lblinversor_Click(object sender, EventArgs e)
        {
            PaineisPrincipais(pnlinv1);
        }

        //Módulos
        private void btnvoltaequipamentos_Click(object sender, EventArgs e)
        {
            PaineisPrincipais(pnlequipamentos);
        }
        private void txtprocuramodulo_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                dgvmodulos.DataSource = func.FiltroMod(txtprocuramodulo.Text, cbxfiltromodulo.Text, Banco);
            }
        }
        private void pbxbuscamodulo_Click(object sender, EventArgs e)
        {
            dgvmodulos.DataSource = func.FiltroMod(txtprocuramodulo.Text, cbxfiltromodulo.Text, Banco);
        }
        private void txtprocuramodulo_Enter(object sender, EventArgs e)
        {
            if (txtprocuramodulo.Text == "O que você procura?")
            {
                txtprocuramodulo.Text = "";
            }
            txtprocuramodulo.ForeColor = Color.Black;
        }
        private void txtprocuramodulo_Leave(object sender, EventArgs e)
        {
            if (txtprocuramodulo.Text == "")
            {
                txtprocuramodulo.Text = "O que você procura?";
                dgvmodulos.DataSource = func.AtualizaPaineis(Banco);
            }
            txtprocuramodulo.ForeColor = Color.Silver;
        }
        private void btnvoltamod1equip_Click(object sender, EventArgs e)
        {
            Limpacampos();
            PaineisPrincipais(pnlmod1);
            btnsalvamodequip.Text = "Cadastrar";
            btnlimpamodequip.Text = "Limpar Campos";
        }
        private void btncadastrarmodulo_Click(object sender, EventArgs e)
        {
            panel18.Visible = true;
        }
        private void btnvisualizarmodulo_Click(object sender, EventArgs e)
        {
            btnsalvamodequip.Text = "Atualizar";
            btnlimpamodequip.Text = "Excluir";

            func.ModeloModulo = dgvmodulos.CurrentRow.Cells[2].Value.ToString();
            func.SelecionaPainel(Banco, func.ModeloModulo);

            cbxmarcamodequipamentos.Text = func.MarcaMod;
            cbxmodelomodequipamentos.Text = func.ModeloModulo;
            txtpotenciamodequipamentos.Text = func.PotenciaMod;
            txtcoefmodequipamentos.Text = func.TemperaturaModulo;
            cbxmaterialmodequipamentos.Text = func.Material;
            cbxcelulasmodequipamentos.Text = func.Celulas;
            txtcompmodequip.Text = func.ComrpimentoMod;
            txtlargmodequip.Text = func.LarguraMod;
            txtgarantiamodequip.Text = func.GarantiaMod;
            txtreginmmodequip.Text = func.RegistroInmetro;
            lblIdModulo.Text = func.Id.ToString();
            panel18.Visible = true;
        }
        private void btnsalvamodequip_Click(object sender, EventArgs e)
        {
            if (btnsalvamodequip.Text == "Atualizar")
            {
                func.MarcaMod = cbxmarcamodequipamentos.Text;
                func.ModeloModulo = cbxmodelomodequipamentos.Text;
                func.PotenciaMod = txtpotenciamodequipamentos.Text;
                func.TemperaturaModulo = txtcoefmodequipamentos.Text;
                func.Material = cbxmaterialmodequipamentos.Text;
                func.Celulas = cbxcelulasmodequipamentos.Text;
                func.ComrpimentoMod = txtcompmodequip.Text;
                func.LarguraMod = txtlargmodequip.Text;
                func.GarantiaMod = txtgarantiamodequip.Text;
                func.RegistroInmetro = txtreginmmodequip.Text;
                func.Id = Convert.ToInt32(lblIdModulo.Text);
                func.id = func.Id.ToString();
                func.AlterarPaineis(Banco);
                dgvmodulos.DataSource = func.AtualizaPaineis(Banco);
                Limpacampos();
                PaineisPrincipais(pnlmod1);
            }
            else
            {
                func.MarcaMod = cbxmarcamodequipamentos.Text;
                func.ModeloModulo = cbxmodelomodequipamentos.Text;
                func.PotenciaMod = txtpotenciamodequipamentos.Text;
                func.TemperaturaModulo = txtcoefmodequipamentos.Text;
                func.Material = cbxmaterialmodequipamentos.Text;
                func.Celulas = cbxcelulasmodequipamentos.Text;
                func.ComrpimentoMod = txtcompmodequip.Text;
                func.LarguraMod = txtlargmodequip.Text;
                func.GarantiaMod = txtgarantiamodequip.Text;
                func.RegistroInmetro = txtreginmmodequip.Text;
                func.InserirPainel(Banco);
                dgvmodulos.DataSource = func.AtualizaPaineis(Banco);
                Limpacampos();
                PaineisPrincipais(pnlmod1);
            }
        }
        private void btnlimpamodequip_Click(object sender, EventArgs e)
        {
            if (btnlimpamodequip.Text == "Excluir")
            {
                var resultado = MessageBox.Show("Tem certeza que deseja excluir o cadastro?", "Atenção", MessageBoxButtons.YesNo);
                if (resultado == DialogResult.Yes)
                {
                    func.ExcluiPainel(Banco, func.ModeloModulo);
                    Limpacampos();
                    CarregaCombobox();
                    dgvmodulos.DataSource = func.AtualizaPaineis(Banco);
                    PaineisPrincipais(pnlmod1);
                }
            }
            else
            {
                cbxmarcamodequipamentos.Text = "";
                cbxmodelomodequipamentos.Text = "";
                txtpotenciamodequipamentos.Text = "";
                txtcoefmodequipamentos.Text = "";
                cbxmaterialmodequipamentos.Text = "";
                cbxcelulasmodequipamentos.Text = "";
                txtcompmodequip.Text = "";
                txtlargmodequip.Text = "";
                txtgarantiamodequip.Text = "";
                txtreginmmodequip.Text = "";
                cbxmaterialmodequipamentos.SelectedIndex = -1;
            }
        }

        //Inversores
        private void btnvoltainv1_Click(object sender, EventArgs e)
        {
            Limpacampos();
            PaineisPrincipais(pnlinv1);
            panel20.Visible = false;
            btnsalvainvequip.Text = "Cadastrar";
            btnlimpainvequip.Text = "Limpar Campos";
        }
        private void txtprocurainversor_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                dgvinversores.DataSource = func.FiltroInv(txtprocurainversor.Text, cbxfiltroinv.Text, Banco);
            }
        }
        private void pbxbuscainversor_Click(object sender, EventArgs e)
        {
            dgvinversores.DataSource = func.FiltroInv(txtprocurainversor.Text, cbxfiltroinv.Text, Banco);
        }
        private void btncadastrarinversor_Click(object sender, EventArgs e)
        {
            panel20.Visible = true;
        }
        private void btnvisualizarinversor_Click(object sender, EventArgs e)
        {
            btnsalvainvequip.Text = "Atualizar";
            btnlimpainvequip.Text = "Excluir";

            //func.id = dgvinversores.CurrentRow.Cells[0].Value.ToString();
            func.ModeloInversor = dgvinversores.CurrentRow.Cells[2].Value.ToString();
            func.SelecionaInversor(Banco,func.ModeloInversor);

            cbxmarcainvequip.Text = func.MarcaInversor;
            cbxmodinvequip.Text = func.ModeloInversor;
            txtpotenciainvequip.Text = func.PotenciaInv;
            cbxfasesinvequip.Text = func.Fases;
            cbxtensaoinvequip.Text = func.Tensao;
            txteficienciainvequip.Text = func.EficienciaInv;
            txtgarantiainvequip.Text = func.GarantiaInv;
            txtreginminvequip.Text = func.RegistroINMETRO;
            txtqtdmpptinvequip.Value = Convert.ToDecimal(func.Qtdmppt);
            lblIdInversor.Text = func.Id.ToString();
            panel20.Visible = true;
        }
        private void button75_Click(object sender, EventArgs e)
        {
            panel20.Visible = false;
            PaineisPrincipais(pnlequipamentos);
        }
        private void txtprocurainversor_Enter(object sender, EventArgs e)
        {
            if (txtprocurainversor.Text == "O que você procura?")
            {
                txtprocurainversor.Text = "";
            }
            txtprocurainversor.ForeColor = Color.Black;
        }
        private void txtprocurainversor_Leave(object sender, EventArgs e)
        {
            if (txtprocurainversor.Text == "")
            {
                txtprocurainversor.Text = "O que você procura?";
                dgvinversores.DataSource = func.AtualizaInversor(Banco);
            }
            txtprocurainversor.ForeColor = Color.Silver;
        }
        private void btnsalvainvequip_Click(object sender, EventArgs e)
        {
            if(btnsalvainvequip.Text == "Atualizar")
            {
                func.MarcaInversor = cbxmarcainvequip.Text;
                func.ModeloInversor = cbxmodinvequip.Text;
                func.PotenciaInv = txtpotenciainvequip.Text;
                func.Fases = cbxfasesinvequip.Text;
                func.Tensao = cbxtensaoinvequip.Text;
                func.EficienciaInv = txteficienciainvequip.Text;
                func.GarantiaInv = txtgarantiainvequip.Text;
                func.RegistroINMETRO = txtreginminvequip.Text;
                func.Qtdmppt = txtqtdmpptinvequip.Value.ToString();
                func.Id = Convert.ToInt32(lblIdInversor.Text);
                func.id = func.Id.ToString();
                func.AlterarInversor(Banco);
                dgvinversores.DataSource = func.AtualizaInversor(Banco);
                Limpacampos();
                PaineisPrincipais(pnlinv1);
            }
            else
            {
                func.MarcaInversor = cbxmarcainvequip.Text;
                func.ModeloInversor = cbxmodinvequip.Text;
                func.PotenciaInv = txtpotenciainvequip.Text;
                func.Fases = cbxfasesinvequip.Text;
                func.Tensao = cbxtensaoinvequip.Text;
                func.EficienciaInv = txteficienciainvequip.Text;
                func.GarantiaInv = txtgarantiainvequip.Text;
                func.RegistroINMETRO = txtreginminvequip.Text;
                func.Qtdmppt = txtqtdmpptinvequip.Value.ToString();
                func.InserirInversor(Banco);
                dgvinversores.DataSource = func.AtualizaInversor(Banco);
                Limpacampos();
                PaineisPrincipais(pnlinv1);
            }
        }
        private void btnlimpainvequip_Click(object sender, EventArgs e)
        {
            if(btnlimpainvequip.Text == "Excluir")
            {
                var resultado = MessageBox.Show("Tem certeza que deseja excluir o cadastro?", "Atenção", MessageBoxButtons.YesNo);
                if (resultado == DialogResult.Yes)
                {
                    func.ExcluiInversor(Banco, func.ModeloInversor);
                    Limpacampos();
                    CarregaCombobox();
                    dgvinversores.DataSource = func.AtualizaInversor(Banco);
                    PaineisPrincipais(pnlinv1);
                }
            }
            else
            {
                cbxmarcainvequip.Text = "";
                cbxmodinvequip.Text = "";
                txtpotenciainvequip.Text = "";
                cbxfasesinvequip.Text = "";
                cbxfasesinvequip.SelectedIndex = -1;
                cbxtensaoinvequip.Text = "";
                txteficienciainvequip.Text = "";
                txtgarantiainvequip.Text = "";
                txtreginminvequip.Text = "";
                txtqtdmpptinvequip.Value = 1;
            }
            
        }

    }
}