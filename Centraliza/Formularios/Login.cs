using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;

namespace Centraliza
{
    public partial class Login : Form
    {
        Thread nt;
        FuncoesBanco funcao = new FuncoesBanco();
        string aux;

        public Login()
        {
            InitializeComponent();
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
        }

        private void Novoform()
        {
            Application.Run(new TelaInicial(aux));
        }

        private void btnautenticar_Click(object sender, EventArgs e)
        {
            //funcao.SelecionaBanco();
            //funcao.PesquisaLogin(txtlogin.Text, funcao.Banco);
            //aux = txtlogin.Text;
            //if((txtlogin.Text == funcao.Login) && (txtsenha.Text == funcao.Senha))
            //{
                if(((TelaInicial)this.Owner).ValidaLogin(txtlogin.Text, txtsenha.Text))
                {
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Usuário ou Senha incorretos", "Atenção");
                }
                /*nt = new Thread(Novoform);
                nt.SetApartmentState(ApartmentState.STA);
                nt.Start();*/
            /*}
            else
            {
                MessageBox.Show("Usuário ou Senha incorretos", "Atenção");
            }*/
        }
        private void txtsenha_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                /*funcao.SelecionaBanco();
                funcao.PesquisaLogin(txtlogin.Text, funcao.Banco);

                if ((txtlogin.Text == funcao.Login) && (txtsenha.Text == funcao.Senha))
                {*/
                    if (((TelaInicial)this.Owner).ValidaLogin(txtlogin.Text, txtsenha.Text))
                    {
                        this.Close();
                    }
                    else
                    {
                        MessageBox.Show("Usuário ou Senha incorretos", "Atenção");
                    }
                    /*TelaInicial t = new TelaInicial(txtlogin.Text);
                    t.Show();*/
                    /*nt = new Thread(Novoform);
                    nt.SetApartmentState(ApartmentState.STA);
                    nt.Start();*/
                /*}
                else
                {
                    MessageBox.Show("Usuário ou Senha incorretos", "Atenção");
                }*/
            }
        }

        private void txtlogin_Enter(object sender, EventArgs e)
        {
            if (txtlogin.Text == "Usuário")
            {
                txtlogin.Text = "";
            }
            txtlogin.ForeColor = Color.Black;
        }
        private void txtlogin_Leave(object sender, EventArgs e)
        {
            if (txtlogin.Text == "")
            {
                txtlogin.Text = "Usuário";
            }
            txtlogin.ForeColor = Color.FromArgb(120, 120, 120);
        }
        private void txtsenha_Enter(object sender, EventArgs e)
        {
            if (txtsenha.Text == "Senha")
            {
                txtsenha.Text = "";
                txtsenha.PasswordChar = '*';
            }
            txtsenha.ForeColor = Color.Black;
        }
        private void txtsenha_Leave(object sender, EventArgs e)
        {
            if (txtsenha.Text == "")
            {
                txtsenha.Text = "Senha";
                txtsenha.PasswordChar= '\0';
            }
            txtsenha.ForeColor = Color.FromArgb(120, 120, 120);
        }
        private void textBox1_Enter(object sender, EventArgs e)
        {
            if (textBox1.Text == "Nome Completo")
            {
                textBox1.Text = "";
            }
            textBox1.ForeColor = Color.Black;
        }
        private void textBox1_Leave(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                textBox1.Text = "Nome Completo";
            }
            textBox1.ForeColor = Color.FromArgb(120, 120, 120);
        }
        private void textBox2_Enter(object sender, EventArgs e)
        {
            if (textBox2.Text == "email")
            {
                textBox2.Text = "";
            }
            textBox2.ForeColor = Color.Black;
        }
        private void textBox2_Leave(object sender, EventArgs e)
        {
            if (textBox2.Text == "")
            {
                textBox2.Text = "email";
            }
            textBox2.ForeColor = Color.FromArgb(120, 120, 120);
        }

        private void lblfazercadastro_Click(object sender, EventArgs e)
        {
            pnlcadastrousuario.Visible = true;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            pnlcadastrousuario.Visible = false;
        }
        private void btnvolta4_Click(object sender, EventArgs e)
        {
            pnlcadastrousuario.Visible = false;
        }
    }
}
