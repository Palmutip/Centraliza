using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System;
using System.IO;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Runtime.InteropServices;
using System.Diagnostics;


namespace Centraliza
{
    public class Orcamento
    {
        public double potenciagerada;
        public double gerano;
        public double total = 0;

        public Orcamento()
        {

        }

        public string EscreverExtenso(decimal valor)
        {
            if (valor <= 0 | valor >= 1000000000000000)
                return "Valor não suportado pelo sistema.";
            else
            {
                string strValor = valor.ToString("000000000000000.00");
                string valor_por_extenso = string.Empty;

                for (int i = 0; i <= 15; i += 3)
                {
                    valor_por_extenso += Escrever_Valor_Extenso(Convert.ToDecimal(strValor.Substring(i, 3)));

                    if (i == 0 & valor_por_extenso != string.Empty)
                    {
                        if (Convert.ToInt32(strValor.Substring(0, 3)) == 1)
                            valor_por_extenso += " trilhão" + ((Convert.ToDecimal(strValor.Substring(3, 12)) > 0) ? " e " : string.Empty);
                        else if (Convert.ToInt32(strValor.Substring(0, 3)) > 1)
                            valor_por_extenso += " trilhões" + ((Convert.ToDecimal(strValor.Substring(3, 12)) > 0) ? " e " : string.Empty);
                    }
                    else if (i == 3 & valor_por_extenso != string.Empty)
                    {
                        if (Convert.ToInt32(strValor.Substring(3, 3)) == 1)
                            valor_por_extenso += " bilhão" + ((Convert.ToDecimal(strValor.Substring(6, 9)) > 0) ? " e " : string.Empty);
                        else if (Convert.ToInt32(strValor.Substring(3, 3)) > 1)
                            valor_por_extenso += " bilhões" + ((Convert.ToDecimal(strValor.Substring(6, 9)) > 0) ? " e " : string.Empty);
                    }
                    else if (i == 6 & valor_por_extenso != string.Empty)
                    {
                        if (Convert.ToInt32(strValor.Substring(6, 3)) == 1)
                            valor_por_extenso += " milhão" + ((Convert.ToDecimal(strValor.Substring(9, 6)) > 0) ? " e " : string.Empty);
                        else if (Convert.ToInt32(strValor.Substring(6, 3)) > 1)
                            valor_por_extenso += " milhões" + ((Convert.ToDecimal(strValor.Substring(9, 6)) > 0) ? " e " : string.Empty);
                    }
                    else if (i == 9 & valor_por_extenso != string.Empty)
                        if (Convert.ToInt32(strValor.Substring(9, 3)) > 0)
                            valor_por_extenso += " mil" + ((Convert.ToDecimal(strValor.Substring(12, 3)) > 0) ? " e " : string.Empty);

                    if (i == 12)
                    {
                        if (valor_por_extenso.Length > 8)
                            if (valor_por_extenso.Substring(valor_por_extenso.Length - 6, 6) == "bilhão" | valor_por_extenso.Substring(valor_por_extenso.Length - 6, 6) == "milhão")
                                valor_por_extenso += " de";
                            else
                                if (valor_por_extenso.Substring(valor_por_extenso.Length - 7, 7) == "bilhões" | valor_por_extenso.Substring(valor_por_extenso.Length - 7, 7) == "milhões"
    | valor_por_extenso.Substring(valor_por_extenso.Length - 8, 7) == "trilhão")
                                valor_por_extenso += " de";
                            else
                                    if (valor_por_extenso.Substring(valor_por_extenso.Length - 8, 8) == "trilhões")
                                valor_por_extenso += " de";

                        if (Convert.ToInt64(strValor.Substring(0, 15)) == 1)
                            valor_por_extenso += " real";
                        else if (Convert.ToInt64(strValor.Substring(0, 15)) > 1)
                            valor_por_extenso += " reais";

                        if (Convert.ToInt32(strValor.Substring(16, 2)) > 0 && valor_por_extenso != string.Empty)
                            valor_por_extenso += " e ";
                    }

                    if (i == 15)
                        if (Convert.ToInt32(strValor.Substring(16, 2)) == 1)
                            valor_por_extenso += " centavo";
                        else if (Convert.ToInt32(strValor.Substring(16, 2)) > 1)
                            valor_por_extenso += " centavos";
                }
                return valor_por_extenso;
            }
        }
        private static string Escrever_Valor_Extenso(decimal valor)
        {
            if (valor <= 0)
                return string.Empty;
            else
            {
                string montagem = string.Empty;
                if (valor > 0 & valor < 1)
                {
                    valor *= 100;
                }
                string strValor = valor.ToString("000");
                int a = Convert.ToInt32(strValor.Substring(0, 1));
                int b = Convert.ToInt32(strValor.Substring(1, 1));
                int c = Convert.ToInt32(strValor.Substring(2, 1));

                if (a == 1) montagem += (b + c == 0) ? "cem" : "cento";
                else if (a == 2) montagem += "duzentos";
                else if (a == 3) montagem += "trezentos";
                else if (a == 4) montagem += "quatrocentos";
                else if (a == 5) montagem += "quinhentos";
                else if (a == 6) montagem += "seiscentos";
                else if (a == 7) montagem += "setecentos";
                else if (a == 8) montagem += "oitocentos";
                else if (a == 9) montagem += "novecentos";

                if (b == 1)
                {
                    if (c == 0) montagem += ((a > 0) ? " e " : string.Empty) + "dez";
                    else if (c == 1) montagem += ((a > 0) ? " e " : string.Empty) + "onze";
                    else if (c == 2) montagem += ((a > 0) ? " e " : string.Empty) + "doze";
                    else if (c == 3) montagem += ((a > 0) ? " e " : string.Empty) + "treze";
                    else if (c == 4) montagem += ((a > 0) ? " e " : string.Empty) + "quatorze";
                    else if (c == 5) montagem += ((a > 0) ? " e " : string.Empty) + "quinze";
                    else if (c == 6) montagem += ((a > 0) ? " e " : string.Empty) + "dezesseis";
                    else if (c == 7) montagem += ((a > 0) ? " e " : string.Empty) + "dezessete";
                    else if (c == 8) montagem += ((a > 0) ? " e " : string.Empty) + "dezoito";
                    else if (c == 9) montagem += ((a > 0) ? " e " : string.Empty) + "dezenove";
                }
                else if (b == 2) montagem += ((a > 0) ? " e " : string.Empty) + "vinte";
                else if (b == 3) montagem += ((a > 0) ? " e " : string.Empty) + "trinta";
                else if (b == 4) montagem += ((a > 0) ? " e " : string.Empty) + "quarenta";
                else if (b == 5) montagem += ((a > 0) ? " e " : string.Empty) + "cinquenta";
                else if (b == 6) montagem += ((a > 0) ? " e " : string.Empty) + "sessenta";
                else if (b == 7) montagem += ((a > 0) ? " e " : string.Empty) + "setenta";
                else if (b == 8) montagem += ((a > 0) ? " e " : string.Empty) + "oitenta";
                else if (b == 9) montagem += ((a > 0) ? " e " : string.Empty) + "noventa";

                if (strValor.Substring(1, 1) != "1" & c != 0 & montagem != string.Empty) montagem += " e ";

                if (strValor.Substring(1, 1) != "1")
                    if (c == 1) montagem += "um";
                    else if (c == 2) montagem += "dois";
                    else if (c == 3) montagem += "três";
                    else if (c == 4) montagem += "quatro";
                    else if (c == 5) montagem += "cinco";
                    else if (c == 6) montagem += "seis";
                    else if (c == 7) montagem += "sete";
                    else if (c == 8) montagem += "oito";
                    else if (c == 9) montagem += "nove";

                return montagem;
            }
        }

        public int UmaUC(string p1)
        {
            int soma = 0;
            if(p1 == "Monofasico")
            {
                soma += 30;
            }
            else if(p1 == "Bifasico")
            {
                soma += 50;
            }
            else if (p1 == "Trifasico")
            {
                soma += 100;
            }
            return soma;
        }
        public int DuasUC(string p1, string p2)
        {
            int soma = 0;
            switch (p1)
            {
                case "Monofasico":
                    soma += 30;
                    break;
                case "Bifasico":
                    soma += 50;
                    break;
                case "Trifasico":
                    soma += 100;
                    break;
                default:
                    break;
            }
            switch (p2)
            {
                case "Monofasico":
                    soma += 30;
                    break;
                case "Bifasico":
                    soma += 50;
                    break;
                case "Trifasico":
                    soma += 100;
                    break;
                default:
                    break;
            }
            return soma;
        }
        public int TresUC(string p1, string p2, string p3)
        {
            int soma = 0;
            switch (p1)
            {
                case "Monofasico":
                    soma += 30;
                    break;
                case "Bifasico":
                    soma += 50;
                    break;
                case "Trifasico":
                    soma += 100;
                    break;
                default:
                    break;
            }
            switch (p2)
            {
                case "Monofasico":
                    soma += 30;
                    break;
                case "Bifasico":
                    soma += 50;
                    break;
                case "Trifasico":
                    soma += 100;
                    break;
                default:
                    break;
            }
            switch (p3)
            {
                case "Monofasico":
                    soma += 30;
                    break;
                case "Bifasico":
                    soma += 50;
                    break;
                case "Trifasico":
                    soma += 100;
                    break;
                default:
                    break;
            }
            return soma;
        }
        public int QuatroUC(string p1, string p2, string p3, string p4)
        {
            int soma = 0;
            switch (p1)
            {
                case "Monofasico":
                    soma += 30;
                    break;
                case "Bifasico":
                    soma += 50;
                    break;
                case "Trifasico":
                    soma += 100;
                    break;
                default:
                    break;
            }
            switch (p2)
            {
                case "Monofasico":
                    soma += 30;
                    break;
                case "Bifasico":
                    soma += 50;
                    break;
                case "Trifasico":
                    soma += 100;
                    break;
                default:
                    break;
            }
            switch (p3)
            {
                case "Monofasico":
                    soma += 30;
                    break;
                case "Bifasico":
                    soma += 50;
                    break;
                case "Trifasico":
                    soma += 100;
                    break;
                default:
                    break;
            }
            switch (p4)
            {
                case "Monofasico":
                    soma += 30;
                    break;
                case "Bifasico":
                    soma += 50;
                    break;
                case "Trifasico":
                    soma += 100;
                    break;
                default:
                    break;
            }
            return soma;
        }
        public int CincoUC(string p1, string p2, string p3, string p4, string p5)
        {
            int soma = 0;
            switch (p1)
            {
                case "Monofasico":
                    soma += 30;
                    break;
                case "Bifasico":
                    soma += 50;
                    break;
                case "Trifasico":
                    soma += 100;
                    break;
                default:
                    break;
            }
            switch (p2)
            {
                case "Monofasico":
                    soma += 30;
                    break;
                case "Bifasico":
                    soma += 50;
                    break;
                case "Trifasico":
                    soma += 100;
                    break;
                default:
                    break;
            }
            switch (p3)
            {
                case "Monofasico":
                    soma += 30;
                    break;
                case "Bifasico":
                    soma += 50;
                    break;
                case "Trifasico":
                    soma += 100;
                    break;
                default:
                    break;
            }
            switch (p4)
            {
                case "Monofasico":
                    soma += 30;
                    break;
                case "Bifasico":
                    soma += 50;
                    break;
                case "Trifasico":
                    soma += 100;
                    break;
                default:
                    break;
            }
            switch (p5)
            {
                case "Monofasico":
                    soma += 30;
                    break;
                case "Bifasico":
                    soma += 50;
                    break;
                case "Trifasico":
                    soma += 100;
                    break;
                default:
                    break;
            }
            return soma;
        }
        public int SeisUC(string p1, string p2, string p3, string p4, string p5, string p6)
        {
            int soma = 0;
            switch (p1)
            {
                case "Monofasico":
                    soma += 30;
                    break;
                case "Bifasico":
                    soma += 50;
                    break;
                case "Trifasico":
                    soma += 100;
                    break;
                default:
                    break;
            }
            switch (p2)
            {
                case "Monofasico":
                    soma += 30;
                    break;
                case "Bifasico":
                    soma += 50;
                    break;
                case "Trifasico":
                    soma += 100;
                    break;
                default:
                    break;
            }
            switch (p3)
            {
                case "Monofasico":
                    soma += 30;
                    break;
                case "Bifasico":
                    soma += 50;
                    break;
                case "Trifasico":
                    soma += 100;
                    break;
                default:
                    break;
            }
            switch (p4)
            {
                case "Monofasico":
                    soma += 30;
                    break;
                case "Bifasico":
                    soma += 50;
                    break;
                case "Trifasico":
                    soma += 100;
                    break;
                default:
                    break;
            }
            switch (p5)
            {
                case "Monofasico":
                    soma += 30;
                    break;
                case "Bifasico":
                    soma += 50;
                    break;
                case "Trifasico":
                    soma += 100;
                    break;
                default:
                    break;
            }
            switch (p6)
            {
                case "Monofasico":
                    soma += 30;
                    break;
                case "Bifasico":
                    soma += 50;
                    break;
                case "Trifasico":
                    soma += 100;
                    break;
                default:
                    break;
            }
            return soma;
        }

        public double Zeradisp(double mes, string padrao)
        {
            double disp=0;
            if(padrao == "Monofasico")
            {
                disp = 30;
            }
            if (padrao == "Bifasico")
            {
                disp = 50;
            }
            if (padrao == "Trifasico")
            {
                disp = 100;
            }

            if((mes-disp)<0)
            {
                mes = 0;
            }
            else
            {
                mes = mes - disp;
            }
            return mes;
        }
        public double CalculaTarifa(double tarifa, double jan, double fev, double mar, double abr, double mai, double jun, double jul, double ago, double set, double outu, double nov, double dez)
        {
            double resultado = 0;

            resultado += jan * tarifa;
            resultado += fev * tarifa;
            resultado += mar * tarifa;
            resultado += abr * tarifa;
            resultado += mai * tarifa;
            resultado += jun * tarifa;
            resultado += jul * tarifa;
            resultado += ago * tarifa;
            resultado += set * tarifa;
            resultado += outu * tarifa;
            resultado += nov * tarifa;
            resultado += dez * tarifa;

            return resultado;
        }
        public double CalculaConsumoTotal(double jan, double fev, double mar, double abr, double mai, double jun, double jul, double ago, double set, double outu, double nov, double dez)
        {
            double total = 0;
            total += jan;
            total += fev;
            total += mar;
            total += abr;
            total += mai;
            total += jun;
            total += jul;
            total += ago;
            total += set;
            total += outu;
            total += nov;
            total += dez;

            return total;
        }
        public double CalculaConsumoComTarifa(double jan, double fev, double mar, double abr, double mai, double jun, double jul, double ago, double set, double outu, double nov, double dez, double tarifa)
        {
            double total = 0;
            total += jan * tarifa;
            total += fev * tarifa;
            total += mar * tarifa;
            total += abr * tarifa;
            total += mai * tarifa;
            total += jun * tarifa;
            total += jul * tarifa;
            total += ago * tarifa;
            total += set * tarifa;
            total += outu * tarifa;
            total += nov * tarifa;
            total += dez * tarifa;

            return total;
        }
        public double CalculaPayback(double tarifao, double custo, double valorsistema, double gerano)
        {
            double payback=0;
            //Payback
            double tarifa = tarifao;

            //Novas
            double CustoInversor = (custo == 0) ? CustoInversor = 0 : CustoInversor = custo;
            int i;
            double porcento = 0.0833333333333333333333;
            double primeiro = (gerano * tarifa) / 12;
            double caixaacumulado = valorsistema + CustoInversor;
            double caixaatual = 0;

            if (caixaatual < caixaacumulado)
            {
                //Primeiro
                for (i = 0; i < 12; i++)
                {
                    if (caixaatual < caixaacumulado)
                    {
                        caixaatual += primeiro;
                        porcento = porcento + 0.0833333333333333333333;
                    }
                    else
                    {
                        payback = porcento - 1;
                        break;
                    }
                }
                if (caixaatual < caixaacumulado)
                {
                    tarifa *= 1.1;
                    double segundo = ((gerano * 0.99) * tarifa) / 12;
                    //Segundo
                    for (i = 0; i < 12; i++)
                    {
                        if (caixaatual < caixaacumulado)
                        {
                            caixaatual += segundo;
                            porcento = porcento + 0.0833333333333333333333;
                        }
                        else
                        {
                            payback = porcento - 1;
                            break;
                        }
                    }
                    if (caixaatual < caixaacumulado)
                    {
                        tarifa *= 1.1;
                        double terceiro = ((gerano * 0.985) * tarifa) / 12;
                        double taxa = (gerano * 0.99) - (gerano * 0.985);
                        //Terceiro
                        for (i = 0; i < 12; i++)
                        {
                            if (caixaatual < caixaacumulado)
                            {
                                caixaatual += terceiro;
                                porcento = porcento + 0.0833333333333333333333;
                            }
                            else
                            {
                                payback = porcento - 1;
                                break;
                            }
                        }
                        if (caixaatual < caixaacumulado)
                        {
                            tarifa *= 1.1;
                            double quatro = ((gerano * 0.985) - taxa);
                            double quarto = (((gerano * 0.985) - taxa) * tarifa) / 12;
                            //Quarto
                            for (i = 0; i < 12; i++)
                            {
                                if (caixaatual < caixaacumulado)
                                {
                                    caixaatual += quarto;
                                    porcento = porcento + 0.0833333333333333333333;
                                }
                                else
                                {
                                    payback = porcento - 1;
                                    break;
                                }
                            }
                            if (caixaatual < caixaacumulado)
                            {
                                tarifa *= 1.1;
                                double cinco = ((quatro) - taxa);
                                double quinto = (((quatro) - taxa) * tarifa) / 12;
                                //Quinto
                                for (i = 0; i < 12; i++)
                                {
                                    if (caixaatual < caixaacumulado)
                                    {
                                        caixaatual += quinto;
                                        porcento = porcento + 0.0833333333333333333333;
                                    }
                                    else
                                    {
                                        payback = porcento - 1;
                                        break;
                                    }
                                }
                                if (caixaatual < caixaacumulado)
                                {
                                    tarifa *= 1.1;
                                    double seis = ((cinco) - taxa);
                                    double sexto = (((cinco) - taxa) * tarifa) / 12;
                                    //Sexto
                                    for (i = 0; i < 12; i++)
                                    {
                                        if (caixaatual < caixaacumulado)
                                        {
                                            caixaatual += sexto;
                                            porcento = porcento + 0.0833333333333333333333;
                                        }
                                        else
                                        {
                                            payback = porcento - 1;
                                            break;
                                        }
                                    }
                                    if (caixaatual < caixaacumulado)
                                    {
                                        tarifa *= 1.1;
                                        double sete = ((seis) - taxa);
                                        double setimo = (((seis) - taxa) * tarifa) / 12;
                                        //Setimo
                                        for (i = 0; i < 12; i++)
                                        {
                                            if (caixaatual < caixaacumulado)
                                            {
                                                caixaatual += setimo;
                                                porcento = porcento + 0.0833333333333333333333;
                                            }
                                            else
                                            {
                                                payback = porcento - 1;
                                                break;
                                            }
                                        }
                                        if (caixaatual < caixaacumulado)
                                        {
                                            tarifa *= 1.1;
                                            double oito = ((sete) - taxa);
                                            double oitavo = (((sete) - taxa) * tarifa) / 12;
                                            //Oitavo
                                            for (i = 0; i < 12; i++)
                                            {
                                                if (caixaatual < caixaacumulado)
                                                {
                                                    caixaatual += oitavo;
                                                    porcento = porcento + 0.0833333333333333333333;
                                                }
                                                else
                                                {
                                                    payback = porcento - 1;
                                                    break;
                                                }
                                            }
                                            if (caixaatual < caixaacumulado)
                                            {
                                                tarifa *= 1.1;
                                                double nove = ((oito) - taxa);
                                                double nono = (((oito) - taxa) * tarifa) / 12;
                                                //Nono
                                                for (i = 0; i < 12; i++)
                                                {
                                                    if (caixaatual < caixaacumulado)
                                                    {
                                                        caixaatual += nono;
                                                        porcento = porcento + 0.0833333333333333333333;
                                                    }
                                                    else
                                                    {
                                                        payback = porcento - 1;
                                                        break;
                                                    }
                                                }
                                                if (caixaatual < caixaacumulado)
                                                {
                                                    tarifa *= 1.1;
                                                    double dez = ((nove) - taxa);
                                                    double decimo = (((nove) - taxa) * tarifa) / 12;
                                                    //Decimo
                                                    for (i = 0; i < 12; i++)
                                                    {
                                                        if (caixaatual < caixaacumulado)
                                                        {
                                                            caixaatual += decimo;
                                                            porcento = porcento + 0.0833333333333333333333;
                                                        }
                                                        else
                                                        {
                                                            payback = porcento - 1;
                                                            break;
                                                        }
                                                    }
                                                    if (caixaatual < caixaacumulado)
                                                    {
                                                        tarifa *= 1.1;
                                                        double oonze = ((dez) - taxa);
                                                        double onze = (((dez) - taxa) * tarifa) / 12;
                                                        //Onze
                                                        for (i = 0; i < 12; i++)
                                                        {
                                                            if (caixaatual < caixaacumulado)
                                                            {
                                                                caixaatual += onze;
                                                                porcento = porcento + 0.0833333333333333333333;
                                                            }
                                                            else
                                                            {
                                                                payback = porcento - 1;
                                                                break;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return payback;
        }
        public void AcharESubstituir(Word.Application wordApp, object ToFindText, object ReplaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundLike = false;
            object matchAllforms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiactitics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref ToFindText,
                ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundLike,
                ref matchAllforms, ref forward,
                ref wrap, ref format, ref ReplaceWithText,
                ref replace, ref matchKashida,
                ref matchDiactitics, ref matchAlefHamza,
                ref matchControl);
        }

    }
}
