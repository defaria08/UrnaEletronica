using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Npgsql;
using System.IO;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.Diagnostics;
using System.Configuration;
using System.Security.Cryptography;

namespace UrnaDeMaria
{
    public partial class Urna : Form
    {
        private string numero = "";
        private int deputadoe;
        private int deputadof;
        private int governador;
        private int presidente;
        private int contador;
        private bool pres;
        private bool gov;


        string ncandidato = "";
        public Urna()
        {
            InitializeComponent();
            IniciaVotacao();

        }

        private void btn_Click(object sender, EventArgs e)
        {

            //ADICIONO O VALOR DO BOTAO SELECIONADO AO LABEL QUE COMPOE O NUMERO DO CANDIDATO
            string nomeBotao = ((Button)sender).Text;

            if (lDigit0.Text == "  ")
            {
                lDigit0.Text = nomeBotao;
                ncandidato += nomeBotao;
            }
            else if (lDigit1.Text == "  ")
            {
                lDigit1.Text = nomeBotao;
                ncandidato += nomeBotao;
                if (contador == 2)
                {
                    gov = true;
                    CarregaCandidato();

                }
                else if (contador == 3)
                {
                    pres = true;
                    CarregaCandidato();

                }
            }
            else if (lDigit2.Text == "  ")
            {
                lDigit2.Text = nomeBotao;
                ncandidato += nomeBotao;
            }
            else if (lDigit3.Text == "  ")
            {
                lDigit3.Text = nomeBotao;
                ncandidato += nomeBotao;
                if (contador == 1)
                {
                    CarregaCandidato();
                }
            }
            else if (lDigit4.Text == "  ")
            {
                lDigit4.Text = nomeBotao;

                ncandidato += nomeBotao;
                CarregaCandidato();
            }

        }

        public void CarregaCandidato()
        {
            //LEITURA DO CANDIDATO NO BANCO PARA APRESENTAR NO PAINEL DA URNA
            lNome.Visible = true;
            lPartido.Visible = true;
            PBFoto.Visible = true;
            Uteis.DAL data = new Uteis.DAL();

            int neleitoral = int.Parse(ncandidato);
            string comando;
            if (pres is true)
                comando = "select * from candidato where id = " + neleitoral + " and cargo = 'PRESIDENTE'";
            else if (gov is true)
                comando = "select * from candidato where id = " + neleitoral + "and cargo = 'GOVERNADOR'";
            else
                comando = "select * from candidato where id = " + neleitoral;
            DataTable dt = data.RetDataTable(comando);

            if (dt.Rows.Count != 0)
            {
                lNome.Text = dt.Rows[0]["nome"].ToString().ToUpper();
                lPartido.Text = dt.Rows[0]["partido"].ToString().ToUpper();
                PBFoto.Load("../../../IMG/" + dt.Rows[0]["foto"].ToString());
            }
            else
            {
                lNome.Text = "Nulo";
                lPartido.Text = "Partido Nulo";
                PBFoto.Load("../../../IMG/nulo.png");
                ncandidato = "0";
            }

        }

        private void BBranco_Click(object sender, EventArgs e)
        {
            //FUNÇÃO QUE ATRIBUI O VALOR DE VOTO EM BRANCO PARA O CANDIDATO
            LimpaDigito();
            lNome.Text = "Branco";
            lPartido.Text = "Branco";
            PBFoto.Load("../../../IMG/branco.png");
            lNome.Visible = true;
            lPartido.Visible = true;
            PBFoto.Visible = true;
            ncandidato = "1";
        }

        private void BCorrige_Click(object sender, EventArgs e)
        {
            //LIMPA OS DADOS
            LimpaDigito();
            lNome.Visible = false;
            lPartido.Visible = false;
            PBFoto.Visible = false;
            ncandidato = "";
        }

        private void BConfirma_Click(object sender, EventArgs e)
        {
            //REALIZO A DIVISAO DE CADA VALOR DE CANDIDATO PARA INSERÇÃO NO BANCO
            if (contador != 4)
            {
                switch (contador)
                {
                    case 0:
                        {

                            deputadoe = int.Parse(ncandidato);
                            lDigit4.Visible = false;
                            LimpaDigito();
                            Lcargo.Text = "DEPUTADO FEDERAL";

                        }
                        break;
                    case 1:
                        {
                            deputadof = int.Parse(ncandidato);
                            lDigit3.Visible = false;
                            lDigit2.Visible = false;
                            LimpaDigito();
                            Lcargo.Text = "GOVERNADOR";

                        }
                        break;
                    case 2:
                        {
                            governador = int.Parse(ncandidato);
                            LimpaDigito();
                            Lcargo.Text = "PRESIDENTE";
                        }
                        break;
                    case 3:
                        {
                            presidente = int.Parse(ncandidato);
                            Uteis.DAL data = new Uteis.DAL();
                            string comando = $"insert into voto (deputadoe,deputadof,governador,presidente) values ({deputadoe},{deputadof},{governador},{presidente})";
                            data.ExecutarComandoSQL(comando);
                            lFim.Visible = true;
                            LimpaDigito();
                            lDigit0.Visible = false;
                            lDigit1.Visible = false;
                            Lcargo.Visible = false;
                            GravaArquivo();
                        }
                        break;



                }
                contador++;
                lNome.Visible = false;
                lPartido.Visible = false;
                PBFoto.Visible = false;
                ncandidato = "";
            }
            else
            {
                IniciaVotacao();
            }


        }

        [Obsolete]
        public void GravaArquivo()
        {
            string arquivo = $"Dep.Federal -{ deputadof}, Dep.Estadual - { deputadoe}, Governador - { governador}, Presidente - { presidente}"; // Editar a primeira linha



            string conteudoTxt, novoConteudoTxt;
            string filtroTipoArquivo = "arquivos txt (*.txt)|*.txt";
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = filtroTipoArquivo;
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = filtroTipoArquivo;

            //###### Primeiro procura o arquivo bruto usando o openFileDialog ######//
            try
            {
                //Pegar o caminho do arquivo escolhido
                string caminhoAbrir = "../../../Arquivos/Test.txt";
                //Jogar o conteúdo do arquivo numa string
                conteudoTxt = File.ReadAllText(caminhoAbrir);

                //Faz as substituições
                novoConteudoTxt = conteudoTxt +"\n"+ arquivo;

                //###### Salva o txt com os valores substituídos onde o usuário escolher ######//
                try
                {
                    string caminhoSalvar = "../../../Arquivos/Test.txt";
                    //Salvar todo o texto no caminho do arquivo escolhido
                    File.WriteAllText(caminhoSalvar, novoConteudoTxt);

                }
                catch (Exception ex) { MessageBox.Show(ex.Message); };
            
            
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); };
        }

        public void IniciaVotacao()
        {
            //INICIAÇÃO DOS CAMPOS 
            ncandidato = "";
            contador = 0;
            Lcargo.Text = "DEPUTADO ESTADUAL";
            Lcargo.Visible = true;
            lFim.Visible = false;
            lNome.Visible = false;
            lPartido.Visible = false;
            PBFoto.Visible = false;
            lDigit0.Visible = true;
            lDigit1.Visible = true;
            lDigit2.Visible = true;
            lDigit3.Visible = true;
            lDigit4.Visible = true;
            deputadoe = 0;
            deputadof = 0;
            governador = 0;
            presidente = 0;
            contador = 0;
            pres = false;
            gov = false;
        }
        public void LimpaDigito()
        {
            //LIMPEZA DOS CAMPOS COM OS NUMEROS DOS CANDIDATOS
            lDigit0.Text = "  ";
            lDigit1.Text = "  ";
            lDigit2.Text = "  ";
            lDigit3.Text = "  ";
            lDigit4.Text = "  ";
        }

        private void Urna_FormClosed(object sender, FormClosedEventArgs e)
        {
            //CRIO O PDF COM O CALCULO DA VOLTAÇÃO

            //ACESSO AO BANDO PARA BUSCA DO RESULTADO
            Uteis.DAL data = new Uteis.DAL();
            string comando;
            DataTable leitor;
            comando = "select count(voto.deputadoe) as nvoto,candidato.nome as ncandidato,candidato.partido as npartido from voto,candidato where voto.deputadoe = candidato.id group by voto.deputadoe,candidato.nome,candidato.partido";
            leitor = data.RetDataTable(comando);
            Resultado result = new Resultado();
            List<Resultado> resultado = new List<Resultado>();

            //CRIO O CABECALHO DO RELATORIOE A FORMATAÇÃO DE PAGINA E LOCAL DE ARMAZENAMENTO
            Document documento = new Document(PageSize.A4);
            documento.SetMargins(3, 2, 3, 2);
            PdfWriter writer = PdfWriter.GetInstance(documento, new FileStream(@"c:\Users\Public\Documents\Relatorio.pdf", FileMode.Create));
            documento.Open();

            PdfPTable table = new PdfPTable(1);

            iTextSharp.text.Font fonte = FontFactory.GetFont(BaseFont.TIMES_ROMAN, 15);

            //-------------Deputado Federal -----------------------------------
            //CRAIACAO DA TABELA DE ACORDO COM O TIPO DO CARGO 
            Paragraph colunacargo = new Paragraph("Número de Votos Deputado Federal", fonte);
            var cell = new PdfPCell();
            cell.AddElement(colunacargo);
            table.AddCell(cell);

            documento.Add(table);
            table = new PdfPTable(3);

            Paragraph colunaqtdvoto = new Paragraph("Número de Votos", fonte);
            Paragraph colunancandidato = new Paragraph("Nome do Candidado", fonte);
            Paragraph colunanpartido = new Paragraph("Nome do Partido", fonte);

            var cell1 = new PdfPCell();
            var cell2 = new PdfPCell();
            var cell3 = new PdfPCell();

            cell1.AddElement(colunaqtdvoto);
            cell2.AddElement(colunancandidato);
            cell3.AddElement(colunanpartido);

            table.AddCell(cell1);
            table.AddCell(cell2);
            table.AddCell(cell3);

            //ATRIBUO O RETORNO DO BANCO PARA UMA LISTA
            for (int i = 0; i < leitor.Rows.Count; i++)
            {
                result = new Resultado();
                result.qtdvoto = leitor.Rows[i]["nvoto"].ToString();
                result.ncandidato = leitor.Rows[i]["ncandidato"].ToString();
                result.npartido = leitor.Rows[i]["npartido"].ToString();

                resultado.Add(result);
            }
            //ADICIONO A LISTA A UM TEMPORARIO PARA ADICIONAR AO RELATORIO
            foreach ( var item in resultado)
            {

                Phrase nvoto = new Phrase(item.qtdvoto);
                var cel = new PdfPCell(nvoto);
                table.AddCell(cel);

                Phrase ncadidato = new Phrase(item.ncandidato);
                var cel1 = new PdfPCell(ncadidato);
                table.AddCell(cel1);

                Phrase npartido = new Phrase(item.npartido);
                var cel2 = new PdfPCell(npartido);
                table.AddCell(cel2);               
            }
            comando = "select count(deputadoe) as nvoto  from voto where deputadoe = 0 ";
            leitor = data.RetDataTable(comando);
            result = new Resultado();
            resultado = new List<Resultado>();

            for (int i = 0; i < leitor.Rows.Count; i++)
            {
                result = new Resultado();
                result.qtdvoto = leitor.Rows[i]["nvoto"].ToString();
                result.ncandidato = "Votos Nulos";
                result.npartido = "--";

                resultado.Add(result);
            }

            foreach (var item in resultado)
            {

                Phrase nvoto = new Phrase(item.qtdvoto);
                var cel = new PdfPCell(nvoto);
                table.AddCell(cel);

                Phrase ncadidato = new Phrase(item.ncandidato);
                var cel1 = new PdfPCell(ncadidato);
                table.AddCell(cel1);

                Phrase npartido = new Phrase(item.npartido);
                var cel2 = new PdfPCell(npartido);
                table.AddCell(cel2);
            }
            comando = "select count(deputadoe) as nvoto  from voto where deputadoe = 1 ";
            leitor = data.RetDataTable(comando);
            result = new Resultado();
            resultado = new List<Resultado>();

            for (int i = 0; i < leitor.Rows.Count; i++)
            {
                result = new Resultado();
                result.qtdvoto = leitor.Rows[i]["nvoto"].ToString();
                result.ncandidato = "Votos em Branco";
                result.npartido = "--";

                resultado.Add(result);
            }

            foreach (var item in resultado)
            {

                Phrase nvoto = new Phrase(item.qtdvoto);
                var cel = new PdfPCell(nvoto);
                table.AddCell(cel);

                Phrase ncadidato = new Phrase(item.ncandidato);
                var cel1 = new PdfPCell(ncadidato);
                table.AddCell(cel1);

                Phrase npartido = new Phrase(item.npartido);
                var cel2 = new PdfPCell(npartido);
                table.AddCell(cel2);
            }

            documento.Add(table);
            //-------------------deputado estadual ------------------------------------------
            comando = "select count(voto.deputadof) as nvoto,candidato.nome as ncandidato,candidato.partido as npartido from voto,candidato where voto.deputadof = candidato.id group by voto.deputadof,candidato.nome,candidato.partido";
            leitor = data.RetDataTable(comando);
            result = new Resultado();
            resultado = new List<Resultado>();

            table = new PdfPTable(1);
            colunacargo = new Paragraph("Número de Votos Deputado Estadual", fonte);
            cell = new PdfPCell();
            cell.AddElement(colunacargo);
            table.AddCell(cell);

            documento.Add(table);
            table = new PdfPTable(3);

            colunaqtdvoto = new Paragraph("Número de Votos", fonte);
            colunancandidato = new Paragraph("Nome do Candidado", fonte);
            colunanpartido = new Paragraph("Nome do Partido", fonte);

            cell1 = new PdfPCell();
            cell2 = new PdfPCell();
            cell3 = new PdfPCell();

            cell1.AddElement(colunaqtdvoto);
            cell2.AddElement(colunancandidato);
            cell3.AddElement(colunanpartido);

            table.AddCell(cell1);
            table.AddCell(cell2);
            table.AddCell(cell3);


            for (int i = 0; i < leitor.Rows.Count; i++)
            {
                result = new Resultado();
                result.qtdvoto = leitor.Rows[i]["nvoto"].ToString();
                result.ncandidato = leitor.Rows[i]["ncandidato"].ToString();
                result.npartido = leitor.Rows[i]["npartido"].ToString();

                resultado.Add(result);
            }

            foreach (var item in resultado)
            {

                Phrase nvoto = new Phrase(item.qtdvoto);
                var cel = new PdfPCell(nvoto);
                table.AddCell(cel);

                Phrase ncadidato = new Phrase(item.ncandidato);
                var cel1 = new PdfPCell(ncadidato);
                table.AddCell(cel1);

                Phrase npartido = new Phrase(item.npartido);
                var cel2 = new PdfPCell(npartido);
                table.AddCell(cel2);
            }
            comando = "select count(deputadoe) as nvoto  from voto where deputadoe = 0 ";
            leitor = data.RetDataTable(comando);
            result = new Resultado();
            resultado = new List<Resultado>();

            for (int i = 0; i < leitor.Rows.Count; i++)
            {
                result = new Resultado();
                result.qtdvoto = leitor.Rows[i]["nvoto"].ToString();
                result.ncandidato = "Votos Nulos";
                result.npartido = "--";

                resultado.Add(result);
            }

            foreach (var item in resultado)
            {

                Phrase nvoto = new Phrase(item.qtdvoto);
                var cel = new PdfPCell(nvoto);
                table.AddCell(cel);

                Phrase ncadidato = new Phrase(item.ncandidato);
                var cel1 = new PdfPCell(ncadidato);
                table.AddCell(cel1);

                Phrase npartido = new Phrase(item.npartido);
                var cel2 = new PdfPCell(npartido);
                table.AddCell(cel2);
            }
            comando = "select count(deputadoe) as nvoto  from voto where deputadoe = 1 ";
            leitor = data.RetDataTable(comando);
            result = new Resultado();
            resultado = new List<Resultado>();

            for (int i = 0; i < leitor.Rows.Count; i++)
            {
                result = new Resultado();
                result.qtdvoto = leitor.Rows[i]["nvoto"].ToString();
                result.ncandidato = "Votos em Branco";
                result.npartido = "--";

                resultado.Add(result);
            }

            foreach (var item in resultado)
            {

                Phrase nvoto = new Phrase(item.qtdvoto);
                var cel = new PdfPCell(nvoto);
                table.AddCell(cel);

                Phrase ncadidato = new Phrase(item.ncandidato);
                var cel1 = new PdfPCell(ncadidato);
                table.AddCell(cel1);

                Phrase npartido = new Phrase(item.npartido);
                var cel2 = new PdfPCell(npartido);
                table.AddCell(cel2);
            }
            documento.Add(table);

            //----------------------- Governador ---------------------------
            comando = "select count(voto.governador) as nvoto,candidato.nome as ncandidato,candidato.partido as npartido from voto,candidato where voto.governador = candidato.id group by voto.governador,candidato.nome,candidato.partido";
            leitor = data.RetDataTable(comando);
            result = new Resultado();
            resultado = new List<Resultado>();

            table = new PdfPTable(1);
            colunacargo = new Paragraph("Número de Votos Governador", fonte);
            cell = new PdfPCell();
            cell.AddElement(colunacargo);
            table.AddCell(cell);

            documento.Add(table);
            table = new PdfPTable(3);

            colunaqtdvoto = new Paragraph("Número de Votos", fonte);
            colunancandidato = new Paragraph("Nome do Candidado", fonte);
            colunanpartido = new Paragraph("Nome do Partido", fonte);

            cell1 = new PdfPCell();
            cell2 = new PdfPCell();
            cell3 = new PdfPCell();

            cell1.AddElement(colunaqtdvoto);
            cell2.AddElement(colunancandidato);
            cell3.AddElement(colunanpartido);

            table.AddCell(cell1);
            table.AddCell(cell2);
            table.AddCell(cell3);


            for (int i = 0; i < leitor.Rows.Count; i++)
            {
                result = new Resultado();
                result.qtdvoto = leitor.Rows[i]["nvoto"].ToString();
                result.ncandidato = leitor.Rows[i]["ncandidato"].ToString();
                result.npartido = leitor.Rows[i]["npartido"].ToString();

                resultado.Add(result);
            }

            foreach (var item in resultado)
            {

                Phrase nvoto = new Phrase(item.qtdvoto);
                var cel = new PdfPCell(nvoto);
                table.AddCell(cel);

                Phrase ncadidato = new Phrase(item.ncandidato);
                var cel1 = new PdfPCell(ncadidato);
                table.AddCell(cel1);

                Phrase npartido = new Phrase(item.npartido);
                var cel2 = new PdfPCell(npartido);
                table.AddCell(cel2);
            }
            comando = "select count(governador) as nvoto  from voto where governador = 0 ";
            leitor = data.RetDataTable(comando);
            result = new Resultado();
            resultado = new List<Resultado>();

            for (int i = 0; i < leitor.Rows.Count; i++)
            {
                result = new Resultado();
                result.qtdvoto = leitor.Rows[i]["nvoto"].ToString();
                result.ncandidato = "Votos Nulos";
                result.npartido = "--";

                resultado.Add(result);
            }

            foreach (var item in resultado)
            {

                Phrase nvoto = new Phrase(item.qtdvoto);
                var cel = new PdfPCell(nvoto);
                table.AddCell(cel);

                Phrase ncadidato = new Phrase(item.ncandidato);
                var cel1 = new PdfPCell(ncadidato);
                table.AddCell(cel1);

                Phrase npartido = new Phrase(item.npartido);
                var cel2 = new PdfPCell(npartido);
                table.AddCell(cel2);
            }
            comando = "select count(governador) as nvoto  from voto where governador = 1 ";
            leitor = data.RetDataTable(comando);
            result = new Resultado();
            resultado = new List<Resultado>();

            for (int i = 0; i < leitor.Rows.Count; i++)
            {
                result = new Resultado();
                result.qtdvoto = leitor.Rows[i]["nvoto"].ToString();
                result.ncandidato = "Votos em Branco";
                result.npartido = "--";

                resultado.Add(result);
            }

            foreach (var item in resultado)
            {

                Phrase nvoto = new Phrase(item.qtdvoto);
                var cel = new PdfPCell(nvoto);
                table.AddCell(cel);

                Phrase ncadidato = new Phrase(item.ncandidato);
                var cel1 = new PdfPCell(ncadidato);
                table.AddCell(cel1);

                Phrase npartido = new Phrase(item.npartido);
                var cel2 = new PdfPCell(npartido);
                table.AddCell(cel2);
            }
            documento.Add(table);

            //----------------------- Presidente ---------------------------
            comando = "select count(voto.presidente) as nvoto,candidato.nome as ncandidato,candidato.partido as npartido from voto,candidato where voto.presidente = candidato.id group by voto.presidente,candidato.nome,candidato.partido";
            leitor = data.RetDataTable(comando);
            result = new Resultado();
            resultado = new List<Resultado>();

            table = new PdfPTable(1);
            colunacargo = new Paragraph("Número de Votos Presidente", fonte);
            cell = new PdfPCell();
            cell.AddElement(colunacargo);
            table.AddCell(cell);

            documento.Add(table);
            table = new PdfPTable(3);

            colunaqtdvoto = new Paragraph("Número de Votos", fonte);
            colunancandidato = new Paragraph("Nome do Candidado", fonte);
            colunanpartido = new Paragraph("Nome do Partido", fonte);

            cell1 = new PdfPCell();
            cell2 = new PdfPCell();
            cell3 = new PdfPCell();

            cell1.AddElement(colunaqtdvoto);
            cell2.AddElement(colunancandidato);
            cell3.AddElement(colunanpartido);

            table.AddCell(cell1);
            table.AddCell(cell2);
            table.AddCell(cell3);


            for (int i = 0; i < leitor.Rows.Count; i++)
            {
                result = new Resultado();
                result.qtdvoto = leitor.Rows[i]["nvoto"].ToString();
                result.ncandidato = leitor.Rows[i]["ncandidato"].ToString();
                result.npartido = leitor.Rows[i]["npartido"].ToString();

                resultado.Add(result);
            }

            foreach (var item in resultado)
            {

                Phrase nvoto = new Phrase(item.qtdvoto);
                var cel = new PdfPCell(nvoto);
                table.AddCell(cel);

                Phrase ncadidato = new Phrase(item.ncandidato);
                var cel1 = new PdfPCell(ncadidato);
                table.AddCell(cel1);

                Phrase npartido = new Phrase(item.npartido);
                var cel2 = new PdfPCell(npartido);
                table.AddCell(cel2);
            }
            comando = "select count(presidente) as nvoto  from voto where presidente = 0 ";
            leitor = data.RetDataTable(comando);
            result = new Resultado();
            resultado = new List<Resultado>();

            for (int i = 0; i < leitor.Rows.Count; i++)
            {
                result = new Resultado();
                result.qtdvoto = leitor.Rows[i]["nvoto"].ToString();
                result.ncandidato = "Votos Nulos";
                result.npartido = "--";

                resultado.Add(result);
            }

            foreach (var item in resultado)
            {

                Phrase nvoto = new Phrase(item.qtdvoto);
                var cel = new PdfPCell(nvoto);
                table.AddCell(cel);

                Phrase ncadidato = new Phrase(item.ncandidato);
                var cel1 = new PdfPCell(ncadidato);
                table.AddCell(cel1);

                Phrase npartido = new Phrase(item.npartido);
                var cel2 = new PdfPCell(npartido);
                table.AddCell(cel2);
            }
            comando = "select count(presidente) as nvoto  from voto where presidente = 1 ";
            leitor = data.RetDataTable(comando);
            result = new Resultado();
            resultado = new List<Resultado>();

            for (int i = 0; i < leitor.Rows.Count; i++)
            {
                result = new Resultado();
                result.qtdvoto = leitor.Rows[i]["nvoto"].ToString();
                result.ncandidato = "Votos em Branco";
                result.npartido = "--";

                resultado.Add(result);
            }

            foreach (var item in resultado)
            {

                Phrase nvoto = new Phrase(item.qtdvoto);
                var cel = new PdfPCell(nvoto);
                table.AddCell(cel);

                Phrase ncadidato = new Phrase(item.ncandidato);
                var cel1 = new PdfPCell(ncadidato);
                table.AddCell(cel1);

                Phrase npartido = new Phrase(item.npartido);
                var cel2 = new PdfPCell(npartido);
                table.AddCell(cel2);
            }
            documento.Add(table);
            documento.Close();


            //ABRE O RELATORIO GERADO
            var p = new Process();
            p.StartInfo = new ProcessStartInfo(@"C:\Users\Public\Documents\Relatorio.pdf")
            {
                UseShellExecute = true
            };
            p.Start();

        }
        
    }
}
