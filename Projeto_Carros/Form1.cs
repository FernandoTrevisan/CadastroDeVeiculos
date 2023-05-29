using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using Org.BouncyCastle.Crypto.Engines;
using System.IO;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;
using System.Xml.Serialization;

namespace Projeto_Carros
{
    public partial class Form1 : Form
    {
        private List<Banco> Carros = new List<Banco>();

        public Form1()
        {
            InitializeComponent();
        }

        public void ExportarDados(DataGridView dgv_Carros)
        {
            Microsoft.Office.Interop.Excel.Application exportarexcel = new Microsoft.Office.Interop.Excel.Application();

            exportarexcel.Application.Workbooks.Add(true);

            int indicecolumn = 0;

            foreach (DataGridViewColumn column in dgv_Carros.Columns)
            {
                indicecolumn++;

                exportarexcel.Cells[1, indicecolumn] = column.Name;
            }

            int indicefila = 0;

            foreach (DataGridViewRow fila in dgv_Carros.Rows)
            {
                indicefila++;

                indicecolumn = 0;

                foreach (DataGridViewColumn column in dgv_Carros.Columns)
                {
                    indicecolumn++;
                    exportarexcel.Cells[indicefila + 1, indicecolumn] = fila.Cells[column.Name].Value;
                }
            }

            exportarexcel.Visible = true;
        }

        int id = 1;
        private void btnCadastrar_Click(object sender, EventArgs e)
        {
            dgv_carros.Rows.Add(id,
            txt_marca.Text,
            txtModelo.Text,
            cmb_Fabricante.Text,
            cmbTipo.Text,
            txtAno.Text,
            cmb_combustivel.Text,
            cmb_Cor.Text,
            txtChassi.Text,
            txtKilometragem.Text,
            txtObservacoes.Text,
            chk_revisao.Checked,
            chk_sinistro.Checked,
            chk_Roubo.Checked,
            chk_Aluguel.Checked,
            chk_Venda.Checked,
            chk_Particular.Checked);



            
            string marca = txt_marca.Text;
            string modelo = txtModelo.Text;
            string fabricante = cmb_Fabricante.Text;
            string tipo = cmbTipo.Text;
            int ano = int.Parse(txtAno.Text);
            string combustivel = cmb_combustivel.Text;
            string cor = cmb_Cor.Text;
            string numero_chassi = txtChassi.Text;
            int kilometragem = int.Parse(txtKilometragem.Text);
            string observacoes = txtObservacoes.Text;



            id++;

            txtAno.Clear();
            txtChassi.Clear();
            cmb_combustivel.Text = "";
            cmb_Cor.Text = "";
            cmb_Fabricante.Text = "";
            txtKilometragem.Clear();
            txtModelo.Clear();
            txtObservacoes.Clear();
            cmbTipo.Text = "";
            txt_marca.Clear();






            chk_Aluguel.Checked = false;
            chk_Particular.Checked = false;
            chk_revisao.Checked = false;
            chk_sinistro.Checked = false;
            chk_Venda.Checked = false;
            chk_Roubo.Checked = false;

            txtCodigo.Text = $"{id.ToString()}";
        }

        private void btnFechar_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnRemover_Click(object sender, EventArgs e)
        {
            int row = dgv_carros.CurrentCell.RowIndex;
            if (dgv_carros.CurrentCell.RowIndex == -1)
            {
                MessageBox.Show("a", "a", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
            else

                dgv_carros.Rows.RemoveAt(row);
        }

        private void btnExportar_Click(object sender, EventArgs e)
        {
            ExportarDados(dgv_carros);
        }

     
     
        private void txtAno_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != (char)8)
            {
                { e.Handled = true; }
            }
        }

        private void txtKilometragem_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != (char)8)
            {
                { e.Handled = true; }
            }
        }



        private void cmb_marca_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != (char)8)
            {
                { e.Handled = true; }
            }
        }

        private void txtCodigo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != (char)8)
            {
                { e.Handled = true; }
            }

        }

        private void txtModelo_KeyPress(object sender, KeyPressEventArgs e)
        {
           if (Char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void txtTipo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void txtObservacoes_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void txt_marca_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            txtCodigo.Text = id.ToString();

            ++id;
        }

     
        private void btnEnviar_Click(object sender, EventArgs e)
        {

            abrirPlanilhaExcel();

        }

        private void abrirPlanilhaExcel()
        {
           
            string folder = @"D:\carros\backup";

            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }

            //File.Create(folder + @"\dump.sql");
            
            

            /*
             
            var arquivo = @"D:\planilha1.xls";

            var planilha = "SELECT * FROM [carros$]";

            var strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" + arquivo + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=1\"";

            var dt = new DataTable();

            using(OleDbConnection con = new OleDbConnection(strCon))
            {
                using (OleDbDataAdapter da = new OleDbDataAdapter(planilha, con))
                {
                    da.Fill(dt);
                    dgv_carros.DataSource = dt;
                }
            }
             
             */
        }

        private void btnEnviar_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Arquivos XML (*.xml)|*.xml";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;

                try
                {
                    XmlSerializer serializer = new XmlSerializer(typeof(List<Banco>));
                    using (StreamReader reader = new StreamReader(filePath))
                    {
                        List<Banco> carrosBackup = (List<Banco>)serializer.Deserialize(reader);

                        int maxID = Carros.Any() ? Carros.Max(c => c.Codigo) : 0;


                       dgv_carros.Rows.Clear();


                        foreach (Banco carro in carrosBackup)
                        {

                            maxID++;
                            carro.Codigo = maxID;

                            DataGridViewRow novaLinha = new DataGridViewRow();
                            novaLinha.CreateCells(dgv_carros);
                            novaLinha.Cells[0].Value = carro.Codigo;
                            novaLinha.Cells[1].Value = carro.Marca;
                            novaLinha.Cells[2].Value = carro.Modelo;
                            novaLinha.Cells[3].Value = carro.Fabricante;
                            novaLinha.Cells[4].Value = carro.Tipo;
                            novaLinha.Cells[5].Value = carro.Ano;
                            novaLinha.Cells[6].Value = carro.Combustivel.ToString();
                            novaLinha.Cells[7].Value = carro.Cor;
                            novaLinha.Cells[8].Value = carro.Chassi.ToString();
                            novaLinha.Cells[9].Value = carro.Km.ToString();
                            novaLinha.Cells[10].Value = carro.Revisão.ToString();
                            novaLinha.Cells[11].Value = carro.Sinistro.ToString();
                            novaLinha.Cells[12].Value = carro.Roubo_Furto.ToString();
                            novaLinha.Cells[13].Value = carro.Aluguel.ToString();
                            novaLinha.Cells[14].Value = carro.Venda.ToString();
                            novaLinha.Cells[15].Value = carro.Particular.ToString();
                            novaLinha.Cells[16].Value = carro.Observacoes;

                            dgv_carros.Rows.Add(novaLinha);
                            
                        }

                        MessageBox.Show("Backup carregado com sucesso.", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao carregar o backup: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnBackup_Click(object sender, EventArgs e)
        {
            List<Banco> CarrosBackup = new List<Banco>();

            foreach (DataGridViewRow linha in dgv_carros.Rows)
            {
                if (!linha.IsNewRow)
                {

                    Banco carro = new Banco();
                    carro.Marca = linha.Cells["Marca"].Value?.ToString();
                    carro.Modelo = linha.Cells["Modelo"].Value?.ToString();
                    carro.Fabricante = linha.Cells["Fabricante"].Value?.ToString();
                    carro.Tipo = linha.Cells["Tipo"].Value?.ToString();
                    carro.Ano = Convert.ToInt32(linha.Cells["Ano"].Value);
                    carro.Combustivel = linha.Cells["Combustivel"].Value?.ToString();
                    carro.Cor = linha.Cells["Cor"].Value?.ToString();
                    carro.Chassi = Convert.ToInt64(linha.Cells["Chassi"].Value);
                    carro.Km = Convert.ToInt32(linha.Cells["KIlometragem"].Value);
                    carro.Revisão = Convert.ToBoolean(linha.Cells["Revisão"].Value).ToString();
                    carro.Sinistro = Convert.ToBoolean(linha.Cells["Sinistro"].Value).ToString();
                    carro.Roubo_Furto = Convert.ToBoolean(linha.Cells["Roubo"].Value).ToString();
                    carro.Aluguel = linha.Cells["Aluguel"].Value?.ToString();
                    carro.Venda = Convert.ToBoolean(linha.Cells["Venda"].Value).ToString();
                    carro.Particular = Convert.ToBoolean(linha.Cells["Particular"].Value).ToString();
                    carro.Observacoes = linha.Cells["Observações"].Value?.ToString();

                    CarrosBackup.Add(carro);
                }
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Arquivos XML|*.xml";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog.FileName;

                    using (var streamWriter = new StreamWriter(filePath))
                    {
                        var serializer = new XmlSerializer(typeof(List<Banco>));
                        serializer.Serialize(streamWriter, CarrosBackup);
                    }

                    MessageBox.Show("Backup salvo com sucesso.", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }

        }

        private void btnPasta_Click(object sender, EventArgs e)
        {

            string folder = @"D:\Carros001\backup";

            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }




        }
    }
}