using System;
using System.Data;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace Projeto_Carros
{

    [Serializable]

    public class Banco
    {

        public int Codigo { get; set; }
        public string Marca { get; set; }
        public string Modelo { get; set; }
        public string Fabricante { get; set; }
        public string Tipo { get; set; }
        public int Ano { get; set; }
        public string Combustivel { get; set; }
        public string Cor { get; set; }
        public long Chassi { get; set; }
        public int Km { get; set; }
        public string Revisão { get; set; }
        public string Sinistro { get; set; }
        public string Roubo_Furto { get; set; }
        public string Aluguel { get; set; }
        public string Venda { get; set; }
        public string Particular { get; set; }
        public string Observacoes { get; set; }

    }
}

