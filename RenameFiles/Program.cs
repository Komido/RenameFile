using iTextSharp.text.pdf.parser;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.IO;
using System.Text;

namespace RenameFiles
{
    class Program
    {
        static void Main(string[] args)
        {
            var CurrentDirectory = Directory.GetCurrentDirectory();
            Console.WriteLine(CurrentDirectory);

            DirectoryInfo d = new DirectoryInfo(CurrentDirectory + "\\files");
            FileInfo[] files = d.GetFiles();
            string planilha = "";

            string aux = "";

            foreach (FileInfo item in files)
            {
                if ((item.Name.Split('.')[1] == "xlsx" || item.Name.Split('.')[1] == "xls") && !item.Name.Contains("~$"))
                {
                    string name = item.Name.Split('.')[0];
                    if (!Directory.Exists(CurrentDirectory + "\\files\\" + item.Name.Split('.')[0]))
                    {
                        Directory.CreateDirectory(CurrentDirectory + "\\files\\" + item.Name.Split('.')[0]);
                    }
                    planilha = item.Name;
                }
            }

            var connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source={CurrentDirectory + "\\files" + "\\" + planilha}; Extended Properties=""Excel 12.0 Xml;HDR=YES""";

            StringBuilder stbQuery = new StringBuilder();
            stbQuery.Append("SELECT * FROM [REMESSA$]");// nome da planilha no arquivo
            OleDbDataAdapter adp = new OleDbDataAdapter(stbQuery.ToString(), connectionString);

            DataTable dtEmpresas = new DataTable();
            adp.Fill(dtEmpresas);

            TirarAcento retirarAcento = new TirarAcento();

            foreach (DataColumn item in dtEmpresas.Columns)
            {
                item.ColumnName = retirarAcento.Remove(item.ColumnName.Replace(" ", "_"));
            }

            string apaga = JsonConvert.SerializeObject(dtEmpresas);

            List<Remessa> json = JsonConvert.DeserializeObject<List<Remessa>>(JsonConvert.SerializeObject(dtEmpresas));

            string[] boleto = null;

            foreach (var file in files)
            {
                if (file.Name.Split('.')[1] == "pdf")
                {

                    using (iTextSharp.text.pdf.PdfReader leitor = new iTextSharp.text.pdf.PdfReader(file.FullName))
                    {
                        StringBuilder texto = new StringBuilder();

                        for (int i = 1; i <= leitor.NumberOfPages; i++)
                        {
                            texto.Append(PdfTextExtractor.GetTextFromPage(leitor, i));
                        }
                        // Esta linha pode mudar de posição
                        string linhaCNPJ = texto.ToString().Split('\n')[26];
                        string documento = linhaCNPJ.Split(' ')[(texto.ToString().Split('\n')[26].Split(' ')).Length - 1];
                        //string documento = texto.ToString().Split('\n')[26].Split(' ')[(texto.ToString().Split('\n')[26].Split(' ')).Length - 1];


                        Remessa remessa = json.Find(element => element.CPF_CNPJ == documento);

                        leitor.Close();

                        DateTime mes = DateTime.Now.AddMonths(-1);
                        string mesRemessa = retirarAcento.Remove(mes.ToString("MMMM").ToUpper());

                        File.Move(file.FullName, file.FullName.Replace(file.Name, "\\" + planilha.Split('.')[0] + "\\BL_" + retirarAcento.Remove(remessa.NOME_DO_PAGADOR.Replace(".", "")) + "_" + retirarAcento.RemoveMascara(remessa.CPF_CNPJ) + "_" + mesRemessa + ".pdf"));
                    }
                }
            }
        }


    }

    public class Remessa
    {
        public float CODIGO { get; set; }
        public DateTime DT_VENC { get; set; }
        public float VALOR { get; set; }
        public string MENSAGEM { get; set; }
        public string ESPECIE_DOC { get; set; }
        public string ESPECIE { get; set; }
        public string ACEITE { get; set; }
        public float NUMDOCUMENTO { get; set; }
        public DateTime DT_PROCESSAMENTO { get; set; }
        public string INSTRUCOES { get; set; }
        public float MULTA { get; set; }
        public float JUROS { get; set; }
        public string CPF_CNPJ { get; set; }
        public string NOME_DO_PAGADOR { get; set; }
        public string ENDERECO { get; set; }
        public float? NUMERO { get; set; }
        public string BAIRRO { get; set; }
        public string CIDADE { get; set; }
        public string UF { get; set; }
        public string CEP { get; set; }
        public string SACADO { get; set; }
        public string CNPJ_SACADO { get; set; }
        public string COMENTARIOS { get; set; }
    }



    public class TirarAcento
    {
        public string Remove(string texto)
        {
            string ComAcentos = "/®~^´`|–ªº§¨ÄÅÁÂÀÃäáâàãÉÊËÈéêëèÍÎÏÌíîïìÖÓÔÒÕöóôòõÜÚÛüúûùÇç";
            string SemAcentos = "____________AAAAAAaaaaaEEEEeeeeIIIIiiiiOOOOOoooooUUUuuuuCc";
            for (int i = 0; i < ComAcentos.Length; i++)
                texto = texto.Replace(ComAcentos[i].ToString(), SemAcentos[i].ToString()).Trim();
            return texto;
        }
        public string RemoveMascara(string cnpj)
        {
            string result = "";
            result = cnpj.Replace(".", "").Replace("-", "").Replace("/", "");
            return result;
        }
    }
}
