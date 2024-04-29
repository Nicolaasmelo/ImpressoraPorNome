using System;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Printing;

class Program
{
    static void Main()
    {

        string stringcon =
                   "Provider=Microsoft.ACE.OLEDB.12.0;" +
                   "Data Source=" + "C:\\Users\\Desc\\Documents\\My Web Sites\\WebSite1\\vet.mdb" + ";" +
                   "Mode=Share Deny None" +
                   ";Jet OLEDB:Database Password=" + "M63b07C42nava";




        OleDbConnection connection = new OleDbConnection(stringcon);

        // Abrir a conexão com o banco de dados
        connection.Open();

        // Executar o comando SELECT para recuperar os dados da tabela
        OleDbCommand command = new OleDbCommand("SELECT *  FROM Raiox2", connection);
        OleDbDataReader reader = command.ExecuteReader();


        string codigo = "";
        string texto = "";
        string impressora = "";
        string nomeImpressora = "";

        if (reader.Read())
        {
            // Campos que vão aparecer na etiqueta 
            texto += $" \n HOSPITAL VETERINARIO DR.HATO \n\n  Cod:{reader["Codigo"]}\n Prop: {reader["Prop"]},{reader["esp"]},\n{reader["nome"]},{reader["raça"]},{reader["sexo"]},{reader["nasc"]} \n Dr: {reader["NomeCompleto"]} \n {reader["nomeExame"]} - {reader["Data"]}";
            codigo += reader["Codigo"];
            impressora += reader["impressora"];
        }

        // Fechar a conexão com o banco de dados
        reader.Close();
        connection.Close();


        if (impressora == "Laboratório")
        {
            nomeImpressora = "Adobe PDF";

        }
        else
        {
            nomeImpressora = "EPSON73B3B8 (L1250 Series)";
        }

        // Informe o nome da impressora desejada

        // Obtém as impressoras instaladas no sistema
        PrinterSettings.InstalledPrinters.Cast<string>().ToList().ForEach(printer =>
        {
            if (printer.Equals(nomeImpressora, StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine($"Impressora encontrada: {printer}");


                //string texto = "Teste";
                // Imprimir o documento com o texto recuperado do banco de dados
                PrintDocument printDoc = new PrintDocument();
                printDoc.PrinterSettings.PrinterName = nomeImpressora;
                printDoc.PrintPage += new PrintPageEventHandler(PrintPage);
                //printDoc.DefaultPageSettings.PaperSize = new PaperSize(texto, 350,300);
                printDoc.DefaultPageSettings.PaperSize = new PaperSize(texto, 350, 150);
                printDoc.PrinterSettings.Copies = 1;
                printDoc.Print();



                // Função para imprimir o conteúdo em cada etiqueta
                void PrintPage(object sender, PrintPageEventArgs e)
                {
                    Font font = new Font("Arial", 20);
                    int numEtiquetasPorPagina = 1;
                    int alturaEtiqueta = 1300;
                    int larguraEtiqueta = 1300;

                    int margemSuperior = 2;
                    int margemEsquerda = 40;

                    int larguraFonte = larguraEtiqueta / texto.Length;
                    Font fonteEtiqueta = new Font(font.FontFamily, larguraFonte);
                    e.Graphics.DrawString(texto, fonteEtiqueta, Brushes.Black, margemEsquerda, margemSuperior);

                    
                }
                return; // Sai do loop assim que encontrar a impressora desejada
            }
        });

        // Se o loop terminar e não encontrar a impressora, exiba uma mensagem
        Console.WriteLine($"Impressora '{nomeImpressora}' não encontrada.");

       
    }
}
