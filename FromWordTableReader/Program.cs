using NPOI.XWPF.UserModel;

namespace FromWordTableReader
{
    class Program
    {
        static void Main(string[] args)
        {
            // path to .docx file to open
            string filePath = "Table.docx";

            // loading Word document
            using (var fs = new System.IO.FileStream(filePath, System.IO.FileMode.Open))
            {
                XWPFDocument document = new XWPFDocument(fs);

                // Getting an array of tables from file
                var tables = document.Tables;

                // Printing all cells data from tables to console
                foreach (var table in tables)
                {
                   for (int i = 0;  i < table.Rows.Count; i++)
                    {
                        var row = table.Rows[i];
                        var cells = row.GetTableICells();
                        for (int j = 0; j < cells.Count; j++)
                        {
                            Console.Write(row.GetCell(j).GetText() + " ");
                        }
                        Console.WriteLine();
                    }
                }
            }
        }
    }
    
}