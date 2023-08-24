using System.Text;
using System.Text.RegularExpressions;

namespace ExcelTransformer.Models
{
    public class Template
    {
        public string Name { get; set; }
        public List<string> Columns { get; set; }
        public string FilePath { get; set; }
        public string Text { get; set;}

        public Template(string filePath)
        {
            FilePath = filePath;
            Name = Path.GetFileNameWithoutExtension(filePath);
            Text = File.ReadAllText(filePath);
            Columns = ExtractColumnNamesFromTemplate();
        }

        private List<string> ExtractColumnNamesFromTemplate()
        {
            List<string> columnNames = new List<string>();

            string[] templateLines = File.ReadAllLines(FilePath, Encoding.UTF8);
            foreach (string line in templateLines)
            {
                MatchCollection matches = Regex.Matches(line, @"\[(.*?)\]");
                foreach (Match match in matches)
                {
                    string columnName = match.Groups[1].Value;
                    if (!columnNames.Contains(columnName))
                    {
                        columnNames.Add(columnName);
                    }
                }
            }

            return columnNames;
        }
    }
}
