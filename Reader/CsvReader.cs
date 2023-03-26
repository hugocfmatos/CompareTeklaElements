using Microsoft.VisualBasic.FileIO;
using System.Data;
using System.Globalization;

namespace Reader
{
    public class CsvReader
    {
        public static DataTable Read(string path)
        {
            DataTable table = CreateDataTable();
            using (TextFieldParser csvParser = new(path))
            {
                csvParser.SetDelimiters(new string[] { "," });
                csvParser.HasFieldsEnclosedInQuotes = true;

                // Skip the row with column names
                csvParser.ReadLine();

                while (!csvParser.EndOfData)
                {
                    string[] fields = csvParser.ReadFields();

                    float[] startCoordinates = DecomposeCoordinates(fields[1]);
                    float[] endCoordinates = DecomposeCoordinates(fields[2]);

                    table.Rows.Add(
                        fields[0],
                        startCoordinates[0],
                        startCoordinates[1],
                        startCoordinates[2],
                        endCoordinates[0],
                        endCoordinates[1],
                        endCoordinates[2],
                        fields[3],
                        Int32.Parse(fields[4]));
                }
            }

            return table;
        }

        private static DataTable CreateDataTable()
        {
            DataTable table = new();
            table.Columns.Add("GUID", typeof(string));
            table.Columns.Add("StartMemberCoordinatesX", typeof(float));
            table.Columns.Add("StartMemberCoordinatesY", typeof(float));
            table.Columns.Add("StartMemberCoordinatesZ", typeof(float));
            table.Columns.Add("EndMemberCoordinatesX", typeof(float));
            table.Columns.Add("EndMemberCoordinatesY", typeof(float));
            table.Columns.Add("EndMemberCoordinatesZ", typeof(float));
            table.Columns.Add("Profile", typeof(string));
            table.Columns.Add("AssemblyPosition", typeof(int));

            return table;
        }

        private static float[] DecomposeCoordinates(string coordinates)
        {
            // Remove special characters and spaces and leave commas ","
            string cleanString = String.Join("", coordinates.Split('"', '(', ')', ' '));

            return cleanString.Split(",").Select(q => float.Parse(q, CultureInfo.InvariantCulture.NumberFormat)).ToArray();
        }
    }
}