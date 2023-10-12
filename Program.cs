using ExcelDataReader;
using System.Text;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

using (var stream = File.OpenRead("RELATORIO_DTB_BRASIL_MUNICIPIO.xlsx"))
{
    using var reader = ExcelReaderFactory.CreateReader(stream);
    reader.Read();

    var headers = new List<string>();

    for (int i = 0; i < reader.FieldCount; i++)
    {
        headers.Add(reader.GetString(i));
    }
    if (headers[0] != "ID" || headers[1] != "State_Code" || headers[2] != "State_Name")
    {
        throw new Exception("Invalid headers");
    }

    do
    {
        while (reader.Read())
        {
            try
            {
                var id = reader.GetValue(0);
                var stateCode = reader.GetValue(1);
                var stateName = reader.GetValue(2);
                Console.WriteLine($"{id} - {stateCode} - {stateName}");
            }
            catch
            {
                // Maybe write a column with error (?)
            }
        }
    } while (reader.NextResult());
}