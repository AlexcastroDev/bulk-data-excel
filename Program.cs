using MiniExcelLibs;
using System.Text;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

var path = "RELATORIO_DTB_BRASIL_MUNICIPIO.xlsx";
using (var stream = File.OpenRead(path))
{
    var rows = stream.Query(useHeaderRow: true).ToList();
    var values = new List<Dictionary<string, object>>();

    for (int i = 0; i < rows.Count; i++)
    {
        var row = rows[i];
        var id = row.ID;
        var State = row.State;
        var City = row.City;
        var reference = i + 1;
        values.Add(
            new Dictionary<string, object> { { "Reference", reference }, { "ID", id }, { "State", State }, { "City", City }, { "Error", "" } }
        );
    }

    Console.WriteLine(values);
    MiniExcel.SaveAs("output.xlsx", values);
}



// using (var stream = File.OpenRead("RELATORIO_DTB_BRASIL_MUNICIPIO.xlsx"))
// {
//     using var reader = ExcelReaderFactory.CreateReader(stream);
//     reader.Read();

//     var headers = new List<string>();

//     for (int i = 0; i < reader.FieldCount; i++)
//     {
//         headers.Add(reader.GetString(i));
//     }
//     if (headers[0] != "ID" || headers[1] != "State_Code" || headers[2] != "State_Name")
//     {
//         throw new Exception("Invalid headers");
//     }

//     do
//     {
//         while (reader.Read())
//         {
//             try
//             {
//                 var id = reader.GetValue(0);
//                 var stateCode = reader.GetValue(1);
//                 var stateName = reader.GetValue(2);
//                 Console.WriteLine($"{id} - {stateCode} - {stateName}");
//             }
//             catch
//             {
//                 // Maybe write a column with error (?)
//             }
//         }
//     } while (reader.NextResult());
// }