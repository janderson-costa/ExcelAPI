using System;
using System.Linq;
using System.Text;
using Newtonsoft.Json;

namespace ExcelAPI {
	class Program {
		static void Main(string[] args) {
			var _workbookPath = GetArg(args, "workbookPath");
			var _excel = new Services.Excel(_workbookPath);
			var _method = GetArg(args, "method");
			var _sheetName = GetArg(args, "sheetName");
			var _controls = GetArg(args, "controls");

			Console.OutputEncoding = Encoding.UTF8;

			// Valida o arquivo do Excel ativo
			if (_workbookPath != "" && !_excel.IsActiveWorkbook(_workbookPath)) {
				return;
			}

			// Retorna os nomes das planilhas do arquivo
			if (_method == "GetSheets") {
				var sheets = _excel.GetSheets();
				var json = JsonConvert.SerializeObject(sheets);

				Console.WriteLine(json);
			}

			// Retorna a célula ou objetos selecionados
			if (_method == "GetControls") {
				var controls = _excel.GetControls();
				var json = JsonConvert.SerializeObject(controls);

				Console.WriteLine(json);

				// items.ForEach(item => {
				// 	Console.WriteLine($"Id: {item.Id}");
				// 	Console.WriteLine($"Name: {item.Name}");
				// 	Console.WriteLine($"Address: {item.Address}");
				// 	Console.WriteLine($"Value: {item.Value}");
				// 	Console.WriteLine($"Text: {item.Text}");
				// 	Console.WriteLine($"Type: {item.Type}");
				// 	Console.WriteLine($"List: {string.Join(", ", item.List)}");
				// 	Console.WriteLine();
				// });
			}

			// Seleciona os elementos do campo na plainlha
			// Comando - Ex.: .\ExcelAPI.exe method=SelectFieldControls sheetName=Relatório_Energia controls='[{"Id":"","Name":"","Address":"$O$9","Value":"Existe rede de energia na mesma propriedade do Site?","Text":"","Type":"cell","List":[],"Object":{}}]'
			if (_method == "SelectFieldControls") {
				//var controls = JsonConvert.DeserializeObject<List<Models.Control>>(_controls);
				//Console.WriteLine(sheetName);
				//excel.SelectFieldControls(sheetName, _controls);
			}

			// Carrega os elementos
			if (_method == "Load") {

			}

			// Limpa o valor dos elementos em todas as planilhas do Excel
			if (_method == "Clear") {

			}

			if (_method == "SaveWorkbook") {
				var result = _excel.SaveWorkbook();

				Console.WriteLine(result);
			}

			// Console.ReadKey();
		}

		static string GetArg(string[] args, string key) {
			var arg = args.FirstOrDefault(x => x.StartsWith(key + "="));
			var value = "";

			if (!string.IsNullOrEmpty(arg)) {
				value = arg.Split('=')[1].Trim();
			}

			return value;
		}
	}
}
