using System;
using System.Collections.Generic;
//using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using _Excel = Microsoft.Office.Interop.Excel;
using VBE = Microsoft.Vbe.Interop.Forms;
// using Office = Microsoft.Office.Core;
// using VB = Microsoft.VisualBasic;
using Microsoft.Office.Core;
using ExcelAPI.Models;

namespace ExcelAPI.Services {
	public class Excel {
		private readonly _Excel.Application _excel = (_Excel.Application)Marshal.GetActiveObject("Excel.Application");

		public Excel(string filePath) {
			filePath = filePath.Replace("/", "\\").Trim();

			_Excel.Workbook target = null;

			foreach (_Excel.Workbook workbook in _excel.Workbooks) {
				// workbook: Representa um arquivo do excel

				if (workbook.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase)) {
					target = workbook;
					break;
				}
			}

			if (target == null)
				return;

			try {
				// Ativa a janela do Excel no z-order
				_Excel.Window window = target.Windows[1];
				window.Activate();
			}
			catch (Exception ex) {
				Console.WriteLine(ex.Message);
			}
		}

		// não usado
		public void OpenWorkbook(string filePath) {
			// Abre o arquivo especificado no Excel.

			filePath = filePath.Replace("/", "\\").Trim();

			// if (Process.GetProcessesByName("EXCEL").Length > 0) {
			// 	// Tenta anexar a uma instância já aberta
			// 	try {
			// 		_excel = (_Excel.Application)Marshal.GetActiveObject("Excel.Application");
			// 	}
			// 	catch (COMException) {
			// 		// Falha ao obter objeto existente
			// 		_excel = new _Excel.Application();
			// 	}
			// }
			// else {
			// 	// Se não há Excel rodando, cria uma nova instância
			// 	_excel = new _Excel.Application();
			// }

			// _excel.Visible = true;

			// // Tenta abrir o arquivo (ou conectar se já estiver aberto)
			// try {
			// 	workbook = _excel.Workbooks.Open(filePath);
			// }
			// catch (COMException) {
			// 	// Se já está aberto, obtém via rota direta
			// 	workbook = _excel.Workbooks[filePath];
			// }
		}

		public string SaveWorkbook() {
			// Salva o arquivo do Excel.
			// Obs.: Retorna falha caso haja alguma janela de dialogo aberta no Excel.

			try {
				_excel.DisplayAlerts = false;

				var workbook = _excel.ActiveWorkbook;

				workbook.Saved = true;
				workbook.Save();

				return "";
			}
			catch (Exception ex) {
				return ex.Message;
			}
			finally {
				try {
					_excel.DisplayAlerts = true;
				}
				catch (Exception) { }
			}
		}

		public bool IsActiveWorkbook(string requestedFilePath) {
			// Retorna se o arquivo do Excel ativo é o mesmo que foi passado como parâmetro.

			var workbookPath = _excel.ActiveWorkbook.FullName.ToLower().TrimEnd('x').Replace("\\", "/").Trim();
			var path = requestedFilePath.ToLower().TrimEnd('x').Replace("\\", "/").Trim();

			return workbookPath.EndsWith(path);
		}

		public List<Sheet> GetSheets() {
			if (_excel == null || _excel.ActiveWorkbook == null) return null;

			var sheets = new List<Sheet>();

			for (int i = 1; i <= _excel.Sheets.Count; i++) {
				var workSheet = (_Excel.Worksheet)_excel.Sheets[i];
				var sheet = new Sheet();

				// Adiciona o ID da planilha como propriedade customizada
				string id = null;

				foreach (_Excel.CustomProperty customProp in workSheet.CustomProperties) {
					if (customProp.Name == "SheetId") {
						id = customProp.Value.ToString();
						break;
					}
				}

				if (id == null) {
					// ! Obs.: Propriedade customizada não é mantida após salvar > fechar > abrir o arquivo
					id = Guid.NewGuid().ToString();
					workSheet.CustomProperties.Add("SheetId", id);
					SaveWorkbook();
				}

				sheet.Id = id;
				sheet.Name = workSheet.Name;
				sheets.Add(sheet);
			}

			return sheets;
		}

		public List<Models.Control> GetControls() {
			// Retorna o objeto que representa a célula ou objeto selecionado no Excel.
			// Obs.: Somente seleção única.

			if (_excel == null || _excel.ActiveWorkbook == null) return null;

			_excel.Visible = true;

			var sheet = (_Excel.Worksheet)_excel.ActiveSheet;
			var excelSelection = _excel.Selection;
			var controls = new List<Models.Control>();

			// ! Célula
			if (excelSelection is _Excel.Range range) {
				var activeCell = _excel.ActiveCell;
				var control = new Models.Control {
					Type = "cell",
					Address = activeCell.Address,
					Value = activeCell.Value?.ToString().Trim() ?? "",
					Object = activeCell
				};

				// ! Dropdown (célula com dropdown usando Dados > Validação de Dados)
				if (activeCell.GetType().GetProperty("Validation") != null && activeCell.Validation.Type == 3) {
					var formula = activeCell.Validation.Formula1;

					if (formula.Contains(":$")) { // Ex.: $A$1:$A$10
						control.List = ConvertRangeToList(_excel, formula);
					}
					else if (formula.Contains(";")) { // Ex.: valor; valor; valor; ...
						control.List = formula.Split(';').ToList();
					}
					else {
						control.List = new List<string>();
					}

					control.List = control.List.Select(x => x?.Trim()).ToList();
					control.List.Sort();
				}

				controls.Add(control);
			}
			else {
				// Forma alternativa de acessar a propriedade selection.ShapeRange
				var shapes = (_Excel.ShapeRange)excelSelection.GetType().InvokeMember(
					"ShapeRange", System.Reflection.BindingFlags.GetProperty, null, excelSelection, null
				);

				foreach (_Excel.Shape shape in shapes) {
					var shapeType = GetShapeType(shape);
					var control = new Models.Control();

					// ! FormControls
					if (shapeType == "checkbox") {
						var fcCheckbox = (_Excel.CheckBox)shape;
						var text = fcCheckbox.Text.Trim();

						if (string.IsNullOrEmpty(text))
							text = shape.AlternativeText.Trim();

						control.Type = shapeType;
						control.Id = shape.ID.ToString();
						control.Name = fcCheckbox.Name;
						control.Text = text;
						control.Value = Convert.ToInt32(fcCheckbox.Value) > 0 ? "1" : "0";
						control.Object = fcCheckbox;

						controls.Add(control);
					}

					else if (shapeType == "radio") {
						var fcRadio = (_Excel.OptionButton)shape;
						var text = fcRadio.Text.Trim();

						if (string.IsNullOrEmpty(text))
							text = shape.AlternativeText.Trim();

						control.Type = shapeType;
						control.Id = shape.ID.ToString();
						control.Name = fcRadio.Name;
						control.Text = text;
						control.Value = Convert.ToInt32(fcRadio.Value) > 0 ? "1" : "0";
						control.Object = fcRadio;

						controls.Add(control);
					}

					else if (shapeType == "dropdown") {
						var fcDropdown = (_Excel.DropDown)shape;

						if (fcDropdown.ListCount <= 0)
							continue;

						control.Type = shapeType;
						control.Id = shape.ID.ToString();
						control.Name = fcDropdown.Name;
						control.Text = GetDropdownText(_excel, fcDropdown).Trim();
						control.Value = fcDropdown.Value.ToString().Trim();
						control.List = ConvertRangeToList(_excel, fcDropdown.ListFillRange);
						control.Object = fcDropdown;

						controls.Add(control);
					}

					// ! OLEObject Controls (ActiveX)
					else if (shapeType == "activeXCheckbox") {
						var oleObject = (_Excel.OLEObject)shape.OLEFormat.Object;
						var _control = (VBE.CheckBox)oleObject.Object;

						control.Type = shapeType;
						control.Id = shape.ID.ToString();
						control.Name = shape.Name;
						control.Text = (!string.IsNullOrEmpty(_control.Caption) ? _control.Caption : shape.AlternativeText).Trim();
						control.Value = (bool)_control.get_Value() ? "1" : "0";
						control.Object = control;

						controls.Add(control);
					}

					else if (shapeType == "activeXRadio") {
						var oleObject = (_Excel.OLEObject)shape.OLEFormat.Object;
						var _control = (VBE.OptionButton)oleObject.Object;

						control.Type = shapeType;
						control.Id = shape.ID.ToString();
						control.Name = shape.Name;
						control.Text = (!string.IsNullOrEmpty(_control.Caption) ? _control.Caption : shape.AlternativeText).Trim();
						control.Value = (bool)_control.get_Value() ? "1" : "0";
						control.Object = control;

						controls.Add(control);
					}

					else if (shapeType == "activeXDropdown") {
						var oleObject = (_Excel.OLEObject)shape.OLEFormat.Object;
						var _control = (VBE.ComboBox)oleObject.Object;

						if (_control.ListCount <= 0) return null;

						control.Type = shapeType;
						control.Id = shape.ID.ToString();
						control.Name = shape.Name;
						control.Text = control.Text.Trim();
						control.Value = _control.get_Value().ToString().Trim();
						control.List = new List<string>();
						control.Object = control;

						for (int i = 0; i < _control.ListCount; i++) {
							var value = _control.get_List(i);

							if (value != null) {
								control.List.Add(value.ToString().Trim());
							}
						}

						controls.Add(control);
					}

					// ! Retângulo (para imagens)
					else if (shapeType == "rectangle") {
						var rectangle = (_Excel.Rectangle)shape.DrawingObject;

						control.Type = shapeType;
						control.Id = shape.ID.ToString();
						control.Name = rectangle.Name;
						control.Object = rectangle;

						controls.Add(control);
					}
				}
			}

			// Garante que todo os elementos da seleção sejam do mesmo tipo
			foreach (var control in controls) {
				if (controls[0].Type != control.Type)
					return null;
			}

			return controls;
		}

		public string SelectFieldControls(string sheetName, List<Models.Control> controls) {
			if (_excel == null || _excel.ActiveWorkbook == null)
				return "Nenhuma planilha aberta";

			var sheet = GetSheetByName(_excel, sheetName);

			if (sheet == null)
				return $"Planilha {sheetName} não encontrada ou com nome diferente.";

			var type = controls[0].Type;

			if (type == "cell") {
				var cell = controls[0];
				var _cell = sheet.Range[cell.Address];

				if (_cell != null) {
					sheet.Select();
					ZoomTo(_excel, _cell.EntireRow);
					// ExcelService.ScrollWindowTo(up: 10);
					// _cell.Select();
					// minimizeForm = true;
				}
				else {
					return "Célula não encontrada.";
				}
			}

			// if (type == "dropdown" || type == "activeXDropdown") {
			//     var dropdown = controls[0];
			// 	shape = ExcelService.GetShape(sheet, dropdown);

			// 	if (shape != null) {
			// 		sheet.Select();
			// 		ExcelService.ZoomTo(shape.TopLeftCell.EntireRow);
			//         ExcelService.ScrollWindowTo(up: 10);
			//         shape.Select();
			// 		minimizeForm = true;
			// 	} else {
			// 		MessageBox.Show(
			// 			"Objeto Dropdown não encontrado.",
			//             cells.Objetos.ToolTipText,
			// 			MessageBoxButtons.OK,
			// 			MessageBoxIcon.Warning
			// 		);
			// 	}
			// }

			// if (type == "checkbox" || type == "activeXCheckbox") {
			//     var checkboxes = controls;

			// 	foreach (var checkbox in checkboxes) {
			// 		shape = ExcelService.GetShape(sheet, checkbox);

			// 		if (shape != null) {
			//             if (count == 0) {
			// 				sheet.Select();
			//                 ExcelService.ZoomTo(shape.TopLeftCell.EntireRow);
			//                 ExcelService.ScrollWindowTo(up: 10);
			//                 shape.Select(false);
			//                 minimizeForm = true;
			//             } else { 
			//                 shape.Select(false);
			//             }
			// 		} else {
			// 			MessageBox.Show(
			// 				"Objeto Checkbox (" + checkbox.Text + ") não encontrado.",
			//                 cells.Objetos.ToolTipText,
			// 				MessageBoxButtons.OK,
			// 				MessageBoxIcon.Warning
			// 			);
			// 		}

			// 		count++;
			// 	}
			// }

			// if (type == "radio" || type == "activeXRadio") {
			//     var radios = controls;

			// 	foreach (var radio in radios) {
			// 		shape = ExcelService.GetShape(sheet, radio);

			// 		if (shape != null) {
			//             if (count == 0) {
			// 				sheet.Select();
			//                 ExcelService.ZoomTo(shape.TopLeftCell.EntireRow);
			//                 ExcelService.ScrollWindowTo(up: 10);
			//                 shape.Select(false);
			//                 minimizeForm = true;
			//             } else { 
			//                 shape.Select(false);
			//             }
			// 		} else {
			// 			MessageBox.Show(
			// 				"Objeto Radio (" + radio.Text + ") não encontrado.",
			//                 cells.Objetos.ToolTipText,
			// 				MessageBoxButtons.OK,
			// 				MessageBoxIcon.Warning
			// 			);
			// 		}

			// 		count++;
			// 	}
			// }

			// if (type == "rectangle") {
			//     var rectangle = controls[0];
			// 	shape = ExcelService.GetShape(sheet, rectangle);

			// 	if (shape != null) {
			// 		sheet.Select();
			// 		ExcelService.ZoomTo(shape.TopLeftCell.EntireRow);
			//         ExcelService.ScrollWindowTo(up: 10);
			//         shape.Select();
			// 		minimizeForm = true;
			// 	} else {
			// 		MessageBox.Show(
			// 			"Objeto Retangulo não encontrado.",
			//             cells.Objetos.ToolTipText,
			// 			MessageBoxButtons.OK,
			// 			MessageBoxIcon.Warning
			// 		);
			// 	}
			// }

			// if (minimizeForm) {
			//     // minimiza o form para que o usuário visualize a célula/controls
			//     WindowState = FormWindowState.Minimized;
			// }

			return "";
		}

		public void UnselectAllShapes() {
			if (_excel == null || _excel.ActiveWorkbook == null) return;

			var activeSheet = (_Excel.Worksheet)_excel.ActiveSheet;

			activeSheet.Range["A1"].Select();
		}

		public void ZoomTo(_Excel._Application excel, _Excel.Range range) {
			// Foca a vizualização no range especificado.

			excel.Goto(range, true);
		}


		// INTERNO

		internal _Excel.Worksheet GetSheetByName(_Excel.Application excel, string sheetName) {
			foreach (_Excel.Worksheet sheet in excel.Worksheets) {
				if (sheet.Name != null && sheet.Name.ToLower().Trim() == sheetName.ToLower().Trim())
					return sheet;
			}

			return null;
		}

		internal string GetShapeType(_Excel.Shape shape) {
			var type = shape.Type;

			if (type == MsoShapeType.msoFormControl) {
				if (shape.FormControlType == _Excel.XlFormControl.xlCheckBox) return "checkbox";
				if (shape.FormControlType == _Excel.XlFormControl.xlOptionButton) return "radio";
				if (shape.FormControlType == _Excel.XlFormControl.xlDropDown) return "dropdown";
			}
			else if (type == MsoShapeType.msoOLEControlObject) {
				var oleObject = (_Excel.OLEObject)shape.OLEFormat.Object;
				var progId = oleObject.progID;

				if (progId == "Forms.CheckBox.1") return "activeXCheckbox";
				if (progId == "Forms.OptionButton.1") return "activeXRadio";
				if (progId == "Forms.ComboBox.1") return "activeXDropdown";
			}
			else if (shape.AutoShapeType == MsoAutoShapeType.msoShapeRectangle) {
				return "rectangle";
			}

			return "";
		}

		internal _Excel.Shape GetShapeByName(_Excel.Application excel, string name) {
			var activeSheet = (_Excel.Worksheet)excel.ActiveSheet;

			foreach (_Excel.Shape shape in activeSheet.Shapes) {
				if (shape.Name.Equals(name)) {
					return shape;
				}
			}

			return null;
		}

		internal string GetDropdownText(_Excel.Application excel, _Excel.DropDown dropdown) {
			// Retorna o valor do dropdown selecionado coõ texto.

			// Obtém o índice do item selecionado (baseado em 1)
			int selectedIndex = dropdown.Value;

			// Verifica se o índice é válido
			if (selectedIndex < 1)
				return string.Empty;

			// Obtém o intervalo de preenchimento da lista, por exemplo, "$A$1:$A$5"
			string listFillRange = dropdown.ListFillRange;

			if (string.IsNullOrEmpty(listFillRange))
				return string.Empty;

			// Remove os símbolos de cifrão para obter uma referência utilizável, por exemplo, "A1:A5"
			listFillRange = listFillRange.Replace("$", "");

			// Acessa o intervalo de células correspondente a partir do objeto Worksheet
			_Excel.Range range = excel.Range[listFillRange];

			// Verifica se o índice está dentro dos limites do intervalo
			if (selectedIndex > range.Rows.Count)
				return string.Empty;

			// Obtém o valor da célula correspondente ao item selecionado
			object cellValue = ((_Excel.Range)range.Cells[selectedIndex, 1]).Value;

			return cellValue?.ToString() ?? string.Empty;
		}

		internal List<string> ConvertRangeToList(_Excel.Application excel, string formula) {
			var range = excel.Range[formula];
			var value = range.Value;

			if (value is object[,] array) {
				return array.Cast<object>()
					.Where(x => x != null)
					.Select(x => x.ToString().Trim())
					.ToList();
			}
			else if (value != null) {
				return new List<string> { value.ToString().Trim() };
			}

			return new List<string>();
		}
	}
}
