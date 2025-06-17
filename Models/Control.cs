using System.Collections.Generic;

namespace ExcelAPI.Models {
	public class Control {
		// Representa uma célula, controle de formulário ou retângulo.

		public string Id = "";
		public string Name = "";
		public string Address = "";
		public string Value = "";
		public string Text = "";
		public string Type = ""; // cell, checkbox, radio, dropdown, activeXCheckbox, activeXRadio, activeXDropdown, rectangle
		public List<string> List = new List<string>();
		public object Object;
	}
}
