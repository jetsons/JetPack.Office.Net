using System;
using System.Collections.Generic;
using System.Text;
using Jetsons.JetPack;

namespace Jetsons.JetPack.Excel {

	public class ExcelResults<T> {

		public bool Success;

		/// <summary>
		/// Strongly typed data of the Excel sheets
		/// </summary>
		public List<ExcelSheet<T>> Sheets = new List<ExcelSheet<T>>();

	}

	public class ExcelSheet<T> {

		public bool Success;

		/// <summary>
		/// Headers that were picked up from the first row of the sheet, or the headers that were given by the caller via columnProps.
		/// </summary>
		public List<string> Headers = new List<string>();

		/// <summary>
		/// Strongly typed data of the CSV records
		/// </summary>
		public List<T> Data = new List<T>();

	}

	public enum ExcelHeaders {

		/// <summary>
		/// The headers are surely on the first row of the Excel file
		/// </summary>
		FirstRow,

		/// <summary>
		/// There are surely no headers in the Excel file
		/// </summary>
		None,

		/// <summary>
		/// Auto-detect if there are headers on the first row of the Excel file
		/// </summary>
		AutoDetect
	}
	
}
