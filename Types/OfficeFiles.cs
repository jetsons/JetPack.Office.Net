using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Jetsons.JetPack;
using Jetsons.Excel;

namespace Jetsons.JetPack
{
	
	public static class OfficeFiles {


		/// <summary>
		/// Parse an XLSX file and convert it into a List of strongly typed Objects.
		/// Never returns null.
		/// If the first line is not headers, and you don't supply any columnProps, then the names of the columns are assumed.
		/// </summary>
		/// <param name="excelPath">Path of the XLSX file</param>
		/// <param name="headers">Read the first line of each sheet as the column headers</param>
		/// <param name="columnProps">Provide the properties per column, if known</param>
		/// <param name="onlySheetsNamed">Only returns the sheets with the given name.</param>
		/// <param name="skipBlankRows">Skips rows where all the cells are blank.</param>
		/// <returns></returns>
		public static ExcelResults<T> LoadXLSX<T>(this string excelPath, ExcelHeaders headers, List<string> columnProps = null, List<string> onlySheetsNamed = null, bool skipBlankRows = true) {
			return ExcelImporter.ImportXlsx<T>(excelPath, headers, columnProps, onlySheetsNamed, skipBlankRows);
		}
		

	}
}
