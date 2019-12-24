using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Jetsons.JetPack;
using OfficeOpenXml;

namespace Jetsons.Excel {
	public static class ExcelImporter {

	
		/// <summary>
		/// Parse an XLSX file and convert it into a List of strongly typed Objects.
		/// Never returns null.
		/// If the first line is not headers, and you don't supply any columnProps, then the names of the columns are assumed.
		/// </summary>
		/// <param name="excelPath">Path of the XLSX file</param>
		/// <param name="headers">Read the first line of each sheet as the column headers?</param>
		/// <param name="columnProps">Provide the properties per column, if known</param>
		/// <param name="onlySheetsNamed">Only returns the sheets with the given name.</param>
		/// <returns></returns>
		public static ExcelResults<T> ImportXlsx<T>(string excelPath, ExcelHeaders headers, List<string> columnProps = null, List<string> onlySheetsNamed = null, bool skipBlankRows = true) {

			var results = new ExcelResults<T>();
			
			// fail if file does not exist
			if (!excelPath.FileExists()){
				return results;
			}

			var workbook = new ExcelPackage(new FileInfo(excelPath)).Workbook;
			
			// fail if workbook cannot be read
			if (workbook == null){
				return results;
			}

			// per sheet
			foreach (var sheet in workbook.Worksheets) {
				
				// filter by worksheet name if wanted
				if (onlySheetsNamed != null && !onlySheetsNamed.Contains(sheet.Name)) {
					continue;
				}

				// skip the sheet if it is blank
				if (sheet.Dimension == null) {
					continue;
				}

				// new sheet
				var sheetResult = new ExcelSheet<T>();
				results.Sheets.Add(sheetResult);

				// if props given, take those
				if (columnProps.Exists()) {
					sheetResult.Headers = columnProps;
				}

				// calc starting row based on header config
				var startRow = headers == ExcelHeaders.None ? 1 : 2;

				// get the size of the sheet
				int rowCount = sheet.Dimension.End.Row;
				int colCount = sheet.Dimension.End.Column;

				// read the headers of the first row if wanted and not given
				if (headers == ExcelHeaders.FirstRow && !sheetResult.Headers.Exists()) {

					// per cell
					for (int c = 1; c <= colCount; c++) {

						// register the header based on the cell text
						var text = sheet.Cells[1, c].Text.Trim();
						if (text.Exists()) {
							sheetResult.Headers.Add(text);
						}
					}
				}

				// per row
				for (int r = startRow; r <= rowCount; r++) {

					// create a new data object for the row
					var row = Activator.CreateInstance<T>();

					// check for blank rows
					var isBlankRow = true;

					// per cell
					for (int c = 1; c <= colCount; c++) {
						
						var prop = GetPropName(c-1, sheetResult);
						var text = sheet.Cells[r, c].Text.Trim();

						// skip if column name not given
						if (prop == null) {
							continue;
						}

						// store this value in the row object
						row.SetPropValue(prop, text);

						// check for blank rows
						if (text.Exists()) {
							isBlankRow = false;
						}
					
					}

					// skip blank rows if wanted
					if (skipBlankRows && isBlankRow) {
					}else{
						sheetResult.Data.Add(row);
					}
				}
			}
			
			// all ok if some sheets were read
			results.Success = results.Sheets.Exists();

			return results;
		}
		
		private static string GetPropName<T>(int column, ExcelSheet<T> results) {

			// if the header has not been registered
			if (!results.Headers.HasSlotAndValue(column)) {

				// if its a dynamic type, generate column name from column index
				/*if (typeof(T).IsDynamicType()) {
					while ((results.Headers.Count - 1) < column) {
						results.Headers.Add("Column" + (results.Headers.Count + 1));
					}
				}*/

				// otherwise mark that we cannot support this prop
				return null;
			}

			// return the registered header
			return results.Headers[column];
		}

	}
}
