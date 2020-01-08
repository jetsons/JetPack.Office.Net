using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
/*using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;*/
using Jetsons.JetPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TikaOnDotNet.TextExtraction;
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
		
		/// <summary>
		/// Returns the plain text of the DOCX file using OpenXMLSDK
		/// 
		/// Code by Microsoft Corporation - https://code.msdn.microsoft.com/office/CSOpenXmlGetPlainText-554918c3/sourcecode?fileId=71592&pathId=851860130
		/// </summary>
		public static string LoadDOCXAsTextFast(this string filepath) {
			try {
				var package = WordprocessingDocument.Open(filepath, true);
				
				OpenXmlElement element = package.MainDocumentPart.Document.Body;
				if (element == null) {
					return "";
				}

				var text = GetPlainText(element).Trim();

				package.Dispose();

				return text;
			}
			catch (Exception) {
				return "";
			}
		}

		/// <summary> 
		/// Read Plain Text in all XmlElements of word document
		/// 
		/// Code by Microsoft Corporation - https://code.msdn.microsoft.com/office/CSOpenXmlGetPlainText-554918c3/sourcecode?fileId=71592&pathId=851860130
		/// </summary> 
		private static string GetPlainText(OpenXmlElement element) {
			StringBuilder PlainTextInWord = new StringBuilder();
			foreach (OpenXmlElement section in element.Elements()) {
				switch (section.LocalName) {
					// Text 
					case "t":
						PlainTextInWord.Append(section.InnerText);
						break;

					case "cr":  // Carriage return 
					case "br":  // Page break 
						PlainTextInWord.Append(Environment.NewLine);
						break;

					// Tab 
					case "tab":
						PlainTextInWord.Append("\t");
						break;

					// Paragraph 
					case "p":
						PlainTextInWord.Append(GetPlainText(section));
						PlainTextInWord.AppendLine(Environment.NewLine);
						break;

					default:
						PlainTextInWord.Append(GetPlainText(section));
						break;
				}
			}

			return PlainTextInWord.ToString();
		}

		/// <summary>
		/// Returns the plain text of the PDF file using iText7, for fast and low-quality text extraction.
		/// 
		/// Code from Code Project - https://www.codeproject.com/Articles/12445/Converting-PDF-to-Text-in-C
		/// </summary>
		/*public static string LoadPDFAsTextFast(string path) {

			// read PDF
			StringBuilder sb = new StringBuilder();

			// get document
			var document = new PdfDocument(new PdfReader(path));

			// per page
			for (int i = 1; i <= document.GetNumberOfPages(); i++) {

				// get text
				var its = new LocationTextExtractionStrategy();
				var page = document.GetPage(i);
				var pageText = PdfTextExtractor.GetTextFromPage(page, its);
				sb.AppendLine(pageText);
			}
			return sb.ToString().Trim();
		}*/

		/// <summary>
		/// Returns the plain text of the document using Tika, for slow but high-quality text extraction.
		/// Supports PDF, RTF, DOC, DOCX, XLS, XLSX, PPT, PPTX, ZIP (file listing), JPG (metadata).
		/// 
		/// Code from Tika - https://github.com/KevM/tikaondotnet
		/// </summary>
		private static string LoadDocumentAsText(this string path) {
			try {
				var textExtractor = new TextExtractor();

				var text = textExtractor.Extract(path).Text;

				return text.Trim();
			}
			catch (Exception) {
				return "";
			}
		}

		/// <summary>
		/// Returns the plain text of the RTF Document using Tika, for slow but high-quality text extraction.
		/// </summary>
		public static string LoadRTFAsText(this string path) {
			return path.LoadDocumentAsText();
		}
		/// <summary>
		/// Returns the plain text of the Word DOC document using Tika, for slow but high-quality text extraction.
		/// </summary>
		public static string LoadDOCAsText(this string path) {
			return path.LoadDocumentAsText();
		}
		/// <summary>
		/// Returns the plain text of the Word DOCX document using Tika, for slow but high-quality text extraction.
		/// </summary>
		public static string LoadDOCXAsText(this string path) {
			return path.LoadDocumentAsText();
		}
		/// <summary>
		/// Returns the plain text of the PDF document using Tika, for slow but high-quality text extraction.
		/// </summary>
		public static string LoadPDFAsText(this string path) {
			return path.LoadDocumentAsText();
		}
		/// <summary>
		/// Returns the plain text of the Excel XLS spreadsheet using Tika, for slow but high-quality text extraction.
		/// </summary>
		public static string LoadXLSAsText(this string path) {
			return path.LoadDocumentAsText();
		}
		/// <summary>
		/// Returns the plain text of the Excel XLSX spreadsheet using Tika, for slow but high-quality text extraction.
		/// </summary>
		public static string LoadXLSXAsText(this string path) {
			return path.LoadDocumentAsText();
		}
		/// <summary>
		/// Returns the plain text of the PowerPoint PPT presentation using Tika, for slow but high-quality text extraction.
		/// </summary>
		public static string LoadPPTAsText(this string path) {
			return path.LoadDocumentAsText();
		}
		/// <summary>
		/// Returns the plain text of the PowerPoint PPTX presentation using Tika, for slow but high-quality text extraction.
		/// </summary>
		public static string LoadPPTXAsText(this string path) {
			return path.LoadDocumentAsText();
		}
		

	}
}
