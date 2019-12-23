using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using iText.Forms;
using iText.Kernel.Pdf;

namespace Pdf_Excel
{
	public class Program
	{
		private static int red, blue, yellow;
		private static string pdfFolderPath = "";

		private static List<string> redKeys = new List<string> { "red", "abc" };
		private static List<string> blueKeys = new List<string> { "red", "abc" };
		private static List<string> yeallowKeys = new List<string> { "red", "abc" };

		[STAThread]
		public static void Main()
		{
			// Console.WriteLine("Hello World");
			// 1. ChooseFolderContainPDFs();
			// Choose successfully -> call Process()
			// else return back to the console.

			StartProject: 

			Console.WriteLine("Choose your options");
			Console.WriteLine("1. Choose Folder which contains PDF");
			var option = Console.ReadLine();
			switch (option)
			{
				case "1":
					// Choose folder to read pdf
					ChooseFolderContainPDFs();

					// Process
					Process();
					break;
				default:
					goto StartProject;
					break;
			}
		}

		public static void ChooseFolderContainPDFs()
		{
			// pdfFolderPath = "...";
			var fbd = new FolderBrowserDialog();
			if (fbd.ShowDialog() == DialogResult.OK)
			{
				//Console.WriteLine(path); // full path
				//Console.WriteLine(System.IO.Path.GetFileName(path)); // file name
				pdfFolderPath = fbd.SelectedPath;
			}

		}

		public static void Process()
		{
			ReadPDF(null);
			WriteToExcel(null);
		}

		private static void ReadPDF(string pathToPDF)
		{
			// Read files in the choosen folder

			foreach (string file in Directory.EnumerateFiles(pdfFolderPath, "*.pdf"))
			{
				/*
				 * Read each pdf file in the folder
				 * Open the file
				 * Read all the fields with key
				 * Step 1: Read and Write File information
				 * Step 2: Read and write complexity
				 * 
				 */
				PdfReader pdfReader = new PdfReader(file);
				PdfDocument pdfDoc = new PdfDocument(pdfReader);
				PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDoc, true);

				var fields = form.GetFormFields();

				var output = "";
				foreach (var field in fields)
				{
					output += $"{field.Key} - {field.Value.GetValue()}\n";
				}
				Console.WriteLine(output);
				pdfDoc.Close();

				// Write to excel
				WriteToExcel(null);
			}

		}

		private static void WriteToExcel(string pathToExcel)
		{
			// Do sth here
			Console.WriteLine(red);
			// 
			Console.WriteLine("Done, succesfully");
		}

	}
}