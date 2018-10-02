using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;

namespace ChangeMultipleTextsIntoMultipleDrawings
{
	class Program {
		private static string[] arquivosDWG;
		//private static string arquivoTXT, numeroAntigo, numeroNovo;

		[STAThread]
		public static void Main(string[] args) {
			SelectFilesDWG();
			
			/*
			SelectFileTXT();
			
			string[] linhas = File.ReadAllLines(arquivoTXT);
			for (int i = 0; i < linhas.Length; i++) {
				numeroAntigo = linhas[i].Split(';')[0];
				numeroNovo = linhas[i].Split(';')[1];
			}
			*/

            AcadApplication acApp = null;

            try {
                acApp = Marshal.GetActiveObject("AutoCAD.Application") as AcadApplication;
            } catch {
                MessageBox.Show("Não foi possível abrir o AutoCAD");
            }
			
        	foreach (var desenho in arquivosDWG) {
        		AcadDocument doc;

                try {
                    doc = acApp.Documents.Open(desenho, false, string.Empty);
                    AcadModelSpace modelSpace = doc.ModelSpace;
                } catch (Exception) {
                    MessageBox.Show("Não foi possível abrir o desenho {0}", desenho);
                    break;
                }

                AcadSelectionSet selset = null;
                selset = doc.SelectionSets.Add("texto");
                short[] ftype = { 0 };
                object[] fdata = { "TEXT" };
                selset.Select(AcSelect.acSelectionSetAll, null, null, ftype, fdata);

                foreach (IAcadText txt in selset) {
                	double[] teste = txt.InsertionPoint as double[];
                	
                	if (teste[0] > 23118 && teste[0] < 23289 && teste[1] > 227 && teste[1] < 376 && txt.TextString == "0A") {
						txt.TextString = "00";
                	}
                	
                	/*
                	if (txt.InsertionPoint) {
                		//
                	}*/
                	
                	/*
                	foreach (var value in linhas) {
                		if (txt.TextString == numeroAntigo) {
							txt.TextString = numeroNovo;
                		}
                	}
                	*/
                }
                
                double[] textLocation = new double[3];
				textLocation[0] = 20180.7576;
				textLocation[1] = 2475.0000;
				textLocation[2] = 0;
				var text = acApp.ActiveDocument.ModelSpace.AddText("00", textLocation, 38.6909);
				text.Alignment = AcAlignment.acAlignmentMiddle;
				text.TextAlignmentPoint = textLocation;
				text.StyleName = "ROMANS";

				textLocation[0] = 20456.5152;
				textLocation[1] = 2475.0000;
				var text1 = acApp.ActiveDocument.ModelSpace.AddText("02/10/18", textLocation, 38.6909);
				text1.Alignment = AcAlignment.acAlignmentMiddle;
				text1.TextAlignmentPoint = textLocation;
				text1.StyleName = "ROMANS";

				textLocation[0] = 20699.1182;
				textLocation[1] = 2475.0000;
				var text2 = acApp.ActiveDocument.ModelSpace.AddText("Documento Aprovado", textLocation, 38.6909);
				text2.Alignment = AcAlignment.acAlignmentMiddleLeft;
				
				text2.TextAlignmentPoint = textLocation;
				text2.StyleName = "ROMANS";

				textLocation[0] = 22819.8955;
				textLocation[1] = 2475.0000;
				var text3 = acApp.ActiveDocument.ModelSpace.AddText("FBS", textLocation, 38.6909);
				text3.Alignment = AcAlignment.acAlignmentMiddle;
				text3.TextAlignmentPoint = textLocation;
				text3.StyleName = "ROMANS";

				textLocation[0] = 23123.9288;
				textLocation[1] = 2475.0000;
				var text4 = acApp.ActiveDocument.ModelSpace.AddText("HSJ", textLocation, 38.6909);
				text4.Alignment = AcAlignment.acAlignmentMiddle;
				text4.TextAlignmentPoint = textLocation;
				text4.StyleName = "ROMANS";

				textLocation[0] = 23427.9727;
				textLocation[1] = 2475.0000;
				var text5 = acApp.ActiveDocument.ModelSpace.AddText("RTO", textLocation, 38.6909);
				text5.Alignment = AcAlignment.acAlignmentMiddle;
				text5.TextAlignmentPoint = textLocation;
				text5.StyleName = "ROMANS";

				acApp.ZoomExtents();
				doc.Save();
				doc.Close();
            }
		}

        static void SelectFilesDWG() {
            // Displays an OpenFileDialog so the user can select a Cursor.
            OpenFileDialog openFileDialog2 = new OpenFileDialog();
            openFileDialog2.Filter = "Drawing Files|*.dwg";
            openFileDialog2.Title = "Selecione os arquivos DWG";
            openFileDialog2.RestoreDirectory = true;
            openFileDialog2.Multiselect = true;

            if (openFileDialog2.ShowDialog() == DialogResult.OK) {
                // Guarda a lista dos caminhos completos dos arquivos selecionados
                arquivosDWG = openFileDialog2.FileNames;
            } else {
				Console.WriteLine("Arquivos não selecionados!");
				Console.ReadKey();
            }
        }
		
		/*
        static void SelectFileTXT() {
            // Displays an OpenFileDialog so the user can select a Cursor.
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "TXT Files|*.txt";
            openFileDialog.Title = "Selecione o arquivo TXT";
            openFileDialog.RestoreDirectory = true;
            openFileDialog.Multiselect = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK) {
                // Guarda a lista dos caminhos completos dos arquivos selecionados
                arquivoTXT = openFileDialog.FileName;
            } else {
				Console.WriteLine("Arquivo não selecionado!");
				Console.ReadKey();
            }
        }
		*/
	}
}