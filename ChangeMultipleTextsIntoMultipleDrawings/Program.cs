using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;
using System.IO;

namespace ChangeMultipleTextsIntoMultipleDrawings
{
	class Program {
		private static string[] arquivosDWG;
		private static string arquivoTXT, numeroAntigo, numeroNovo;

		[STAThread]
		public static void Main(string[] args) {
			SelectFilesDWG();
			SelectFileTXT();
			
			string[] linhas = File.ReadAllLines(arquivoTXT);
			for (int i = 0; i < linhas.Length; i++) {
				numeroAntigo = linhas[i].Split(';')[0];
				numeroNovo = linhas[i].Split(';')[1];
			}

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
                	foreach (var value in linhas) {
                		if (txt.TextString == numeroAntigo) {
							txt.TextString = numeroNovo;
                		}
                	}
                }
                
                double[] textLocation = new double[3];
				textLocation[0] = 7278.6720;
				textLocation[1] = 635.8620;
				textLocation[2] = 0;
				var text = acApp.ActiveDocument.ModelSpace.AddText("0B", textLocation, 46.3125);
				text.Alignment = AcAlignment.acAlignmentMiddle;
				text.TextAlignmentPoint = textLocation;
				text.StyleName = "style2";

				textLocation[0] = 7469.2051;
				textLocation[1] = 635.8620;
				var text1 = acApp.ActiveDocument.ModelSpace.AddText("ALTERADO NÚMERO ERB1", textLocation, 46.3125);
				text1.Alignment = AcAlignment.acAlignmentMiddleLeft;
				text1.TextAlignmentPoint = textLocation;
				text1.StyleName = "style2";

				textLocation[0] = 9895.2593;
				textLocation[1] = 635.8620;
				var text2 = acApp.ActiveDocument.ModelSpace.AddText("FBS", textLocation, 46.3125);
				text2.Alignment = AcAlignment.acAlignmentMiddle;
				text2.TextAlignmentPoint = textLocation;
				text2.StyleName = "style2";

				textLocation[0] = 10262.7772;
				textLocation[1] = 635.8620;
				var text3 = acApp.ActiveDocument.ModelSpace.AddText("HSJ", textLocation, 46.3125);
				text3.Alignment = AcAlignment.acAlignmentMiddle;
				text3.TextAlignmentPoint = textLocation;
				text3.StyleName = "style2";

				textLocation[0] = 10630.2951;
				textLocation[1] = 635.8620;
				var text4 = acApp.ActiveDocument.ModelSpace.AddText("JRS", textLocation, 46.3125);
				text4.Alignment = AcAlignment.acAlignmentMiddle;
				text4.TextAlignmentPoint = textLocation;
				text4.StyleName = "style2";

				textLocation[0] = 10997.8131;
				textLocation[1] = 635.8620;
				var text5 = acApp.ActiveDocument.ModelSpace.AddText("19/09/18", textLocation, 46.3125);
				text5.Alignment = AcAlignment.acAlignmentMiddle;
				text5.TextAlignmentPoint = textLocation;
				text5.StyleName = "style2";

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
	}
}