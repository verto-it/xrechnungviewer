using SharpShell.Attributes;
using SharpShell.SharpContextMenu;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Xsl;

namespace xrechnungviewer
{
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.ClassOfExtension, ".xml")]
    public class CountLinesExtension : SharpContextMenu
    {
        protected override bool CanShowMenu()
        {
            return true;
        }

        protected override ContextMenuStrip CreateMenu()
        {
            var menu = new ContextMenuStrip();

            var item = new ToolStripMenuItem
            {
                Text = "xRechnung als PDF öffnen"
            };

            item.Click += (sender, args) => OpenXRechnungAsPDF();

            menu.Items.Add(item);

            return menu;
        }


        private void ConvertXRechnungToPDF(string xRechnungPath, string pdfPath)
        {
            var xslPath = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "xrechnung-3.0.2-xrechnung-visualization-2024-06-20", "xsl", "xr-pdf.xsl");

            var xsl = new XslCompiledTransform();

            xsl.Load(xslPath);

            var settings = new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Parse
            };

            using (var reader = XmlReader.Create(xRechnungPath, settings))
            {
                using (var writer = XmlWriter.Create(pdfPath))
                {
                    xsl.Transform(reader, writer);
                }
            }
        }

        private void OpenXRechnungAsPDF()
        {
            foreach (var filePath in SelectedItemPaths)
            {
                var tempPath = Path.GetTempPath();
                var tempFile = Path.Combine(tempPath, Path.GetFileNameWithoutExtension(filePath) + ".pdf");

                ConvertXRechnungToPDF(filePath, tempFile);

                System.Diagnostics.Process.Start(tempFile);

                System.Threading.Thread.Sleep(5000);

                File.Delete(tempFile);
            }
        }

    }
}
