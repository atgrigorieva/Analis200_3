using System;
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
using System.Threading;
using System.IO;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.Xml;
using SWF = System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Drawing.Printing;
using Spire.PdfViewer.Forms;
using Spire.Pdf;
using Spire.Pdf.Widget;
using Spire.PdfViewer.Drawing;
using Spire.PdfViewer.Asp;
namespace Analis200
{
    public partial class Preview : Form
    {
        Analis _Analis;
        public Preview(Analis parent)
        {
            InitializeComponent();
            this._Analis = parent;
        }

        private void Preview_Load(object sender, EventArgs e)
        {
            Spire.Pdf.PdfDocument pdfdocument = new Spire.Pdf.PdfDocument();
            if (File.Exists(_Analis.filename))
            {
             //   pdfDocumentViewer1.LoadFromFile(_Analis.filename);
            }
            pdfdocument.PrintDocument.Print();
            pdfdocument.Dispose();
        }
      /*  private void btnPrint_Click(object sender, EventArgs e)
        {
            if (pdfdocument.PrintDocument.PageCount > 0)
            {
                pdfdocument.PrintDocument.Print();
            }
        }


        private void pdfDocumentViewer1_PdfLoaded(object sender, EventArgs args)
        {
            this.comBoxPages.Items.Clear();
            int totalPage = this.printPreviewDialog1.PageCount;

            for (int i = 1; i <= totalPage; i++)
            {
                this.comBoxPages.Items.Add(i.ToString());
            }

            this.comBoxPages.SelectedIndex = 0;
        }

        private void pdfDocumentViewer1_PageNumberChanged(object sender, EventArgs args)
        {
            if (this.comBoxPages.Items.Count <= 0)
                return;
            if (this.printPreviewDialog1.CurrentPageNumber != this.comBoxPages.SelectedIndex + 1)
            {
                this.comBoxPages.SelectedIndex = this.printPreviewDialog1.CurrentPageNumber - 1;
            }
        }

        private void comBoxPages_SelectedIndexChanged(object sender, EventArgs e)
        {
            int soucePage = this.printPreviewDialog1.CurrentPageNumber;
            int targetPage = this.comBoxPages.SelectedIndex + 1;
            if (soucePage != targetPage)
            {
                this.printPreviewDialog1.GoToPage(targetPage);
            }
        }*/
    }
}
