using Microsoft.Office.Interop.Word;
using System;
using System.Windows.Forms;

namespace Deneme2
{
    public partial class Form1 : Form
    {
        private readonly IWordDocumentService _wordDocumentService;
        private readonly IWordTableService _wordTableService;

        public Form1()
        {
            InitializeComponent();
            _wordDocumentService = InstanceFactory.GetInstance<IWordDocumentService>();
            _wordTableService = InstanceFactory.GetInstance<IWordTableService>();
        }


        private void CreateDocumentBtn_Click(object sender, EventArgs e)
        {
            CreateTitlePage();
        }



        private void CreateTitlePage()
        {
            Microsoft.Office.Interop.Word.Application application=null;
            Document document=null;

            try
            {
                var result=_wordDocumentService.CreateDocument();
                application = result.Application;
                document = result.Document;

                int totalRows=16;
                int totalColumns=21;

                Table contentTable=_wordTableService.CreateTable(document,totalRows,totalColumns);

                string filePath=@"C:\Users\Metraj Merkezi\Desktop\WordYoutubeDenemeler";
                _wordDocumentService.SaveDocumentWithAutoIncrement(_wordDocumentService, document, filePath, "Antet_Deneme");
                MessageBox.Show("Word Dokümanı Başarıyla Oluşturdu.");
            }
            finally
            {

            }
        }
    }
}
