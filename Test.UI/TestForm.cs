using MetrajMerkezi.Office.Word.Abstraction.File;
using MetrajMerkezi.Office.Word.Abstraction.Manipulation;
using MetrajMerkezi.Office.Word.Abstraction.Object._Page;
using MetrajMerkezi.Office.Word.Abstraction.Object._Table;
using MetrajMerkezi.Office.Word.Enum;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using Test.UI.DependencyResolver;

namespace Test.UI
{
    public partial class TestForm : Form
    {
        private readonly IWordDocumentService _wordDocumentService;
        private readonly IWordTableService _wordTableService;
        private readonly IWordRowService _wordRowService;
        private readonly IWordCellService _wordCellService;
        private readonly IWordPageSetupService _wordPageSetupService;

        public const string IssuedForApprovalEN = "سبب الإصدار \nISSUED FOR";
        public const string REV = "تنقيح \nREV";
        public const string PREPARED = "بواسط \nPREPARED";
        public const string CHECKED = "راجعها \nCHECKED";
        public const string APPROVED = "وافق عليها \nAPPROVED";
        public const string DATE = "التاريخ \nDATE";


        public TestForm()
        {
            InitializeComponent();
            _wordDocumentService = InstanceFactory.GetInstance<IWordDocumentService>();
            _wordTableService = InstanceFactory.GetInstance<IWordTableService>();
            _wordRowService = InstanceFactory.GetInstance<IWordRowService>();
            _wordCellService = InstanceFactory.GetInstance<IWordCellService>();
            _wordPageSetupService = InstanceFactory.GetInstance<IWordPageSetupService>();
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
                ///<summary>
                /// Word Uygulaması ve bu uygulama içerisine eklenmiş olan dokümanının oluşturulması ve kullanılmak üzere result isimli değişkene atanması
                /// </summary>
                var result=_wordDocumentService.CreateDocument();

                ///<summary>
                /// result kısmından gelen Application ve Document nesnelerinin metodun ilk kısmında null olarak tanımlanan application ve document isimli değişkenlere atanması
                /// </summary>
                application = result.Application;
                document = result.Document;

                ///<summary>
                /// Oluşturulacak olan tablonun satır ve sütun sayılarının integer türünden değişkenler olarak atanması
                /// </summary>
                int totalRows=18;
                int totalColumns=21;

                ///<summary>
                /// Oluşturulan word application içerisine eklenen document nesnesin ilk sayfasının firstSection ismiyle bir değişkene atanması
                /// </summary>
                Section firstSection = document.Sections[1];

                ///<summary>
                /// firstSection isimli değişkene atanan sayfanın kenar boşluklarının ayarlanması
                /// </summary>
                _wordPageSetupService.SetPageMargins(firstSection, 40.0f, 30.0f, 40.0f, 40.0f);

                ///<summary>
                /// Oluşturulan word dokümanı içerisinde belirtilen satır ve sütun adedince tablonun oluşturulması 
                /// ve oluşturulan tablonun manipüle edilebilmesi için contentTable isimli bir değişkene atanması
                /// </summary>
                Table contentTable=_wordTableService.CreateTable(document,totalRows,totalColumns);


                ///<summary>
                /// Oluşturulan tablonun satır yüksekliklerinin ayarlanması
                /// </summary>
                for (int i = 1; i <= 10; i++)
                {
                    if (i == 1)
                        _wordRowService.SetRowHeight(contentTable, i, 80);
                    if (i > 1 && i <= 6)
                        _wordRowService.SetRowHeight(contentTable, i, 14);
                    if (i == 7)
                        _wordRowService.SetRowHeight(contentTable, i, 35);
                    if (i == 8)
                        _wordRowService.SetRowHeight(contentTable, i, 70);
                    if (i == 9)
                        _wordRowService.SetRowHeight(contentTable, i, 90);
                    if (i == 10)
                        _wordRowService.SetRowHeight(contentTable, i, 65);
                }

                ///<summary>
                /// Oluşturulan tablonun 1. sütun genişliğinin ayarlanması
                /// </summary>
                contentTable.Columns [1].Width = 35;

                ///<summary>
                /// 1. Satırdaki hücrelerin birleştirme işlemleri, içerik ayarlama ve hizalama işlemlerinin yapılması
                /// </summary>
                _wordCellService.MergeAndSetTextToSpesicifCell(contentTable,
                    1, 1, 1, 9,
                    1, 1, "Engineer Sticker's",
                    11.50f,
                    HorizontalTextAlignment.Left,
                    VerticalTextAlignment.Top
                    );

                _wordCellService.MergeAndSetTextToSpesicifCell(contentTable,
                    1, 2, 1, 13,
                    1, 2, "Contractor's Sticker",
                    11.50f,
                    HorizontalTextAlignment.Left,
                    VerticalTextAlignment.Top);


                ///<summary>
                /// 1. Sütun 4-6. Satırlar Arasında yer alan revizyon numaralarının ilgili hücrelere eklenmesi
                /// </summary>
                string [] RevLetter=["C","B","A"];
                for (int i = 4; i <= 6; i++)
                    _wordCellService.SetTextToCell(contentTable, i, 1, RevLetter [i - 4], 12, false, WdParagraphAlignment.wdAlignParagraphCenter);

                ///<summary>
                /// 2. Satırda 2-9 arasındaki hücrelerin birleştirilmesi
                /// </summary>
                for (int i = 2; i <= 6; i++)
                {
                    _wordCellService.MergeCells(contentTable, i, 2, i, 9);
                }


                ///<summary>
                /// 2. Satır 3. Sütundan başlayarak sona doğru 3'er sütun atlayarak 6. Satıra kadar ilgili hücrelerin birleştirilmesi
                /// </summary>
                for (int columnOffset = 3; columnOffset <= 6; columnOffset++)
                {
                    for (int i = 2; i <= 6; i++)
                    {
                        _wordCellService.MergeCells(contentTable, i, columnOffset, i, columnOffset + 2);
                    }
                }

                ///<summary>
                /// 7. Satır 1. Sütuna REV. text değerinin yazdırılması
                /// </summary>
                _wordCellService.SetTextToCell(contentTable, 7, 1, REV, 12, true, WdParagraphAlignment.wdAlignParagraphCenter);

                ///<summary>
                /// 7. Satır 2-9. Sütunların birleştirilmesi ve Text değerinin IssuedForApprovalEN olarak ayarlanması
                /// </summary>
                _wordCellService.MergeAndSetTextToSpesicifCell(contentTable,
                    7, 2, 7, 9,
                    7, 2, IssuedForApprovalEN,
                    11.50f,
                    HorizontalTextAlignment.Center,
                    VerticalTextAlignment.Center);

                ///<summary>
                /// Hücre içeriklerininin iterasyon için liste içerisine tanımlanması
                /// </summary>
                var cellData = new List<(int startColumn,int endColumn,string text)>
                {
                    (3,5,PREPARED),
                    (4,6,CHECKED),
                    (5,7,APPROVED),
                    (6,8,DATE),
                };

                ///<summary>
                /// Liste türünden değişkene atanmış olan hücre birleştirme değerleri ve içeriğinin İTERASYON ile hücrelere eklenmesi
                /// </summary>
                foreach (var (startColumn, endColumn, text) in cellData)
                {
                    _wordCellService.MergeAndSetTextToSpesicifCell(contentTable,
                        7, startColumn, 7, endColumn,
                        7, startColumn, text,
                        11.50f,
                        HorizontalTextAlignment.Center,
                        VerticalTextAlignment.Center);
                }

                ///<summary>
                /// 7. Satırda yer alan hücrelerin satır yüksekliklerinin ayarlanması
                /// </summary>
                for (int i = 1; i <= 6; i++)
                {
                    var cellRange=contentTable.Cell(7,i).Range;
                    cellRange.ParagraphFormat.LineSpacing = 0.70f;
                    cellRange.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                }

                _wordCellService.MergeCells(contentTable, 8, 1, 8, 5);

                string imagePath=@"C:\Users\Metraj Merkezi\Desktop\ÇŞB_Logo.png";
                _wordCellService.InsertImageToCell(contentTable, 8, 1, imagePath, 153, 49);








                string filePath=@"C:\Users\Metraj Merkezi\Desktop\WordYoutubeDenemeler";
                _wordDocumentService.SaveDocumentWithAutoIncrement(_wordDocumentService, document, filePath, "Antet_Deneme");
                MessageBox.Show("Word Dokümanı Başarıyla Oluşturdu.");
            }
            finally
            {
                try
                {
                    if (document != null)
                    {
                        Marshal.FinalReleaseComObject(document);
                        document = null;
                    }

                    try
                    {
                        if (application != null)
                        {
                            var version=application.Version;
                            application.Quit();
                            Marshal.FinalReleaseComObject(application);
                        }
                    }
                    catch (COMException)
                    {

                    }
                    finally
                    {
                        application = null;
                    }
                }
                catch (Exception ex)
                {

                    MessageBox.Show($"Hata Meydana Geldi! {ex.Message}");
                }
                finally
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
        }

        public void MergeCells(Table table,
        int startRow,
        int startColumn,
        int endRow,
        int endColumn)
        {
            var startCell = table.Cell(startRow, startColumn);
            var endCell = table.Cell(endRow, endColumn);
            startCell.Merge(endCell);
        }
    }
}
