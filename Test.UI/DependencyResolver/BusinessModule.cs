using MetrajMerkezi.Office.Word.Abstraction.File;
using MetrajMerkezi.Office.Word.Abstraction.Manipulation;
using MetrajMerkezi.Office.Word.Abstraction.Object._Page;
using MetrajMerkezi.Office.Word.Abstraction.Object._Table;
using MetrajMerkezi.Office.Word.Services.File;
using MetrajMerkezi.Office.Word.Services.Manipulation;
using MetrajMerkezi.Office.Word.Services.Object._Page;
using MetrajMerkezi.Office.Word.Services.Object._Table;
using Ninject.Modules;

namespace Test.UI.DependencyResolver
{
    public class BusinessModule : NinjectModule
    {
        public override void Load()
        {
            Bind<IWordDocumentService>().To<WordDocumentService>();
            Bind<IWordTableService>().To<WordTableService>();
            Bind<IWordRowService>().To<WordRowService>();
            Bind<IWordCellService>().To<WordCellService>();
            Bind<IWordPageSetupService>().To<WordPageSetupService>();
        }
    }
}
