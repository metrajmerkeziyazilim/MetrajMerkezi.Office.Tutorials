using MetrajMerkezi.Office.Word.Abstraction.File;
using MetrajMerkezi.Office.Word.Abstraction.Object._Table;
using MetrajMerkezi.Office.Word.Services.File;
using MetrajMerkezi.Office.Word.Services.Object._Table;
using Ninject.Modules;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetrajMerkezi.Office.Word.DependencyResolver
{
    public class BusinessModule : NinjectModule
    {
        public override void Load()
        {
            Bind<IWordDocumentService>().To<WordDocumentService>();
            Bind<IWordTableService>().To<WordTableService>();
        }
    }
}
