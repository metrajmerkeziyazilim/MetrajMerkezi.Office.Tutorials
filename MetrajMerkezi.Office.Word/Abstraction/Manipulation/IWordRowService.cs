using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetrajMerkezi.Office.Word.Abstraction.Manipulation
{
    public interface IWordRowService
    {
        void SetRowHeight(Table table,int rowIndex,float height);
    }
}
