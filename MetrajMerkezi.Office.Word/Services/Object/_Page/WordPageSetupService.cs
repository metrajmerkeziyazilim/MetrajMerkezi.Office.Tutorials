using MetrajMerkezi.Office.Word.Abstraction.Object._Page;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetrajMerkezi.Office.Word.Services.Object._Page
{
    public class WordPageSetupService : IWordPageSetupService
    {
        public void SetPageMargins(Section section, float topMargin, float bottomMargin, float leftMargin, float rightMargin)
        {
            section.PageSetup.TopMargin = topMargin;
            section.PageSetup.BottomMargin = bottomMargin;
            section.PageSetup.LeftMargin = leftMargin;
            section.PageSetup.RightMargin = rightMargin;
        }
    }
}
