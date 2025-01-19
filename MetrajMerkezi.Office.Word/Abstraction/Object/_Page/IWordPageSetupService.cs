using Microsoft.Office.Interop.Word;

namespace MetrajMerkezi.Office.Word.Abstraction.Object._Page
{
    public interface IWordPageSetupService
    {
        void SetPageMargins(Section section, float topMargin, float bottomMargin, float leftMargin, float rightMargin);
    }
}
