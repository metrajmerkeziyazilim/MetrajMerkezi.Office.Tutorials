using MetrajMerkezi.Office.Word.Abstraction.Manipulation;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetrajMerkezi.Office.Word.Services.Manipulation
{
    public class WordRowService : IWordRowService
    {

        public void SetRowHeight(Table table, int rowIndex, float height)
        {
            try
            {
                if (rowIndex > 0 && rowIndex <= table.Rows.Count)
                {
                    Row row = table.Rows[rowIndex];

                    row.HeightRule = WdRowHeightRule.wdRowHeightExactly;
                    row.Height = height;

                    foreach (Cell cell in row.Cells)
                    {
                        cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom;
                    }
                }
                else
                {
                    Console.WriteLine("Lütfen geçerli bir satır numarası giriniz...");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Satır yüksekliği ayarlanırken bir hata meydana geldi. {ex}");
            }
        }
    }
}
