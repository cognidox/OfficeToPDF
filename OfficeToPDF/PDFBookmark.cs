using System.Collections.Generic;

namespace OfficeToPDF
{
    class PDFBookmark
    {
        public int page { get; set; }
        public string title { get; set; }
        public List<PDFBookmark> children { get; set; }   
    }
}
