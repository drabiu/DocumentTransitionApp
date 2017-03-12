using DocumentTransitionUniversalApp.Data_Structures;

namespace DocumentTransitionUniversalApp.Helpers
{
    public class TreeElementIcon
    {
        public ElementTypes ElementType { get; private set; }

        public const string ExcelSheet = "ms-appx:///Assets/layers_2_icon&48.png";
        public const string WordParagraph = "ms-appx:///Assets/align_right_icon&48.png";
        public const string PresentationSlide = "ms-appx:///Assets/doc_empty_icon&48.png";

        public TreeElementIcon(ElementTypes elmentType)
        {
            ElementType = elmentType;
        }

        public string GetIcon()
        {
            string result = WordParagraph;

            if (ElementType is WordElementType)
            {
                if (ElementType is WordElementType.ParagraphElementSubType)
                    result = WordParagraph;

                if (ElementType is WordElementType.PictureElementSubType)
                    result = "";

                if (ElementType is WordElementType.TableElementSubType)
                    result = "";
            }
            else if (ElementType is ExcelElementType)
            {
                if (ElementType is ExcelElementType.ColumnElementSubType)
                    result = "";

                if (ElementType is ExcelElementType.RowElementSubType)
                    result = "";

                if (ElementType is ExcelElementType.SheetElementSubType)
                    result = ExcelSheet;

            }
            else if (ElementType is PresentationElementType)
            {
                if (ElementType is PresentationElementType.SlideElementSubType)
                    result = PresentationSlide;
            }

            return result;
        }
    }
}
