using DocumentTransitionUniversalApp.TransitionAppServices;

namespace DocumentTransitionUniversalApp.Helpers
{
    public class TreeElementIcon
    {
        public ElementType ElementType { get; private set; }

        public const string ExcelSheet = "ms-appx:///Assets/layers_2_icon&48.png";
        public const string WordParagraph = "ms-appx:///Assets/align_right_icon&48.png";
        public const string PresentationSlide = "ms-appx:///Assets/doc_empty_icon&48.png";
        public const string Default = "ms-appx:///Assets/cancel_icon & 48.png";
        public const string ExcelRow = "ms-appx:///Assets/checkbox_unchecked_icon&48.png";
        public const string ExcelCell = "ms-appx:///Assets/table_selection_row.png";

        public TreeElementIcon(ElementType elmentType)
        {
            ElementType = elmentType;
        }

        public string GetIcon()
        {
            switch (ElementType)
            {
                case ElementType.Paragraph:
                    return WordParagraph;
                case ElementType.Picture:
                    return "";
                case ElementType.Table:
                    return "";
                case ElementType.Sheet:
                    return ExcelSheet;
                case ElementType.Row:
                    return ExcelRow;
                case ElementType.Column:
                    return "";
                case ElementType.Cell:
                    return ExcelCell;
                case ElementType.Slide:
                    return PresentationSlide;
                default:
                    return Default;
            }
        }
    }
}
