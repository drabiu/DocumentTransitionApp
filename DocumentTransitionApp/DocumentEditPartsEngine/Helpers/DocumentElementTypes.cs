using System;

namespace DocumentEditPartsEngine.Helpers
{
    [Serializable]
    public abstract class ElementTypes
    {
    }

    public class WordElementType : ElementTypes
    {
        public class ParagraphElementSubType : ElementTypes
        {
            public const string Prefix = "[Par]";
        }

        public class TableElementSubType : ElementTypes
        {
            public const string Prefix = "[Tab]";
        }

        public class PictureElementSubType : ElementTypes
        {
            public const string Prefix = "[Pic]";
        }
    }

    public class ExcelElementType : ElementTypes
    {
        public class SheetElementSubType : ElementTypes
        {
            public const string Prefix = "[Sht]";
        }

        public class RowElementSubType : ElementTypes
        {
            public const string Prefix = "[Row]";
        }

        public class ColumnElementSubType : ElementTypes
        {
            public const string Prefix = "[Col]";
        }
    }

    public class PresentationElementType : ElementTypes
    {
        public class SlideElementSubType : ElementTypes
        {
            public const string Prefix = "[Sld]";
        }

    }
}
