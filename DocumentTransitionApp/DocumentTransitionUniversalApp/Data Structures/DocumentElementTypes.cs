namespace DocumentTransitionUniversalApp.Data_Structures
{
    public abstract class ElementTypes
    {
    }

    public class WordElementType : ElementTypes
    {
        public static readonly string Paragraph = "[Par]";
        public static readonly string Table = "[Tab]";
        public static readonly string Picture = "[Pic]";
    }

    public class ExcelElementType : ElementTypes
    {
        public static readonly string Sheet = "[Sht]";
    }

    public class PresentationElementType : ElementTypes
    {
        public static readonly string Slide = "[Sld]";
    }
}
