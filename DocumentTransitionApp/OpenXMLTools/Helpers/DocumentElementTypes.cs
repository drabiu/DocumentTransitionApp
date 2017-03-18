using System;

namespace OpenXMLTools.Helpers
{
    [Serializable]
    public enum ElementType
    {
        Paragraph,
        BulletList,
        NumberedList,
        Hyperlink,
        Table,
        Picture,
        Sheet,
        Column,
        Row,
        Cell,
        Slide
    }
}
