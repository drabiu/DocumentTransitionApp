namespace OpenXMLTools.Word
{
    public class WordPartRelId
    {
        public string OldId { get; set; }
        public string NewId { get; set; }

        public WordPartRelId(string oldId, string newId)
        {
            OldId = oldId;
            NewId = newId;
        }
    }
}
