using System.Xml;

namespace FileTools.Models.Xml
{
    public class Node
    {
        public int Index { get; set; }
        public string Name { get; set; }
        public XmlNode XmlNode { get; set; }
    }
}