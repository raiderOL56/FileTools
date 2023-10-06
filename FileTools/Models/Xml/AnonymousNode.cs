namespace FileTools.Models.Xml
{
    public class AnonymousRootNode
    {
        public string NameNodeRoot { get; set; }
        public Dictionary<string, object>? Attributes { get; set; } = new Dictionary<string, object>();
        public Dictionary<string, AnonymousNode> ChildNodes { get; set; } = new Dictionary<string, AnonymousNode>();
        public string Value { get; set; }
    }
    public class AnonymousNode
    {
        public Dictionary<string, object>? Attributes { get; set; } = new Dictionary<string, object>();
        public Dictionary<string, AnonymousNode> ChildNodes { get; set; } = new Dictionary<string, AnonymousNode>();
        public string Value { get; set; }
    }
    public class Obj
    {
        public string Name { get; set; }
        public int Index { get; set; }
    }
}