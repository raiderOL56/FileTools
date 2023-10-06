using System.Xml;
using System.Xml.Serialization;
using FileTools.Models.Xml;

namespace FileTools.Services.Xml
{
    public class XmlService
    {
        private readonly XmlDocument _xmlDoc = new XmlDocument();
        private readonly string _fullpath;
        private readonly XmlElement? _rootElement;
        public XmlService(string fullpath)
        {
            try
            {
                _fullpath = fullpath;
                _xmlDoc.Load(fullpath);
                _rootElement = _xmlDoc.DocumentElement;
            }
            catch (Exception)
            {
                throw;
            }
        }
        public T DeserializeNode<T>(string nodeName)
        {
            try
            {
                foreach (XmlNode nodeChild in _rootElement.ChildNodes)
                {
                    XmlNode? xmlNode = GetNode(nodeName, nodeChild);

                    if (xmlNode != null)
                    {
                        XmlSerializer serializer = new XmlSerializer(typeof(T));

                        using (XmlReader xmlReader = new XmlNodeReader(xmlNode))
                            return (T)serializer.Deserialize(xmlReader);
                    }
                }

                return default(T);
            }
            catch (Exception)
            {
                throw;
            }
        }
        public AnonymousNode DeserializeNode(string nodeName)
        {
            try
            {
                foreach (XmlNode nodeChild in _rootElement.ChildNodes)
                {
                    XmlNode? xmlNode = GetNode(nodeName, nodeChild);

                    if (xmlNode != null)
                    {
                        return GetAnonymousNode(xmlNode);
                    }
                }

                return new AnonymousNode();
            }
            catch (Exception)
            {
                throw;
            }
        }
        public T DeserializeXml<T>()
        {
            try
            {
                XmlSerializer xmlSerializer = new XmlSerializer(typeof(T));

                using (FileStream fileStream = new FileStream(_fullpath, FileMode.Open))
                    return (T)xmlSerializer.Deserialize(fileStream);
            }
            catch (Exception)
            {
                throw;
            }
        }
        public AnonymousRootNode DeserializeXml()
        {
            try
            {
                AnonymousRootNode anonymousRootNode = new AnonymousRootNode();
                anonymousRootNode.NameNodeRoot = _rootElement.LocalName;

                if (_rootElement.HasAttributes)
                    foreach (XmlAttribute attribute in _rootElement.Attributes)
                        anonymousRootNode.Attributes[attribute.LocalName] = attribute.Value;
                else
                    anonymousRootNode.Attributes = null;

                bool isEndNode = _rootElement.ChildNodes.Cast<XmlNode>().Any(xmlNode => xmlNode.NodeType == XmlNodeType.Text);

                if (_rootElement.HasChildNodes && !isEndNode)
                {
                    List<Node> nodes = new List<Node>();
                    for (int position = 0; position < _rootElement.ChildNodes.Count; position++)
                    {
                        XmlNode xmlNode = _rootElement.ChildNodes[position];

                        nodes.Add(new Node() { Index = position, Name = xmlNode.LocalName, XmlNode = xmlNode });
                    }

                    foreach (IGrouping<string, Node> groupNodes in nodes.GroupBy(node => node.Name))
                    {
                        List<Node> nodesByName = groupNodes.ToList();

                        for (int position = 0; position < nodesByName.Count(); position++)
                        {
                            Node node = nodesByName[position];
                            if (position > 0)
                                nodes[node.Index].Name = $"{groupNodes.Key} ({position})";
                        }
                    }

                    foreach (Node node in nodes)
                        anonymousRootNode.ChildNodes[node.Name] = GetAnonymousNode(node.XmlNode);
                }
                else
                {
                    XmlNode xmlNode = _rootElement.ChildNodes.Cast<XmlNode>().FirstOrDefault();

                    anonymousRootNode.ChildNodes = null;
                    anonymousRootNode.Value = xmlNode != null ? xmlNode.Value : string.Empty;
                }

                return anonymousRootNode;
            }
            catch (Exception)
            {
                throw;
            }
        }
        private AnonymousNode GetAnonymousNode(XmlNode xmlNode)
        {
            AnonymousNode anonymousXmlObj = new AnonymousNode();

            if (xmlNode.Attributes.Count > 0)
                foreach (XmlAttribute attribute in xmlNode.Attributes)
                    anonymousXmlObj.Attributes[attribute.LocalName] = attribute.Value;
            else
                anonymousXmlObj.Attributes = null;

            bool isEndNode = xmlNode.ChildNodes.Cast<XmlNode>().Any(xmlNode => xmlNode.NodeType == XmlNodeType.Text);
            if (xmlNode.HasChildNodes && !isEndNode)
            {
                List<Node> nodes = new List<Node>();
                for (int position = 0; position < xmlNode.ChildNodes.Count; position++)
                {
                    XmlNode node = xmlNode.ChildNodes[position];

                    nodes.Add(new Node() { Index = position, Name = node.LocalName, XmlNode = node });
                }

                foreach (IGrouping<string, Node> groupNodes in nodes.GroupBy(node => node.Name))
                {
                    List<Node> nodesByName = groupNodes.ToList();
                    for (int position = 0; position < nodesByName.Count(); position++)
                    {
                        Node node = nodesByName[position];
                        if (position > 0)
                            nodes[node.Index].Name = $"{groupNodes.Key} ({position})";
                    }
                }

                foreach (Node node in nodes)
                    anonymousXmlObj.ChildNodes[node.Name] = GetAnonymousNode(node.XmlNode);
            }
            else
            {
                XmlNode xmlNodeFirst = xmlNode.ChildNodes.Cast<XmlNode>().FirstOrDefault();
                anonymousXmlObj.Value = xmlNodeFirst != null ? xmlNodeFirst.Value : string.Empty;
                anonymousXmlObj.ChildNodes = null;
            }

            return anonymousXmlObj;
        }
        private XmlNode GetNode(string nodeName, XmlNode xmlNode)
        {
            try
            {
                if (xmlNode.Name.Equals(nodeName))
                    return xmlNode;

                foreach (XmlNode nodeChild in xmlNode.ChildNodes)
                {
                    XmlNode result = GetNode(nodeName, nodeChild);

                    if (result != null)
                        return result;
                }

                return null;
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}