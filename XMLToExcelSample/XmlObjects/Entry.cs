using System.Xml.Serialization;

namespace XMLToExcellConverter.XmlObjects
{
    [XmlRoot(ElementName = "entry")]
    public class Entry
    {
        [XmlElement(ElementName = "Id")]
        public string Id { get; set; }
        [XmlElement(ElementName = "IsElite")]
        public string IsElite { get; set; }
        [XmlElement(ElementName = "GoldCost")]
        public string GoldCost { get; set; }
        [XmlElement(ElementName = "Exp")]
        public string Exp { get; set; }
    }
}
