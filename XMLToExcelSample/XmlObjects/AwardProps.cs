using System.Xml.Serialization;

namespace XMLToExcellConverter.XmlObjects
{
    [XmlRoot(ElementName = "AwardProps")]
    public class AwardProps
    {
        [XmlElement(ElementName = "AwardPropRecord")]
        public AwardPropRecord AwardPropRecord { get; set; }
        [XmlAttribute(AttributeName = "xsi", Namespace = "http://www.w3.org/2000/xmlns/")]
        public string Xsi { get; set; }
    }
}
