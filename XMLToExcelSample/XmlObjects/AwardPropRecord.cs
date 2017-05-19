using System.Xml.Serialization;

namespace XMLToExcellConverter.XmlObjects
{

    [XmlRoot(ElementName = "AwardPropRecord")]
    public class AwardPropRecord
    {
        [XmlElement(ElementName = "g_AwardProps")]
        public G_AwardProps G_AwardProps { get; set; }
    }
}
