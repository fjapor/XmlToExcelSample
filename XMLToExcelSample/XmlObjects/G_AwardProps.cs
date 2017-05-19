using System.Collections.Generic;
using System.Xml.Serialization;

namespace XMLToExcellConverter.XmlObjects
{

    [XmlRoot(ElementName = "g_AwardProps")]
    public class G_AwardProps
    {
        [XmlElement(ElementName = "entry")]
        public List<Entry> Entry { get; set; }
    }

}
