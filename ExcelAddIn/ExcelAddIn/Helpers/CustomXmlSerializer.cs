using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace ExcelAddIn.Helpers
{
	public static class CustomXmlSerializer
	{
		public static string SerializeObject<T>(this T toSerialize, XmlRootAttribute xmlRootAttribute)
		{
			XmlSerializer xmlSerializer = new XmlSerializer(toSerialize.GetType(), xmlRootAttribute);
			XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces();
			namespaces.Add("", "");
			using (StringWriter textWriter = new StringWriter())
			{
				using (XmlWriter xmlWriter = XmlWriter.Create(textWriter, new XmlWriterSettings() { OmitXmlDeclaration = true, Indent = true }))
				{
					xmlSerializer.Serialize(xmlWriter, toSerialize, namespaces);
					return textWriter.ToString();
				}
			}
		}
	}
}
