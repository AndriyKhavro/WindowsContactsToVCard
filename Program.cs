using System.Xml.Linq;
using VCard.Models;

var folder = @"C:\windows_contacts";

var contactFiles = Directory.GetFiles(folder, "*.contact");

var contacts = contactFiles.Select(file => ParseContactFile(file))
    .ToArray();

var contactVCards = new List<string>();

foreach (var contact in contacts.Where(contact => contact.Phones.Any()))
{
    contactVCards.Add(VCard.Helpers.CardHelper.CreateVCard(contact));
}

var fileContent = string.Join(Environment.NewLine, contactVCards);

File.WriteAllText(@"C:\code\windows_contacts.vcf", fileContent);

Contact ParseContactFile(string file)
{
    var fileXml = File.ReadAllText(file);
    var xml = XElement.Parse(fileXml.Replace(":", ""));

    var numbers = ReadProperties("cNumber")
        .Concat(ReadProperties("MSWABMAPIPropTag0x800A001F"));

    var contact = new Contact
    {
        FormattedName = ReadProperty("cFormattedName"),
        FirstName = ReadProperty("cGivenName"),
        LastName = ReadProperty("cFamilyName"),
        Phones = numbers.Select((number, i) => new Phone
        {
            Number = number,
        }).ToList(),
        Email = new List<Email>(),
        Addresses = new List<Address>()
    };

    return contact;

    string? ReadProperty(string name)
    {
        return xml.Descendants(name).FirstOrDefault()?.Value;
    }

    IEnumerable<string> ReadProperties(string name)
    {
        return xml.Descendants(name)
            .Select(node => node.Value);
    }
}