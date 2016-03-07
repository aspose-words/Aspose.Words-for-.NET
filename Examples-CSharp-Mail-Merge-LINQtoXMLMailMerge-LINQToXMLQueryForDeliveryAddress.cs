// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
var deliveryAddress =
from delivery in orderXml.Elements("Address")
where ((string)delivery.Attribute("Type") == "Shipping")
select new
{
    Name = (string)delivery.Element("Name"),
    Country = (string)delivery.Element("Country"),
    Zip = (string)delivery.Element("Zip"),
    State = (string)delivery.Element("State"),
    City = (string)delivery.Element("City"),
    Street = (string)delivery.Element("Street")
};
