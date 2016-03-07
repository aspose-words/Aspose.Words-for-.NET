' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' Query the delivery (shipping) address using LINQ.
Dim deliveryAddress = From delivery In orderXml.Elements("Address") _
Where (CStr(delivery.Attribute("Type")) = "Shipping") _
'                        Select New With {Key .Name = CStr(delivery.Element("Name")), Key .Country = CStr(delivery.Element("Country")), Key .Zip = CStr(delivery.Element("Zip")), Key .State = CStr(delivery.Element("State")), Key .City = CStr(delivery.Element("City")), Key .Street = CStr(delivery.Element("Street"))}
