<script language="javascript" runat="server">

  (function()
	{
		XmlSerializer = function(rootElement)
		{
			return {
				Serialize: function(obj)
				{
					return rootElement.Serialize(obj);
				}
			};
		};

		XmlElement = function(name, propertyName, attributes, elements, fixedValue)
		{
			if (elements !== null && elements !== undefined) elements = new VBArray(elements).toArray();
			if (attributes !== null && attributes !== undefined) attributes = new VBArray(attributes).toArray();

			return {

				Serialize: function(obj, value)
				{
					value = value || fixedValue;
				
					var buff = [];

					buff.push("<");
					buff.push(name || propertyName);
					buff.push(this.SerializeAttributes(obj));
					buff.push(">")

					if (value !== null && value !== undefined)
					{
						buff.push("<![CDATA[");
						buff.push(value.toString());
						buff.push("]]>");
					}
					else if (propertyName)
					{
						buff.push("<![CDATA[");
						buff.push(obj[propertyName]);
						buff.push("]]>");
					}
					else if (elements)
					{
						var i, l;

						for (i = 0, l = elements.length; i < l; i++)
						{
							buff.push(elements[i].Serialize(obj));
						}
					}

					buff.push("</");
					buff.push(name || propertyName);
					buff.push(">");

					return buff.join("");
				},

				SerializeAttributes: function(obj)
				{
					if (!attributes) return "";

					var buff = [];

					var i, l;

					for (i = 0, l = attributes.length; i < l; i++)
					{
						buff.push(attributes[i].Serialize(obj));
					}

					return " " + buff.join(" ");
				}
			};
		};

		XmlAttribute = function(name, propertyName, value)
		{
			return {
				Serialize: function(obj)
				{
					return (name || propertyName) + '="' + (value || obj[propertyName]) + '"'
				}
			};
		};

		XmlArray = function(name, propertyName, elementTemplate)
		{
			return {
				Serialize: function(obj)
				{
					var arr = VBArray(obj[propertyName]).toArray();
					var buff = ["<" + name + ">"];

					for (var i = 0; i < arr.length; i++)
					{
						buff.push(elementTemplate.Serialize(null, arr[i]));
					}

					buff.push("</" + name + ">");

					return buff.join("");
				}
			};
		};

	})();

</script>
