<script type="text/jscript" runat="server" language="jscript">

	XmlSerializer = {};

	(function ()
	{
		XmlSerializer.serialize = function (obj, rootElementName, ommitXmlDeclaration)
		{
			return new _XmlSerializer().serialize(obj, rootElementName, ommitXmlDeclaration);
		};

		XmlSerializer.deserialize = function (obj, isJsObject, doNotParseValues)
		{
			return new _XmlSerializer().deserialize(obj, isJsObject, doNotParseValues);
		};

		//_XmlSerializer ===================================================================================================
		function getType(obj)
		{
			return VbJsInterop.getTypeName(obj) || typeof obj;
		}

		//MSXML Constants -----------------------------------
		//http://msdn.microsoft.com/en-us/library/windows/desktop/ms753745(v=vs.85).aspx
		var MSXML_NODE_ELEMENT                = 1;
		var MSXML_NODE_ATTRIBUTE              = 2;
		var MSXML_NODE_TEXT                   = 3;
		var MSXML_NODE_CDATA_SECTION          = 4; 
		var MSXML_NODE_ENTITY_REFERENCE       = 5;
		var MSXML_NODE_ENTITY                 = 6;
		var MSXML_NODE_PROCESSING_INSTRUCTION = 7;
		var MSXML_NODE_COMMENT                = 8;
		var MSXML_NODE_DOCUMENT               = 9;
		var MSXML_NODE_DOCUMENT_TYPE          = 10;
		var MSXML_NODE_DOCUMENT_FRAGMENT      = 11;
		var MSXML_NODE_NOTATION               = 12;
		//---------------------------------------------------

		function _XmlSerializer()
		{
			this.xmlWriter = new XmlWriter();
		}

		_XmlSerializer.prototype.serialize = function (obj, rootElementName, ommitXmlDeclaration)
		{
			if (rootElementName === undefined)
				rootElementName = rootElementName || getType(obj);

			if (obj === null || obj === undefined) return "";

			var jsObj = VbJsInterop.toJsObject(obj);

			if (!ommitXmlDeclaration)
				this.xmlWriter.writeXmlDeclarationTag();

			if (rootElementName) this.xmlWriter.openElement(rootElementName);

			this.writeXmlValue(jsObj);

			if (rootElementName) this.xmlWriter.closeElement();

			return this.xmlWriter.toString();
		};

		_XmlSerializer.prototype.deserialize = function (xml, isJsObject, doNotParseValues)
		{
			if (!xml) return undefined;

			var xmlDom;

			if (typeof xml === "string")
			{
				xmlDom = new ActiveXObject("Msxml2.DOMDocument.6.0");
				xmlDom.async = false;				

				if (!xmlDom.loadXML(xml))
					throw new Error("Error parsing XML string:\n" + xmlDom.parseError.ErrorCode + ": " + xmlDom.parseError.reason);
			}
			else
			{
				xmlDom = xml;
			}

			var result;

			if (xmlDom.nodeType === MSXML_NODE_DOCUMENT)
			{
				var childNodes = xmlDom.childNodes;

				for (var i = 0; i < childNodes.length; i++)
				{
					if (childNodes[i].nodeType !== MSXML_NODE_ELEMENT && childNodes[i] !== MSXML_NODE_ATTRIBUTE)
					{
						continue;
					}

					result = parseNode(childNodes[i], doNotParseValues);
				}
			}
			else
			{
				result = parseNode(xmlDom, doNotParseValues);
			}

			if (!isJsObject)
				return VbJsInterop.toVbObject(result);

			return result;
		};

		function parseNode(node, doNotParseValues)
		{
			if (node.hasChildNodes)
			{
				var firstChild = node.firstChild;

				if (node.childNodes.length == 1 && (firstChild.nodeType == MSXML_NODE_TEXT || firstChild.nodeType == MSXML_NODE_CDATA_SECTION))
				{
					var value = firstChild.nodeValue;

					if (!doNotParseValues)
					{
						var isoDateStringRegex = /(\d{4})\-(\d{2})\-(\d{2})T(\d{2}):(\d{2}):(\d{2})(\.(\d+))?Z?([\+\-]\d{4})?/gi;

						if (isoDateStringRegex.test(value))
							return new Date(Date.parse(value));

						if (!isNaN(value))
							return parseFloat(value);

						if (value.toLowerCase() === "true" || value.toLowerCase() === "false")
							return (value.toLowerCase() === "true");
					}

					return value;
				}
				else
				{
					var similarChildren = node.selectNodes("./" + firstChild.nodeName);

					if (similarChildren.length > 1)
					{
						var result = [];

						for (var i = 0; i < similarChildren.length; i++)
						{
							result.push(parseNode(similarChildren[i], doNotParseValues));
						}

						return result;
					}
					else
					{
						var result = {};

						for (var i = 0; i < node.childNodes.length; i++)
						{
							var child = node.childNodes[i];

							if (child.nodeType !== MSXML_NODE_ELEMENT && child.nodeType !== MSXML_NODE_ATTRIBUTE)
								continue;

							result[child.nodeName] = parseNode(child, doNotParseValues);
						}

						return result;
					}
				}
			}
		}

		_XmlSerializer.prototype.writeXmlValue = function (value)
		{
			if (value === null || value === undefined)
				return;

			if (value instanceof Array)
			{
				this.writeArray(value);
			}
			else if (value instanceof Date)
			{
				this.xmlWriter.write(value.toJSON());
			}
			else if (typeof value === "string")
			{
				this.xmlWriter.write('<![CDATA[' + value + ']]>');
			}
			else if (typeof value === "object")
			{
				this.writeObject(value);
			}
			else
			{
				this.xmlWriter.write(value);
			}
		}

		_XmlSerializer.prototype.writeObject = function (obj)
		{
			for (var key in obj)
			{
				var value;

				try
				{
					value = VbJsInterop.toJsObject(obj[key]);
				}
				catch (e)
				{
					//As propriedades e métodos privados de uma classe também são listados no for-in do JScript, porém
					//ao tentar acessa-las um erro é lançado. Neste caso, continua com a próxima propriedade.
					continue;
				}

				if (typeof value === "function") continue;

				this.xmlWriter.openElement(key);

				this.writeXmlValue(value);

				this.xmlWriter.closeElement();
			}
		}

		_XmlSerializer.prototype.writeArray = function (arr)
		{
			for (var i = 0; i < arr.length; i++)
			{
				this.xmlWriter.openElement(getType(arr[i]));

				this.writeXmlValue(arr[i]);

				this.xmlWriter.closeElement();
			}
		}
		//==================================================================================================================


		//XmlWriter =========================================================================================
		function XmlWriter()
		{
			this.buffer = [];
			this.tagStack = [];
		}

		XmlWriter.prototype.writeXmlDeclarationTag = function ()
		{
			this.buffer.push('<?xml version="1.0" encoding="Utf-8"?>');
		};		

		XmlWriter.prototype.openElement = function (elementName)
		{
			var lineNumber = this.buffer.push("<" + elementName) - 1;

			this.tagStack.push({ tagName: elementName, lineNumber: lineNumber });
		};

		XmlWriter.prototype.closeElement = function ()
		{
			var tagInfo = this.tagStack.pop();

			this.buffer[tagInfo.lineNumber] += ">";

			this.buffer.push("</" + tagInfo.tagName + ">");
		};

		XmlWriter.prototype.writeAttribute = function (attributeName, value)
		{
			var tagInfo = this.tagStack[this.tagStack.length - 1];

			this.buffer[tagInfo.lineNumber] += " " + attributeName + '="' + value +'"';
		};

		XmlWriter.prototype.write = function (str)
		{
			this.buffer.push(str);
		};

		XmlWriter.prototype.toString = function ()
		{
			return this.buffer.join("");
		};
		//===================================================================================================

	})();

</script>