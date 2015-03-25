<script type="text/jscript" runat="server" language="jscript">

	(function ()
	{
		SoapClient = function ()
		{
			var _responseSoapPrefix = null;

			this.url       = null;
			this.namespace = null;
			this.headers   = null;

			this.addHeader = function (headerName)
			{
				this.headers = this.headers || {};

				this.headers[headerName] = new SoapHeader(headerName);

				return this.headers[headerName];
			};

			this.useWsSecurity = function (username, password)
			{
				this.headers = this.headers || {};

				this.headers["UsernameTokenHeader"] = new WsSecurityUsernameTokenHeader(username, password);
			};

			this.invoke = function (operationName, data)
			{
				PerformanceMeter.start("[soap call] " + buildUri(this.url, operationName))


				validateFields(this.url, operationName);

				this.namespace = this.namespace || "http://tempuri.org/";
				var preparedInput = prepareInput(data);

				var message = createSoapEnvelope(this, operationName, preparedInput);

				var response = sendRequest(this.url, operationName, message, this.namespace);

				checkResponse(message, response);


				PerformanceMeter.print("[soap call] " + buildUri(this.url, operationName));

				return response;
			};

			function validateFields(url, operationName)
			{
				if (!url) throw new Error("The service URL must be specified.");
				if (!operationName) throw new Error("The operation name must be specified.");
			}

			function sendRequest(url, operationName, message, namespace)
			{
				var xmlHttp;
				var action;

				try
				{
					xmlHttp = createServerObject("Msxml2.ServerXMLHTTP", "Microsoft.XMLHTTP");

					action = buildUri(namespace, operationName);

					xmlHttp.open("post", url, false);

					xmlHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8");
					xmlHttp.setRequestHeader("SOAPAction", action);

					xmlHttp.send(message);

					var result = getSafeResponse(xmlHttp);
					result.requestMessage = message;
					result.toObject = function (doNotParseValues, isJsObject) { return resultToObject.call(this, operationName, url, action, message, doNotParseValues, isJsObject); };

					return result;
				}
				catch (e)
				{
					throwSoapError(message, xmlHttp, url, action, e.message);
				}
			}

			function resultToObject(operationName, url, action, requestMessage, doNotParseValues, isJsObject)
			{
				var contentNode;

				/*
					First, checks if the SOAP response message is in the pattern
					
					[...]
					<soap:Body>
						<ActionNameResponse>
							<ActionNameResult>[...]</ActionNameResult>
						</ActionNameResponse>
					</soap:Body>
					[...]

					and then in the pattern 

					[...]
					<soap:Body>
						<ActionNameResponse>[...]</ActionNameResponse>
					</soap:Body>
					[...]

					, which are common in default .NET and Java web services.
				*/
				contentNode = this.responseXML.selectSingleNode("//" + operationName + "Result");

				if (contentNode) return XmlSerializer.deserialize(contentNode, isJsObject, doNotParseValues);

				contentNode = this.responseXML.selectSingleNode("//" + operationName + "Response");

				if (contentNode) return XmlSerializer.deserialize(contentNode, isJsObject, doNotParseValues);

				/*
					Serializes the entire soap:Body.
				*/

				//Gets the SOAP namespace prefix, because the XPath in XMLHTTP threw an
				//error when using either the XPath 1 function local-name() and the XPath 2 wildcard *:ElementName .
				//The prefix is store within this client instance until it is not found within the response SOAP message.
				var soapPrefix = getPrefix(this);

				if (soapPrefix)
					contentNode = this.responseXML.selectSingleNode("//" + soapPrefix + ":Body");
				else
					contentNode = this.responseXML.selectSingleNode("//Body");


				if (contentNode) return XmlSerializer.deserialize(contentNode, isJsObject, doNotParseValues);


				//If the soap:Body node is not found, refreshes the soap prefix cache and try again
				soapPrefix = getPrefix(xmlHttp, true);

				if (!soapPrefix) //If the prefix is not found, and the Body element is not found, throws and error.
				{
					throwSoapError(requestMessage, this, url, action, "Invalid SOAP response message");
				}

				resultToObject(operationName, url, action, requestMessage, doNotParseValues, isJsObject);
			}

			function getPrefix(response, reset)
			{
				if (!reset && _responseSoapPrefix) return _responseSoapPrefix;


				var soap11Match = response.responseText.match(/xmlns\:([\w\-]+)=\"http\:\/\/schemas\.xmlsoap\.org\/soap\/envelope/);

				if (soap11Match.length >= 2) return soap11Match[1];


				var soap12Match = response.responseText.match(/xmlns\:([\w\-]+)=\"http:\/\/www\.w3\.org\/2003\/05\/soap-envelope/);

				if (soap12Match.length >= 2) return soap12Match[1];

				
				return null;
			}

		}

		SoapClient.create = function (url, namespace)
		{
			var client = new SoapClient();

			client.url       = url       || client.url;
			client.namespace = namespace || client.namespace;

			return client;
		}

		function SoapError(message)
		{
			this.name = "SoapError";
			this.message = message;
		}
		SoapError.prototype = new Error();
		SoapError.prototype.constructor = SoapError;


		// Helpers --------------------------------------------------------------------------------------------------------------------------------------------
		function replaceAll(str, oldStr, newStr)
		{
			return str.split(oldStr).join(newStr);
		}

		function buildUri(part1, part2)
		{
			if (part1.lastIndexOf("/") != part1.length - 1) part1 += "/";

			return part1 + part2;
		}
		//-----------------------------------------------------------------------------------------------------------------------------------------------------


		//Internals -------------------------------------------------------------------------------------------------------------------------------------------

		function createSoapEnvelope(request, operationName, data)
		{
			var soapEnvelope = 
			[
				'<?xml version="1.0" encoding="utf-8"?>',
				'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">',
				'	{header}',
				'	<soap:Body>',
				'		<{operationName} xmlns="{namespace}">',
				'			{body}',
				'        </{operationName}>',
				'	</soap:Body>',
				'</soap:Envelope>'
			]
			.join("\n");

			soapEnvelope = replaceAll(soapEnvelope, "{header}",        buildHeader(request));
			soapEnvelope = replaceAll(soapEnvelope, "{operationName}", operationName);
			soapEnvelope = replaceAll(soapEnvelope, "{namespace}",     request.namespace);

			var params = [];

			for (var key in data)
			{
				params.push(XmlSerializer.serialize(data[key], key, true));
			}

			soapEnvelope = replaceAll(soapEnvelope, "{body}", params.join(""));


			return soapEnvelope;
		}

		function buildHeader(request)
		{
			if (!request.headers) return "";

			var buffer = [];
			var headers = request.headers;

			buffer.push("<soap:Header>");			

			for (var headerName in headers)
			{
				if (!headers.hasOwnProperty(headerName)) continue;

				buffer.push(headers[headerName].getHeaderString(request.namespace));
			}

			buffer.push("</soap:Header>");

			return buffer.join("");
		}

		function prepareInput(data)
		{
			var arr = VbJsInterop.toJsArray(data);

			if (!arr) return data;

			var result = ArrayToObject(arr);

			for (var key in result)
			{
				if (!result.hasOwnProperty(key)) continue;

				result[key] = prepareInput(result[key]);
			}

			return result;
		}

		function checkResponse(requestMessage, response, url, action)
		{
			var statusFamily = response.statusCode.toString().charAt(0);

			if (statusFamily != "4" && statusFamily != "5") return;

			throwSoapError(requestMessage, response, url, action);
		}

		function throwSoapError(requestMessage, xmlHttp, url, action, errorMessage)
		{
			var response = getSafeResponse(xmlHttp);

			var errorString = [];

			errorString.push("There was an error invoking the web service: " + (errorMessage || ""));

			errorString.push("HTTP " + (response.statusCode || "[no HTTP status code]") + ": " + (response.statusText || "[no HTTP status text]"));

			if (response.responseXML)
			{
				var faultNode = response.responseXML.selectSingleNode("//faultstring");

				if (faultNode)
				{
					errorString.push("\n----------\n");
					errorString.push(faultNode.text);
				}
			}

			errorString.push("\n----------\n");
			errorString.push("Service URL: " + url);
			errorString.push("Invoked action: " + action);

			errorString.push("\n----------\n");
			errorString.push("Request message:");
			errorString.push(requestMessage);

			errorString.push("\n----------\n");
			errorString.push("Response message:");
			errorString.push(response.responseText);

			var error = new SoapError(errorString.join("\n"));

			error.url             = url;
			error.action          = action;
			error.httpStatusCode  = response.statusCode;
			error.httpStatusText  = response.statusText;
			error.requestMessage  = requestMessage;
			error.responseMessage = response.responseText;

			throw error;
		}

		function getSafeResponse(xmlHttp)
		{
			return {
				statusCode:   getResponseProperty(xmlHttp, "status") || getResponseProperty(xmlHttp, "statusCode"),
				statusText:   getResponseProperty(xmlHttp, "statusText"),
				responseText: getResponseProperty(xmlHttp, "responseText"),
				responseXML:  getResponseProperty(xmlHttp, "responseXML")
			};
		}

		function getResponseProperty(xmlHttp, propertyName)
		{
			try
			{
				return xmlHttp[propertyName];
			}
			catch (e) { }

			return null;
		}

		//-----------------------------------------------------------------------------------------------------------------------------------------------------

		function SoapHeader(name)
		{
			this.values = {};

			this.addField = function (fieldName, value)
			{
				this.values[fieldName] = value;

				return this;
			};

			this.getHeaderString = function (namespace)
			{
				var buffer = [];

				buffer.push("<" + name + ' xmlns="' + namespace + '" >');

				for (var fieldName in this.values)
				{
					if (!this.values.hasOwnProperty(fieldName)) continue;

					var field = this.values[fieldName];

					buffer.push('<' + fieldName + '>');
					buffer.push(field);
					buffer.push('</' + fieldName + '>');
					buffer.push('\n');
				}

				buffer.push("</" + name + '>');

				return buffer.join("");
			};
		}

		function WsSecurityUsernameTokenHeader(user, password)
		{
			var template =
			[
				'<wsse:Security soap:mustUnderstand="1" xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd" xmlns:wsu="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">',
				'	<wsse:UsernameToken wsu:Id="UsernameToken">',
				'		<wsse:Username>{username}</wsse:Username>',
				'		<wsse:Password Type="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-username-token-profile-1.0#PasswordText">{password}</wsse:Password>',
				'		<wsse:Nonce EncodingType="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-soap-message-security-1.0#Base64Binary">{nonce}</wsse:Nonce>',
				'		<wsu:Created>{createdTime}</wsu:Created>',
				'	</wsse:UsernameToken>',
				'</wsse:Security>'
			]
			.join("\n");


			this.getHeaderString = function ()
			{
				var createdTime = new Date();

				return (
					template
					.replace("{username}", user)
					.replace("{password}", password)
					.replace("{createdTime}", createdTime.toJSON())
					.replace("{nonce}", createNonce(createdTime))
				);					
			};


			function createNonce(createdTime)
			{
				var rand = Math.round(Math.random() * 1000);
				var created = createdTime.toJSON();

				return Base64.encode(Sha1.hash(rand + created));
			}
		}

	})();

</script>