<script type="text/jscript" runat="server" language="jscript">

	WebApiClient          = {};
	WebApiClientError     = null;
	WebApiClientHttpError = null;

	(function ()
	{
		var url;
		var async;

		var headers = {};

		WebApiClient.DATA_FORMAT_FORM_DATA           = "application/x-www-form-urlencoded";
		WebApiClient.DATA_FORMAT_MULTIPART_FORM_DATA = "multipart/form-data";
		WebApiClient.DATA_FORMAT_JSON                = "application/json";
		WebApiClient.DATA_FORMAT_XML                 = "text/xml";
		WebApiClient.DATA_FORMAT_ANY                 = null;

		WebApiClient.HTTP_VERB_GET    = "GET";
		WebApiClient.HTTP_VERB_POST   = "POST";
		WebApiClient.HTTP_VERB_PUT    = "PUT";
		WebApiClient.HTTP_VERB_PATCH  = "PATCH";
		WebApiClient.HTTP_VERB_DELETE = "DELETE";

		WebApiClient.create = function (apiUrl, isAsync)
		{
			return new _WebApiClient(apiUrl, isAsync);
		};


		function _WebApiClient(apiUrl, isAsync)
		{
			async = isAsync || false;

			if (!apiUrl) throw new Error("The API URL must be specified.");

			url = apiUrl.replace(/\/$/, "");			
		}


		_WebApiClient.prototype.addHeader = function (name, value)
		{
			headers[name] = value;
		};

		_WebApiClient.prototype.removeHeader = function (name)
		{
			delete headers[name];
		};

		//Post methods ===========================================================================================
		_WebApiClient.prototype.post = function (action, parameters)
		{
			return this.postJson(action, parameters, WebApiClient.DATA_FORMAT_JSON);
		};

		_WebApiClient.prototype.postFormData = function (action, data, responseFormat)
		{
			responseFormat = responseFormat || WebApiClient.DATA_FORMAT_FORM_DATA;

			var result = makeRequest(
				action,
				WebApiClient.HTTP_VERB_POST,
				data,
				WebApiClient.DATA_FORMAT_FORM_DATA,
				responseFormat
			);

			return getResponseValue(result, responseFormat);
		};

		_WebApiClient.prototype.postJson = function (action, data, responseFormat)
		{
			responseFormat = responseFormat || WebApiClient.DATA_FORMAT_JSON;

			var result = makeRequest(
				action,
				WebApiClient.HTTP_VERB_POST,
				data,
				WebApiClient.DATA_FORMAT_JSON,
				responseFormat
			);

			return getResponseValue(result, responseFormat);
		};

		_WebApiClient.prototype.postXml = function (action, data, responseFormat)
		{
			responseFormat = responseFormat || WebApiClient.DATA_FORMAT_XML;

			var result = makeRequest(
				action,
				WebApiClient.HTTP_VERB_POST,
				data,
				WebApiClient.DATA_FORMAT_XML,
				responseFormat
			);

			return getResponseValue(result, responseFormat);
		};
		//========================================================================================================

		//Get methods ===========================================================================================
		_WebApiClient.prototype.get = function (action, parameters)
		{
			return this.getJson(action, parameters);
		};

		_WebApiClient.prototype.getJson = function (action, parameters)
		{
			var result = makeRequest(
				action,
				WebApiClient.HTTP_VERB_GET,
				parameters,
				WebApiClient.DATA_FORMAT_FORM_DATA,
				WebApiClient.DATA_FORMAT_JSON
			);

			return getResponseValue(result, WebApiClient.DATA_FORMAT_JSON);
		};

		_WebApiClient.prototype.getXml = function (action, data, responseFormat)
		{
			responseFormat = responseFormat || WebApiClient.DATA_FORMAT_XML;

			var result = makeRequest(
				action,
				WebApiClient.HTTP_VERB_POST,
				data,
				WebApiClient.DATA_FORMAT_XML,
				responseFormat
			);

			return getResponseValue(result, responseFormat);
		};
		//========================================================================================================

		function makeRequest(action, verb, data, requestDataFormat, responseDataFormat)
		{
			requestDataFormat  = requestDataFormat  || WebApiClient.DATA_FORMAT_JSON;
			responseDataFormat = responseDataFormat || WebApiClient.DATA_FORMAT_JSON;

			var actionUrl = getActionUrl(action);


			var xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");

			if (requestDataFormat === WebApiClient.DATA_FORMAT_FORM_DATA || verb === "GET")
			{
				actionUrl += "?" + serializeData(data, WebApiClient.DATA_FORMAT_FORM_DATA);
			}

			xmlHttp.open(verb, actionUrl, async);
			
			xmlHttp.setRequestHeader("Content-Type", getContentType(requestDataFormat));

			if (responseDataFormat != WebApiClient.DATA_FORMAT_ANY)
				xmlHttp.setRequestHeader("Accept", getContentType(responseDataFormat));

			for (var headerName in headers)
				xmlHttp.setRequestHeader(headerName, headers[headerName]);

			try
			{
				xmlHttp.send(serializeData(data, requestDataFormat));
			}
			catch (e)
			{
				throw new WebApiClientError(e.message);
			}

			var responseContentType = xmlHttp.getAllResponseHeaders().match(/Content\-Type\:\s*([\w\/\-]+)/gi)[0];

			return {
				statusCode:   xmlHttp.status,
				statusText:   xmlHttp.statusText,
				responseText: xmlHttp.responseText,
				responseXml:  xmlHttp.responseXml,
				contentType:  responseContentType
			};
		}

		function getContentType(format)
		{
			if (format == WebApiClient.DATA_FORMAT_FORM_DATA)
				return "text/plain";

			if (format == WebApiClient.DATA_FORMAT_JSON)
				return "application/json; charset=utf-8";

			if (format == WebApiClient.DATA_FORMAT_JSON)
				return "text/xml; charset=utf-8";

			throw new RangeError("Invalid format type: " + format);
		}

		function serializeData(data, format)
		{
			if (format == WebApiClient.DATA_FORMAT_FORM_DATA)
				return FormDataSerializer.serialize(data);

			if (format == WebApiClient.DATA_FORMAT_JSON)
				return Json.toJson(data);

			if (format == WebApiClient.DATA_FORMAT_XML)
				return XmlSerializer.serialize(data);

			throw new RangeError("Invalid request format type.");
		}

		function getResponseValue(response, format)
		{
			checkResponse(response);

			if (!response.responseText && !response.responseXml)
				return undefined;

			if (format == WebApiClient.DATA_FORMAT_JSON)
				return Json.fromJson(response.responseText);

			if (format == WebApiClient.DATA_FORMAT_XML)
				return XmlSerializer.deserialize(response.responseXml);

			return response.responseText;
		}

		function getActionUrl(action)
		{
			action = action.replace(/^\//, "");

			return url + "/" + action;
		}

		function checkResponse(response)
		{
			var codeFamily = (~~(response.statusCode / 100)) * 100;

			if (codeFamily === 400 || codeFamily === 500)
				throw new WebApiClientHttpError(response.statusCode, response.statusText);
		}


		//WebApiClientError ====================================================
		WebApiClientError = function (message)
		{
			this.name = "WebApiClientError";
			this.message = message || "Erro ao invocar serviço";
		};
		WebApiClientError.prototype = new Error();
		WebApiClientError.prototype.constructor = WebApiClientError;
		//======================================================================


		//WebApiClientHttpError ================================================
		WebApiClientHttpError = function (httpStatusCode, message)
		{
			this.name = "WebApiClientHttpError";
			this.message = message || "Erro ao invocar serviço";
			this.httpStatusCode = httpStatusCode;
		};
		WebApiClientHttpError.prototype = new WebApiClientError();
		WebApiClientHttpError.prototype.constructor = WebApiClientHttpError;

		WebApiClientHttpError.prototype.toString = function ()
		{
			return "HTTP " + this.httpStatusCode + ": " + this.message;
		};
		//======================================================================
	})();

</script> 