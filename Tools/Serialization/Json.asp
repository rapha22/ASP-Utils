<script type="text/jscript" runat="server" language="jscript">

	(function ()
	{
		Json = {};

		Json.toJson = function (obj)
		{
			if (obj === null || obj === undefined) return "";

			var jsObj = VbJsInterop.toJsObject(obj);
			
			if (typeof jsObj !== "object" || jsObj instanceof Date || jsObj instanceof Array)
			{
				return getJsonValue(jsObj);
			}
			else
			{
				var result = ["{"];

				for (var key in jsObj)
				{
					try
					{
						var value = VbJsInterop.toJsObject(jsObj[key]);
					}
					catch (e)
					{
						//As propriedades e métodos privados de uma classe também são listados no for-in do JScript, porém
						//ao tentar acessa-las um erro é lançado. Neste caso, continua com a próxima propriedade.
						continue;
					}

					if (typeof value === "function") continue;

					result.push('"' + key + '"');
					result.push(":");

					result.push(getJsonValue(value));

					result.push(",");
				}

				result.pop(); //Removes trailing extra comma

				result.push("}")

				return result.join("");
			}
		};
		
		Json.escapeString = function (str)
		{
			return str
				.split("'").join("'")
				.split('"').join('\\"')
				.split('\r\n').join('\\n')
				.split('\n').join('\\n')
				.split('\r').join('\\r')
				.split('\t').join('\\t');
		};

		Json.fromJson = function (jsonString, isJsObject)
		{
			var parsedObject = safeParse(jsonString);

			if (isJsObject)
				return parsedObject;
			else
				return VbJsInterop.toVbObject(parsedObject);
		};

		function safeParse(jsonString)
		{
			if (jsonString === "" || jsonString === null || jsonString === undefined)
			{
				return undefined;
			}

			//Solução encontrada em http://stackoverflow.com/questions/2583472/regex-to-validate-json

			var unsafeJsonTestRegex = /[^,:{}\[\]0-9.\-+Eaeflnr-u \n\r\t]/;
			var unsafeJsonTestString = jsonString.replace(/"(\\.|[^"\\])*"/g, "");

			if (unsafeJsonTestRegex.test(unsafeJsonTestString))
				throw new Error("A string JSON a ser deserializada não é segura.");

			var isoStringRegex = /"(\d{4})\-(\d{2})\-(\d{2})T(\d{2}):(\d{2}):(\d{2})(\.(\d+))?Z?([\+\-]\d{4})?"/gi

			jsonString = jsonString.replace(isoStringRegex, function (match) { return "new Date(Date.parseISODate(" + match + "))" });

			return eval("(" + jsonString + ")");
		}

		function getJsonValue(value)
		{
			if (value === null || value === undefined)
			{
				return "null";
			}
			else if (value instanceof Array)
			{
				return serializeArray(value);
			}
			else if (value instanceof Date)
			{
				return '"' + value.toISOString() + '"';
			}
			else if (typeof value === "string")
			{
				return '"' + Json.escapeString(value) + '"';
			}
			else if (typeof value === "boolean")
			{
				return value ? "true" : "false";
			}
			else if (typeof value === "object")
			{
				return Json.toJson(value);
			}
			else
			{
				return value;
			}
		}

		function serializeArray(array)
		{
			var buff = [];

			for (var i = 0; i < array.length; i++)
				buff.push(getJsonValue(array[i]));

			return "[" + buff.join(",") + "]";
		}

	})();

</script>