<script type="text/jscript" runat="server" language="jscript">

	FormDataSerializer = {};

	(function ()
	{
		var vbScriptSupport = null;
	
		FormDataSerializer.serialize = function (obj, prefix)
		{
			if (obj === null || obj === undefined) return "";
		
			var buffer = [];
			
			var jsObject = VbJsInterop.toJsObject(obj);
			
			if (jsObject instanceof Array)
			{
				return serializeArray(array);
			}

			for (var key in obj)
			{
				var value = VbJsInterop.toJsObject(obj[key]);

				if (typeof value === "function")
				{
					continue;
				}
				else if (value instanceof Array)
				{
					buffer.push(serializeArray(value, key));
				}
				else if (typeof value === "object")
				{
					buffer.push(getQSValue((prefix || "") + key, value));
				}
				else
				{
					buffer.push((prefix || "") + key + "=" + getQSValue(key, value));
				}
			}
					
			return buffer.join("&");
		};

		FormDataSerializer.deserialize = function (queryString, isJsObject)
		{
			throw new Error("Método de deserialização não implementado.");
		};
		
		function getQSValue(key, value)
		{
			if (value === null || value === undefined)
			{
				return "";
			}
			else if (value instanceof Date)
			{
				return value.toISOString();
			}
			else if (typeof value === "date")
			{
				var d = convertVbDate(value);

				return getQSValue(d);
			}
			else if (typeof value === "object")
			{
				return FormDataSerializer.serialize(value, key + ".");
			}
			else
			{
				return encodeURIComponent(value);
			}
		}

		function serializeArray(array, key)
		{
			var buff = [];

			key = key || "";

			for (var i = 0; i < array.length; i++)
				buff.push(key + "=" + getQSValue(key, array[i]));

			return buff.join("&");
		}		
		
	})();

</script>