<script type="text/javascript" language="jscript" runat="server">

	VbJsInterop = {};

	(function ()
	{
		var vbJsArrayConverter;

		VbJsInterop.toJsObject = function (vbObj, recursive)
		{
			//Array
			var arr = VbJsInterop.toJsArray(vbObj);

			if (arr) return arr;

			//Date
			if (typeof vbObj === "date" && !(vbObj instanceof Date))
				return VbJsInterop.toJsDate(vbObj);

			//Dictionary
			if (VbJsInterop.getTypeName(vbObj).toUpperCase() === "DICTIONARY")
				return VbJsInterop.convertVbDictionary(vbObj, recursive);

			//Object
			if (recursive && typeof vbObj === "object")
			{
				var result = {};

				for (var key in vbObj)
				{
					result[key] = VbJsInterop.toJsObject(vbObj, recursive);
				}

				return result;
			}

			return vbObj;
		};

		VbJsInterop.toVbObject = function (jsObj)
		{
			if (jsObj === null || jsObj === undefined)
			{
				return undefined;
			}
			else if (jsObj instanceof Array)
			{
				return VbJsInterop.toVbArray(jsObj);
			}
			else if (jsObj instanceof Date)
			{
				return VbJsInterop.toVbDate(jsObj)
			}
			else if (typeof jsObj === "object")
			{
				var result = {};

				for (var key in jsObj)
				{
					result[key] = VbJsInterop.toVbObject(jsObj[key]);
				}

				return result;
			}

			return jsObj; //O objeto JScript é compatível ou é um tipo primitivo
		};

		VbJsInterop.toJsArray = function (vbArray)
		{
			//Verifica se o array já é um array do JScript.
			if (typeof vbArray == "object" && vbArray instanceof [].constructor)
			{
				return vbArray;
			}

			try
			{
				return VBArray(vbArray).toArray();
			}
			catch (e)
			{
				return undefined;
			}
		};

		VbJsInterop.toVbArray = function (jsArray)
		{
			vbJsArrayConverter = vbJsArrayConverter || Create_VbJsInterop_ArrayConverter();


			vbJsArrayConverter.CreateVbArray(jsArray.length);

			for (var i = 0; i < jsArray.length; i++)
			{
				vbJsArrayConverter.SetVBArrayValue(i, jsArray[i]);
			}

			return vbJsArrayConverter.GetVBArray();
		};

		VbJsInterop.toJsDate = function (vbDate)
		{
			var jsDate = new Date(VbJsInterop_GetJsTimestamp(vbDate));

			//Arruma o fuso-horário----------------------------------------------
			var dummy = new Date();
			var timezoneOffset = dummy.getTimezoneOffset();

			var timezoneOffsetHours = timezoneOffset / 60;
			var timezoneOffsetMinutes = timezoneOffset % 60

			jsDate.setHours(jsDate.getHours() + timezoneOffsetHours);
			jsDate.setMinutes(jsDate.getMinutes() + timezoneOffsetMinutes);

			jsDate.getTimezoneOffset = function ()
			{
				return timezoneOffset;
			};
			//--------------------------------------------------------------------

			return jsDate;
		};

		VbJsInterop.toVbDate = function (jsDate)
		{
			var d2 = new Date(jsDate.getTime());

			//Arruma o fuso-horário----------------------------------------------
			var timezoneOffset = jsDate.getTimezoneOffset();

			var timezoneOffsetHours = timezoneOffset / 60;
			var timezoneOffsetMinutes = timezoneOffset % 60

			d2.setHours(d2.getHours() - timezoneOffsetHours);
			d2.setMinutes(d2.getMinutes() - timezoneOffsetMinutes);
			//-------------------------------------------------------------------

			return VbJsInterop_GetVbDate(d2.getTime());
		}

		VbJsInterop.convertVbDictionary = function (dict, recursive)
		{
			var keys  = VbJsInterop.toJsArray(dict.keys());
			var items = VbJsInterop.toJsArray(dict.items());

			var result = {};

			for (var i = 0; i < keys.length; i++)
			{
				result[keys[i]] = recursive ? VbJsInterop.toJsObject(items[i]) : items[i];
			}

			return result;
		};

		VbJsInterop.getTypeName = function (obj)
		{
			if (typeof VbJsInterop_GetTypeName === "undefined")
				return "";

			return VbJsInterop_GetTypeName(obj);
		};

	})();

</script>

<%
Class cVbJsInterop_ArrayConverter

	Dim Json_currentArray_

	Public Sub CreateVBArray(size)
	
		Dim arr()
		ReDim arr(size - 1)
		
		Json_currentArray_ = arr
	
	End Sub
	
	Public Sub SetVBArrayValue(index, value)

		If IsObject(value) Then
			Set Json_currentArray_(index) = value		
		Else		
			Json_currentArray_(index) = value		
		End If	
	
	End Sub
	
	Public Function GetVBArray()
	
		GetVBArray = Json_currentArray_
	
	End Function
	
End Class

Function Create_VbJsInterop_ArrayConverter()

	Set Create_VbJsInterop_ArrayConverter = New cVbJsInterop_ArrayConverter

End Function


Function VbJsInterop_GetJsTimestamp(d)

	VbJsInterop_GetJsTimestamp = DateDiff("s", #1970-01-01#, d) * 1000 '01/01/1970 = Epoch = Data mínima do ECMAScript.

End Function

Function VbJsInterop_GetVbDate(jsMilliseconds)

	VbJsInterop_GetVbDate = DateAdd("s", jsMilliseconds / 1000, #1970-01-01#) '01/01/1970 = Epoch = Data mínima do ECMAScript.

End Function

Function VbJsInterop_GetTypeName(obj)

	VbJsInterop_GetTypeName = TypeName(obj)

End Function

%>