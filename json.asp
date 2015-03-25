<script language="vbscript" runat="server">
Function TFunc(arr)

  Response.Write arr

End Function
</script>

<script type="text/jscript" runat="server" language="jscript">

	function toJson(obj)
	{
		var esc = function (string)
		{
			return string
				.split("'").join("\\'")
				.split('"').join('\\"')
				.split('\r\n').join('\\n')
				.split('\n').join('\\n')
				.split('\r').join('\\r');
		};
		
		var getJsonValue = function (value)
		{
			if (typeof value === "object")
			{
				return toJson(value);
			}
			else if (typeof value === "string")
			{
				return '"' + esc(value) + '"';
			}
			else if (typeof value === "date")
			{
				return 'new Date("' + value + '")';
			}
			else
			{
				try
				{
					var convertedArray = VBArray(value).toArray();
					
					var arr = [];

					for (var i = 0; i < convertedArray.length; i++)
						arr.push(getJsonValue(convertedArray[i]));
					
					return "[" + arr.join(",") + "]";
				}
				catch(e)
				{
					return value;	
				}			
			}		
		};
		
		
	
		var result = ["{"];
		
		for (var key in obj)
		{
			result.push('"' + key + '"');
			result.push(":");
			
			result.push(getJsonValue(obj[key]));
			
			result.push(",");
		}
		
		result.pop(); //Removes trailing extra comma
		
		result.push("}")
			
		return result.join("");
	}

	function fromJson(string)
	{
		return eval("(" + string + ")");
	}
	
	TFunc("OI!!!!!!!!!!");

</script>

<%

Class MiniMe

	Public Property Get Nome
		Nome = "AWESO" & vbNewLine & "MENESS"
	End Property

End Class

Class MinhaClasse

	Public Property Get Prop1
		Prop1 = Array(1, 2, #22/05/1988#, 3, 4, new MiniMe, Array(9, 8, 7), True)
	End Property

	Public Property Get Prop2
		Prop2 = "BEH!"
	End Property

	Public Property Get Prop3
		Prop3 = "CEH! ""ou não"""
	End Property

	Public Property Get Prop4
		Set Prop4 = New MiniMe
	End Property
	
End Class

Dim x : Set x = New MinhaClasse

Dim json : json = toJson(x)

Response.Write json
Response.Write "<br><br>"
Response.Write fromJson(json).Prop1(2)
%>
