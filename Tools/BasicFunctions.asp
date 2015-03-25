<script type="text/jscript" runat="server" language="jscript">

	//Verifica se um objeto está vazio
	//noEmptyStringCheck: caso seja false ou não seja especificado, retorna true se o valor for uma string vazia ("")
	function IsNotEmpty(value, noEmptyStringCheck)
	{
		if (!noEmptyStringCheck && typeof value === "string")
		{
			return value !== "";
		}

		return value != null;
	}

	//Retorna o primeiro valor não-vazio da lista de parâmetros (similar ao IsNull ou Coalesce do SQL Server)
	//Recebe qualquer número parâmetros (ex.: NEmpty(a, b, c); NEmpty(); NEmpty(a); NEmpty(a, b, c, e, f, g, h, i);
	function NEmpty()
	{
		for (var i = 0; i < arguments.length; i++)
		{
			if (IsNotEmpty(arguments[i])) return arguments[i];
		}

		return undefined;
	}

	//Retorna o primeiro valor não-vazio de um array
	//ex.: (VBScript) NEmpty2(Array("", Null Empty, 3, Nothing)) retorna 3  |  (JScript) NEmpty2(["", null undefined, 3]) retorna 3 
	function NEmpty2(arr)
	{
		return NEmpty.apply(this, VbJsInterop.toJsArray(arr));
	}

	function DebugWriteLine(str)
	{
		Debug.writeln(str);
	}

	//Converte um Array do VBScript em forma de dicionário (ex.: Array("chave1", valor1, "chave2", valor2) ) em um objeto.
	function ArrayToObject(arr)
	{
		var jsArray = VbJsInterop.toJsArray(arr);

		if (!jsArray) throw new TypeError("O parâmetro deve ser um array.");

		var result = {};
		var key = null;

		for (var i = 0; i < jsArray.length; i++)
		{
			if (i % 2 == 0)
				key = jsArray[i];
			else
				result[key] = jsArray[i];
		}

		return result;
	}

	//Cria um objeto com os parâmetros no formato ("chave1", valor1, "chave2", valor2)
	function BuildObject(/*...*/)
	{
		var result = {};

		for (var i = 0; i < arguments.length; i += 2)
		{
			result[arguments[i]] = arguments[i + 1];
		}

		return result;
	}



	function createServerObject()
	{
		for (var i = 0; i < arguments.length; i++)
		{
			try
			{
				var obj = new ActiveXObject(arguments[i]);

				if (obj) return obj;
			}
			catch (e) { }
		}

		return null;
	}

	function GetProperty(obj, propertyName)
	{
		return obj[propertyName];
	}

	//Funções de extensão do JS =====================================================================================================
	String.prototype.padLeft = function (padChar, charCount)
	{
		charCount = Number(charCount);

		var newStr = this.substring(0);

		while (newStr.length < charCount)
			newStr = padChar + newStr;

		return newStr;
	}
	
	Date.prototype.toISOString = function ()
	{
		var values =
		{
			year           : this.getFullYear(),
			month          : this.getMonth() + 1,
			day            : this.getDate(),
			hours          : this.getHours(),
			minutes        : this.getMinutes(),
			seconds        : this.getSeconds(),
			milliseconds   : this.getMilliseconds(),
			timezoneOffset : ~~(this.getTimezoneOffset() / 60),
			timezoneSignal : this.getTimezoneOffset() < 0 ? "+" : "-"
		};

		var format = "{year}-{month:2}-{day:2}T{hours:2}:{minutes:2}:{seconds:2}.{milliseconds}{timezoneSignal}{timezoneOffset:2}00";

		var isoDate = format.replace(/\{(.*?)\}/gi, function (match, group)
		{
			var pad = group.split(":");

			var value = values[pad[0]];

			if (pad.length > 1)
				value = String(value).padLeft("0", pad[1]);

			return value;				
		});

		return isoDate;
	};

	Date.isISODateString = function (strDate)
	{
		return /^(\d{4})\-(\d{2})\-(\d{2})T(\d{2}):(\d{2}):(\d{2})(\.(\d+))?Z?([\+\-]\d{4})?$/gi.test(strDate);
	};

	Date.parseISODate = function (strDate)
	{
		if (Date.isISODateString(strDate))
		{
			var parts = strDate.split("T");

			var dateParts = parts[0].split("-");
			var timeParts = parts[1].split(":");

			var lastPart = timeParts[2].split(".");

			var year         = parseInt(dateParts[0], 10);
			var month        = parseInt(dateParts[1], 10);
			var day          = parseInt(dateParts[2], 10);
			var hours        = parseInt(timeParts[0], 10);
			var minutes      = parseInt(timeParts[1], 10);
			var seconds      = parseInt(timeParts[2], 10);
			var milliseconds = 0;
			var timezone     = null;

			if (lastPart[1])
			{
				milliseconds = parseInt(lastPart[1], 10);
				timezone = lastPart[1].substr(1);

				if (/^[\-\+][0-9]{4}$/.test(timezone))
				{
					var signal = timezone.charAt(0) == "+" ? "-" : "+";

					var tzHours = parseInt(signal + timezone.substring(1, 3), 10);
					var tzMinutes = parseInt(signal + timezone.substring(3, 6), 10);

					hours += tzHours;
					minutes += tzMinutes;
				}
			}

			return Date.UTC(year, month - 1, day, hours, minutes, seconds, milliseconds);
		}

		throw new Error("Invalid ISO date format");
	};

	Date.prototype.toJSON = function ()
	{
		var values =
		{
			year:         this.getUTCFullYear(),
			month:        this.getUTCMonth() + 1,
			day:          this.getUTCDate(),
			hours:        this.getUTCHours(),
			minutes:      this.getUTCMinutes(),
			seconds:      this.getUTCSeconds(),
			milliseconds: this.getUTCMilliseconds()
		};

		var format = "{year}-{month:2}-{day:2}T{hours:2}:{minutes:2}:{seconds:2}.{milliseconds}Z";

		var isoDate = format.replace(/\{(.*?)\}/gi, function (match, group)
		{
			var pad = group.split(":");

			var value = values[pad[0]];

			if (pad.length > 1)
				value = String(value).padLeft("0", pad[1]);

			return value;
		});

		return isoDate;
	};

	(function ()
	{
		var originalFunc = Date.parse;		

		Date.parse = function (strDate)
		{
			var isoDateRegex = /^(\d{4})\-(\d{2})\-(\d{2})T(\d{2}):(\d{2}):(\d{2})(\.(\d+))?Z?([\+\-]\d{4})?$/gi;

			if (isoDateRegex.test(strDate))
			{
				return Date.parseISODate(strDate);
			}

			return originalFunc(strDate);
		};
	})();

	Array.prototype.contains = function (item)
	{
		for (var i = 0; i < this.length; i++)
			if (this[i] === item) return true;

		return false;
	};

	if (typeof Function.prototype.bind !== "undefined")
	{
		Function.prototype.bind = function (thisObject)
		{
			var self = this;

			return function ()
			{
				return self.apply(thisObject, arguments);
			};
		};
	}
	//===============================================================================================================================

</script>

<script runat="server" language="vbscript">

	Function GetJsTimestamp(d)

		GetJsTimestamp = DateDiff("s", #1970-01-01#, d) * 1000

	End Function

	Function GetVbDate(jsMilliseconds)

		GetVbDate = DateAdd("s", jsMilliseconds / 1000, #1970-01-01#)

	End Function

</script>

