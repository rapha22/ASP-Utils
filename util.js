(function ()
{
	function safelyAddProperty(obj, propertyName, value)
	{
		if (typeof obj[propertyName] == "undefined")
		{
			obj[propertyName] = value;
		}
	}

	/***********************
	Localization
	***********************/
	var Localization =
	{
		current: "br",

		getCurrent: function ()
		{
			return this.localizations[this.current];
		},

		localizations:
		{
			br:
			{
				months:
				[
					"janeiro",
					"fevereiro",
					"março",
					"abril",
					"maio",
					"junho",
					"julho",
					"agosto",
					"setembro",
					"outubro",
					"novembro",
					"dezembro"
				],

				weekdays:
				[
					"domingo",
					"segunda-feira",
					"terça-feira",
					"quarta-feira",
					"quinta-feira",
					"sexta-feira",
					"sábado"
				],

				currencySymbol: "R$",
				decimalSeparator: ",",
				thousandsSeparator: "."
			},

			en:
			{
				months:
				[
					"january",
					"february",
					"march",
					"april",
					"may",
					"june",
					"july",
					"august",
					"september",
					"october",
					"november",
					"december"
				],

				weekdays:
				[
					"sunday",
					"monday",
					"tuesday",
					"wednesday",
					"thursday",
					"friday",
					"saturnday"
				],

				currencySymbol: "US$",
				decimalSeparator: ".",
				thousandsSeparator: ","
			}
		}
	};

	/***********************
	String
	***********************/
	safelyAddProperty(String.prototype, "replaceAll", function (oldString, newString)
	{
		return this.split(oldString).join(newString);
	});

	safelyAddProperty(String.prototype, "trim", function (oldString, newString)
	{
		return this.replace(/^\s+|\s+$/g, '');
	});

	safelyAddProperty(String.prototype, "trimLeft", function (oldString, newString)
	{
		return this.replace(/^\s+/, '');
	});

	safelyAddProperty(String.prototype, "trimRight", function (oldString, newString)
	{
		return this.replace(/\s+$/, '');
	});

	safelyAddProperty(String.prototype, "padLeft", function (padString, length)
	{
		var str = this;

		while (str.length < length)
			str = padString + this;

		return str;
	});

	safelyAddProperty(String.prototype, "padRight", function (padString, length)
	{
		var str = this;

		while (str.length < length)
			str = this + padString;

		return str;
	});

	safelyAddProperty(String, "isNullOrWhitespace", function (str)
	{
		return Boolean(!str || !str.trim());
	});

	safelyAddProperty(String.prototype, "interpolate", function (data)
	{
		var result = this;

		//Matches ${propertyName}
		result = this.replace(/\$\{([\s\S]*?)\}/gi, function (match, capture)
		{
			return data[capture];
		});

		//Matches $!{expression}
		result = result.replace(/\$\!\{([\s\S]*?)\}/gi, function (match, capture)
		{
			return eval(capture) || "";
		});

		return result;
	});

	//Adiciona um nome a mais para o método interpolate
	safelyAddProperty(String.prototype, "format", String.prototype.interpolate);


	/***********************
	Date
	***********************/
	safelyAddProperty(Date.prototype, "getMonthName", function ()
	{
		return Localization.getCurrent().months[this.getMonth()];
	});

	safelyAddProperty(Date.prototype, "getAbbreviatedMonthName", function ()
	{
		return this.getMonthName().substr(0, 3);
	});

	safelyAddProperty(Date.prototype, "getWeekDayName", function ()
	{
		return Localization.getCurrent().weekdays[this.getDay()];
	});

	safelyAddProperty(Date.prototype, "getAmPm", function ()
	{
		return this.getHours() <= 12 ? "AM" : "PM";
	});

	safelyAddProperty(Date.prototype, "getHours12", function ()
	{
		var h = this.getHours();

		return h <= 12 ? h : h - 12;
	});

	safelyAddProperty(Date.prototype, "getTwoDigitYear", function ()
	{
		var year = this.getYear();

		return year - ~~(year / 100) * 100;
	});

	safelyAddProperty(Date.prototype, "getMilleniumYear", function ()
	{
		var year = this.getFullYear();

		return ~~(year / 1000) * 1000;
	});

	safelyAddProperty(Date.prototype, "format", function (strFormat)
	{
		//Nota: não sei por que, mas dar um return direto com o código abaixo faz a função retornar undefined.
		//Tive que armazenar o resultado numa variável antes (result).

		var result =
			strFormat
			.replaceAll("yyyy", this.getFullYear().toString().padLeft("0", 4))
			.replaceAll("YYYY", this.getFullYear().toString().padLeft("0", 4))
			.replaceAll("yy", this.getTwoDigitYear().toString().padLeft("0", 2))
			.replaceAll("YY", this.getTwoDigitYear().toString().padLeft("0", 2))
			.replaceAll("MMMM", this.getMonthName())
			.replaceAll("MMM", this.getAbbreviatedMonthName())
			.replaceAll("MM", (this.getMonth() + 1).toString().padLeft("0", 2))
			.replaceAll("dddd", this.getWeekDayName())
			.replaceAll("DDDD", this.getWeekDayName())
			.replaceAll("ddd", this.getWeekDayName().substr(0, 3))
			.replaceAll("DDD", this.getWeekDayName().substr(0, 3))
			.replaceAll("dd", this.getDate().toString().padLeft("0", 2))
			.replaceAll("DD", this.getDate().toString().padLeft("0", 2))
			.replaceAll("HH", this.getHours().toString().padLeft("0", 2))
			.replaceAll("hh", this.getHours12().toString().padLeft("0", 2))
			.replaceAll("mm", this.getMinutes().toString().padLeft("0", 2))
			.replaceAll("ss", this.getSeconds().toString().padLeft("0", 2))
			.replaceAll("SS", this.getSeconds().toString().padLeft("0", 2))
			.replaceAll("AP", this.getAmPm())
			.replaceAll("ap", this.getAmPm());

		return result;
	});

	safelyAddProperty(Date, "fromJsonDate", function (jsonDate)
	{
		return new Date(parseInt(jsonDate.substr(6), 10));
	});

	safelyAddProperty(Date, "parseFormat", function (dateString, format)
	{
		if (!dateString || !format)
		{
			return undefined;
		}

		//Obtem as diferentes partes da string de formato
		var formatParts = [];
		var partHolder = "";

		for (var i = 0; i < format.length; i++)
		{
			var c = format.charAt(i);

			//Caso o caractere atual seja diferente do último armazenado, isso significa que uma outra parte do formato começou
			if (partHolder && c != partHolder.charAt(partHolder.length - 1))
			{
				//Caso seja "AP" (formato para o indicador AM/PM), a letra pode ser diferente
				if (partHolder.toUpperCase() != "A" && c.toUpperCase() != "P")
				{
					formatParts.push(partHolder);
					partHolder = c;

					continue;
				}
			}

			partHolder += c;
		}

		formatParts.push(partHolder);


		var year = 0;
		var month = 0;
		var day = 0;
		var hours = 0;
		var minutes = 0;
		var seconds = 0;

		for (var i = 0; i < formatParts.length; i++)
		{
			var part = formatParts[i];

			//Caso a string não seja um placeholder ("dd", "MM", etc), apenas remove o separador da string de data
			if (!/\w+/.test(part))
				dateString = dateString.replace(part, "");

			switch (part)
			{
				case "yyyy":
				case "YYYY":
					year = parseInt(dateString.substr(0, 4), 10);
					dateString = dateString.substr(4);
					break;

				case "YY":
				case "yy":
					year = parseInt(dateString.substr(0, 2), 10);

					var now = new Date();
					var millenium = ~ ~(now.getFullYear() / 100) * 100;

					if (year > now.getTwoDigitYear())
						year += millenium - 100;
					else
						year += millenium;

					dateString = dateString.substr(3);
					break;

				case "MM":
					month = parseInt(dateString.substr(0, 2), 10);
					dateString = dateString.substr(2);
					break;

				case "dd":
				case "DD":
					day = parseInt(dateString.substr(0, 2), 10);
					dateString = dateString.substr(2);
					break;

				case "HH":
				case "hh":
					hours = parseInt(dateString.substr(0, 2), 10);
					dateString = dateString.substr(2);
					break;

				case "mm":
					minutes = parseInt(dateString.substr(0, 2), 10);
					dateString = dateString.substr(2);
					break;

				case "SS":
				case "ss":
					seconds = parseInt(dateString.substr(0, 2), 10);
					dateString = dateString.substr(2);
					break;

				case "AP":
				case "ap":
					var amPm = dateString.substr(0, 2);

					if (amPm.toUpperCase() == "PM")
						hours += 12;

					dateString = dateString.substr(2);
					break;
			}
		}

		return new Date(year, month - 1, day, hours, minutes, seconds, 0);
	});

	safelyAddProperty(Date, "today", function (jsonDate)
	{
		var now = new Date();

		return new Date(now.getFullYear(), now.getMonth(), now.getDate());
	});


	/***********************
	Number/Math
	***********************/
	safelyAddProperty(Number.prototype, "toNumberString", function (decimalDigits)
	{
		var loc = Localization.getCurrent();

		decimalDigits = decimalDigits || 0;

		var str = this.toFixed(decimalDigits); 											//Converte o número para uma string com a quantidade de dígitos especificada
		str = str.replace(".", loc.decimalSeparator); 									//Troca o ponto de número decimal (EUA) por vírgula (BR)
		str = str.replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1" + loc.thousandsSeparator); 	//Adiciona pontos de milhar

		return str;
	});

	//Math.round com parâmetro para precisão
	(function ()
	{
		var roundFunc = Math.round;

		Math.round = function (number, precision)
		{
			if (precision === undefined)
			{
				return roundFunc(number);
			}
			else
			{
				return parseFloat(number.toFixed(precision));
			}
		};
	})();

	safelyAddProperty(Number.prototype, "round", function (precision)
	{
		return Math.round(this, precision || 0);
	});


	/***********************
	Array
	***********************/
	safelyAddProperty(Array.prototype, "removeFirst", function (value)
	{
		var index = this.indexOf(value);

		if (index >= 0)
		{
			this.splice(index, 1);
		}

		return this;
	});

	safelyAddProperty(Array.prototype, "removeLast", function (value)
	{
		var index = this.lastIndexOf(value);

		if (index >= 0)
		{
			this.splice(index, 1);
		}

		return this;
	});

	safelyAddProperty(Array.prototype, "removeAll", function (value)
	{
		var index = this.indexOf(value);

		while (index >= 0)
		{
			this.splice(index, 1);

			index = this.indexOf(value);
		}

		return this;
	});

	safelyAddProperty(Array.prototype, "map", function (func)
	{
		if (!func)
		{
			return this.slice();
		}

		var result = [];

		for (var i = 0; i < this.length; i++)
		{
			result.push(func(this[i]));
		}

		return result;
	});

	safelyAddProperty(Array.prototype, "filter", function (predicate)
	{
		if (!predicate)
		{
			return this.slice();
		}

		var results = [];

		for (var i = 0; i < this.length; i++)
		{
			if (predicate(this[i]))
				results.push(this[i]);
		}

		return results;
	});

	safelyAddProperty(Array.prototype, "some", function (predicate, thisObject)
	{
		if (!predicate)
		{
			return this.slice();
		}

		var results = [];

		for (var i = 0; i < this.length; i++)
		{
			if (predicate.call(thisObject, this[i]))
				return true;				
		}

		return false;
	});

	safelyAddProperty(Array.prototype, "sortBy", function (valueFunc)
	{
		if (!valueFunc)
		{
			return this.slice();
		}

		return this.sort(function (a, b)
		{
			var valueA = valueFunc(a);
			var valueB = valueFunc(b);

			if (valueA > valueB)
			{
				return 1;
			}
			else if (valueA == valueB)
			{
				return 0;
			}
			else
			{
				return -1;
			}
		});
	});

	safelyAddProperty(Array.prototype, "sortByDesc", function (valueFunc)
	{
		if (!valueFunc)
		{
			return this.slice();
		}

		return this.sort(function (a, b)
		{
			var valueA = valueFunc(a);
			var valueB = valueFunc(b);

			if (valueA > valueB)
			{
				return -1;
			}
			else if (valueA == valueB)
			{
				return 0;
			}
			else
			{
				return 1;
			}
		});
	});

	safelyAddProperty(Array.prototype, "sortDesc", function ()
	{
		return this.sort().reverse();
	});

	safelyAddProperty(Array.prototype, "distinct", function ()
	{
		var results = [];

		for (var i = 0; i < this.length; i++)
		{
			if (results.indexOf(this[i]) < 0)
				results.push(this[i]);
		}

		return results;
	});

	safelyAddProperty(Array.prototype, "toStringArray", function ()
	{
		return this.map(
			function (i)
			{
				return (i !== null && i !== undefined) ? i.toString() : i;
			}
		);
	});

	safelyAddProperty(Array.prototype, "reduce", function (func)
	{
		if (!func)
		{
			return undefined;
		}

		var result = this[0];

		for (var i = 1, len = this.length; i < len; i++)
			result = func(result, this[i]);

		return result;
	});

	safelyAddProperty(Array.prototype, "sum", function (selectorFunc)
	{
		if (this.length === 0) return 0;

		return this.reduce(function (total, item) { return total + item; });
	});

	safelyAddProperty(Array.prototype, "max", function ()
	{
		if (this.length === 0) return undefined;

		return this.reduce(function (total, item) { return total >= item ? total : item; });
	});

	safelyAddProperty(Array.prototype, "min", function ()
	{
		if (this.length === 0) return undefined;

		return this.reduce(function (total, item) { return total <= item ? total : item; });
	});

	safelyAddProperty(Array.prototype, "avg", function ()
	{
		if (this.length === 0) return 0;

		return this.sum() / this.length;
	});

	//TODO: Implementar método que aceite arrays com tamanhos diferentes
	safelyAddProperty(Array, "flatten", function (arrays, func)
	{
		if (!arrays || !func)
		{
			return undefined;
		}

		if (!arrays.length)
		{
			return [];
		}

		var results = [];

		var maxLen = arrays.map(function (arr) { return arr.length; }).max();

		for (var item = 0; item < maxLen; item++)
		{
			var arr;
			var result;

			/*
			Obtem o valor inicial, percorrendo todos os arrays e verificando se a posição atual está dentro do tamanho de cada um.
			Caso esteja, o valor inicial é o item na posição atual do array na posição arr.
			*/
			for (arr = 0; arr < arrays.length; arr++)
			{
				if (item < arrays[arr].length)
				{
					result = arrays[arr][item];
					break;
				}
			}

			++arr;

			//Executa a função de agregação e armazena os resultados
			for (; arr < arrays.length; arr++) //Inicia o loop à partir do array logo após o array anterior, que já foi agregado (++arr)
			{
				if (item >= arrays[arr].length)
					break;

				result = func(result, arrays[arr][item]);
			}

			results[item] = result;
		}

		return results;
	});

	safelyAddProperty(Array, "sum", function (arrays)
	{
		return Array.flatten(arrays, function (total, item) { return total + item; });
	});

	/***********************
	window.location
	***********************/
	safelyAddProperty(window.location, "getParts", function ()
	{
		var root = "${0}//${1}${2}".interpolate([location.protocol, location.hostname, (location.port ? ":" + location.port : "")]);

		var pathname = this.pathname;

		if (pathname.indexOf("/") === 0)
			pathname = pathname.substr(1);

		var path = pathname.split("/");

		return [root].concat(path);
	});

	safelyAddProperty(window.location, "root", function (deepness)
	{
		return this.getParts().slice(0, 1 + (deepness || 0)).join("/");
	});

	safelyAddProperty(window.location, "map", function (relativeUrl, rootDeepness)
	{
		return relativeUrl.replaceAll("~", this.root(rootDeepness));
	});
})();
