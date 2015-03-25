<script type="text/jscript" runat="server" language="jscript">

	(function ()
	{
		var globalStartTime = new Date();
		var measures = {};

		/*global*/ PerformanceMeter =
		{
			printTotalTime: function ()
			{
				Debug.writeln("Total time: " + (new Date() - globalStartTime) + "ms");
			},

			start: function (description)
			{
				measures[description] = new Date();
			},

			print: function (description)
			{
				Debug.writeln(description + ": " + (new Date() - measures[description]) + "ms");
			},

			log: function (message)
			{
				Debug.writeln(message);
			}
		}

	})();

</script>