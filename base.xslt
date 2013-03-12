<?xml version="1.0" encoding="utf-16"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="msxsl">
  
	<xsl:output indent="no" encoding="utf-8" omit-xml-declaration="yes"/>

	<xsl:param name="siteUrl" />
	<xsl:param name="now" />
	<xsl:param name="timestamp" />
    <xsl:param name="siteUrlImg" />
    <xsl:param name="isFirstBulk" />
    <xsl:param name="isLastBulk" />

	<xsl:decimal-format name="br" decimal-separator="," grouping-separator="." />


	<!-- Global variables -->
	<xsl:variable name="lowerCaseChars" select="'abcdefghijklmnopqrstuvwxyzáàâãéêiíóôõuúûçñ'"/>
	<xsl:variable name="upperCaseChars" select="'ABCDEFGHIJKLMNOPQRSTUVWXYZÁÀÂÃÉÊIÍÓÔÕUÚÛÇÑ'"/>
	<xsl:variable name="upperCaseSpecialChars" select="'ÁÀÂÃÄÉÊÈËIÍÌÏÓÒÔÕÖUÚÛÜÇÑ'"/>
	<xsl:variable name="lowerCaseSpecialChars" select="'áàâãäéêèëiíìïóòôõöuúûüçñ'"/>
	<xsl:variable name="lineBreak">
		<xsl:text>
</xsl:text>
	</xsl:variable>
	<xsl:variable name="tab">
		<xsl:text>	</xsl:text>
	</xsl:variable>



	<!-- function replace(string text, string oldString, string newString) -->
	<xsl:template name="replace">
		
		<xsl:param name="text"></xsl:param>
		<xsl:param name="oldString"></xsl:param>
		<xsl:param name="newString"></xsl:param>

		<xsl:choose>
			
			<xsl:when test="not(contains($text, $oldString))">
				<xsl:value-of select="$text"/>
			</xsl:when>
			
			<xsl:otherwise>
				
				<xsl:variable name="firstPart" select="substring-before($text, $oldString)"/>
				<xsl:variable name="lastPart" select="substring-after($text, $oldString)"/>
				<xsl:variable name="replacedText" select="concat($firstPart, $newString, $lastPart)"/>
				
				<xsl:choose>
					<xsl:when test="contains($replacedText, $oldString)">
						<xsl:call-template name="replace">
							<xsl:with-param name="text" select="$replacedText"/>
							<xsl:with-param name="oldString" select="$oldString"/>
							<xsl:with-param name="newString" select="$newString"/>
						</xsl:call-template>
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select="$replacedText"/>
					</xsl:otherwise>
				</xsl:choose>
				
			</xsl:otherwise>
			
		</xsl:choose>
	</xsl:template>

	
	<!-- function format-date(string date, string format) -->
	<xsl:template name="format-date">
		
		<xsl:param name="date" />
		<xsl:param name="format" />

		<xsl:variable name="year" select="substring($date, 1, 4)"/>
		<xsl:variable name="month" select="substring($date, 6, 2)"/>
		<xsl:variable name="day" select="substring($date, 9, 2)"/>
		<xsl:variable name="hours" select="substring($date, 12, 2)"/>
		<xsl:variable name="minutes" select="substring($date, 15, 2)"/>
		<xsl:variable name="seconds" select="substring($date, 18, 2)"/>

		<xsl:variable name="replacedYear">
			<xsl:call-template name="replace">
				<xsl:with-param name="text" select="$format"/>
				<xsl:with-param name="oldString" select="'yyyy'"/>
				<xsl:with-param name="newString" select="$year"/>
			</xsl:call-template>
		</xsl:variable>

		<xsl:variable name="replacedMonth">
			<xsl:call-template name="replace">
				<xsl:with-param name="text" select="$replacedYear"/>
				<xsl:with-param name="oldString" select="'MM'"/>
				<xsl:with-param name="newString" select="$month"/>
			</xsl:call-template>
		</xsl:variable>

		<xsl:variable name="replacedDay">
			<xsl:call-template name="replace">
				<xsl:with-param name="text" select="$replacedMonth"/>
				<xsl:with-param name="oldString" select="'dd'"/>
				<xsl:with-param name="newString" select="$day"/>
			</xsl:call-template>
		</xsl:variable>

		<xsl:variable name="replacedHour">
			<xsl:call-template name="replace">
				<xsl:with-param name="text" select="$replacedDay"/>
				<xsl:with-param name="oldString" select="'HH'"/>
				<xsl:with-param name="newString" select="$hours"/>
			</xsl:call-template>
		</xsl:variable>

		<xsl:variable name="replacedMinutes">
			<xsl:call-template name="replace">
				<xsl:with-param name="text" select="$replacedHour"/>
				<xsl:with-param name="oldString" select="'mm'"/>
				<xsl:with-param name="newString" select="$minutes"/>
			</xsl:call-template>
		</xsl:variable>

		<xsl:variable name="replacedSeconds">
			<xsl:call-template name="replace">
				<xsl:with-param name="text" select="$replacedMinutes"/>
				<xsl:with-param name="oldString" select="'ss'"/>
				<xsl:with-param name="newString" select="$seconds"/>
			</xsl:call-template>
		</xsl:variable>

		<xsl:value-of select="$replacedSeconds"/>
		
	</xsl:template>

	<xsl:template name="remove-accents">
		<xsl:param name="text"/>

		<xsl:variable name="upperCaseNormalChars" select="'AAAAAEEEEIIIIOOOOOUUUUCN'"/>
		<xsl:variable name="lowerCaseNormalChars" select="'aaaaaeeeeiiiiooooouuuucn'"/>

		<xsl:value-of select="translate(translate($text, $upperCaseSpecialChars, $upperCaseNormalChars), $lowerCaseSpecialChars, $lowerCaseNormalChars)"/>
		
	</xsl:template>

	<xsl:template name="remove-line-break">
		<xsl:param name="text"/>
		<xsl:call-template name="replace">
			<xsl:with-param name="text" select="$text"/>
			<xsl:with-param name="oldString" select="$lineBreak"/>
			<xsl:with-param name="newString" select="''"/>
		</xsl:call-template>
	</xsl:template>

	<xsl:template name="remove-space">
		<xsl:param name="text"/>
		<xsl:call-template name="replace">
			<xsl:with-param name="text" select="normalize-space($text)"/>
			<xsl:with-param name="oldString" select="' '"/>
			<xsl:with-param name="newString" select="''"/>
		</xsl:call-template>
	</xsl:template>
	
</xsl:stylesheet>
