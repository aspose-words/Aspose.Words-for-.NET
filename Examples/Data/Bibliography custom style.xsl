<!--Based on http://lenzconsulting.com/xml-to-string/ -->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography">
	<xsl:output omit-xml-declaration="yes"/>
	<xsl:param name="dev-version" select="'0.1'"/>
	<xsl:param name="dev-name" select="'TestBibliographyCustomStyle'"/>
	<xsl:template match="/">
		<xsl:choose>
			<xsl:when test="b:Version">
				<xsl:value-of select="$dev-version"/>
			</xsl:when>
			<xsl:when test="b:OfficeStyleKey">
				<xsl:value-of select="$dev-name"/>
			</xsl:when>
			<xsl:when test="b:XslVersion">
				<xsl:value-of select="$dev-version"/>
			</xsl:when>
			<xsl:when test="b:StyleNameLocalized">
				<xsl:value-of select="$dev-name"/>
			</xsl:when>
			<xsl:when test="b:GetImportantFields">
				<b:ImportantFields/>
			</xsl:when>
			<xsl:otherwise>
				<html>
					<body>
						<xsl:for-each select="//b:Source">
							<p>
								<xsl:text>Tag: </xsl:text>
								<b><xsl:value-of select="b:Tag"/></b>
								<xsl:text>; Source type: </xsl:text>
								<b><xsl:value-of select="b:SourceType"/></b>
								<xsl:text>; Title: </xsl:text>
								<b><xsl:value-of select="b:Title"/></b>
								<xsl:text>;</xsl:text>
							</p>
						</xsl:for-each>
					</body>
				</html>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
</xsl:stylesheet>