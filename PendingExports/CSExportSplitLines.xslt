<?xml version="1.0"?>

<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
  <xsl:template match="//cs-object">
    <xsl:copy-of select="."/>
	<xsl:text>&#10;</xsl:text>
  </xsl:template>
</xsl:stylesheet>