<?xml version="1.0"?>
<!-- MIM ServiceToCSV.xslt
	 Written by Carol Wapshere, www.wapshere.com/missmiis
	 
	 Used to convert the policy.xml and schema.xml files exported from the MIM Service to CSV format.
-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:a="http://www.microsoft.com/mms/mmsml/v2" xmlns:data="http://example.com/data">
<xsl:output media-type="text" omit-xml-declaration="yes"  encoding="ANSI" indent="no" />
<xsl:template match="/">
<xsl:call-template name="header" />
</xsl:template>

<xsl:template name="header">id;type;attribute;value
	<xsl:for-each select="Results/ExportObject/ResourceManagementObject">
		<xsl:call-template name="ResourceManagementObject" />
	</xsl:for-each>
</xsl:template>

<xsl:template name="ResourceManagementObject">
	<xsl:for-each select="ResourceManagementAttributes/ResourceManagementAttribute">
		<xsl:choose>
			<xsl:when test="IsMultiValue='false'">
				<xsl:call-template name="SingleValue">
					<xsl:with-param name="id" select="../../ObjectIdentifier" />
					<xsl:with-param name="type" select="../../ObjectType" />
				</xsl:call-template>
			</xsl:when>	
			<xsl:when test="IsMultiValue='true'">
				<xsl:call-template name="MultiValue">
					<xsl:with-param name="id" select="../../ObjectIdentifier" />
					<xsl:with-param name="type" select="../../ObjectType" />
				</xsl:call-template>
			</xsl:when>	
		</xsl:choose>
	</xsl:for-each>			
</xsl:template>

<xsl:template name="SingleValue">
    <xsl:param name="id" />	
    <xsl:param name="type" />	
<!--Debug SingleValue id:<xsl:value-of select="$id" />-->
<!--Debug SingleValue type:<xsl:value-of select="$type" />-->
	<xsl:choose>
		<xsl:when test="AttributeName='Filter'">
			<xsl:variable name="filter" select="substring-before(substring-after(Value,'&gt;'),'&lt;')"/>
			<xsl:text>&#10;</xsl:text><xsl:value-of select="$id"/>;<xsl:value-of select="$type" />;<xsl:value-of select="AttributeName" />;<xsl:value-of select="$filter"/>
		</xsl:when>	
		<xsl:when test="contains(Value,';') or contains(Value,'&#10;') or contains(Value,'&lt;') or contains(Value,'&gt;')">
			<!-- Skip values that don't work in CSV -->
		</xsl:when>	
		<xsl:otherwise>
			<xsl:text>&#10;</xsl:text><xsl:value-of select="$id"/>;<xsl:value-of select="$type" />;<xsl:value-of select="AttributeName" />;<xsl:value-of select="Value"/>
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>

<xsl:template name="MultiValue">
    <xsl:param name="id" />	
    <xsl:param name="type" />	
<!--Debug MultiValue id:<xsl:value-of select="$id" />-->
<!--Debug MultiValue type:<xsl:value-of select="$type" />-->
    <xsl:variable name="values">
        <xsl:call-template name="join">
			<xsl:with-param name="nodeset" select="Values/string" />
        </xsl:call-template>
    </xsl:variable>
	<xsl:text>&#10;</xsl:text><xsl:value-of select="$id"/>;<xsl:value-of select="$type" />;<xsl:value-of select="AttributeName" />;<xsl:value-of select="$values"/>
</xsl:template>

<xsl:template name="join">
	<xsl:param name="nodeset" />
	<xsl:param name="delimiter" />
	<xsl:param name="result" select="''" />
	<xsl:choose>
		<xsl:when test="$nodeset">
			<xsl:call-template name="join">
				<xsl:with-param name="nodeset" select="$nodeset[position() &gt; 1]" />
				<xsl:with-param name="delimiter" select="','" />
				<xsl:with-param name="result" select="concat($result,$delimiter,$nodeset[1])" />
			</xsl:call-template>
		</xsl:when>
		<xsl:otherwise>
			<xsl:value-of select="$result" disable-output-escaping="no" />
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>

</xsl:stylesheet>
