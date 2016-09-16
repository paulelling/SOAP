<?xml version="1.0"?> 
<xsl:stylesheet xmlns:xsl="http://www.w3.org/TR/WD-xsl">

  <xsl:template match="/">
    <HTML>
      <xsl:apply-templates/>
    </HTML>
  </xsl:template>
  
  <xsl:template match="EMAIL">
  <CENTER>
    <B><xsl:value-of select="TITLE"/></B>
    <P/>
    <FORM ACTION="send.asp" METHOD="post">
    <TABLE>
      <TR>
        <TD><xsl:value-of select="FROM"/></TD>
        <TD><INPUT TYPE="text" NAME="strFrom"/></TD>
      </TR><TR>
        <TD><xsl:value-of select="TO"/></TD>
        <TD><INPUT TYPE="text" NAME="strTo"/></TD>
      </TR><TR>
        <TD><xsl:value-of select="SUBJECT"/></TD>
        <TD><INPUT TYPE="text" NAME="strSubject" SIZE="40"/></TD>
      </TR><TR>
        <TD><xsl:value-of select="MESSAGE"/></TD>
        <TD><TEXTAREA NAME="strMessage" COLS="40" ROWS="5"></TEXTAREA></TD>
      </TR>
    </TABLE>
    <BR/><INPUT TYPE="submit" VALUE="send"/>
    </FORM>
  </CENTER>
  </xsl:template>
  
</xsl:stylesheet>
