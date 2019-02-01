function xmlRequestFunction(SP_USER, SP_PASS, SP_URL) 
{
    return  '<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope" xmlns:a="http://www.w3.org/2005/08/addressing" xmlns:u="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">\n'
            + '  <s:Header>'
            + '    <a:Action s:mustUnderstand="1">http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</a:Action>'
            + '    <a:ReplyTo><a:Address>http://www.w3.org/2005/08/addressing/anonymous</a:Address></a:ReplyTo>'
            + '    <a:To s:mustUnderstand="1">https://login.microsoftonline.com/extSTS.srf</a:To>'
            + '    <o:Security s:mustUnderstand="1" xmlns:o="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">'
            + '      <o:UsernameToken>'
            + '        <o:Username>'+ SP_USER +'</o:Username>'
            + '        <o:Password>'+ SP_PASS +'</o:Password>'
            + '      </o:UsernameToken>'
            + '    </o:Security>'
            + '  </s:Header>'
            + '  <s:Body>'
            + '    <t:RequestSecurityToken xmlns:t="http://schemas.xmlsoap.org/ws/2005/02/trust">'
            + '      <wsp:AppliesTo xmlns:wsp="http://schemas.xmlsoap.org/ws/2004/09/policy">'
            
            + '      </wsp:AppliesTo>'
            + '      <t:KeyType>http://schemas.xmlsoap.org/ws/2005/05/identity/NoProofKey</t:KeyType>'
            + '      <t:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</t:RequestType>'
            + '      <t:TokenType>urn:oasis:names:tc:SAML:1.0:assertion</t:TokenType>'
            + '    </t:RequestSecurityToken>'
            + '  </s:Body>'
            + '</s:Envelope>';
}


function optionsXmlRequestFunction(xmlLength)
{
    return {
        hostname: 'login.microsoftonline.com',
        path    : "/extSTS.srf",
        method  : 'POST',
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
            'Content-Length': xmlLength
        }
    };
}


function optionsAuthRequest(SP_HOST, SP_TOKEN)
{
    return {
        hostname: SP_HOST,
        agent: false,
        path: "/_forms/default.aspx?wa=wsignin1.0",
        method: 'POST',
        headers: 
        {
            'Content-Type': 'application/x-www-form-urlencoded',
            'Content-Length': SP_TOKEN.length,
            'Host': SP_HOST
        }
    };
}


module.exports.xmlRequestFunction = xmlRequestFunction;
module.exports.requestOptions = optionsXmlRequestFunction;
module.exports.optionsAuthRequest = optionsAuthRequest;