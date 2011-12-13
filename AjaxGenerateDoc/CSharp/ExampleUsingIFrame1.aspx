<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ExampleUsingIFrame1.aspx.cs" Inherits="AjaxGenerateDoc.ExampleUsingIFrame1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Aspose.Words with AJAX - Example 1</title>
<script language="javascript" type="text/javascript">
<!--

//Invoke document generation.
function generate_onclick() 
{
    //Show the progress message.
    var idicator = document.getElementById('<%= generatingIndicatorPanel.ClientID %>');
    idicator.style.display = "inline";
    
    // Create an IFRAME.
    var iframe = document.createElement("iframe");

    // Point the IFRAME to GenerateFile.
    iframe.src = "GenerateFile.aspx";
    // This makes the IFRAME invisible to the user.
    iframe.style.display = "none";
    // Add the IFRAME to the page.  This will trigger a request to GenerateFile now.
    document.body.appendChild(iframe); 

    checkFrameLoad();   
}

//Chech if generation is completed (this method calls a Web method).
function checkFrameLoad() 
{
    //Check if generating is completed
    PageMethods.CheckCompleted(onComplete);
}

//Show or hide message.
function onComplete(result) 
{
    if(result==false)
    {
        //Run the same check again after 1 sec.
        setTimeout('checkFrameLoad()', 1000);
    }
    else
    {
        //Hide the progress message.
        var idicator = document.getElementById('<%= generatingIndicatorPanel.ClientID %>');
        idicator.style.display = "none";
    }
}
// -->
</script>
</head>
<body>
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods="true" />
        <br />
            <span>
                This example demonstrates how to show a progress message while invoking Aspose.Words
                to generate a document. In this example an IFrame is used. 
                <br />
                <br />
            </span>
            <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                <ContentTemplate>
                    <input id="buttonDownload" type="button" value="Generate Document" onclick="generate_onclick()" />
                </ContentTemplate>
            </asp:UpdatePanel>
            <div id="generatingIndicatorPanel" style="display:none;" runat="server">
                <span>Please wait while system generates a document.</span>
                <img src="images/indicator.gif" />
            </div>
    </form>
</body>
</html>
