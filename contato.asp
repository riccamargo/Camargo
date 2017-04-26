<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!doctype html>
<html>
<head>
<meta charset="utf-8">
    <title>Envio de Mensagem - Camargo & Associados - Terceirização & Gestão :::.....</title>
    <!-- ===== Google Fonts ===== -->
    <link rel="stylesheet" href="http://fonts.googleapis.com/css?family=Source+Sans+Pro:400,700,400italic|Raleway:500,600,700" />
    <!-- ===== Favicon Icon ===== -->
    <link rel="icon" href="images/favicon.ico"/>
    <!-- ===== Bootstrap ===== -->
    <link rel="stylesheet" href="css/bootstrap.min.css" />
    <!-- ===== Font Icons ===== -->
    <link rel="stylesheet" href="assets/font-awesome/css/font-awesome.min.css" />
    <!-- ===== Colors ===== -->
    <link rel="stylesheet" href="css/colors/color3.css" />
    <!-- ===== Preloader ===== -->
    <link rel="stylesheet" href="css/preloader.css" />
    <!-- ===== style.css ===== -->
    <link rel="stylesheet" href="css/style.css" />
    <!-- ===== Responsive CSS ===== -->
    <link rel="stylesheet" href="css/responsive.css" />
</head>

<%
Dim vNome, vEmail, vSubject, vMensagem, vTitulo
vNome     =  request.form("contact-form-name")
vEmail     =  request.form("contact-form-email")
vSubject     =  request.form("contact-form-subject")
vMensagem =  request("contact-form-message")
vTitulo     =  "Consulta recebida pelo WebSite"



msgBody = "A Empresa/Pessoa abaixo efetuou a seguinte consulta a nosso Web-site e foi incluída em " & vbCrLf 
msgBody = msgBody & "nosso bando de dados, para recebimento de nosssa newsletter!" & vbCrLf 
msgBody = msgBody & vbCrLf
msgBody = msgBody & "Nome - " & vNome & vbCrLf 
msgBody = msgBody & "E-mail - " & vEmail & vbCrLf 
msgBody = msgBody & vbCrLf
'msgBody = msgBody & "Endereço - " & endereco & vbCrLf 
'msgBody = msgBody & "Cidade -   " & cidade & "   Estado : " & estado & vbCrLf 
'msgBody = msgBody & "CEP - " & cep & " Telefone : " & fone & vbCrLf 
msgBody = msgBody & "Título : " & vSubject & vbCrLf 
msgBody = msgBody & vbCrLf
msgBody = msgBody & "A consulta foi : " & vMensagem & vbCrLf  
msgBody = msgBody & vbCrLf
msgBody = msgBody & vbCrLf
msgBody = msgBody & "Necessitamos apenas para concluirmos o processo que você clique no link " & vbCrLf 
msgBody = msgBody & "abaixo para confirmar seu email." & vbCrLf
msgBody = msgBody & vbCrLf
msgBody = msgBody & "http://www.gerenciaonline.com.br/consultasql.asp?id="& idmembro & vbCrLf


'Response.write(msgBody) & "<br>"

' definindo uma variavel auxiliar
sch = "http://schemas.microsoft.com/cdo/configuration/"
' criando o objeto de configuração do CDO
Set cdoConfig = Server.CreateObject("CDO.Configuration") 
' definindo as configurações
cdoConfig.Fields.Item(sch & "sendusing") = 2
cdoConfig.Fields.Item(sch & "smtpauthenticate") = 1
cdoConfig.Fields.Item(sch & "smtpserver") = "smtp.gerenciaonline.com.br"
cdoConfig.Fields.Item(sch & "sendusername") = "suporte@gerenciaonline.com.br"
cdoConfig.Fields.Item(sch & "sendpassword") = "rcamargo5558"
cdoConfig.Fields.Item(sch & "smtpserverport") = 587
cdoConfig.fields.update
 
' criando o objeto de msg do CDO
Set cdoMessage = Server.CreateObject("CDO.Message")
 
' associando as configurações ao obj Mensagem
Set cdoMessage.Configuration = cdoConfig 
' definido variaveis da msg
cdoMessage.From = "suporte@gerenciaonline.com.br"
cdoMessage.To = "camargo@gerenciaonline.com.br"
cdoMessage.Subject = vTitulo
cdoMessage.TextBody = msgBody & vbCrLf 
cdoMessage.Send
Set cdoMessage = Nothing
Set cdoConfig = Nothing

%>


<body>

<!-- ===== Header ===== -->
<header id="home">

<!-- ===== Contact Us ===== -->
<section id="section-contact" class="section-contact bgc-one">
<div class="container">
	
	<div class="row">
		<div class="col-md-8 col-md-offset-2">
        
            <div class="package background-color-white">
                <div class="header-dark">
                        <h3>Sua Mensagem foi enviada com sucesso, em breve estaremos atendendo sua solicitação. </h3>
                </div>
				<div class="package-features">

                    <p>Lembramos que o preenchimento correto de suas informações principalmente e-mail é que possibilitará entramos em contato.</p>
                    
                    <p>O E-mail que você informou foi : <%=vEmail%>.</p>
                    
                    <p>Caso encontre alguma falha volte e refaça sua consulta.</p>
                    
                    <p>Seus dados foram incluídos em nosso banco de dados para facilitar suas futuras consultas e para recebimento de nossa newsletter.</p>
                    
                    <p>Muito Obrigado!</p>
                                               
	                <p><input type="button" value="Voltar" name="vFechar" onClick="javascript:self.close();" class="btn standard-button"></p>
				</div>
			</div>
		</div>
	</div>
</div>

</header>
</body>
</html>










