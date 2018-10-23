<%
'*****************************************************************************
'TÍTULO PÁGINA
function GL_FTPG(inParam, infuncao_id)
'
dim titulo
dim Usuario_Internet_id, empresa_Id, nome_produto, banco_dados, msg_ret
dim objlog, retorno
'
nome_produto = session("nome_produto") 
banco_dados = session("nome_banco") 
empresa_id = session("empresa_id")
if session("empresa_id") = "" then
	empresa_id = "1"
end if
Usuario_Internet_id = session("gl_usuario_internet_id")
'
set objlog = server.createobject("ECRACS24.ACS")
RETORNO = OBJLOG.gConsulta_Titulo_Principal(Usuario_Internet_id, Empresa_id, Banco_Dados, Nome_Produto, _
        infuncao_id, Titulo, msg_ret)
GL_FTPG = session("gl_titulo_pagina_ini") & inParam & " " & Titulo & "&nbsp;" & session("gl_titulo_pagina_fim")
'
end function
'*****************************************************************************
'SUB TÍTULO
function GL_FSPG(inParam)
'
if inParam <> "" then
	GL_FSPG = session("gl_sub_titulo_ini") & inParam & session("gl_sub_titulo_fim")
else
	GL_FSPG = "&nbsp;"
end if
'
end function
'*****************************************************************************
'TÍTULO CAMPO
function GL_FTCA(inParam)
'
if inParam <> "" then
	GL_FTCA = session("gl_titulo_campo_ini") & inParam & session("gl_titulo_campo_fim")
else
	GL_FTCA = "&nbsp;"
end if
'
end function
'*****************************************************************************
'CONTEÚDO CAMPO
function GL_FCTC(inParam)
'
if inParam <> "" then
	GL_FCTC = session("gl_conteudo_campo_ini") & inParam & session("gl_conteudo_campo_fim")
else
	GL_FCTC = "&nbsp;"
end if
'
end function
'*****************************************************************************
'TÍTULO TABELA
function GL_FTTB(inParam)
'
if inParam <> "" then
	GL_FTTB = session("gl_titulo_tabela_ini") & inParam & session("gl_titulo_tabela_fim")
else
	GL_FTTB = "&nbsp;"
end if
'
end function
'*****************************************************************************
'CONTEÚDO TABELA LINHA PAR
function GL_FCTP(inParam)
'
if inParam <> "" then
	GL_FCTP = session("gl_cont_tabela_par_ini") & inParam & session("gl_cont_tabela_par_fim")
else
	GL_FCTP = "&nbsp;"
end if
'
end function
'*****************************************************************************
'CONTEÚDO TABELA LINHA ÍMPAR
function GL_FCTI(inParam)
'
if inParam <> "" then
	GL_FCTI = session("gl_cont_tabela_imp_ini") & inParam & session("gl_cont_tabela_imp_fim")
else
	GL_FCTI = "&nbsp;"
end if
'
end function
'*****************************************************************************
'MENSAGEM ERRO
function GL_FMER(inParam)
'
if inParam <> "" then
	GL_FMER = SESSION("GL_MSG_ERRO_INI") & inParam & SESSION("GL_MSG_ERRO_FIM")
else
	GL_FMER = "&nbsp;"
end if
'
end function
'*****************************************************************************
'MENSAGEM AVISO
function GL_FMAV(inParam)
'
if inParam <> "" then
	GL_FMAV = SESSION("GL_MSG_AVISO_INI") & inParam & SESSION("GL_MSG_AVISO_FIM")
else
	GL_FMAV = "&nbsp;"
end if
'
end function

'*****************************************************************************
'LINK
function GL_FLIN(inParam)
'
if inParam <> "" then
	GL_FLIN = session("gl_link_ini") & inParam & session("gl_link_fim")
else
	GL_FLIN = "&nbsp;"
end if
'
end function
'*****************************************************************************
'COR DE FUNDO DA PÁGINA
function GL_CRFP()
'
GL_CRFP = session("gl_cor_fundo_pag")
'
end function
'*****************************************************************************
'COR DE FUNDO DO TÍTULO DA TABELA
function GL_CRTT()
'
GL_CRTT = session("gl_cor_fundo_tit_tabela")
'
end function
'*****************************************************************************
'COR DE FUNDO TABELA LINHA PAR
function GL_CRTP()
'
GL_CRTP = session("gl_cor_fundo_tabela_par")
'
end function
'*****************************************************************************
'COR DE FUNDO TABELA LINHA ÍMPAR
function GL_CRTI()
'
GL_CRTI = session("gl_cor_fundo_tabela_imp")
'
end function
'*****************************************************************************
'INDIC. ANTES PRAZO
function GL_FCAP(inParam)
'
if inParam <> "" then
	GL_FCAP = session("gl_cor_ant_prazo_ini") & inParam & session("gl_cor_ant_prazo_fim")
else
	GL_FCAP = "&nbsp;"
end if

'
end function
'*****************************************************************************
'INDIC. NO PRAZO
function GL_FCNP(inparam)
'
if inParam <> "" then
	GL_FCNP = session("gl_cor_no_prazo_ini") & inParam & session("gl_cor_no_prazo_fim")
else
	GL_FCNP = "&nbsp;"
end if
'
end function
'*****************************************************************************
'INDIC. 1 DIA ATRASO
function GL_FC1A(inParam)
'
if inParam <> "" then
	GL_FC1A = session("gl_cor_1_atraso_ini") & inParam & session("gl_cor_1_atraso_fim")
else
	GL_FC1A = "&nbsp;"
end if
'
end function
'*****************************************************************************
'INDIC. 2 DIA ATRASO
function GL_FC2A(inParam)
'
if inParam <> "" then
	GL_FC2A = session("gl_cor_2_atraso_ini") & inParam & session("gl_cor_2_atraso_fim")
else
	GL_FC2A = "&nbsp;"
end if
'
end function
'*****************************************************************************
'INDIC. + 2 DIA ATRASO
function GL_FCM2(inParam)
'
if inParam <> "" then
	GL_FCM2 = session("gl_cor_mais_2_atraso_INI") & inParam & session("gl_cor_mais_2_atraso_FIM")
else
	GL_FC2A = "&nbsp;"
end if
'
end function
'*****************************************************************************
'CONFIGURAÇÕES DA PÁGINA
function GL_TIPG()
'
GL_TIPG = session("gl_config_pagina")
'
end function
'*****************************************************************************
'CONFIGURAÇÕES DA TABELA
function GL_TITB()
'
GL_TITB = SESSION("GL_CONFIG_TABELA")
'
end function
'*****************************************************************************
'CONTEÚDO DO CAMPO  EM DESTAQUE
function GL_FCTD(inParam)
'
if inParam <> "" then
	GL_FCTD = session("gl_CONT_CAMPO_DESTAQUE_INI") & inParam & session("gl_CONT_CAMPO_DESTAQUE_FIM")
else
	GL_FCTD = "&nbsp;"
end if
'
end function
'*****************************************************************************
'CONTEÚDO DA LINHA PAR EM DESTAQUE
function GL_FLPD(inParam)
'
if inParam <> "" then
	GL_FLPD = session("gl_CONT_LINHA_PAR_DESTAQUE_INI") & inParam & session("gl_CONT_LINHA_PAR_DESTAQUE_FIM")
else
	GL_FLPD = "&nbsp;"
end if
'
end function
'*****************************************************************************
'CONTEÚDO DA LINHA ÍMPAR EM DESTAQUE
function GL_FLID(inParam)
'
if inParam <> "" then
	GL_FLID = session("gl_CONT_LINHA_IMP_DESTAQUE_INI") & inParam & session("gl_CONT_LINHA_IMP_DESTAQUE_FIM")
else
	GL_FLID = "&nbsp;"
end if
'
end function
'*****************************************************************************
'MENSAGEM COMERCIAL
function GL_FMC1(inParam)
'
if inParam <> "" then
	GL_FMC1 = session("gl_MENSAGEM_COMERCIAL_INI") & inParam & session("gl_MENSAGEM_COMERCIAL_FIM")
else
	GL_FMC1 = "&nbsp;"
end if
'
end function
'*****************************************************************************
'COR DE FUNDO DA PÁGINA INICIAL
function GL_CRPI()
'
GL_CRPI = session("gl_COR_FUNDO_PAG_INICIAL")
'
end function
'*****************************************************************************
'COR DE FUNDO MENSAGEM COMERCIAL
function GL_CRMS()
'
GL_CRMS = session("gl_COR_FUNDO_MSG_COMERCIAL")
'
end function
'*****************************************************************************
'LINK DA PÁGINA INICIAL
function GL_FLPG(inParam)
'
if inParam <> "" then
	GL_FLPG = SESSION("GL_LINK_PAG_INICIAL_INI") & inParam & SESSION("GL_LINK_PAG_INICIAL_FIM")
else
	GL_FLPG = "&nbsp;"
end if
'
end function
'*****************************************************************************
'LIMK PÁGINA INICIAL
'function GL_FLPG(inParam)
''
'if inParam <> "" then
'	GL_FLPG = session("gl_LINK_PAG_INCIAL_INI") & inParam & session("gl_LINK_PAG_INCIAL_FIM")
'else
'	GL_FLPG = "&nbsp;"
'end if
''
'end function
'*****************************************************************************
'TÍTULO PÁGINA INICIAL
function GL_FTPI(inParam)
'
if inParam <> "" then
	GL_FTPI = session("gl_TITULO_PAGINA_INICIAL_INI") & inParam & session("gl_TITULO_PAGINA_INICIAL_FIM")
else
	GL_FTPI = "&nbsp;"
end if
'
end function
'*****************************************************************************
'SUB TÍTULO PÁGINA INICIAL
function GL_FSPI(inParam)
'
if inParam <> "" then
	GL_FSPI = session("gl_SUB_TIT_PAGINA_INICIAL_INI") & inParam & session("gl_SUB_TIT_PAGINA_INICIAL_FIM")
else
	GL_FSPI = "&nbsp;"
end if
'
end function
'*****************************************************************************
'COR DE FUNDO DO SUB TÍTULO PAG. INICIAL
function GL_CSPI()
'
GL_CSPI = session("gl_FUNDO_SUB_PAG_INICIAL") 
'
end function
'*****************************************************************************
'COR DE FUNDO DO LOGO DO MENU
function GL_CFLM()
'
GL_CFLM = session("gl_COR_FUNDO_LOGO") 
'
end function
'*****************************************************************************
'DADOS DO USUÁRIO
function GL_FDUS(inParam)
'
if inParam <> "" then
	GL_FDUS = session("gl_DADOS_DO_USUARIO_INI") & inParam & session("gl_DADOS_DO_USUARIO_FIM")
else
	GL_FDUS = "&nbsp;"
end if
'
end function
'*****************************************************************************
'COR DE FUNDO DO MENU
function GL_CRMN()
'
GL_CRMN = session("gl_COR_FUNDO_MENU") 
'
end function
'*****************************************************************************
'TÍTULO MENU
function GL_FTMN(inParam)
'
if inParam <> "" then
	GL_FTMN = session("gl_TITULO_MENU_INI") & inParam & session("gl_TITULO_MENU_FIM")
else
	GL_FTMN = "&nbsp;"
end if
'
end function
'*****************************************************************************
'SUB TÍTULO MENU
function GL_FSMN(inParam)
'
if inParam <> "" then
	GL_FSMN = session("gl_SUB_TITULO_MENU_INI") & inParam & session("gl_SUB_TITULO_MENU_FIM")
else
	GL_FSMN = "&nbsp;"
end if
'
end function
'*****************************************************************************
'LINK LINHA PAR
function GL_FLLP(inParam)
'
if inParam <> "" then
	GL_FLLP = session("gl_LINK_LINHA_PAR_INI") & inParam & session("gl_LINK_LINHA_PAR_FIM")
else
	GL_FLLP = "&nbsp;"
end if
'
end function
'*****************************************************************************
'LINK LINHA ÍMPAR
function GL_FLLI(inParam)
'
if inParam <> "" then
	GL_FLLI = session("gl_LINK_LINHA_IMP_INI") & inParam & session("gl_LINK_LINHA_IMP_FIM")
else
	GL_FLLI = "&nbsp;"
end if
'
end function
'*****************************************************************************
'MENSAGEM COMERCIAL 2
function GL_FMC2(inParam)
'
if inParam <> "" then
	GL_FMC2 = session("gl_MENSAGEM_COMERCIAL2_INI") & inParam & session("gl_MENSAGEM_COMERCIAL2_FIM")
else
	GL_FMC2 = "&nbsp;"
end if
'
end function
'*****************************************************************************
'CONFIGURAÇÕES DO SUB TÍTULO DA PÁGINA
function GL_TIST()
'
GL_TIST = SESSION("GL_CONF_SUB_TITULO")
'
end function
'*****************************************************************************
'CONFIGURAÇÕES DO TÍTULO DA PÁGINA
function GL_TITP()
'
GL_TITP = SESSION("GL_CONF_TITULO_PAG")
'
end function
'*****************************************************************************
'COR DE FUNDO CAMPO DESTAQUE
function GL_CRCD()
'
GL_CRCD = SESSION("GL_COR_FUNDO_CAMPO_DEST")
'
end function
'*****************************************************************************
'IMAGEM TITULO PÁGINA
function GL_IMTP()
'
GL_IMTP = SESSION("GL_IMAGEM_TITULO")
'
end function
'*****************************************************************************
'IMAGEM SUB TITULO PÁGINA
function GL_IMST()
'
GL_IMST = SESSION("GL_IMAGEM_SUB_TITULO")
'
end function
'*****************************************************************************
'IMAGEM LINK LINHA PAR
function GL_IMLP()
'
GL_IMLP = SESSION("GL_IMAGEM_LINK_PAR")
'
end function
'*****************************************************************************
'IMAGEM LINK LINHA ÍMPAR
function GL_IMLI()
'
GL_IMLI = SESSION("GL_IMAGEM_LINK_IMPAR") 
'
end function
'*****************************************************************************
'configuração e-cargo
FUNCTION GL_CONFIG()
'
DIM EMPRESA_ID, OBJLOG, retorno, msg_ret
DIM nome_banco, Nome_Produto
dim nome_asp_ini
DIM TITULO_PAGINA_INI, TITULO_PAGINA_FIM
DIM SUB_TITULO_INI, SUB_TITULO_FIM
DIM TITULO_CAMPO_INI, TITULO_CAMPO_FIM
DIM CONTEUDO_CAMPO_INI, CONTEUDO_CAMPO_FIM
DIM TITULO_TABELA_INI, TITULO_TABELA_FIM
DIM CONT_TABELA_PAR_INI, CONT_TABELA_PAR_FIM
DIM CONT_TABELA_IMP_INI, CONT_TABELA_IMP_FIM
DIM LINK_INI, LINK_FIM
DIM COR_FUNDO_PAG, FIM
DIM COR_FUNDO_TIT_TABELA
DIM COR_FUNDO_TABELA_PAR
DIM COR_FUNDO_TABELA_IMP
DIM MSG_ERRO_INI, MSG_ERRO_FIM
DIM MSG_AVISO_INI, MSG_AVISO_FIM
DIM COR_ANT_PRAZO
DIM COR_NO_PRAZO
DIM COR_1_ATRASO
DIM COR_2_ATRASO
DIM COR_MAIS_2_ATRASO
DIM CONFIG_PAGINA
DIM CONFIG_TABELA_INI, CONFIG_TABELA_FIM
DIM COR_ANT_PRAZO_INI, COR_ANT_PRAZO_FIM
DIM COR_NO_PRAZO_INI, COR_NO_PRAZO_FIM
DIM COR_1_ATRASO_INI, COR_1_ATRASO_FIM
DIM COR_2_ATRASO_INI, COR_2_ATRASO_FIM
DIM COR_MAIS_2_ATRASO_INI, COR_MAIS_2_ATRASO_FIM
'**************
DIM CONT_CAMPO_DESTAQUE_INI, CONT_CAMPO_DESTAQUE_FIM
DIM CONT_LINHA_PAR_DESTAQUE_INI, CONT_LINHA_PAR_DESTAQUE_FIM
DIM CONT_LINHA_IMP_DESTAQUE_INI, CONT_LINHA_IMP_DESTAQUE_FIM
DIM MENSAGEM_COMERCIAL_INI, MENSAGEM_COMERCIAL_FIM
DIM COR_FUNDO_PAG_INICIAL
DIM COR_FUNDO_MSG_COMERCIAL
DIM LINK_PAG_INICIAL_INI, LINK_PAG_INICIAL_FIM
DIM TITULO_PAGINA_INICIAL_INI, TITULO_PAGINA_INICIAL_FIM
DIM SUB_TIT_PAGINA_INICIAL_INI, SUB_TIT_PAGINA_INICIAL_FIM
DIM FUNDO_SUB_PAG_INICIAL
DIM COR_FUNDO_LOGO
DIM DADOS_DO_USUARIO_INI, DADOS_DO_USUARIO_FIM
DIM COR_FUNDO_MENU
DIM TITULO_MENU_INI, TITULO_MENU_FIM
DIM SUB_TITULO_MENU_INI, SUB_TITULO_MENU_FIM
DIM LINK_LINHA_PAR_INI, LINK_LINHA_PAR_FIM
DIM LINK_LINHA_IMP_INI, LINK_LINHA_IMP_FIM
DIM MENSAGEM_COMERCIAL2_INI, MENSAGEM_COMERCIAL2_FIM
DIM CONF_TITULO_PAG, CONF_SUB_TITULO
dim IMAGEM_TITULO, IMAGEM_SUB_TITULO, IMAGEM_LINK_PAR, IMAGEM_LINK_IMPAR, COR_FUNDO_CAMPO_DEST
'
set objlog = server.createobject("ECRACS24.ACS")

IF session("grupo_empresa_id") = 0 THEN
	if session("gl_nome_asp_ini") = "" then
		session("gl_nome_asp_ini") = "INDEX.ASP"
		NOME_ASP_INI = "INDEX.ASP"
	ELSE
		NOME_ASP_INI = session("gl_nome_asp_ini")
	END IF
	'
	RETORNO = OBJLOG.gBusca_Empresa_id(nome_banco, nome_produto, Nome_ASP_INI, Empresa_id, msg_ret)
	IF RETORNO <> 0 THEN
		session("msg_erro") = msg_ret
	END IF
ELSE
	Empresa_id = SESSION("EMPRESA_ID")
END IF

if session("empresa_id") = "" then
	empresa_id = "1"
end if

nome_produto = session("nome_produto") 
nome_banco = session("nome_banco") 

'TÍTULO PÁGINA
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 1, TITULO_PAGINA_INI, TITULO_PAGINA_FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
END IF

'SUB TÍTULO 
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 2, SUB_TITULO_INI, SUB_TITULO_FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
END IF

'TÍTULO CAMPO
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 3, TITULO_CAMPO_INI, TITULO_CAMPO_FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
END IF

'CONTEUDO CAMPO
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 4, CONTEUDO_CAMPO_INI, CONTEUDO_CAMPO_FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
END IF

'TÍTULO TABELA
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 5, TITULO_TABELA_INI, TITULO_TABELA_FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
	
END IF

'CONTEUDO TABELA LINHA PAR
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 6, CONT_TABELA_PAR_INI, CONT_TABELA_PAR_FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
	
END IF

'CONTEUDO TABELA LINHA ÍMPAR
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 7, CONT_TABELA_IMP_INI, CONT_TABELA_IMP_FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
	
END IF

'LINK
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 8, LINK_INI, LINK_FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
	
END IF

'COR DE FUNDO  DA PÁGINA
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 9, COR_FUNDO_PAG, FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
	
END IF

'COR DE FUNDO  DO TÍTULO DA TABELA
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 10, COR_FUNDO_TIT_TABELA, FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
	
END IF

'COR DE FUNDO TABELA LINHA PAR
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 11, COR_FUNDO_TABELA_PAR, FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
	
END IF

'COR DE FUNDO TABELA LINHA ÍMPAR
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 12, COR_FUNDO_TABELA_IMP, FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
	
END IF

'MENSAGEM ERRO
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 13, MSG_ERRO_INI, MSG_ERRO_FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
	
END IF

'MENSAGEM AVISO
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 14, MSG_AVISO_INI, MSG_AVISO_FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
	
END IF

'INDIC ANTES PRAZO
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 15, COR_ANT_PRAZO_INI, COR_ANT_PRAZO_FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
	
END IF

'COR INDIC NO PRAZO
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 16, COR_NO_PRAZO_INI, COR_NO_PRAZO_FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
	
END IF

'COR INDIC 1 DIA ATRASO
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 17, COR_1_ATRASO_INI, COR_1_ATRASO_FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
	
END IF

'COR INDIC 2 DIA ATRASO
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 18, COR_2_ATRASO_INI, COR_2_ATRASO_FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
	
END IF

'COR INDIC +2 DIA ATRASO
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 19, COR_MAIS_2_ATRASO_INI, COR_MAIS_2_ATRASO_FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
	
END IF

'CONFIGURAÇÕES DA PÁGINA
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 20, CONFIG_PAGINA, FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
	
END IF

'CONFIGURAÇÕES DA TABELA
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 21, CONFIG_TABELA_INI, CONFIG_TABELA_FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
	
END IF
'**************************************************
'Conteúdo campo em destaque
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 22, CONT_CAMPO_DESTAQUE_INI, CONT_CAMPO_DESTAQUE_FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
END IF

'Conteúdo tabela linha par em destaque
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 23, CONT_LINHA_PAR_DESTAQUE_INI, CONT_LINHA_PAR_DESTAQUE_FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
END IF

'Conteúdo tabela linha ímpar em destaque
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 24, CONT_LINHA_IMP_DESTAQUE_INI, CONT_LINHA_IMP_DESTAQUE_FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
END IF

'Mensagem Comercial
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 25, MENSAGEM_COMERCIAL_INI, MENSAGEM_COMERCIAL_FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
END IF

'Cor de fundo página inicial
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 26, COR_FUNDO_PAG_INICIAL, FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
END IF

'Cor de fundo Mensagem Comercial
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 27, COR_FUNDO_MSG_COMERCIAL, FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
END IF

'Link da Página inicial
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 28, LINK_PAG_INICIAL_INI, LINK_PAG_INICIAL_FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
END IF

'Título da Página inicial
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 29, TITULO_PAGINA_INICIAL_INI, TITULO_PAGINA_INICIAL_FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
END IF

'Sub Título da Página inicial
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 30, SUB_TIT_PAGINA_INICIAL_INI, SUB_TIT_PAGINA_INICIAL_FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
END IF

'Cor de Fundo Sub Titulo Pág. inicial
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 31, FUNDO_SUB_PAG_INICIAL, FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
END IF

'Cor de Fundo do Logo do Menu
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 32, COR_FUNDO_LOGO, FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
END IF

'DADOS DO USUÁRIO
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 33, DADOS_DO_USUARIO_INI, DADOS_DO_USUARIO_FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
END IF

'
'COR DE FUNDO DO MENU
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 34, COR_FUNDO_MENU, FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
END IF
'

'TÍTULO MENU
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 35, TITULO_MENU_INI, TITULO_MENU_FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
END IF


'SUB TÍTULO MENU
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 36, SUB_TITULO_MENU_INI, SUB_TITULO_MENU_FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
END IF


'LINK LINHA PAR
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 37, LINK_LINHA_PAR_INI, LINK_LINHA_PAR_FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
END IF


'LINK LINHA ÍMPAR
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 38, LINK_LINHA_IMP_INI, LINK_LINHA_IMP_FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
END IF

'MENSAGEM COMERCIAL 2
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 39, MENSAGEM_COMERCIAL2_INI, MENSAGEM_COMERCIAL2_FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
END IF
'
'CONFIGURAÇÃO DO SUB TÍTULO DA PÁGINA
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 40, CONF_SUB_TITULO, FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
END IF
'
'CONFIGURAÇÃO DO TÍTULO DA PÁGINA
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 41, CONF_TITULO_PAG, FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
END IF
  

'IMAGEM TÍTULO PÁGINA
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 42, IMAGEM_TITULO,  FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
END IF
'
'IMAGEM SUB TÍTULO PÁGINA
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 43, IMAGEM_SUB_TITULO,  FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
END IF
'
'COR DE FUNDO CAMPO DESTAQUE
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 44, COR_FUNDO_CAMPO_DEST,  FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
END IF
'
'IMAGEM LINK LINHA PAR
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 45, IMAGEM_LINK_PAR, FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
END IF
'
'IMAGEM LINK LINHA ÍMPAR
RETORNO = OBJLOG.gConfiguracao(Empresa_id, nome_banco, Nome_Produto, 46, IMAGEM_LINK_IMPAR,  FIM, msg_ret)
IF RETORNO <> 0 THEN
	session("msg_erro") = msg_ret
END IF
'
SESSION("GL_TITULO_PAGINA_INI") = TITULO_PAGINA_INI  
SESSION("GL_TITULO_PAGINA_FIM") = TITULO_PAGINA_FIM 
SESSION("GL_SUB_TITULO_INI") = SUB_TITULO_INI 
SESSION("GL_SUB_TITULO_FIM") = SUB_TITULO_FIM 
SESSION("GL_TITULO_CAMPO_INI") = TITULO_CAMPO_INI 
SESSION("GL_TITULO_CAMPO_FIM") = TITULO_CAMPO_FIM 
SESSION("GL_CONTEUDO_CAMPO_INI") = CONTEUDO_CAMPO_INI
SESSION("GL_CONTEUDO_CAMPO_FIM") = CONTEUDO_CAMPO_FIM 
SESSION("GL_TITULO_TABELA_INI") = TITULO_TABELA_INI 
SESSION("GL_TITULO_TABELA_FIM") = TITULO_TABELA_FIM 
SESSION("GL_CONT_TABELA_PAR_INI")=  CONT_TABELA_PAR_INI 
SESSION("GL_CONT_TABELA_PAR_FIM") = CONT_TABELA_PAR_FIM 
SESSION("GL_CONT_TABELA_IMP_INI") = CONT_TABELA_IMP_INI 
SESSION("GL_CONT_TABELA_IMP_FIM") = CONT_TABELA_IMP_FIM 
SESSION("GL_LINK_INI") = LINK_INI 
SESSION("GL_LINK_FIM") = LINK_FIM 
SESSION("GL_COR_FUNDO_PAG") = COR_FUNDO_PAG 
SESSION("GL_COR_FUNDO_TIT_TABELA") = COR_FUNDO_TIT_TABELA 
SESSION("GL_COR_FUNDO_TABELA_PAR") = COR_FUNDO_TABELA_PAR 
SESSION("GL_COR_FUNDO_TABELA_IMP") = COR_FUNDO_TABELA_IMP 
SESSION("GL_MSG_ERRO_INI") = MSG_ERRO_INI
SESSION("GL_MSG_ERRO_FIM") = MSG_ERRO_FIM
SESSION("GL_MSG_AVISO_INI") = MSG_AVISO_INI
SESSION("GL_MSG_AVISO_FIM") = MSG_AVISO_FIM
SESSION("GL_COR_ANT_PRAZO_INI") = COR_ANT_PRAZO_INI 
SESSION("GL_COR_ANT_PRAZO_FIM") = COR_ANT_PRAZO_FIM
SESSION("GL_COR_NO_PRAZO_INI") = COR_NO_PRAZO_INI
SESSION("GL_COR_NO_PRAZO_FIM") = COR_NO_PRAZO_FIM
SESSION("GL_COR_1_ATRASO_INI") = COR_1_ATRASO_INI
SESSION("GL_COR_1_ATRASO_FIM") = COR_1_ATRASO_FIM
SESSION("GL_COR_2_ATRASO_INI") = COR_2_ATRASO_INI 
SESSION("GL_COR_2_ATRASO_FIM") = COR_2_ATRASO_FIM
SESSION("GL_COR_MAIS_2_ATRASO_INI") = COR_MAIS_2_ATRASO_INI 
SESSION("GL_COR_MAIS_2_ATRASO_FIM") = COR_MAIS_2_ATRASO_FIM
SESSION("GL_CONFIG_PAGINA") = CONFIG_PAGINA 
SESSION("GL_CONFIG_TABELA") = CONFIG_TABELA_INI 
'/**********************
SESSION("GL_CONT_CAMPO_DESTAQUE_INI") = CONT_CAMPO_DESTAQUE_INI  
SESSION("GL_CONT_CAMPO_DESTAQUE_FIM") = CONT_CAMPO_DESTAQUE_FIM 
SESSION("GL_CONT_LINHA_PAR_DESTAQUE_INI") = CONT_LINHA_PAR_DESTAQUE_INI 
SESSION("GL_CONT_LINHA_PAR_DESTAQUE_FIM") = CONT_LINHA_PAR_DESTAQUE_FIM 
SESSION("GL_CONT_LINHA_IMP_DESTAQUE_INI") = CONT_LINHA_IMP_DESTAQUE_INI 
SESSION("GL_CONT_LINHA_IMP_DESTAQUE_FIM") = CONT_LINHA_IMP_DESTAQUE_FIM 
SESSION("GL_MENSAGEM_COMERCIAL_INI") = MENSAGEM_COMERCIAL_INI 
SESSION("GL_MENSAGEM_COMERCIAL_FIM") = MENSAGEM_COMERCIAL_FIM 
SESSION("GL_COR_FUNDO_PAG_INICIAL") = COR_FUNDO_PAG_INICIAL 
SESSION("GL_COR_FUNDO_MSG_COMERCIAL") = COR_FUNDO_MSG_COMERCIAL 
SESSION("GL_LINK_PAG_INICIAL_INI") = LINK_PAG_INICIAL_INI 
SESSION("GL_LINK_PAG_INICIAL_FIM") = LINK_PAG_INICIAL_FIM 
SESSION("GL_TITULO_PAGINA_INICIAL_INI") = TITULO_PAGINA_INICIAL_INI 
SESSION("GL_TITULO_PAGINA_INICIAL_FIM") = TITULO_PAGINA_INICIAL_FIM 
SESSION("GL_SUB_TIT_PAGINA_INICIAL_INI") = SUB_TIT_PAGINA_INICIAL_INI 
SESSION("GL_SUB_TIT_PAGINA_INICIAL_FIM") = SUB_TIT_PAGINA_INICIAL_FIM 
SESSION("GL_FUNDO_SUB_PAG_INICIAL") = FUNDO_SUB_PAG_INICIAL 
SESSION("GL_COR_FUNDO_LOGO") = COR_FUNDO_LOGO 
SESSION("GL_DADOS_DO_USUARIO_INI") = DADOS_DO_USUARIO_INI 
SESSION("GL_DADOS_DO_USUARIO_FIM") = DADOS_DO_USUARIO_FIM 
SESSION("GL_COR_FUNDO_MENU") = COR_FUNDO_MENU 
SESSION("GL_TITULO_MENU_INI") = TITULO_MENU_INI 
SESSION("GL_TITULO_MENU_FIM") = TITULO_MENU_FIM 
SESSION("GL_SUB_TITULO_MENU_INI") = SUB_TITULO_MENU_INI 
SESSION("GL_SUB_TITULO_MENU_FIM") = SUB_TITULO_MENU_FIM 
SESSION("GL_LINK_LINHA_PAR_INI") = LINK_LINHA_PAR_INI 
SESSION("GL_LINK_LINHA_PAR_FIM") = LINK_LINHA_PAR_FIM 
SESSION("GL_LINK_LINHA_IMP_INI") = LINK_LINHA_IMP_INI  
SESSION("GL_LINK_LINHA_IMP_FIM") = LINK_LINHA_IMP_FIM 
SESSION("GL_MENSAGEM_COMERCIAL2_INI") = MENSAGEM_COMERCIAL2_INI
SESSION("GL_MENSAGEM_COMERCIAL2_FIM") = MENSAGEM_COMERCIAL2_FIM

SESSION("GL_IMAGEM_TITULO") = IMAGEM_TITULO
SESSION("GL_IMAGEM_SUB_TITULO") = IMAGEM_SUB_TITULO
SESSION("GL_IMAGEM_LINK_PAR") = IMAGEM_LINK_PAR
SESSION("GL_IMAGEM_LINK_IMPAR") = IMAGEM_LINK_IMPAR
SESSION("GL_COR_FUNDO_CAMPO_DEST") = COR_FUNDO_CAMPO_DEST
'*************
SESSION("GL_CONF_TITULO_PAG")  = CONF_TITULO_PAG
SESSION("GL_CONF_SUB_TITULO") = CONF_SUB_TITULO
END FUNCTION
%>