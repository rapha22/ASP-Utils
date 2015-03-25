<%
'-------------------------------------------------------------------------------------------
'Adiciona o comportamento de "namespaces" � sess�o. Um namespace pode ser utilizado
'por determinadas p�ginas, e quando o usu�rio acessar uma p�gina que n�o utiliza
'aquele namespace, as chaves dentro dele s�o limpas.
'
'Utiliza��o:
'
'	TempSession("MeuNamespace")("NomeDoParticipante") = "Fulano"
'
'	Response.Write TempSession("MeuNamespace")("NomeDoParticipante") 'Imprime "Fulano"
'
'-------------------------------------------------------------------------------------------


'Prefixo dos �tens que est�o em "Temp Session"
Const TempSession_prefix_ = "__TEMP_SESSION__"

'Guarda os namespaces utilizados
Dim TempSession_UsedNamespaces_
Set TempSession_UsedNamespaces_ = CreateObject("Scripting.Dictionary")


'Representa um namespace (uma cole��o de itens relacionados) na sess�o.
'CLASSE INTERNA - N�O DEVE SER UTILIZADA DIRETAMENTE (utilize a fun��o "TempSession")
Class cTempSessionNamespace_

	Private namespace_


	Public Default Property Get Item(key)	
		Dim fullKey : fullKey = GetFullKey(key)
	
		If IsObject(Session(fullKey)) Then
			Set Item = Session(fullKey)
		Else
			Item = Session(fullKey)
		End If
	End Property

	Public Property Let Item(key, value)
		Session(GetFullKey(key)) = value
	End Property

	Public Property Set Item(key, value)
		Set Session(GetFullKey(key)) = value
	End Property
	
	Public Function Init(namespace)

		If namespace = "" Then
			Call Err.Raise(-1, "cTempSession.Using", "O par�metro namespace (primeiro par�metro) deve ser informado.")
		End If

		If namespace_ <> "" Then
			Call Err.Raise(-1, "cTempSession.Using", "O namespace desta Temp Session j� foi inicializado.")
		End If

		namespace_ = namespace

		Set Init = Me

	End Function


	Private Function GetFullKey(key)

		If Not IsNotEmpty(namespace_) Then
			Call Err.Raise(-1, "cTempSession.GetFullKey", "O namespace desta Temp Session n�o foi inicializado.")
		End If

		GetFullKey = TempSession_prefix_ & namespace_ & "__" & key

	End Function

End Class


'Fun��o que d� acesso � funcionalidade de sess�o tempor�ria
Function TempSession(key)

	Dim tempSessionNs

	If Not TempSession_UsedNamespaces_.Exists(key) Then 'Cria o namespace caso ele n�o exista

		Set tempSessionNs = New cTempSessionNamespace_
		tempSessionNs.Init(key)

		Set TempSession_UsedNamespaces_(key) = tempSessionNs

	Else
		Set tempSessionNs = TempSession_UsedNamespaces_(key)
	End If

	Set TempSession = tempSessionNs

End Function


'Limpa todas as sess�es tempor�rias que n�o foram utilizadas. Deve ser invocada no final de todas as p�ginas (foi inclu�da em footer.asp)
Function TempSession_CleanUnused()

	Dim i

	For i = 1 To Session.Contents.Count
	
		Dim sessionKey : sessionKey = Session.Contents.Key(i)

		'Verifica se esta chave est� em Temp Session (ou seja, se possui o prefixo de Temp Session)
		If InStr(sessionKey, TempSession_prefix_) = 1 Then

			Dim deveRemover : deveRemover = True

			'Verifica se esta chave est� em um namespace que foi utilizado
			For Each tempSessionkey In TempSession_UsedNamespaces_

				If InStr(sessionKey, TempSession_prefix_ & tempSessionkey) = 1 Then
					deveRemover = False
					Exit For
				End If

			Next

			If deveRemover Then 
				Session.Contents.Remove(sessionKey)
				i = i - 1
			End If

		End If

	Next

End Function

'For�a a remo��o de um namespace espec�fico
Function TempSession_Remove(namespace)

	Dim i

	For i = 1 To Session.Contents.Count
	
		Dim sessionKey : sessionKey = Session.Contents.Key(i)

		If InStr(sessionKey, TempSession_prefix_ & namespace) = 1 Then
		
			Session.Contents.Remove(sessionKey)
			i = i - 1
			
		End If
		
	Next

End Function

%>