<%
'-------------------------------------------------------------------------------------------
'Adiciona o comportamento de "namespaces" р sessуo. Um namespace pode ser utilizado
'por determinadas pсginas, e quando o usuсrio acessar uma pсgina que nуo utiliza
'aquele namespace, as chaves dentro dele sуo limpas.
'
'Utilizaчуo:
'
'	TempSession("MeuNamespace")("NomeDoParticipante") = "Fulano"
'
'	Response.Write TempSession("MeuNamespace")("NomeDoParticipante") 'Imprime "Fulano"
'
'-------------------------------------------------------------------------------------------


'Prefixo dos эtens que estуo em "Temp Session"
Const TempSession_prefix_ = "__TEMP_SESSION__"

'Guarda os namespaces utilizados
Dim TempSession_UsedNamespaces_
Set TempSession_UsedNamespaces_ = CreateObject("Scripting.Dictionary")


'Representa um namespace (uma coleчуo de itens relacionados) na sessуo.
'CLASSE INTERNA - NУO DEVE SER UTILIZADA DIRETAMENTE (utilize a funчуo "TempSession")
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
			Call Err.Raise(-1, "cTempSession.Using", "O parтmetro namespace (primeiro parтmetro) deve ser informado.")
		End If

		If namespace_ <> "" Then
			Call Err.Raise(-1, "cTempSession.Using", "O namespace desta Temp Session jс foi inicializado.")
		End If

		namespace_ = namespace

		Set Init = Me

	End Function


	Private Function GetFullKey(key)

		If Not IsNotEmpty(namespace_) Then
			Call Err.Raise(-1, "cTempSession.GetFullKey", "O namespace desta Temp Session nуo foi inicializado.")
		End If

		GetFullKey = TempSession_prefix_ & namespace_ & "__" & key

	End Function

End Class


'Funчуo que dс acesso р funcionalidade de sessуo temporсria
Function TempSession(key)

	Dim tempSessionNs

	If Not TempSession_UsedNamespaces_.Exists(key) Then 'Cria o namespace caso ele nуo exista

		Set tempSessionNs = New cTempSessionNamespace_
		tempSessionNs.Init(key)

		Set TempSession_UsedNamespaces_(key) = tempSessionNs

	Else
		Set tempSessionNs = TempSession_UsedNamespaces_(key)
	End If

	Set TempSession = tempSessionNs

End Function


'Limpa todas as sessѕes temporсrias que nуo foram utilizadas. Deve ser invocada no final de todas as pсginas (foi incluэda em footer.asp)
Function TempSession_CleanUnused()

	Dim i

	For i = 1 To Session.Contents.Count
	
		Dim sessionKey : sessionKey = Session.Contents.Key(i)

		'Verifica se esta chave estс em Temp Session (ou seja, se possui o prefixo de Temp Session)
		If InStr(sessionKey, TempSession_prefix_) = 1 Then

			Dim deveRemover : deveRemover = True

			'Verifica se esta chave estс em um namespace que foi utilizado
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

'Forчa a remoчуo de um namespace especэfico
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