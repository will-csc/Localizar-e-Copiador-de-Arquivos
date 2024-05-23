Attribute VB_Name = "Main"
Sub AcharDocumentos()

    '------------ VERIFICAÇÃO DE ERROS --------------
    PenteFino = Verificar_Erros()
    If PenteFino = False Then
        Exit Sub
    End If
    
    '--------------- DEFINIÇÃO DE VARIAVEIS ----------------
    Set ws = ThisWorkbook.Sheets("LOCALIZAR DOC")
    Caminho = ws.Range("A:E").Find("CAMINHO ORIGINAL").Offset(1, 0).Value
    CaminhoFinal = ws.Range("A:E").Find("PASTA ONDE").Offset(1, 0).Value
    Set Key_words = New Collection
    Set Fso = CreateObject("Scripting.FileSystemObject")
    
    'OPÇÕES
    Opção_Pastas = ws.Range("A:E").Find("CAM. ORIG.?").Offset(1, 0).Value
    Opção_CopRem = ws.Range("A:E").Find("OU MOVER?").Offset(1, 0).Value
    
    '------------ CONFERENCIA PALAVRAS -------------
    For i = 1 To 10
        Procurar = "PALAVRA CHAVE " & i
        Palavra = Trim(ws.Range("A:D").Find(Procurar).Offset(1, 0).Value)
        If Palavra <> "" Then
            Key_words.Add Palavra
        End If
    Next i
    
    TotalWords = Key_words.Count
    '------------ LOOPING P/ PEGAR OS ARQUIVOS -------------
    Set Folder = Fso.GetFolder(Caminho)
    
    If Opção_Pastas = "NÃO" Then
        Call AcharDoc(Folder, Key_words, Opção_CopRem, TotalWords, CaminhoFinal)
    Else
        For Each Pasta In Folder.SubFolders
            Call AcharDoc(Folder, Key_words, Opção_CopRem, TotalWords, CaminhoFinal)
            
            Set Folder1 = Fso.GetFolder(Pasta)
            If Folder1.SubFolders.Count > 0 Then
                For Each Pasta1 In Folder1.SubFolders
                    Call AcharDoc(Pasta1, Key_words, Opção_CopRem, TotalWords, CaminhoFinal)
                    
                    Set Folder2 = Fso.GetFolder(Pasta1)
                    For Each Pasta2 In Folder2.SubFolders
                        Call AcharDoc(Pasta2, Key_words, Opção_CopRem, TotalWords, CaminhoFinal)
                         
                        Set Folder3 = Fso.GetFolder(Pasta2)
                        For Each Pasta3 In Folder3.SubFolders
                            Call AcharDoc(Pasta3, Key_words, Opção_CopRem, TotalWords, CaminhoFinal)
                        Next Pasta3
                    Next Pasta2
                Next Pasta1
            End If
        Next Pasta
    End If
    
End Sub

Sub AcharDoc(Folder, Key_words, Opção_CopRem, TotalWords, CaminhoFinal)
    Set Fso = CreateObject("Scripting.FileSystemObject")
    Set Pasta = Fso.GetFolder(Folder)
    For Each File In Pasta.Files
            NomeArq = File.Name
            CaminhoArq = File.Path
                
            For Each Palavra In Key_words
                If InStr(1, NomeArq, Palavra) > 0 Then
                    Cont = Cont + 1
                If Cont = TotalWords And Opção_CopRem = "COPIAR" Then
                    Fso.CopyFile CaminhoArq, CaminhoFinal & "\" & NomeArq
                ElseIf Cont = TotalWords And Opção_CopRem = "MOVER" Then
                    Fso.mOVEFile CaminhoArq, CaminhoFinal & "\" & NomeArq
                End If
            End If
        Next Palavra
        Cont = 0
                
    Next File
End Sub

