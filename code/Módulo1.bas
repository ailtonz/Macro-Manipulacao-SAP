Attribute VB_Name = "Módulo1"

Sub btn_executar_clique()

    Dim int_linha                         As Integer
    
    Dim int_qtd_repeticoes_total          As Integer
    Dim int_qtd_repeticoes_realizadas     As Integer
    
    Dim tms_inicio                        As Date
    Dim tms_fim                           As Date
    
    
    int_qtd_repeticoes_realizadas = 0
    int_qtd_repeticoes_total = Cells(15, 1)
    
    tms_inicio = Now
    Cells(9, 1) = Format(tms_inicio, "yyyy-mm-dd hh:mm:ss")

    For int_linha = 5 To int_qtd_repeticoes_total + 5
    
        int_qtd_repeticoes_realizadas = int_qtd_repeticoes_realizadas + 1
        
    
        SendKeys "%{TAB}", True
        
        SendKeys "{HOME}", True
        SendKeys "+{END}", True
        SendKeys "{DEL}", True
        SendKeys "{TAB}", True
        
        SendKeys "{HOME}", True
        SendKeys "+{END}", True
        SendKeys "{DEL}", True
        SendKeys "{TAB}", True
        
        SendKeys "{HOME}", True
        SendKeys "+{END}", True
        SendKeys "{DEL}", True
        SendKeys "{TAB}", True
        
        SendKeys "{HOME}", True
        SendKeys "+{END}", True
        SendKeys "{DEL}", True
        SendKeys "{TAB}", True
        
        SendKeys "{HOME}", True
        SendKeys "+{END}", True
        SendKeys "{DEL}", True
        SendKeys "{TAB}", True
        
        SendKeys "{HOME}", True
        SendKeys "+{END}", True
        SendKeys "+{DEL}", True
        
        SendKeys "+{TAB}", True
        SendKeys "+{TAB}", True
        SendKeys "+{TAB}", True
        SendKeys "+{TAB}", True
        SendKeys "+{TAB}", True
        
        SendKeys "{TAB}", True
        SendKeys "{TAB}", True

        SendKeys Cells(5, 3), True
        SendKeys "{TAB}", True

        SendKeys Cells(5, 4), True
        SendKeys "{TAB}", True

        SendKeys Cells(5, 5), True
        SendKeys "{F5}", True

        Application.Wait (Now + TimeValue("00:00:01"))

        SendKeys "{TAB}", True
        SendKeys "{TAB}", True
        SendKeys Cells(5, 6), True

        SendKeys "{TAB}", True
        SendKeys "{TAB}", True
        SendKeys "{TAB}", True
        SendKeys "{TAB}", True
        SendKeys "{TAB}", True
        SendKeys "{TAB}", True
        SendKeys Cells(5, 7), True

        SendKeys "{TAB}", True
        SendKeys Cells(5, 8), True

        SendKeys "{TAB}", True
        SendKeys "{TAB}", True
        SendKeys "{TAB}", True
        SendKeys Cells(5, 9), True

        SendKeys "{TAB}", True
        SendKeys "{TAB}", True
        SendKeys "{TAB}", True
        SendKeys "{TAB}", True
        SendKeys Cells(5, 10), True

        SendKeys "{TAB}", True
        SendKeys "{TAB}", True
        SendKeys "{TAB}", True
        SendKeys "{TAB}", True
        SendKeys Cells(5, 11), True

        SendKeys "{TAB}", True
        SendKeys "{TAB}", True
        SendKeys "{TAB}", True
        SendKeys "{TAB}", True
        SendKeys "{TAB}", True
        SendKeys "{TAB}", True
        SendKeys Cells(5, 12), True

        SendKeys "^s", True
        Application.Wait (Now + TimeValue("00:00:01"))

        SendKeys "+{F3}", True
        Application.Wait (Now + TimeValue("00:00:01"))

        SendKeys "{HOME}", True
        SendKeys "+{END}", True
        SendKeys "^c", True

        SendKeys "%{TAB}", True
        Range("M" & int_linha).Select
        Range("M" & int_linha).PasteSpecial
    Next

    tms_fim = Now
    Cells(10, 1) = Format(tms_fim, "yyyy-mm-dd hh:mm:ss")
    
    tms_fim = Now
    Cells(11, 1) = Format(tms_fim - tms_inicio, "hh:mm:ss")
End Sub
