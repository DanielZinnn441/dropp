# Encoding UTF-8 with BOM
$OutputEncoding = [System.Text.Encoding]::UTF8

$targetDir = "c:\Users\d4nie\OneDrive\Imagens\PRODTOS ENCONTARDO"
$csvFile = Join-Path $targetDir "planilha_pesquisa_produtos_dropship.csv"
$xlsxFile = Join-Path $targetDir "planilha_pesquisa_produtos_dropship.xlsx"

# 1. Create CSV File
$csvContent = @"
Nº;Data da pesquisa;Produto;Categoria;Link do anúncio (Amazon);Preço de venda na Amazon (R$);Vendas estimadas/mês;Nº de avaliações;Nota média (0-5);Nº de concorrentes;Fornecedor (loja na Shopee);Link do fornecedor;Preço de custo unitário (R$);Frete até o cliente (R$) - estimado;Custo entregue no cliente (R$);Qtd. mínima por pedido;Prazo de entrega até o cliente (dias);Envia com nota/embalagem própria?;Aceita endereço de entrega diferente?;Preço de venda no anúncio (R$);Frete cobrado do cliente (R$);Taxa Amazon (%);Taxa Amazon (R$);Outros custos (R$);Custo total por venda (R$);Lucro líquido por unidade (R$);Margem de lucro (%);ROI (%);Status;Prioridade;Observações
1;10/09/2026;Kit 3 Potes Organizadores para Geladeira Premium;Cozinha / Organização;https://www.amazon.com.br/Organizadores-Geladeira-Escorredores-Herm%C3%A9ticas-Conserva/dp/B0H9P9QR57;43,90;N/I;1;5,0;1;Shopee (Cruzeiro, SP);https://shopee.com.br/Kit-3-potes-Organizadores-Geladeira-Organizador-Alimentos-Vegetais-Cozinha-Moderna-i.629805123.58205197736;32,98;12,04;45,02;1;5 a 10 dias;Não;Sim;43,90;0,00;15%;6,59;0,00;51,61;-7,71;-17,56%;-17,13%;Em análise;Alta;Nota 5,0 na Amazon. Com cupom frete grátis Shopee o lucro fica em +R$ 4,33 (9,86%).
2;10/09/2026;Pré Treino Insane Clown 350g Demons Lab - Blue Crystal;Suplementos / Saúde;https://www.amazon.com.br/Treino-Insane-Clown-350g-Demons/dp/B0GYQM5WFT;116,90;N/I;N/I;N/I;N/I;Shopee (Cruzeiro, SP);https://shopee.com.br/INSANE-CLOWN-350G-BLUE-CRYSTAL-i.384424516.22391686410;99,75;10,83;110,58;1;5 a 10 dias;Não;Sim;116,90;0,00;15%;17,54;0,00;128,12;-11,22;-9,60%;-10,15%;Em análise;Média;Recomenda-se ajustar preço na Amazon para R$ 139,90 para garantir margem positiva.
3;10/09/2026;Kit 6 Pares Meias Cano Invisível Lupo Sport Sapatilha;Moda / Roupas;https://www.amazon.com.br/Lupo-esportivas-algod%C3%A3o-respir%C3%A1vel-pares/dp/B00PROD03;74,90;N/I;1449;4,8;3;Shopee Lojas Oficiais;https://shopee.com.br/Kit-De-6-Pares-Meias-Cano-Invis%C3%ADvel-Lupo-Sport-Algod%C3%A3o-Soquete-Sapatilha-Unissex;49,23;9,62;58,85;1;5 a 10 dias;Não;Sim;74,90;0,00;15%;11,24;0,00;70,09;4,82;6,43%;8,18%;Em análise;Alta;Alta demanda na Amazon (4,8★ com 1,4k+ avaliações). Lucro sobe para R$ 14,44 (19,27%) com frete grátis.
4;10/09/2026;Carregador Turbo Duo 168W USB-C e USB-A com Cabo USB-C;Eletrônicos / Acessórios;https://www.amazon.com.br/dp/B0H5XWN3WD;79,90;N/I;N/I;N/I;N/I;Shopee (Cruzeiro, SP);https://shopee.com.br/product/1572491828/22094298275;25,88;9,62;35,50;1;5 a 10 dias;Não;Sim;79,90;0,00;15%;11,99;0,00;47,49;32,42;40,57%;91,31%;Em análise;Média;Produto Destaque! Lucro de R$ 32,42 a R$ 42,04 por unidade (Margem > 40% e ROI até 162%).
5;10/09/2026;Jogo de Talheres Faqueiro Aço Inox 24 Peças;Cozinha / Utensílios;https://www.amazon.com.br/Jogo-Talheres-Faqueiro-Inox-Pecas/dp/B00PROD05;52,69;N/I;4;2,6;1;Shopee (Cruzeiro, SP);https://shopee.com.br/Jogo-de-Talheres-24-Pe%C3%A7as-Faqueiro-A%C3%A7o-Inox-Faca-Garfo-Colher-de-Mesa-e-Colheres-de-Sobremesa;25,40;9,62;35,02;1;5 a 10 dias;Não;Sim;52,69;0,00;15%;7,90;0,00;42,92;9,77;18,54%;27,89%;Em análise;Média;Boa margem de lucro (18% a 36%), porém requer atenção pela nota 2,6 na Amazon.
"@

$utf8WithBom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($csvFile, $csvContent, $utf8WithBom)
Write-Host "CSV gerado com sucesso: $csvFile"

# 2. Try Excel COM
try {
    $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "Viabilidade Produtos"

    # Title
    $ws.Cells.Item(1, 1) = "PLANILHA DE PESQUISA E VIABILIDADE DE PRODUTOS - Dropship: Shopee (fornecedor) envia direto ao cliente da Amazon"
    $ws.Range("A1:AE1").Merge()
    $ws.Range("A1:AE1").Font.Bold = $true
    $ws.Range("A1:AE1").Font.Size = 13
    $ws.Range("A1:AE1").Font.ColorIndex = 2
    $ws.Range("A1:AE1").Interior.Color = 0x5C301B
    $ws.Range("A1:AE1").HorizontalAlignment = -4108

    # Headers Row 2
    $ws.Range("A2:D2").Merge(); $ws.Range("A2:D2").Value2 = "PRODUTO"; $ws.Range("A2:D2").Interior.Color = 0x7030A0; $ws.Range("A2:D2").Font.Bold = $true; $ws.Range("A2:D2").Font.ColorIndex = 2; $ws.Range("A2:D2").HorizontalAlignment = -4108
    $ws.Range("E2:J2").Merge(); $ws.Range("E2:J2").Value2 = "PESQUISA NA AMAZON (mercado)"; $ws.Range("E2:J2").Interior.Color = 0x0066FF; $ws.Range("E2:J2").Font.Bold = $true; $ws.Range("E2:J2").Font.ColorIndex = 2; $ws.Range("E2:J2").HorizontalAlignment = -4108
    $ws.Range("K2:O2").Merge(); $ws.Range("K2:O2").Value2 = "FORNECEDOR NA SHOPEE - ENVIO DIRETO"; $ws.Range("K2:O2").Interior.Color = 0x339900; $ws.Range("K2:O2").Font.Bold = $true; $ws.Range("K2:O2").Font.ColorIndex = 2; $ws.Range("K2:O2").HorizontalAlignment = -4108
    $ws.Range("P2:S2").Merge(); $ws.Range("P2:S2").Value2 = "ENVIO DIRETO AO CLIENTE"; $ws.Range("P2:S2").Interior.Color = 0x339900; $ws.Range("P2:S2").Font.Bold = $true; $ws.Range("P2:S2").Font.ColorIndex = 2; $ws.Range("P2:S2").HorizontalAlignment = -4108
    $ws.Range("T2:Y2").Merge(); $ws.Range("T2:Y2").Value2 = "PRECIFICAÇÃO E CUSTOS"; $ws.Range("T2:Y2").Interior.Color = 0x00A5FF; $ws.Range("T2:Y2").Font.Bold = $true; $ws.Range("T2:Y2").Font.ColorIndex = 2; $ws.Range("T2:Y2").HorizontalAlignment = -4108
    $ws.Range("Z2:AB2").Merge(); $ws.Range("Z2:AB2").Value2 = "RESULTADO"; $ws.Range("Z2:AB2").Interior.Color = 0x800080; $ws.Range("Z2:AB2").Font.Bold = $true; $ws.Range("Z2:AB2").Font.ColorIndex = 2; $ws.Range("Z2:AB2").HorizontalAlignment = -4108
    $ws.Range("AC2:AE2").Merge(); $ws.Range("AC2:AE2").Value2 = "GESTÃO"; $ws.Range("AC2:AE2").Interior.Color = 0x595959; $ws.Range("AC2:AE2").Font.Bold = $true; $ws.Range("AC2:AE2").Font.ColorIndex = 2; $ws.Range("AC2:AE2").HorizontalAlignment = -4108

    # Headers Row 3
    $headers = @(
        "Nº", "Data da pesquisa", "Produto", "Categoria",
        "Link do anúncio (Amazon)", "Preço de venda na Amazon (R$)", "Vendas estimadas/mês", "Nº de avaliações", "Nota média (0-5)", "Nº de concorrentes",
        "Fornecedor (loja na Shopee)", "Link do fornecedor", "Preço de custo unitário (R$)", "Frete até o cliente (R$) - estimado", "Custo entregue no cliente (R$)",
        "Qtd. mínima por pedido", "Prazo de entrega até o cliente (dias)", "Envia com nota/embalagem própria?", "Aceita endereço de entrega diferente?",
        "Preço de venda no anúncio (R$)", "Frete cobrado do cliente (R$)", "Taxa Amazon (%)", "Taxa Amazon (R$)", "Outros custos (R$)", "Custo total por venda (R$)",
        "Lucro líquido por unidade (R$)", "Margem de lucro (%)", "ROI (%)",
        "Status", "Prioridade", "Observações"
    )

    for ($i = 0; $i -lt $headers.Length; $i++) {
        $c = $ws.Cells.Item(3, $i + 1)
        $c.Value2 = $headers[$i]
        $c.Font.Bold = $true
        $c.Interior.Color = 0xE0E0E0
        $c.HorizontalAlignment = -4108
    }

    # Rows 4 to 8
    $items = @(
        @{ Num=1; Date="10/09/2026"; Prod="Kit 3 Potes Organizadores para Geladeira Premium"; Cat="Cozinha / Organização"; AmzL="https://www.amazon.com.br/Organizadores-Geladeira-Escorredores-Herm%C3%A9ticas-Conserva/dp/B0H9P9QR57"; AmzP=43.90; Sales="N/I"; Rev=1; Rating=5.0; Comp=1; ShpS="Shopee (Cruzeiro, SP)"; ShpL="https://shopee.com.br/Kit-3-potes-Organizadores-Geladeira-Organizador-Alimentos-Vegetais-Cozinha-Moderna-i.629805123.58205197736"; ShpC=32.98; ShpF=12.04; MinQ=1; Del="5 a 10 dias"; Pkg="Não"; Addr="Sim"; FrtC=0.00; TaxP=0.15; OthC=0.00; St="Em análise"; Prio="Alta"; Obs="Nota 5,0 na Amazon. Com cupom frete grátis Shopee o lucro fica em +R$ 4,33 (9,86%)." },
        @{ Num=2; Date="10/09/2026"; Prod="Pré Treino Insane Clown 350g Demons Lab - Blue Crystal"; Cat="Suplementos / Saúde"; AmzL="https://www.amazon.com.br/Treino-Insane-Clown-350g-Demons/dp/B0GYQM5WFT"; AmzP=116.90; Sales="N/I"; Rev="N/I"; Rating="N/I"; Comp="N/I"; ShpS="Shopee (Cruzeiro, SP)"; ShpL="https://shopee.com.br/INSANE-CLOWN-350G-BLUE-CRYSTAL-i.384424516.22391686410"; ShpC=99.75; ShpF=10.83; MinQ=1; Del="5 a 10 dias"; Pkg="Não"; Addr="Sim"; FrtC=0.00; TaxP=0.15; OthC=0.00; St="Em análise"; Prio="Média"; Obs="Recomenda-se ajustar preço na Amazon para R$ 139,90 para garantir margem positiva." },
        @{ Num=3; Date="10/09/2026"; Prod="Kit 6 Pares Meias Cano Invisível Lupo Sport Sapatilha"; Cat="Moda / Roupas"; AmzL="https://www.amazon.com.br/Lupo-esportivas-algod%C3%A3o-respir%C3%A1vel-pares/dp/B00PROD03"; AmzP=74.90; Sales="N/I"; Rev=1449; Rating=4.8; Comp=3; ShpS="Shopee Lojas Oficiais"; ShpL="https://shopee.com.br/Kit-De-6-Pares-Meias-Cano-Invis%C3%ADvel-Lupo-Sport-Algod%C3%A3o-Soquete-Sapatilha-Unissex"; ShpC=49.23; ShpF=9.62; MinQ=1; Del="5 a 10 dias"; Pkg="Não"; Addr="Sim"; FrtC=0.00; TaxP=0.15; OthC=0.00; St="Em análise"; Prio="Alta"; Obs="Alta demanda na Amazon (4,8★ com 1,4k+ avaliações). Lucro sobe para R$ 14,44 (19,27%) com frete grátis." },
        @{ Num=4; Date="10/09/2026"; Prod="Carregador Turbo Duo 168W USB-C e USB-A com Cabo USB-C"; Cat="Eletrônicos / Acessórios"; AmzL="https://www.amazon.com.br/dp/B0H5XWN3WD"; AmzP=79.90; Sales="N/I"; Rev="N/I"; Rating="N/I"; Comp="N/I"; ShpS="Shopee (Cruzeiro, SP)"; ShpL="https://shopee.com.br/product/1572491828/22094298275"; ShpC=25.88; ShpF=9.62; MinQ=1; Del="5 a 10 dias"; Pkg="Não"; Addr="Sim"; FrtC=0.00; TaxP=0.15; OthC=0.00; St="Em análise"; Prio="Média"; Obs="Produto Destaque! Lucro de R$ 32,42 a R$ 42,04 por unidade (Margem > 40% e ROI até 162%)." },
        @{ Num=5; Date="10/09/2026"; Prod="Jogo de Talheres Faqueiro Aço Inox 24 Peças"; Cat="Cozinha / Utensílios"; AmzL="https://www.amazon.com.br/Jogo-Talheres-Faqueiro-Inox-Pecas/dp/B00PROD05"; AmzP=52.69; Sales="N/I"; Rev=4; Rating=2.6; Comp=1; ShpS="Shopee (Cruzeiro, SP)"; ShpL="https://shopee.com.br/Jogo-de-Talheres-24-Pe%C3%A7as-Faqueiro-A%C3%A7o-Inox-Faca-Garfo-Colher-de-Mesa-e-Colheres-de-Sobremesa"; ShpC=25.40; ShpF=9.62; MinQ=1; Del="5 a 10 dias"; Pkg="Não"; Addr="Sim"; FrtC=0.00; TaxP=0.15; OthC=0.00; St="Em análise"; Prio="Média"; Obs="Boa margem de lucro (18% a 36%), porém requer atenção pela nota 2,6 na Amazon." }
    )

    for ($idx = 0; $idx -lt $items.Count; $idx++) {
        $r = $idx + 4
        $it = $items[$idx]

        $ws.Cells.Item($r, 1) = $it.Num
        $ws.Cells.Item($r, 2) = $it.Date
        $ws.Cells.Item($r, 3) = $it.Prod
        $ws.Cells.Item($r, 4) = $it.Cat
        $ws.Cells.Item($r, 5) = $it.AmzL
        $ws.Cells.Item($r, 6) = $it.AmzP
        $ws.Cells.Item($r, 7) = $it.Sales
        $ws.Cells.Item($r, 8) = $it.Rev
        $ws.Cells.Item($r, 9) = $it.Rating
        $ws.Cells.Item($r, 10) = $it.Comp
        $ws.Cells.Item($r, 11) = $it.ShpS
        $ws.Cells.Item($r, 12) = $it.ShpL
        $ws.Cells.Item($r, 13) = $it.ShpC
        $ws.Cells.Item($r, 14) = $it.ShpF
        $ws.Cells.Item($r, 15) = "=M$r+N$r"
        $ws.Cells.Item($r, 16) = $it.MinQ
        $ws.Cells.Item($r, 17) = $it.Del
        $ws.Cells.Item($r, 18) = $it.Pkg
        $ws.Cells.Item($r, 19) = $it.Addr
        $ws.Cells.Item($r, 20) = "=F$r"
        $ws.Cells.Item($r, 21) = $it.FrtC
        $ws.Cells.Item($r, 22) = $it.TaxP
        $ws.Cells.Item($r, 23) = "=T$r*V$r"
        $ws.Cells.Item($r, 24) = $it.OthC
        $ws.Cells.Item($r, 25) = "=O$r+W$r+X$r"
        $ws.Cells.Item($r, 26) = "=(T$r+U$r)-Y$r"
        $ws.Cells.Item($r, 27) = "=Z$r/T$r"
        $ws.Cells.Item($r, 28) = "=Z$r/O$r"
        $ws.Cells.Item($r, 29) = $it.St
        $ws.Cells.Item($r, 30) = $it.Prio
        $ws.Cells.Item($r, 31) = $it.Obs

        $ws.Cells.Item($r, 6).NumberFormat = "R$ #,##0.00"
        $ws.Cells.Item($r, 13).NumberFormat = "R$ #,##0.00"
        $ws.Cells.Item($r, 14).NumberFormat = "R$ #,##0.00"
        $ws.Cells.Item($r, 15).NumberFormat = "R$ #,##0.00"
        $ws.Cells.Item($r, 20).NumberFormat = "R$ #,##0.00"
        $ws.Cells.Item($r, 21).NumberFormat = "R$ #,##0.00"
        $ws.Cells.Item($r, 22).NumberFormat = "0.0%"
        $ws.Cells.Item($r, 23).NumberFormat = "R$ #,##0.00"
        $ws.Cells.Item($r, 24).NumberFormat = "R$ #,##0.00"
        $ws.Cells.Item($r, 25).NumberFormat = "R$ #,##0.00"
        $ws.Cells.Item($r, 26).NumberFormat = "R$ #,##0.00"
        $ws.Cells.Item($r, 27).NumberFormat = "0.00%"
        $ws.Cells.Item($r, 28).NumberFormat = "0.00%"
    }

    $ws.UsedRange.Columns.AutoFit()
    $wb.SaveAs($xlsxFile, 51)
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Host "XLSX gerado com sucesso: $xlsxFile"
} catch {
    Write-Host "Aviso: Excel COM não está instalado/disponível neste sistema, o arquivo CSV foi gerado e pode ser aberto no Excel/Google Sheets."
}
