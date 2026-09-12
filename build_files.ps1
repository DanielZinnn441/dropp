# PowerShell script to generate Excel (.xlsx) and CSV (.csv) files for Dropship Product Analysis

$targetDir = "c:\Users\d4nie\OneDrive\Imagens\PRODTOS ENCONTARDO"
$csvFile = Join-Path $targetDir "planilha_pesquisa_produtos_dropship.csv"
$xlsxFile = Join-Path $targetDir "planilha_pesquisa_produtos_dropship.xlsx"

# 1. Create CSV with UTF-8 BOM and Semicolon separator
$csvLines = @(
    "Nº;Data da pesquisa;Produto;Categoria;Link do anúncio (Amazon);Preço de venda na Amazon (R$);Vendas estimadas/mês;Nº de avaliações;Nota média (0-5);Nº de concorrentes;Fornecedor (loja na Shopee);Link do fornecedor;Preço de custo unitário (R$);Frete até o cliente (R$) - estimado;Custo entregue no cliente (R$);Qtd. mínima por pedido;Prazo de entrega até o cliente (dias);Envia com nota/embalagem própria?;Aceita endereço de entrega diferente?;Preço de venda no anúncio (R$);Frete cobrado do cliente (R$);Taxa Amazon (%);Taxa Amazon (R$);Outros custos (R$);Custo total por venda (R$);Lucro líquido por unidade (R$);Margem de lucro (%);ROI (%);Status;Prioridade;Observações",
    "1;10/09/2026;Kit 3 Potes Organizadores para Geladeira Premium;Cozinha / Organização;https://www.amazon.com.br/Organizadores-Geladeira-Escorredores-Herm%C3%A9ticas-Conserva/dp/B0H9P9QR57;43,90;N/I;1;5,0;1;Shopee (Cruzeiro, SP);https://shopee.com.br/Kit-3-potes-Organizadores-Geladeira-Organizador-Alimentos-Vegetais-Cozinha-Moderna-i.629805123.58205197736;32,98;12,04;45,02;1;5 a 10 dias;Não;Sim;43,90;0,00;15%;6,59;0,00;51,61;-7,71;-17,56%;-17,13%;Em análise;Alta;Nota 5,0 na Amazon. Com cupom frete grátis Shopee o lucro fica em +R$ 4,33 (9,86%).",
    "2;10/09/2026;Pré Treino Insane Clown 350g Demons Lab - Blue Crystal;Suplementos / Saúde;https://www.amazon.com.br/Treino-Insane-Clown-350g-Demons/dp/B0GYQM5WFT;116,90;N/I;N/I;N/I;N/I;Shopee (Cruzeiro, SP);https://shopee.com.br/INSANE-CLOWN-350G-BLUE-CRYSTAL-i.384424516.22391686410;99,75;10,83;110,58;1;5 a 10 dias;Não;Sim;116,90;0,00;15%;17,54;0,00;128,12;-11,22;-9,60%;-10,15%;Em análise;Média;Recomenda-se ajustar preço na Amazon para R$ 139,90 para garantir margem positiva.",
    "3;10/09/2026;Kit 6 Pares Meias Cano Invisível Lupo Sport Sapatilha;Moda / Roupas;https://www.amazon.com.br/Lupo-esportivas-algod%C3%A3o-respir%C3%A1vel-pares/dp/B00PROD03;74,90;N/I;1449;4,8;3;Shopee Lojas Oficiais;https://shopee.com.br/Kit-De-6-Pares-Meias-Cano-Invis%C3%ADvel-Lupo-Sport-Algod%C3%A3o-Soquete-Sapatilha-Unissex;49,23;9,62;58,85;1;5 a 10 dias;Não;Sim;74,90;0,00;15%;11,24;0,00;70,09;4,82;6,43%;8,18%;Em análise;Alta;Alta demanda na Amazon (4,8★ com 1,4k+ avaliações). Lucro sobe para R$ 14,44 (19,27%) com frete grátis.",
    "4;10/09/2026;Carregador Turbo Duo 168W USB-C e USB-A com Cabo USB-C;Eletrônicos / Acessórios;https://www.amazon.com.br/dp/B0H5XWN3WD;79,90;N/I;N/I;N/I;N/I;Shopee (Cruzeiro, SP);https://shopee.com.br/product/1572491828/22094298275;25,88;9,62;35,50;1;5 a 10 dias;Não;Sim;79,90;0,00;15%;11,99;0,00;47,49;32,42;40,57%;91,31%;Em análise;Média;Produto Destaque! Lucro de R$ 32,42 a R$ 42,04 por unidade (Margem > 40% e ROI até 162%).",
    "5;10/09/2026;Jogo de Talheres Faqueiro Aço Inox 24 Peças;Cozinha / Utensílios;https://www.amazon.com.br/Jogo-Talheres-Faqueiro-Inox-Pecas/dp/B00PROD05;52,69;N/I;4;2,6;1;Shopee (Cruzeiro, SP);https://shopee.com.br/Jogo-de-Talheres-24-Pe%C3%A7as-Faqueiro-A%C3%A7o-Inox-Faca-Garfo-Colher-de-Mesa-e-Colheres-de-Sobremesa;25,40;9,62;35,02;1;5 a 10 dias;Não;Sim;52,69;0,00;15%;7,90;0,00;42,92;9,77;18,54%;27,89%;Em análise;Média;Boa margem de lucro (18% a 36%), porém requer atenção pela nota 2,6 na Amazon."
)

$utf8WithBom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllLines($csvFile, $csvLines, $utf8WithBom)
Write-Host "CSV criado em: $csvFile"

# 2. Try Excel COM to generate native .xlsx with styles and formulas
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $ws = $workbook.Worksheets.Item(1)
    $ws.Name = "Viabilidade Products Dropship"

    # Header Row 1 - Title
    $ws.Cells.Item(1, 1) = "PLANILHA DE PESQUISA E VIABILIDADE DE PRODUTOS — Dropship: Shopee (fornecedor) envia direto ao cliente da Amazon"
    $ws.Range("A1:AE1").Merge()
    $ws.Range("A1:AE1").Font.Bold = $true
    $ws.Range("A1:AE1").Font.Size = 13
    $ws.Range("A1:AE1").Font.ColorIndex = 2
    $ws.Range("A1:AE1").Interior.Color = 0x5C301B # Dark Navy
    $ws.Range("A1:AE1").HorizontalAlignment = -4108 # Center

    # Header Row 2 - Group Names
    $groups = @(
        @{ Range = "A2:D2"; Text = "PRODUTO"; Color = 0x7030A0 },
        @{ Range = "E2:J2"; Text = "PESQUISA NA AMAZON (mercado)"; Color = 0x0066FF },
        @{ Range = "K2:O2"; Text = "FORNECEDOR NA SHOPEE — ENVIO DIRETO"; Color = 0x339900 },
        @{ Range = "P2:S2"; Text = "ENVIO DIRETO AO CLIENTE"; Color = 0x339900 },
        @{ Range = "T2:Y2"; Text = "PRECIFICAÇÃO E CUSTOS"; Color = 0x00A5FF },
        @{ Range = "Z2:AB2"; Text = "RESULTADO"; Color = 0x800080 },
        @{ Range = "AC2:AE2"; Text = "GESTÃO"; Color = 0x595959 }
    )

    foreach ($g in $groups) {
        $rng = $ws.Range($g.Range)
        $rng.Merge()
        $rng.Value2 = $g.Text
        $rng.Font.Bold = $true
        $rng.Font.ColorIndex = 2
        $rng.Interior.Color = $g.Color
        $rng.HorizontalAlignment = -4108
    }

    # Header Row 3 - Columns
    $cols = @(
        "Nº", "Data da pesquisa", "Produto", "Categoria",
        "Link do anúncio (Amazon)", "Preço de venda na Amazon (R$)", "Vendas estimadas/mês", "Nº de avaliações", "Nota média (0-5)", "Nº de concorrentes",
        "Fornecedor (loja na Shopee)", "Link do fornecedor", "Preço de custo unitário (R$)", "Frete até o cliente (R$) - estimado", "Custo entregue no cliente (R$)",
        "Qtd. mínima por pedido", "Prazo de entrega até o cliente (dias)", "Envia com nota/embalagem própria?", "Aceita endereço de entrega diferente?",
        "Preço de venda no anúncio (R$)", "Frete cobrado do cliente (R$)", "Taxa Amazon (%)", "Taxa Amazon (R$)", "Outros custos (R$)", "Custo total por venda (R$)",
        "Lucro líquido por unidade (R$)", "Margem de lucro (%)", "ROI (%)",
        "Status", "Prioridade", "Observações"
    )

    for ($i = 0; $i -lt $cols.Count; $i++) {
        $cell = $ws.Cells.Item(3, $i + 1)
        $cell.Value2 = $cols[$i]
        $cell.Font.Bold = $true
        $cell.Font.Size = 9
        $cell.Interior.Color = 0xF2F2F2
        $cell.HorizontalAlignment = -4108
    }

    # Data Rows
    $data = @(
        @{ Num=1; Date="10/09/2026"; Prod="Kit 3 Potes Organizadores para Geladeira Premium"; Cat="Cozinha / Organização"; AmzLink="https://www.amazon.com.br/Organizadores-Geladeira-Escorredores-Herm%C3%A9ticas-Conserva/dp/B0H9P9QR57"; AmzPrice=43.90; Sales="N/I"; Reviews=1; Rating=5.0; Comp=1; ShpStore="Shopee (Cruzeiro, SP)"; ShpLink="https://shopee.com.br/Kit-3-potes-Organizadores-Geladeira-Organizador-Alimentos-Vegetais-Cozinha-Moderna-i.629805123.58205197736"; ShpCost=32.98; ShpFreight=12.04; MinQty=1; Delivery="5 a 10 dias"; OwnPkg="Não"; DiffAddr="Sim"; FreightCust=0.00; AmzTaxPct=0.15; OtherCosts=0.00; Status="Em análise"; Priority="Alta"; Obs="Nota 5,0 na Amazon. Com cupom frete grátis Shopee o lucro fica em +R$ 4,33 (9,86%)." },
        @{ Num=2; Date="10/09/2026"; Prod="Pré Treino Insane Clown 350g Demons Lab - Blue Crystal"; Cat="Suplementos / Saúde"; AmzLink="https://www.amazon.com.br/Treino-Insane-Clown-350g-Demons/dp/B0GYQM5WFT"; AmzPrice=116.90; Sales="N/I"; Reviews="N/I"; Rating="N/I"; Comp="N/I"; ShpStore="Shopee (Cruzeiro, SP)"; ShpLink="https://shopee.com.br/INSANE-CLOWN-350G-BLUE-CRYSTAL-i.384424516.22391686410"; ShpCost=99.75; ShpFreight=10.83; MinQty=1; Delivery="5 a 10 dias"; OwnPkg="Não"; DiffAddr="Sim"; FreightCust=0.00; AmzTaxPct=0.15; OtherCosts=0.00; Status="Em análise"; Priority="Média"; Obs="Recomenda-se ajustar preço na Amazon para R$ 139,90 para garantir margem positiva." },
        @{ Num=3; Date="10/09/2026"; Prod="Kit 6 Pares Meias Cano Invisível Lupo Sport Sapatilha"; Cat="Moda / Roupas"; AmzLink="https://www.amazon.com.br/Lupo-esportivas-algod%C3%A3o-respir%C3%A1vel-pares/dp/B00PROD03"; AmzPrice=74.90; Sales="N/I"; Reviews=1449; Rating=4.8; Comp=3; ShpStore="Shopee Lojas Oficiais"; ShpLink="https://shopee.com.br/Kit-De-6-Pares-Meias-Cano-Invis%C3%ADvel-Lupo-Sport-Algod%C3%A3o-Soquete-Sapatilha-Unissex"; ShpCost=49.23; ShpFreight=9.62; MinQty=1; Delivery="5 a 10 dias"; OwnPkg="Não"; DiffAddr="Sim"; FreightCust=0.00; AmzTaxPct=0.15; OtherCosts=0.00; Status="Em análise"; Priority="Alta"; Obs="Alta demanda na Amazon (4,8★ com 1,4k+ avaliações). Lucro sobe para R$ 14,44 (19,27%) com frete grátis." },
        @{ Num=4; Date="10/09/2026"; Prod="Carregador Turbo Duo 168W USB-C e USB-A com Cabo USB-C"; Cat="Eletrônicos / Acessórios"; AmzLink="https://www.amazon.com.br/dp/B0H5XWN3WD"; AmzPrice=79.90; Sales="N/I"; Reviews="N/I"; Rating="N/I"; Comp="N/I"; ShpStore="Shopee (Cruzeiro, SP)"; ShpLink="https://shopee.com.br/product/1572491828/22094298275"; ShpCost=25.88; ShpFreight=9.62; MinQty=1; Delivery="5 a 10 dias"; OwnPkg="Não"; DiffAddr="Sim"; FreightCust=0.00; AmzTaxPct=0.15; OtherCosts=0.00; Status="Em análise"; Priority="Média"; Obs="Produto Destaque! Lucro de R$ 32,42 a R$ 42,04 por unidade (Margem > 40% e ROI até 162%)." },
        @{ Num=5; Date="10/09/2026"; Prod="Jogo de Talheres Faqueiro Aço Inox 24 Peças"; Cat="Cozinha / Utensílios"; AmzLink="https://www.amazon.com.br/Jogo-Talheres-Faqueiro-Inox-Pecas/dp/B00PROD05"; AmzPrice=52.69; Sales="N/I"; Reviews=4; Rating=2.6; Comp=1; ShpStore="Shopee (Cruzeiro, SP)"; ShpLink="https://shopee.com.br/Jogo-de-Talheres-24-Pe%C3%A7as-Faqueiro-A%C3%A7o-Inox-Faca-Garfo-Colher-de-Mesa-e-Colheres-de-Sobremesa"; ShpCost=25.40; ShpFreight=9.62; MinQty=1; Delivery="5 a 10 dias"; OwnPkg="Não"; DiffAddr="Sim"; FreightCust=0.00; AmzTaxPct=0.15; OtherCosts=0.00; Status="Em análise"; Priority="Média"; Obs="Boa margem de lucro (18% a 36%), porém requer atenção pela nota 2,6 na Amazon." }
    )

    for ($r = 0; $r -lt $data.Count; $r++) {
        $rowIdx = $r + 4
        $item = $data[$r]

        $ws.Cells.Item($rowIdx, 1) = $item.Num
        $ws.Cells.Item($rowIdx, 2) = $item.Date
        $ws.Cells.Item($rowIdx, 3) = $item.Prod
        $ws.Cells.Item($rowIdx, 4) = $item.Cat
        $ws.Cells.Item($rowIdx, 5) = $item.AmzLink
        $ws.Cells.Item($rowIdx, 6) = $item.AmzPrice
        $ws.Cells.Item($rowIdx, 7) = $item.Sales
        $ws.Cells.Item($rowIdx, 8) = $item.Reviews
        $ws.Cells.Item($rowIdx, 9) = $item.Rating
        $ws.Cells.Item($rowIdx, 10) = $item.Comp
        $ws.Cells.Item($rowIdx, 11) = $item.ShpStore
        $ws.Cells.Item($rowIdx, 12) = $item.ShpLink
        $ws.Cells.Item($rowIdx, 13) = $item.ShpCost
        $ws.Cells.Item($rowIdx, 14) = $item.ShpFreight
        $ws.Cells.Item($rowIdx, 15) = "=M$rowIdx+N$rowIdx"
        $ws.Cells.Item($rowIdx, 16) = $item.MinQty
        $ws.Cells.Item($rowIdx, 17) = $item.Delivery
        $ws.Cells.Item($rowIdx, 18) = $item.OwnPkg
        $ws.Cells.Item($rowIdx, 19) = $item.DiffAddr
        $ws.Cells.Item($rowIdx, 20) = "=F$rowIdx"
        $ws.Cells.Item($rowIdx, 21) = $item.FreightCust
        $ws.Cells.Item($rowIdx, 22) = $item.AmzTaxPct
        $ws.Cells.Item($rowIdx, 23) = "=T$rowIdx*V$rowIdx"
        $ws.Cells.Item($rowIdx, 24) = $item.OtherCosts
        $ws.Cells.Item($rowIdx, 25) = "=O$rowIdx+W$rowIdx+X$rowIdx"
        $ws.Cells.Item($rowIdx, 26) = "=(T$rowIdx+U$rowIdx)-Y$rowIdx"
        $ws.Cells.Item($rowIdx, 27) = "=Z$rowIdx/T$rowIdx"
        $ws.Cells.Item($rowIdx, 28) = "=Z$rowIdx/O$rowIdx"
        $ws.Cells.Item($rowIdx, 29) = $item.Status
        $ws.Cells.Item($rowIdx, 30) = $item.Priority
        $ws.Cells.Item($rowIdx, 31) = $item.Obs

        # Formatting
        $ws.Cells.Item($rowIdx, 6).NumberFormat = "R$ #,##0.00"
        $ws.Cells.Item($rowIdx, 13).NumberFormat = "R$ #,##0.00"
        $ws.Cells.Item($rowIdx, 14).NumberFormat = "R$ #,##0.00"
        $ws.Cells.Item($rowIdx, 15).NumberFormat = "R$ #,##0.00"
        $ws.Cells.Item($rowIdx, 20).NumberFormat = "R$ #,##0.00"
        $ws.Cells.Item($rowIdx, 21).NumberFormat = "R$ #,##0.00"
        $ws.Cells.Item($rowIdx, 22).NumberFormat = "0.0%"
        $ws.Cells.Item($rowIdx, 23).NumberFormat = "R$ #,##0.00"
        $ws.Cells.Item($rowIdx, 24).NumberFormat = "R$ #,##0.00"
        $ws.Cells.Item($rowIdx, 25).NumberFormat = "R$ #,##0.00"
        $ws.Cells.Item($rowIdx, 26).NumberFormat = "R$ #,##0.00"
        $ws.Cells.Item($rowIdx, 27).NumberFormat = "0.00%"
        $ws.Cells.Item($rowIdx, 28).NumberFormat = "0.00%"
    }

    $ws.UsedRange.Columns.AutoFit()
    $workbook.SaveAs($xlsxFile, 51) # 51 = xlOpenXMLWorkbook (.xlsx)
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Host "XLSX criado em: $xlsxFile"
} catch {
    Write-Host "Aviso ao criar XLSX via Excel COM: $_"
}
