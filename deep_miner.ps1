# PowerShell Deep Scraper & Attribute Extractor for Product Sourcing

param(
    [string]$Url
)

function Scrape-ProductPage {
    param([string]$targetUrl)
    
    Write-Host "Iniciando raspagem profunda para URL: $targetUrl"
    
    $req = [System.Net.HttpWebRequest]::Create($targetUrl)
    $req.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    $req.Timeout = 10000
    
    try {
        $res = $req.GetResponse()
        $stream = $res.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($stream)
        $html = $reader.ReadToEnd()
        $res.Close()

        # Extract OpenGraph Title
        $title = ""
        if ($html -match '<meta\s+property="og:title"\s+content="([^"]+)"') {
            $title = $Matches[1]
        } elseif ($html -match '<title>([^<]+)</title>') {
            $title = $Matches[1]
        }

        # Extract OpenGraph Image
        $image = ""
        if ($html -match '<meta\s+property="og:image"\s+content="([^"]+)"') {
            $image = $Matches[1]
        }

        # Extract Price (Regex pattern)
        $price = 0.00
        if ($html -match 'R\$\s*([\d\.,]+)') {
            $rawPrice = $Matches[1].Replace('.', '').Replace(',', '.')
            [double]::TryParse($rawPrice, [ref]$price) | Out-Null
        }

        # Detect Color / Specs from Title & HTML
        $colors = @('Preto', 'Branco', 'Azul', 'Vermelho', 'Rosa', 'Cinza', 'Transparente', 'Inox', 'Verde', 'Amarelo')
        $detectedColor = "Geral"
        foreach ($c in $colors) {
            if ($title -like "*$c*" -or $html -like "*$c*") {
                $detectedColor = $c
                break
            }
        }

        return @{
            Success = $true
            Title = [System.Net.WebUtility]::HtmlDecode($title)
            ImageUrl = $image
            Price = $price
            Color = $detectedColor
            SourceUrl = $targetUrl
        }
    } catch {
        return @{
            Success = $false
            Error = $_.Exception.Message
            SourceUrl = $targetUrl
        }
    }
}

if ($Url) {
    $result = Scrape-ProductPage -targetUrl $Url
    $result | ConvertTo-Json
}
