function Convert-Base64 {
    [CmdletBinding()]
    param (
        # base 64 string
        [Parameter(Mandatory,ValueFromPipeline)]
        [string]
        $Base64String
    )

    process {
        $bytes = [System.Convert]::FromBase64String($Base64String)
        $content = [System.Text.Encoding]::UTF8.GetString($bytes)
        return $content
    }
}
