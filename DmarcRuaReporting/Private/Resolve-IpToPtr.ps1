Function Resolve-IpToPtr {
    Param (
        [Parameter(Mandatory)]
        [System.String]
        $IPAddress
    )

    Write-Debug -Message "Processing IP '$IPAddress'"
    try {
        $ptrs = @(Resolve-DnsName -Name $IPAddress -Type PTR -ErrorAction Stop -Verbose:$false | Select-Object -ExpandProperty NameHost)
        Write-Debug -Message "  Resolved to $($ptrs.Count) PTR record(s)."

        #If the IP address resolves to a single PTR record then use this.
        if ($ptrs.Count -eq 1) {
            $ptr = $ptrs[0]

        #If the IP address resolves to multiple PTR records...
        } else {

            #Filter out generic PTR records.
            $nonArpa = @($ptrs | Where-Object {$_ -notlike '*.in-addr.arpa*'})

            #If we're now left with a single, non-generic PTR record then use this.
            if ($nonArpa.Count -eq 1) {
                $ptr = $nonArpa[0]

            #If this filters everything out then fall back to using the generic PTRs, converted into a combined string.
            } elseif ($nonArpa.Count -eq 0) {
                $ptr = ([String]$ptrs) -Replace ' ', ','

            #If there's more than 1 non-generic PTR, then convert these into a combined string.
            } elseif ($nonArpa.Count -gt 1) {
                $ptr = ([String]$nonArpa) -Replace ' ', ','
            }
        }
        $resolvedToPtr = $true

    } catch {
        Write-Debug -Message "  Failed to resolve to a PTR record."
        $ptr = '<DNS PTR lookup failure>'
        $resolvedToPtr = $false
    }

    [PSCustomObject]@{
        IPAddress     = $IPAddress
        ResolvedToPtr = $resolvedToPtr = $true
        ResolvedPtr   = $ptr
    } | Write-Output
}
