function Invoke-RemoteZipExtraction {
    [CmdletBinding()]
    param (
        [String] $remoteUrl,
        [String] $fileName,
        [String] $extractPath
    )
    function getBytes($url, $rangeStart, $rangeEnd) {
        $webRequest = [System.Net.HttpWebRequest]::Create($url)
        $webRequest.Method = "GET"
        $webRequest.AddRange($rangeStart, $rangeEnd)
        $response = $webRequest.GetResponse()
        $responseStream = $response.GetResponseStream()
        # Read the response stream into a MemoryStream
        $memoryStream = New-Object System.IO.MemoryStream
        $responseStream.CopyTo($memoryStream)
    
        # Convert the MemoryStream to a byte array
        $fileBytes = $memoryStream.ToArray()
    
        # Close the response, response stream, and memory stream
        $response.Close()
        $responseStream.Close()
        $memoryStream.Close()
        $fileBytes
    }
    
    function Decompress-Data($compressedData) {
        # Create a memory stream from the compressed data
        $compressedStream = new-object System.IO.MemoryStream(, $compressedData)
    
        # Create a DeflateStream to decompress the data
        $deflateStream = New-Object System.IO.Compression.DeflateStream($compressedStream, [System.IO.Compression.CompressionMode]::Decompress)
    
        # Create a memory stream to store the decompressed data
        $decompressedStream = New-Object System.IO.MemoryStream
    
        # Copy the decompressed data from the DeflateStream to the memory stream
        $deflateStream.CopyTo($decompressedStream)
    
        # Convert the decompressed memory stream to a byte array
        $decompressedData = $decompressedStream.ToArray()
    
        # Close the streams and release resources
        $deflateStream.Close()
        $compressedStream.Close()
        $decompressedStream.Close()
    
        $decompressedData
    }
    
    function Fetch-EndOfCentralDirectory($url, $zipByteLength) {
        $EOCD_MAX_BYTES = 128
        $eocdInitialOffset = [Math]::Max(0, $zipByteLength - $EOCD_MAX_BYTES)
    
        $fileBytes = getBytes -url $url -rangeStart $eocdInitialOffset -rangeEnd $zipByteLength
    
        $eocdSignature = [System.BitConverter]::GetBytes(0x06054B50) # End of Central Directory signature
        $eocdOffset = $null
        for ($i = $fileBytes.Length - 22; $i -ge 0; $i--) {
            if ([System.BitConverter]::ToUInt32($fileBytes, $i) -eq [System.BitConverter]::ToUInt32($eocdSignature, 0)) {
                $eocdOffset = $i
                break
            }
        }
    
        # Verify that the EOCD signature was found
        if ($eocdOffset -eq $null) {
            Write-Output "End of Central Directory signature not found."
        }
    
        Write-Output ($zipByteLength - $EOCD_MAX_BYTES + $eocdOffset)
    }
    
    function Extract-FileFromRemoteZip($url, $eocdOffset, $fileNameToExtract, $extractedFilePath) {
        # Read the number of entries in the Central Directory
        $entriesCountOffset = $eocdOffset + 10
        $entriesCountBytes = getBytes -url $url -rangeStart $entriesCountOffset -rangeEnd ($entriesCountOffset + 1)
        $entriesCount = [System.BitConverter]::ToUInt16($entriesCountBytes, 0)
    
        # Read the Central Directory and extract the file names
        $centralDirectoryOffset = [System.BitConverter]::ToUInt32((getBytes -url $url -rangeStart ($eocdOffset + 16) -rangeEnd ($eocdOffset + 16 + 3)), 0)
        $centralDirectorySize = [System.BitConverter]::ToUInt32((getBytes -url $url -rangeStart ($eocdOffset + 12) -rangeEnd ($eocdOffset + 12 + 3)), 0)
        
        $centralDirectoryBytes = getBytes -url $url -rangeStart $centralDirectoryOffset -rangeEnd ($centralDirectoryOffset + $centralDirectorySize - 1)
        $position = 0
        for ($i = 0; $i -lt $entriesCount; $i++) {
            $headerSignature = [System.BitConverter]::GetBytes(0x02014B50) # Central File Header signature
    
            $headerOffset = -1
            $signatureLength = $headerSignature.Length
            $centralDirectoryLength = $centralDirectoryBytes.Length
    
            for ($j = $position; $j -lt $centralDirectoryLength - $signatureLength; $j++) {
                $match = $true
                for ($k = 0; $k -lt $signatureLength; $k++) {
                    if ($centralDirectoryBytes[$j + $k] -ne $headerSignature[$k]) {
                        $match = $false
                        break
                    }
                }
    
                if ($match) {
                    $headerOffset = $j
                    break
                }
            }
    
            if ($headerOffset -eq -1) {
                Write-Output "Central File Header signature not found."
                exit
            }
    
            $nameLengthOffset = $headerOffset + 28
            $nameLengthBytes = $centralDirectoryBytes[$nameLengthOffset..($nameLengthOffset + 1)]
            $nameLength = [System.BitConverter]::ToUInt16($nameLengthBytes, 0)
    
            $nameOffset = $headerOffset + 46
            $nameBytes = $centralDirectoryBytes[$nameOffset..($nameOffset + $nameLength - 1)]
            $fileName = [System.Text.Encoding]::ASCII.GetString($nameBytes)
    
            if ($fileName -eq $fileNameToExtract) {
                $compressedSizeOffset = $headerOffset + 20
                $compressedSizeBytes = $centralDirectoryBytes[$compressedSizeOffset..($compressedSizeOffset + 3)]
                $compressedSize = [System.BitConverter]::ToUInt32($compressedSizeBytes, 0)
                
                $localFileHeaderOffset = [System.BitConverter]::ToUInt32($centralDirectoryBytes, $headerOffset + 42)
                
                $compressedDataOffset = 30 # Offset for Local File Header before compressed data
    
                $compressedData = getBytes -url $url -rangeStart ($localFileHeaderOffset + $nameLength + $compressedDataOffset) -rangeEnd ($localFileHeaderOffset + $nameLength + $compressedDataOffset + $compressedSize - 1)
    
                $fileData = Decompress-Data -compressedData $compressedData
    
                [System.IO.File]::WriteAllBytes($extractedFilePath, $fileData)
    
                Write-Host "File '$fileNameToExtract' extracted successfully to '$extractedFilePath'."
                break
            }
            $position = $headerOffset + 1
        }
    }
    
    $response = Invoke-WebRequest -Uri $remoteUrl -Method Head
    $contentLength = [Convert]::ToInt32($response.Headers.'Content-Length')
    $eocdOffset = Fetch-EndOfCentralDirectory -url $remoteUrl -zipByteLength $contentLength
    
    Extract-FileFromRemoteZip -url $remoteUrl -eocdOffset $eocdOffset -fileNameToExtract $fileName -extractedFilePath $extractPath    
}