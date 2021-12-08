#----------------------------------------------[Authentication]----------------------------------------------

function Get-GraphAPIAccessToken {
    $resource = "?resource=https://graph.microsoft.com/"
    $url = $env:IDENTITY_ENDPOINT + $resource
    $Headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $Headers.Add("X-IDENTITY-HEADER", $env:IDENTITY_HEADER)
    $Headers.Add("Metadata", "True")
    $accessToken = Invoke-RestMethod -Uri $url -Method 'GET' -Headers $Headers
    return $accessToken.access_token
}

function Get-KeyVaultCredential {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [String] $Name,

        [Parameter(Mandatory = $true, Position = 1)]
        [String] $KeyVault
    )

    $userName = Get-AutomationVariable -Name $Name
    $secure = Get-AzKeyVaultSecret -VaultName $KeyVault -Name $Name
    [PSCredential]$cred = New-Object System.Management.Automation.PSCredential ($userName, $secure.SecretValue)
    return $cred
}

#------------------------------------------------[EMail]-----------------------------------------------------

function Get-ContentTypeFromFileName {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string] $FileName
    )
    $mimeMappings = @{
        '.323'         = 'text/h323'
        '.aaf'         = 'application/octet-stream'
        '.aca'         = 'application/octet-stream'
        '.accdb'       = 'application/msaccess'
        '.accde'       = 'application/msaccess'
        '.accdt'       = 'application/msaccess'
        '.acx'         = 'application/internet-property-stream'
        '.afm'         = 'application/octet-stream'
        '.ai'          = 'application/postscript'
        '.aif'         = 'audio/x-aiff'
        '.aifc'        = 'audio/aiff'
        '.aiff'        = 'audio/aiff'
        '.application' = 'application/x-ms-application'
        '.art'         = 'image/x-jg'
        '.asd'         = 'application/octet-stream'
        '.asf'         = 'video/x-ms-asf'
        '.asi'         = 'application/octet-stream'
        '.asm'         = 'text/plain'
        '.asr'         = 'video/x-ms-asf'
        '.asx'         = 'video/x-ms-asf'
        '.atom'        = 'application/atom+xml'
        '.au'          = 'audio/basic'
        '.avi'         = 'video/x-msvideo'
        '.axs'         = 'application/olescript'
        '.bas'         = 'text/plain'
        '.bcpio'       = 'application/x-bcpio'
        '.bin'         = 'application/octet-stream'
        '.bmp'         = 'image/bmp'
        '.c'           = 'text/plain'
        '.cab'         = 'application/octet-stream'
        '.calx'        = 'application/vnd.ms-office.calx'
        '.cat'         = 'application/vnd.ms-pki.seccat'
        '.cdf'         = 'application/x-cdf'
        '.chm'         = 'application/octet-stream'
        '.class'       = 'application/x-java-applet'
        '.clp'         = 'application/x-msclip'
        '.cmx'         = 'image/x-cmx'
        '.cnf'         = 'text/plain'
        '.cod'         = 'image/cis-cod'
        '.cpio'        = 'application/x-cpio'
        '.cpp'         = 'text/plain'
        '.crd'         = 'application/x-mscardfile'
        '.crl'         = 'application/pkix-crl'
        '.crt'         = 'application/x-x509-ca-cert'
        '.csh'         = 'application/x-csh'
        '.css'         = 'text/css'
        '.csv'         = 'application/octet-stream'
        '.cur'         = 'application/octet-stream'
        '.dcr'         = 'application/x-director'
        '.deploy'      = 'application/octet-stream'
        '.der'         = 'application/x-x509-ca-cert'
        '.dib'         = 'image/bmp'
        '.dir'         = 'application/x-director'
        '.disco'       = 'text/xml'
        '.dll'         = 'application/x-msdownload'
        '.dll.config'  = 'text/xml'
        '.dlm'         = 'text/dlm'
        '.doc'         = 'application/msword'
        '.docm'        = 'application/vnd.ms-word.document.macroEnabled.12'
        '.docx'        = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        '.dot'         = 'application/msword'
        '.dotm'        = 'application/vnd.ms-word.template.macroEnabled.12'
        '.dotx'        = 'application/vnd.openxmlformats-officedocument.wordprocessingml.template'
        '.dsp'         = 'application/octet-stream'
        '.dtd'         = 'text/xml'
        '.dvi'         = 'application/x-dvi'
        '.dwf'         = 'drawing/x-dwf'
        '.dwp'         = 'application/octet-stream'
        '.dxr'         = 'application/x-director'
        '.eml'         = 'message/rfc822'
        '.emz'         = 'application/octet-stream'
        '.eot'         = 'application/octet-stream'
        '.eps'         = 'application/postscript'
        '.etx'         = 'text/x-setext'
        '.evy'         = 'application/envoy'
        '.exe'         = 'application/octet-stream'
        '.exe.config'  = 'text/xml'
        '.fdf'         = 'application/vnd.fdf'
        '.fif'         = 'application/fractals'
        '.fla'         = 'application/octet-stream'
        '.flr'         = 'x-world/x-vrml'
        '.flv'         = 'video/x-flv'
        '.gif'         = 'image/gif'
        '.gtar'        = 'application/x-gtar'
        '.gz'          = 'application/x-gzip'
        '.h'           = 'text/plain'
        '.hdf'         = 'application/x-hdf'
        '.hdml'        = 'text/x-hdml'
        '.hhc'         = 'application/x-oleobject'
        '.hhk'         = 'application/octet-stream'
        '.hhp'         = 'application/octet-stream'
        '.hlp'         = 'application/winhlp'
        '.hqx'         = 'application/mac-binhex40'
        '.hta'         = 'application/hta'
        '.htc'         = 'text/x-component'
        '.htm'         = 'text/html'
        '.html'        = 'text/html'
        '.htt'         = 'text/webviewhtml'
        '.hxt'         = 'text/html'
        '.ico'         = 'image/x-icon'
        '.ics'         = 'application/octet-stream'
        '.ief'         = 'image/ief'
        '.iii'         = 'application/x-iphone'
        '.inf'         = 'application/octet-stream'
        '.ins'         = 'application/x-internet-signup'
        '.isp'         = 'application/x-internet-signup'
        '.IVF'         = 'video/x-ivf'
        '.jar'         = 'application/java-archive'
        '.java'        = 'application/octet-stream'
        '.jck'         = 'application/liquidmotion'
        '.jcz'         = 'application/liquidmotion'
        '.jfif'        = 'image/pjpeg'
        '.jpb'         = 'application/octet-stream'
        '.jpe'         = 'image/jpeg'
        '.jpeg'        = 'image/jpeg'
        '.jpg'         = 'image/jpeg'
        '.js'          = 'application/x-javascript'
        '.jsx'         = 'text/jscript'
        '.latex'       = 'application/x-latex'
        '.lit'         = 'application/x-ms-reader'
        '.lpk'         = 'application/octet-stream'
        '.lsf'         = 'video/x-la-asf'
        '.lsx'         = 'video/x-la-asf'
        '.lzh'         = 'application/octet-stream'
        '.m13'         = 'application/x-msmediaview'
        '.m14'         = 'application/x-msmediaview'
        '.m1v'         = 'video/mpeg'
        '.m3u'         = 'audio/x-mpegurl'
        '.man'         = 'application/x-troff-man'
        '.manifest'    = 'application/x-ms-manifest'
        '.map'         = 'text/plain'
        '.mdb'         = 'application/x-msaccess'
        '.mdp'         = 'application/octet-stream'
        '.me'          = 'application/x-troff-me'
        '.mht'         = 'message/rfc822'
        '.mhtml'       = 'message/rfc822'
        '.mid'         = 'audio/mid'
        '.midi'        = 'audio/mid'
        '.mix'         = 'application/octet-stream'
        '.mmf'         = 'application/x-smaf'
        '.mno'         = 'text/xml'
        '.mny'         = 'application/x-msmoney'
        '.mov'         = 'video/quicktime'
        '.movie'       = 'video/x-sgi-movie'
        '.mp2'         = 'video/mpeg'
        '.mp3'         = 'audio/mpeg'
        '.mpa'         = 'video/mpeg'
        '.mpe'         = 'video/mpeg'
        '.mpeg'        = 'video/mpeg'
        '.mpg'         = 'video/mpeg'
        '.mpp'         = 'application/vnd.ms-project'
        '.mpv2'        = 'video/mpeg'
        '.ms'          = 'application/x-troff-ms'
        '.msi'         = 'application/octet-stream'
        '.mso'         = 'application/octet-stream'
        '.mvb'         = 'application/x-msmediaview'
        '.mvc'         = 'application/x-miva-compiled'
        '.nc'          = 'application/x-netcdf'
        '.nsc'         = 'video/x-ms-asf'
        '.nws'         = 'message/rfc822'
        '.ocx'         = 'application/octet-stream'
        '.oda'         = 'application/oda'
        '.odc'         = 'text/x-ms-odc'
        '.ods'         = 'application/oleobject'
        '.one'         = 'application/onenote'
        '.onea'        = 'application/onenote'
        '.onetoc'      = 'application/onenote'
        '.onetoc2'     = 'application/onenote'
        '.onetmp'      = 'application/onenote'
        '.onepkg'      = 'application/onenote'
        '.osdx'        = 'application/opensearchdescription+xml'
        '.p10'         = 'application/pkcs10'
        '.p12'         = 'application/x-pkcs12'
        '.p7b'         = 'application/x-pkcs7-certificates'
        '.p7c'         = 'application/pkcs7-mime'
        '.p7m'         = 'application/pkcs7-mime'
        '.p7r'         = 'application/x-pkcs7-certreqresp'
        '.p7s'         = 'application/pkcs7-signature'
        '.pbm'         = 'image/x-portable-bitmap'
        '.pcx'         = 'application/octet-stream'
        '.pcz'         = 'application/octet-stream'
        '.pdf'         = 'application/pdf'
        '.pfb'         = 'application/octet-stream'
        '.pfm'         = 'application/octet-stream'
        '.pfx'         = 'application/x-pkcs12'
        '.pgm'         = 'image/x-portable-graymap'
        '.pko'         = 'application/vnd.ms-pki.pko'
        '.pma'         = 'application/x-perfmon'
        '.pmc'         = 'application/x-perfmon'
        '.pml'         = 'application/x-perfmon'
        '.pmr'         = 'application/x-perfmon'
        '.pmw'         = 'application/x-perfmon'
        '.png'         = 'image/png'
        '.pnm'         = 'image/x-portable-anymap'
        '.pnz'         = 'image/png'
        '.pot'         = 'application/vnd.ms-powerpoint'
        '.potm'        = 'application/vnd.ms-powerpoint.template.macroEnabled.12'
        '.potx'        = 'application/vnd.openxmlformats-officedocument.presentationml.template'
        '.ppam'        = 'application/vnd.ms-powerpoint.addin.macroEnabled.12'
        '.ppm'         = 'image/x-portable-pixmap'
        '.pps'         = 'application/vnd.ms-powerpoint'
        '.ppsm'        = 'application/vnd.ms-powerpoint.slideshow.macroEnabled.12'
        '.ppsx'        = 'application/vnd.openxmlformats-officedocument.presentationml.slideshow'
        '.ppt'         = 'application/vnd.ms-powerpoint'
        '.pptm'        = 'application/vnd.ms-powerpoint.presentation.macroEnabled.12'
        '.pptx'        = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        '.prf'         = 'application/pics-rules'
        '.prm'         = 'application/octet-stream'
        '.prx'         = 'application/octet-stream'
        '.ps'          = 'application/postscript'
        '.psd'         = 'application/octet-stream'
        '.psm'         = 'application/octet-stream'
        '.psp'         = 'application/octet-stream'
        '.pub'         = 'application/x-mspublisher'
        '.qt'          = 'video/quicktime'
        '.qtl'         = 'application/x-quicktimeplayer'
        '.qxd'         = 'application/octet-stream'
        '.ra'          = 'audio/x-pn-realaudio'
        '.ram'         = 'audio/x-pn-realaudio'
        '.rar'         = 'application/octet-stream'
        '.ras'         = 'image/x-cmu-raster'
        '.rf'          = 'image/vnd.rn-realflash'
        '.rgb'         = 'image/x-rgb'
        '.rm'          = 'application/vnd.rn-realmedia'
        '.rmi'         = 'audio/mid'
        '.roff'        = 'application/x-troff'
        '.rpm'         = 'audio/x-pn-realaudio-plugin'
        '.rtf'         = 'application/rtf'
        '.rtx'         = 'text/richtext'
        '.scd'         = 'application/x-msschedule'
        '.sct'         = 'text/scriptlet'
        '.sea'         = 'application/octet-stream'
        '.setpay'      = 'application/set-payment-initiation'
        '.setreg'      = 'application/set-registration-initiation'
        '.sgml'        = 'text/sgml'
        '.sh'          = 'application/x-sh'
        '.shar'        = 'application/x-shar'
        '.sit'         = 'application/x-stuffit'
        '.sldm'        = 'application/vnd.ms-powerpoint.slide.macroEnabled.12'
        '.sldx'        = 'application/vnd.openxmlformats-officedocument.presentationml.slide'
        '.smd'         = 'audio/x-smd'
        '.smi'         = 'application/octet-stream'
        '.smx'         = 'audio/x-smd'
        '.smz'         = 'audio/x-smd'
        '.snd'         = 'audio/basic'
        '.snp'         = 'application/octet-stream'
        '.spc'         = 'application/x-pkcs7-certificates'
        '.spl'         = 'application/futuresplash'
        '.src'         = 'application/x-wais-source'
        '.ssm'         = 'application/streamingmedia'
        '.sst'         = 'application/vnd.ms-pki.certstore'
        '.stl'         = 'application/vnd.ms-pki.stl'
        '.sv4cpio'     = 'application/x-sv4cpio'
        '.sv4crc'      = 'application/x-sv4crc'
        '.swf'         = 'application/x-shockwave-flash'
        '.t'           = 'application/x-troff'
        '.tar'         = 'application/x-tar'
        '.tcl'         = 'application/x-tcl'
        '.tex'         = 'application/x-tex'
        '.texi'        = 'application/x-texinfo'
        '.texinfo'     = 'application/x-texinfo'
        '.tgz'         = 'application/x-compressed'
        '.thmx'        = 'application/vnd.ms-officetheme'
        '.thn'         = 'application/octet-stream'
        '.tif'         = 'image/tiff'
        '.tiff'        = 'image/tiff'
        '.toc'         = 'application/octet-stream'
        '.tr'          = 'application/x-troff'
        '.trm'         = 'application/x-msterminal'
        '.tsv'         = 'text/tab-separated-values'
        '.ttf'         = 'application/octet-stream'
        '.txt'         = 'text/plain'
        '.u32'         = 'application/octet-stream'
        '.uls'         = 'text/iuls'
        '.ustar'       = 'application/x-ustar'
        '.vbs'         = 'text/vbscript'
        '.vcf'         = 'text/x-vcard'
        '.vcs'         = 'text/plain'
        '.vdx'         = 'application/vnd.ms-visio.viewer'
        '.vml'         = 'text/xml'
        '.vsd'         = 'application/vnd.visio'
        '.vss'         = 'application/vnd.visio'
        '.vst'         = 'application/vnd.visio'
        '.vsto'        = 'application/x-ms-vsto'
        '.vsw'         = 'application/vnd.visio'
        '.vsx'         = 'application/vnd.visio'
        '.vtx'         = 'application/vnd.visio'
        '.wav'         = 'audio/wav'
        '.wax'         = 'audio/x-ms-wax'
        '.wbmp'        = 'image/vnd.wap.wbmp'
        '.wcm'         = 'application/vnd.ms-works'
        '.wdb'         = 'application/vnd.ms-works'
        '.wks'         = 'application/vnd.ms-works'
        '.wm'          = 'video/x-ms-wm'
        '.wma'         = 'audio/x-ms-wma'
        '.wmd'         = 'application/x-ms-wmd'
        '.wmf'         = 'application/x-msmetafile'
        '.wml'         = 'text/vnd.wap.wml'
        '.wmlc'        = 'application/vnd.wap.wmlc'
        '.wmls'        = 'text/vnd.wap.wmlscript'
        '.wmlsc'       = 'application/vnd.wap.wmlscriptc'
        '.wmp'         = 'video/x-ms-wmp'
        '.wmv'         = 'video/x-ms-wmv'
        '.wmx'         = 'video/x-ms-wmx'
        '.wmz'         = 'application/x-ms-wmz'
        '.wps'         = 'application/vnd.ms-works'
        '.wri'         = 'application/x-mswrite'
        '.wrl'         = 'x-world/x-vrml'
        '.wrz'         = 'x-world/x-vrml'
        '.wsdl'        = 'text/xml'
        '.wvx'         = 'video/x-ms-wvx'
        '.x'           = 'application/directx'
        '.xaf'         = 'x-world/x-vrml'
        '.xaml'        = 'application/xaml+xml'
        '.xap'         = 'application/x-silverlight-app'
        '.xbap'        = 'application/x-ms-xbap'
        '.xbm'         = 'image/x-xbitmap'
        '.xdr'         = 'text/plain'
        '.xla'         = 'application/vnd.ms-excel'
        '.xlam'        = 'application/vnd.ms-excel.addin.macroEnabled.12'
        '.xlc'         = 'application/vnd.ms-excel'
        '.xlm'         = 'application/vnd.ms-excel'
        '.xls'         = 'application/vnd.ms-excel'
        '.xlsb'        = 'application/vnd.ms-excel.sheet.binary.macroEnabled.12'
        '.xlsm'        = 'application/vnd.ms-excel.sheet.macroEnabled.12'
        '.xlsx'        = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        '.xlt'         = 'application/vnd.ms-excel'
        '.xltm'        = 'application/vnd.ms-excel.template.macroEnabled.12'
        '.xltx'        = 'application/vnd.openxmlformats-officedocument.spreadsheetml.template'
        '.xlw'         = 'application/vnd.ms-excel'
        '.xml'         = 'text/xml'
        '.xof'         = 'x-world/x-vrml'
        '.xpm'         = 'image/x-xpixmap'
        '.xps'         = 'application/vnd.ms-xpsdocument'
        '.xsd'         = 'text/xml'
        '.xsf'         = 'text/xml'
        '.xsl'         = 'text/xml'
        '.xslt'        = 'text/xml'
        '.xsn'         = 'application/octet-stream'
        '.xtp'         = 'application/octet-stream'
        '.xwd'         = 'image/x-xwindowdump'
        '.z'           = 'application/x-compress'
        '.zip'         = 'application/x-zip-compressed'
    }

    $extension = [System.IO.Path]::GetExtension($FileName)
    $contentType = $mimeMappings[$extension]
    if ([string]::IsNullOrEmpty($contentType)) {
        return New-Object System.Net.Mime.ContentType
    }
    else {
        return New-Object System.Net.Mime.ContentType($contentType)
    }
}

function Send-O365MailMessage {
    [CmdletBinding()]
    param (
        [Parameter(Position = 1, Mandatory = $true)]
        [PSCredential] $Credential,

        [Parameter(Position = 2, Mandatory = $true)]
        [String] $ClientId,

        [Parameter(Position = 2, Mandatory = $true)]
        [String] $TenantId,

        [Parameter(Position = 4, Mandatory = $true)]
        [Object] $To,

        [Parameter(Position = 5, Mandatory = $true)]
        [String] $Subject,

        [Parameter(Position = 6, Mandatory = $true)]
        [String] $Body,

        [Parameter(Mandatory = $false)]
        [String] $RedirectURI,

        [Parameter(Mandatory = $false)]
        [String[]] $Attachments,

        [Parameter(Mandatory = $false)]
        [Collections.HashTable] $InlineAttachments,

        [Parameter(Mandatory = $false)]
        [String[]] $Cc,

        [Parameter(Mandatory = $false)]
        [String[]] $Bcc,

        [Parameter(Mandatory = $false)]
        [Switch] $BodyAsHTML,

        [Parameter(Mandatory = $false)]
        [String] $Encoding = "UTF8",

        [Parameter(Mandatory = $false)]
        [String] $Priority = "Normal"
    )
    Process {
        if ([String]::IsNullOrEmpty($RedirectURI)) {
            $RedirectURI = "msal" + $ClientId + "://auth"
        }

        # send with Graph API
        # we're creating an object to convert it to json later.
        $MessageObj = [PSCustomObject]@{}
        $MessageObj | Add-Member -MemberType NoteProperty -Name 'subject' -Value $subject

        # Build the body section
        if ($BodyAsHTML) {
            $contentType = "html"
        }
        else {
            $contentType = "text"
        }
        $BodyObj = [PSCustomObject]@{
            contentType = $contentType
            content     = $body
        }

        $MessageObj | Add-Member -MemberType NoteProperty -Name 'body' -Value $BodyObj

        # build recipient section
        # to section
        $emailAddressArr = @()
        foreach ($item in $to) {
            if ($item.Contains("<")) {
                $ADDRESSParts = $item.Split("<")
                $NamePart = $ADDRESSParts[0].Trim(" ")
                $mailPart = $ADDRESSParts[1].Trim(" ").TrimEnd(">")
                $emailObj = [PSCustomObject]@{
                    address = $mailPart
                    name    = $NamePart
                }
            }
            else {
                $mailPart = $item.Trim("<").Trim(">").Trim(" ")
                $emailObj = [PSCustomObject]@{
                    address = $mailPart
                    name    = ""
                }
            }
            $emailAddressArr += [PSCustomObject]@{emailAddress = $emailObj }
        } # foreach: to section
        $toRecipientsObj = @(
            $emailAddressArr
        )
        $MessageObj | Add-Member -MemberType NoteProperty -Name 'toRecipients' -Value $toRecipientsObj
        # cc section
        if (![String]::IsNullOrEmpty($Cc)) {
            $emailAddressArr = @()
            foreach ($item in $Cc) {
                if ($item.Contains("<")) {
                    $ADDRESSParts = $item.Split("<")
                    $NamePart = $ADDRESSParts[0].Trim(" ")
                    $mailPart = $ADDRESSParts[1].Trim(" ").TrimEnd(">")
                    $emailObj = [PSCustomObject]@{
                        address = $mailPart
                        name    = $NamePart
                    }
                }
                else {
                    $mailPart = $item.Trim("<").Trim(">").Trim(" ")
                    $emailObj = [PSCustomObject]@{
                        address = $mailPart
                        name    = ""
                    }
                }
                $emailAddressArr += [PSCustomObject]@{emailAddress = $emailObj }
            } # foreach
            $ccRecipientsObj = @(
                $emailAddressArr
            )
            $MessageObj | Add-Member -MemberType NoteProperty -Name 'ccRecipients' -Value $ccRecipientsObj
        } # if: cc section
        # bcc section
        if (![String]::IsNullOrEmpty($Bcc)) {
            $emailAddressArr = @()
            foreach ($item in $Bcc) {
                if ($item.Contains("<")) {
                    $ADDRESSParts = $item.Split("<")
                    $NamePart = $ADDRESSParts[0].Trim(" ")
                    $mailPart = $ADDRESSParts[1].Trim(" ").TrimEnd(">")
                    $emailObj = [PSCustomObject]@{
                        address = $mailPart
                        name    = $NamePart
                    }
                }
                else {
                    $mailPart = $item.Trim("<").Trim(">").Trim(" ")
                    $emailObj = [PSCustomObject]@{
                        address = $mailPart
                        name    = ""
                    }
                }
                $emailAddressArr += [PSCustomObject]@{emailAddress = $emailObj }
            } # foreach
            $bccRecipientsObj = @(
                $emailAddressArr
            )
            $MessageObj | Add-Member -MemberType NoteProperty -Name 'bccRecipients' -Value $bccRecipientsObj
        } # if: bcc section
        # create email attachments section
        $hasAttachments = $false
        if ($null -ne $Attachments) {
            $AttachmentsArr = @()
            $hasAttachments = $true
            foreach ($item in $Attachments) {
                $EncodedFile = [convert]::ToBase64String( [system.io.file]::readallbytes($item))
                $FileType = (Get-ContentTypeFromFileName -FileName $item).MediaType
                $attachedFileObj = [PSCustomObject]@{
                    '@odata.type' = "#microsoft.graph.fileAttachment"
                    name          = $item.Split("\")[$item.Split("\").count - 1]
                    contentBytes  = $EncodedFile
                    contentType   = "application/octet-stream" #$FileType
                }
                $AttachmentsArr += $attachedFileObj
            } # foreach
        } # if: create email attachments section
        # create section for inline attachemnts aka embedded files
        if ($null -ne $InlineAttachments) {
            $InlineAttachmentsArr = @()
            $hasAttachments = $true
            foreach ($item in $InlineAttachments.GetEnumerator()) {
                $FileName = $item.Value.ToString()
                $EncodedFile = [convert]::ToBase64String( [system.io.file]::readallbytes($FileName))
                $FileType = (Get-ContentTypeFromFileName -FileName $FileName).MediaType
                $InlineattachedFileObj = [PSCustomObject]@{
                    '@odata.type' = "#microsoft.graph.fileAttachment"
                    name          = $FileName.Split("\")[$FileName.Split("\").count - 1]
                    contentBytes  = $EncodedFile
                    contentID     = $item.Key
                    isInline      = $true
                    contentType   = $FileType
                }
                $InlineAttachmentsArr += $InlineattachedFileObj
            } #foreach
        } #if: create section for inline attachemnts aka embedded files
        # both tyes are attachements from a technical view so we combine both
        if ($hasAttachments) {
            $allAttachments = $AttachmentsArr + $InlineAttachmentsArr
            $MessageObj | Add-Member -MemberType NoteProperty -Name 'attachments' -Value $allAttachments
        }

        # message priority part
        $ChangePrority = $false
        switch ($Priority ) {
            "High" { $ChangePrority = $true }
            "Low" { $ChangePrority = $true }
        }
        if ($ChangePrority) {
            $MessageObj | Add-Member -MemberType NoteProperty -Name 'importance' -Value $Priority
        }
        $fullObj = [PSCustomObject]@{message = $MessageObj }

        # convert all the stuff from above to json
        $jsonObj = $fullObj | ConvertTo-Json -Depth 50

        $ApiUrl = "https://graph.microsoft.com/v1.0/me/sendMail"

        try {
            # get the Token
            $token = (Get-MsalToken -UserCredential $Credential -ClientId $ClientId -TenantId $TenantId -Scopes "Mail.Send").AccessToken

            # ...aaaaaand send
            Invoke-RestMethod -Headers @{Authorization = "Bearer $token" } -Uri $ApiUrl -Method Post -Body $jsonObj -ContentType "application/json"
        }
        catch {
            Write-Error $_.Exception.Message
        }

    }
}

#------------------------------------------------[Logging]---------------------------------------------------

function New-LogFile {
    param (
        [Parameter(Mandatory = $true)]
        [String] $Path,

        [Parameter(Mandatory = $true)]
        [String] $FileName,

        [Parameter(Mandatory = $false)]
        [String] $TimeStamp,

        [Parameter(Mandatory = $false)]
        [String] $Key
    )

    if (!(Test-Path -Path $Path)) { New-Item -Path $Path -ItemType Directory -Force }
    if ($TimeStamp -eq "") {
        $PathLogFile = $Path + $(Get-Date -Format "yyyyMMdd-HHmmss") + "_$FileName.log"
    }
    else {
        if ($Key -ne "") {
            $PathLogFile = $Path + $TimeStamp + "_" + $FileName + "_" + $Key + ".log"
        }
        else {
            $PathLogFile = $Path + $TimeStamp + "_$FileName.log"
        }
    }
    return $PathLogFile
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [String]$LogMessage,

        [Parameter(Mandatory = $true)]
        [String]$LogFile,

        [Parameter(Mandatory = $false)]
        [ValidateSet("Error", "Warn", "Info", "Ok", "Failed", "Success")][String]$LogLevel
    )

    $arr = @()

    # Set Date/Time
    $dateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $arr += $dateTime
    # Set log level
    switch ($LogLevel) {
        'Error' {
            $arr += 'ERROR    '
        }
        'Warn' {
            $arr += 'WARN     '
        }
        'Info' {
            $arr += 'INFO     '
        }
        'Ok' {
            $arr += 'OK       '
        }
        'Failed' {
            $arr += 'FAILED   '
        }
        'Success' {
            $arr += 'SUCCESS  '
        }
        Default {
            $arr += '         '
        }
    } # switch: Set log level

    # Set message
    $arr += $LogMessage

    # Build line from array
    $line = [System.String]::Join(" ", $arr)

    # Write to log
    if ($LogFile -ne "") { $line | Out-File -FilePath $LogFile -Append }

    # Write to host
    Write-Host $line
}

function Add-LogFileToSpo {
    param(
        [Parameter(Mandatory = $true)]
        [String] $DriveId,

        [Parameter(Mandatory = $true)]
        [String] $Folder,

        [Parameter(Mandatory = $true)]
        [String] $LogFile,

        [Parameter(Mandatory = $true)]
        [HashTable] $GraphApiHeader
    )
    $split = $LogFile.Split("\")
    $fileName = $split[$split.Length - 1]

    $uri = "$global:graphApiBaseUrl/drives/$DriveId/root:/$Folder/$($fileName):/content"

    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $GraphApiHeader -Method Put -InFile $logFile -ContentType "application/txt"
        return $response.webUrl
    }
    catch {
        return $null
    }
}