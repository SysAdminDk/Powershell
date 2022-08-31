<#
    .DESCRIPTION
    Function to create random password that will be accepted by Active Directory password policy

    .PARAMETER length 
    Defines the length of the gennerated password.
    - Lenght have to be between 10 and 30 (Can be changed in the Create-Password function)

    .PARAMETER Numbers
    Defines the amount of numbers used in the gennerated password.
    - Must be between 1 and 4 (Can be changed in the Create-Password function)

    .PARAMETER SpecialChars
    Defines the amount of special ascii chars to be used in gennerated password.
    - Must be between 1 and 4 (Can be changed in the Create-Password function)

    .EXAMPLE
    Create-Password -length 30 -Numbers 4 -SpecialChars 4 -verbose

#>

function Get-RandomCharacters($length, $characters) {
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length }
    $private:ofs=""
    return [String]$characters[$random]
}

function Scramble-String([string]$inputString){     
    $characterArray = $inputString.ToCharArray()   
    $scrambledStringArray = $characterArray | Get-Random -Count $characterArray.Length     
    $outputString = -join $scrambledStringArray
    return $outputString 
}

function Create-Password() {
    Param(
        [Parameter(Mandatory)][ValidateRange(10,30)][Int]$length,
        [Parameter(Mandatory)][ValidateRange(1,4)][Int]$Numbers,
        [Parameter(Mandatory)][ValidateRange(1,4)][Int]$SpecialChars
    )

    Write-Verbose "Selected Password length: $length"
    Write-Verbose "Selected number of numbers : $Numbers"
    Write-Verbose "Selected number of special characters : $SpecialChars"

    $AsciCharLength = $length-$Numbers-$SpecialChars
    Write-Verbose "Ascii Chars in password : $AsciCharLength"

    $LoverCaseLength = [math]::Round($(Get-Random -Maximum $AsciCharLength))
    Write-Verbose "Random lovercase length : $LoverCaseLength"

    $UpperCaseLength = [math]::Round($AsciCharLength-$LoverCaseLength)
    Write-Verbose "Calculated upercase length : $UpperCaseLength"

    $password = Get-RandomCharacters -length $LoverCaseLength -characters 'abcdefghiklmnoprstuvwxyz'
    $password += Get-RandomCharacters -length $UpperCaseLength -characters 'ABCDEFGHKLMNOPRSTUVWXYZ'
    $password += Get-RandomCharacters -length $Numbers -characters '1234567890'
    $password += Get-RandomCharacters -length $SpecialChars -characters '!#$%&/()=?'
    
    $password = Scramble-String $($password.Substring(0,$length))
    Write-Verbose "Random generated password : $password"
    
    return $password
}
