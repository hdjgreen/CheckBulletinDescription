###################################################################################################
#                                                                                                 #
# The script is used to check bulletin number and description in catalog whether they are correct #
# by comparing with website                                                                       #
#                                                                                                 #
# Parameter: $CatalogPath                                                                         #
# Usage: -CatalogPath "..\Catalog REL15-12 Version2.xls"   #
#        -Username "XXX"                                                             #
#        -Password "*********"                                                                    #
#        -OnlyBulletin $true                                                                      #
#                                                                                                 #
# Time: 2015/12/18                                                                                #
# Modified Time: 2015/12/29                                                                       #
#                2016/4/13 Add this step created certificate using user name and password
#                2016/4/25 Add a method, which can automatically change bulletin number and description in catalog
#                2016/5/13 Add a feature that only check and modify bulletin number in catalog if parameter OnlyBulletin sets as $true
#                                                                                                 #
###################################################################################################



param($CatalogPath, $Username = $null, $Password = $null, $OnlyBulletin = $False)


# Read bulletin number and description column from catalog
# C column: KB Number
# D column: Security Number, namely, bulletin number
# E column: Description

Function Get-CatalogData($catalog)
{
    
    # Determine whether catalog path is correct, exit if not.
    if(!(Test-Path $catalog))
    {
        $Prompt = "!!!Fail: Catalog path:`"$catalog`" is invalid, please check!"
        Output-Info $Prompt $False        

        exit -1
    }

    $xl=New-Object -ComObject "Excel.Application"

    $wb=$xl.Workbooks.Open($catalog)
    $ws=$wb.ActiveSheet

    $Row=2

    While (1)
    {

        $data=$ws.Range("C$Row").Text
    
        if($data)
        {
            #Write-Host "C$Row is $data"

            $KB_Bulletins.$data   = $ws.Range("D$Row").Text
            $KB_Description.$data = $ws.Range("E$Row").Text

            $Row++
        }
        else
        {
            break
        }
    }

    # Determine how many KBs
    $KBNumber = $Row - 2

    $Prompt = "There are $KBNumber KBs in Catalog"

    Output-Info $Prompt $true

    $xl.displayAlerts=$False
    $wb.Close()
    $xl.Application.Quit()

}


# Check bulletin number of KBs in WinCXE and Support Microsoft website, then compare them with bulletin number in catalog
# WinCXE website can't find "Update" and "Silverlight" type, the two kinds of types can be found by Support Microsoft
Function Check-BulletinNumber($Cred)
{

    foreach($kb in $KB_Bulletins.Keys)
    {
        
        # The link is used to search KB
        $WinCXEKBPage = "http://wsered-iis/segdr/reports/search.aspx?q={0}" -f $kb

        if($Cred -ne $null)
        {
            $BucketInfo = Invoke-WebRequest $WinCXEKBPage -Credential $Cred
        }
        else
        {
            $BucketInfo = Invoke-WebRequest $WinCXEKBPage -UseDefaultCredentials
        }

        # Find out Bucket ID
        $BucketInfo.Links | Where-Object{ $_.href -match "\/segdr\/BucketDetails\.aspx\?ID=(\d+)"} | Out-Null

        if($Matches[1] -eq $null)  # May be Silverlight
        {
            # If code enter into here, then the KB may be a silverlight,
            # Because silverlight can't be found in WinCXE website
            
            $Prompt = "@@@Warning: KB{0} can't get any related info from WinCXE website. Trying from Microsoft Support website" -f $kb
            Output-Info $Prompt $False 

            
            # Go to Support Microsoft website to search for the KB
            $ExitCode = Handle-IndividualKB $kb

            $Pattern =  "(\bMS\d+-\d+)"
            if($ExitCode -eq -1)  # The KB doesn't exist
            {
                $Prompt = "!!!Fail: KB$KB doesn't exist in Microsoft Support website either"
                Output-Info $Prompt $False 

                $Script:ErrorNumber++

                continue
            }
            elseif($ExitCode -match $Pattern)   # To see whether $ExitCode returned is bulletin number
            {
                $BulletinNumber_ = $ExitCode
            }
            else  # Handle the KB as update type
            {
                Compare-Update $KB $ExitCode
                    
                continue              
            }
    
        }
        else # if found bucket ID
        {
            # The link which can get bulletin number of the KB from WinCXE website
            $RelasePage = "http://wsered-iis/segdr/ContentProposal.aspx?ID={0}" -f $Matches[1] # $Matches[1] is bucket ID

            if($Cred -ne $null)
            {
                $BulletinInfo = Invoke-WebRequest $RelasePage -Credential $Cred
            }
            else
            {
                $BulletinInfo = Invoke-WebRequest $RelasePage -UseDefaultCredentials
            }

            $BulletinNumberItem = $BulletinInfo.InputFields | Where-Object {$_.name -match ".*MSRCBulletin.*"} | Select-Object -Property value

            $BulletinNumber_ = $BulletinNumberItem.value

            if($BulletinNumber_ -eq $null)
            {
                # if the variable $BulletinNumber is null, then the KB may be update type, but not security update
            
                $Prompt = "@@@Warning: Can't find bulletin number of KB{0} from WinCXE website. Trying from Microsoft Support website" -f $kb
            
                Output-Info $Prompt $False 

                $ExitCode = Handle-IndividualKB $kb

                $Pattern =  "(\bMS\d+-\d+)"
                if($ExitCode -eq -1)  # The KB doesn't exist
                {
                    $Prompt = "!!!Fail: KB$KB doesn't exist in Microsoft Support website either"
                    Output-Info $Prompt $False 

                    $Script:ErrorNumber++

                    continue
                }
                elseif($ExitCode -match $Pattern)   # To see whether $ExitCode returned is bulletin number
                {  
                    $BulletinNumber_ = $ExitCode
                }
                else # Handle the KB as update type
                {
                    Compare-Update $KB $ExitCode
                    
                    continue
                }
            }
        }
        
        $Prompt = "In catalog Bulletin number of KB{0} is {1}" -f $kb,$KB_Bulletins[$kb]
        Output-Info $Prompt $true

        $BulletinNumber = ([string]$BulletinNumber_).Trim()
        $Prompt = "In website Bulletin number of KB{0} is {1}" -f  $kb,$BulletinNumber
        Output-Info $Prompt $true

        # Compare bulletin number
        if($KB_Bulletins[$kb] -eq $BulletinNumber)
        {
            $Prompt = "Pass: Bulletin number of {0} is matched" -f $kb
            Output-Info $Prompt $true
        }
        else
        {
            $Prompt = "!!!Fail: Bulletin number of {0} doesn't match, trying to change the bulletin number of the KB from catalog" -f $kb
            Output-Info $Prompt $False 

            $Script:ErrorNumber++
			
			Change-Catalog $CatalogPath "1" $KB $BulletinNumber
        }

        if($OnlyBulletin -eq $False)
        {
            # Return result which check description
            if((Check-Description $kb $BulletinNumber) -eq $true)
            {
                $Prompt = "Pass: Description of {0} is matched" -f $kb
                Output-Info $Prompt $true
            }
            else
            {
                $Prompt = "!!!Fail: Description of {0} doesn't match, trying to change the description of the KB from catalog" -f $kb
                Output-Info $Prompt $False 

                $Script:ErrorNumber++
				
            }
        }

    }
}


# Using bulletin number WinCXE websit got to spell out a link like https://technet.microsoft.com/library/security/MS15-120
# then to find out description by the link, then compare them with description in catalog 
Function Check-Description($KB,$BulletinNumber)
{
    
    # The link which can get description of the KB 
    $BulletinPage = "https://technet.microsoft.com/library/security/{0}" -f $BulletinNumber

    $BulletinInfo = Invoke-WebRequest $BulletinPage -UseDefaultCredentials    

    # Convert format first, if doesn't convert the content, it will be all of the content match and will be also get all content
    # After converting, content will be divided into some segments, such matching is the content we want
    $BulletinInfo.Content | ConvertFrom-Csv | Where-Object{$_ -match ".*<h2 class=`"subheading`">(.*)\s\(\d+\)<\/h2>"}

    #PS C:\WINDOWS\system32> $Matches

    #Name                           Value                                                                                                         
    #----                           -----                                                                                                         
    #1                              Security Update for Microsoft Graphics Component to Address Remote Code Execution                             
    #0                              @{<!DOCTYPE html>=<h2 class="subheading">Security Update for Microsoft Graphics Component to Address 
    
    $Prompt = "In catalog description of KB{0} is {1}" -f $KB,$KB_Description[$KB]
    Output-Info $Prompt $true
     
    $Prompt = "In website description of KB{0} is {1}" -f  $KB,$Matches[1].Trim() # Trim() remove trailing spaces for the string 
    Output-Info $Prompt $true
    
    
    # Compare description
    if($KB_Description[$KB] -ne $Matches[1].Trim())
    {
        Change-Catalog $CatalogPath "2" $KB $Matches[1].Trim()
		return $false   
    }

    return $true

}

# Handle KB which can't be searched from WinCXE, these KBs will be searched from Support Microsoft website
# if returns -1, indicate that the KB doesn't exist in website
# if returns bulletin number, indicate that the KB should be a security update
# if returns a string which doesn't match -1 or bulletin number, indicate that the KB should be a update  
Function Handle-IndividualKB($KB)
{
    $url = "https://support.microsoft.com/en-us/kb/{0}" -f $KB

    $ie = New-Object -COM InternetExplorer.Application
    $ie.Navigate($url)

    while ($ie.ReadyState -ne 4) 
    {
        sleep 50 
    }

    $html = $ie.Document

    
    # The KB doesn't exist if $html.title is "Microsoft Support"
    if($html.title -eq "Microsoft Support") 
    {
        return -1
    }
    else
    {
        $Pattern = "(?<BulletinNumber>\bMS\d+-\d+):.*"

        if($html.title -match $Pattern)
        {
            return $Matches.BulletinNumber
        }
        else
        {
            # KB type is "Update"
            return $html.title
        }

    }

}

# Compare description in catalog and Support Microsoft website that KB type is update
Function Check-IndividualDescription($KB,$DesForWebsite)
{
    if($KB_Description[$KB] -eq $DesForWebsite)
    {
        return $true
    }
    else
    {
        return $false
    }
}

# Check  bulletin number in catalog that type is update
# the bulletin number column of the KB should be ""
Function Check-UpdateTypeBulletinNumber($KB)
{
    if($KB_Bulletins[$KB] -eq "")
    {
        return $true    
    }
    else
    {
        return $false
    }    
}

# Compare KB that type is update
Function Compare-Update($KB, $DesForWebsite_)
{

    $Prompt = "KB{0} may be a update but not security update, so its bulletin should be NULL" -f $KB
    Output-Info $Prompt $true

    $Prompt = "In catalog Bulletin number of KB{0} is {1}" -f $KB,$KB_Bulletins[$KB]
    Output-Info $Prompt $true


    $Prompt = "In website Bulletin number of KB{0} is " -f  $KB
    Output-Info $Prompt $true

    
    # To see whether bulletin number column in catalog is "", it is ture if yes, otherwise it is false
    if((Check-UpdateTypeBulletinNumber $KB) -eq $true)
    {
        $Prompt = "Pass: Bulletin number of {0} is matched" -f $KB
        Output-Info $Prompt $true
    }
    else
    {
        $Prompt = "!!!Fail: Bulletin number of {0} doesn't match, trying to change the bulletin number of the KB from catalog" -f $KB
        Output-Info $Prompt $False 
        $Script:ErrorNumber++
		
		Change-Catalog $CatalogPath "1" $KB $null
    }    
    
    if($OnlyBulletin -eq $False)
    {
        $Prompt = "In catalog description of KB{0} is {1}" -f $KB,$KB_Description[$KB].Trim() # Trim() remove trailing spaces for the string 
        Output-Info $Prompt $true

        $DesForWebsite = ([string]$DesForWebsite_).Trim()
        $Prompt = "In website description of KB{0} is {1}" -f  $KB, $DesForWebsite
        Output-Info $Prompt $true
    
        # Compare description of website and catalog
        if((Check-IndividualDescription $KB $DesForWebsite) -eq $true)
        {
            $Prompt = "Pass: Description of {0} is matched" -f $KB
            Output-Info $Prompt $true
        }
        else
        {
            $Prompt = "!!!Fail: Description of {0} doesn't match, trying to change the description of the KB from catalog" -f $KB
            Output-Info $Prompt $False 
            $Script:ErrorNumber++
		
		    Change-Catalog $CatalogPath "2" $KB $DesForWebsite
        }
    }

}

Function Output-Info($Prompt,$Result)
{
    if($Result -eq $true)
    {
        Write-Host $Prompt -ForegroundColor Green

        $Script:LogArray.add($Prompt) | Out-Null
    }
    else
    {
        Write-Host $Prompt -BackgroundColor Red

        $Script:LogArray.add($Prompt) | Out-Null
    }
}


function Out-Log($LogFile, $Header, $Footer, $LogArray)
{
	$Log = New-Object System.Collections.ArrayList

	#add all of $Header array as long as it isn't null or empty
	if($Header -ne $null -and $Header -ne "")
	{
		$Log.AddRange(@($Header))
	}

    #add all of $LogArray array to $log array as long as it isn't null or empty
	if($LogArray -ne $null -and $LogArray -ne "")
	{
		$Log.AddRange(@($LogArray))
	}

	#add all of $footer array to $log array as long as it isn't null or empty
	if($Footer -ne $null -and $Footer -ne "")
	{
		$Log.AddRange(@($Footer))
	}
	
	#-force command for out-file doesn't work properly and returns an error instead of making folders 
	#for the given path when they aren't present.
	#using new-item with -force works properly , so using it to make a dummy file as a workaground
	new-item -Force -path $LogFile -value "dummy file" -type File | Out-Null
	
	#output file in $LogFile path containing $Log
	Out-File -filePath $LogFile -InputObject $Log
}

#Change bulletin number and description in the catalog if they are different with website
#parameter: 
# $Type: change bulletin if is 1, change description if is 2
# $KB:   The KB will be changed
# $Content: Content of the KB will be changed
Function Change-Catalog($CatalogPath, $Type, $KB, $Content)
{
    $xl=New-Object -ComObject "Excel.Application"

    $wb=$xl.Workbooks.Open($CatalogPath)
    $ws=$wb.ActiveSheet

    $Row=2

    While (1)
    {

        $data=$ws.Range("C$Row").Text   
        
        if($data)
        {
			if($data -eq $KB)
			{
				if(($Type -eq "1") -and ($Content -ne $null)) #Change Bulletin number
				{
					#$ws.Range("D$Row").value2 = $Content
                    $ws.HyperLinks.Add(
                    $ws.Range("D$Row"),
                    "https://technet.microsoft.com/library/security/$Content",
                    "","",
                    $Content
                    ) | Out-Null
                    break
				}
                elseif(($Type -eq "1") -and ($Content -eq $null)) #Due to the KB is update type, so the cell should be empty
                {
                    $ws.Range("D$Row").value2 = ""
                    break
                }
				elseif($Type -eq "2") #Change description
				{
					$ws.Range("E$Row").value2 = $Content
                    break
				}
                else
                {
                    Write-Host "the parameter Type of the method Change-Catalog is incorrect, please check"
                    break
                }
			}

            $Row++
        }
        else
        {
            break
        }
    }
	
	$xl.displayAlerts=$False
    $wb.SaveAs($CatalogPath)
    $wb.Close()
    $xl.Application.Quit()
}

###################################################################################################################
##############################   Main   ###########################################################################
###################################################################################################################

# KB and its bulletin in the catalog
$KB_Bulletins = @{}

# KB and its description in the catalog
$KB_Description = @{}

# Path of log file
$LogFile = "CheckBulletinDescription.txt"

# Number of no matching
$Script:ErrorNumber = 0

#log arrays
$Script:LogArray = New-Object System.Collections.ArrayList
$Script:Header   = New-Object System.Collections.ArrayList
$Script:Footer   = New-Object System.Collections.ArrayList

#log header
$msg = ("Start Date: " + (get-date).year + "/" + (get-date).month + "/" + (get-date).day + " " + (get-date).timeofday)
$header.add($msg) | Out-Null

#Creat certificate if neither user name nor password are null
if(($Username -ne $null) -and ($Password -ne $null))
{
    $Pwd = ConvertTo-SecureString $Password -AsPlainText -Force
    $Cred = New-Object System.Management.Automation.PSCredential($Username, $Pwd)
}
else
{
    $Cred = $null
}

# Collect bulletin number and description in the catalog
Get-CatalogData $CatalogPath

Check-BulletinNumber $Cred

$Prompt = "Error number: $Script:ErrorNumber"

#log footer
$msg = ("End Date: " + (get-date).year + "/" + (get-date).month + "/" + (get-date).day + " " + (get-date).timeofday)
$Footer.add($msg) | Out-Null


if($Script:ErrorNumber -ne 0)
{
    Output-Info $Prompt $False 

    Out-Log $LogFile $Header $Footer $LogArray

    exit -1
}
else
{
    Output-Info $Prompt $true

    Out-Log $LogFile $Header $Footer $LogArray

    exit 0
}
