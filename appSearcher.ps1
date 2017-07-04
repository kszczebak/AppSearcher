[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

 

try

{

    $server_list = get-content "server_list.txt" -ErrorAction Stop;

    #$output_file = "RESULT.txt"

}

catch

{

    #$_.Exception.Message;

    [System.Windows.Forms.MessageBox]::Show("File not found or file name is incorrect!`n`nCorrect file name is: server_list.txt",'File Error','OK','Error');

    Exit 1;

}

 

$global:output = "";

$global:found = $false;

 

function SearchApplication($app)

{


 while($app -eq "")

    {

        Write-host "Please choose proper name!";

        $app = Read-Host;     

    }

        Write-host "`nSearching for: " -NoNewline; Write-host $app -ForegroundColor Green; 

        Write-Host "`nPlease wait...`n`n`n";

        foreach($server in $server_list)

        {

 

            Try

            {

                #$server_fqdn = (nslookup $server)[3];

                #$server_fqdn = $server_fqdn.Substring(9);

                #sleep -s 2;

                $server_fqdn = SearchFQDN -server $server;

 

                $apps_from_registry = Invoke-Command -ComputerName $server_fqdn {Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*} -ErrorAction Stop | Select-Object DisplayName, DisplayVersion;

               

                #sleep -s 1; #wait for execution

                

                $apps_from_registry += Invoke-Command -ComputerName $server_fqdn {Get-ItemProperty HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*} -ErrorAction Stop | Select-Object DisplayName, DisplayVersion;

 

                Write-host "Software installed on " -NoNewline ;  Write-Host $server_fqdn  -ForegroundColor Red;

 

                $global:output += "Software installed on " + $server_fqdn + "`n";

 

                SearchInList -list $apps_from_registry -server $server_fqdn;

               

                if($global:found -eq $false)

                {

                    Write-host "Application not found.";

                    $global:output += "Application not found";

                }

                  

                Write-host "`n`n";

                               

            }

            Catch

            {

                Write-host "$_.Exception.Message`n";

                Write-host "Unable to connect to host " $server_fqdn ". Please check if hostname is correct and try agin!`n`n";

                $global:output += "Unable to connect to host " + $server_fqdn + ". Please check if hostname is correct and try agin!`n`n";

            } 

        }

 

        $global:output | Out-File "application_list.csv";

        #[System.Windows.Forms.MessageBox]::Show("Checking complete!",'Task done','OK','Information');

        Write-Output "End of procedure`n`n";

 

}

 

 

function SearchInList($list, $server)

{

 

    foreach($application in $list)

    {

         Try

         {

            

              if($application.DisplayName -match $searched_application )

              {

                    if($application.DisplayName -eq "Configuration Manager Client")

                    {

                        $sccm_components = Invoke-Command -ComputerName $server {Get-WmiObject -Namespace root/ccm -class ccm_installedcomponent} -ErrorAction Stop | Select-Object DisplayName, Version;

 

                        foreach($component in $sccm_components)

                        {

                            if($component.DisplayName -match "Configmgr remote")

                            {

                                Write-host $component.DisplayName "version:" $component.Version"`n";

                                $global:output += $component.DisplayName + "`tversion: " + $component.Version + "`n";

                            }

                        }

 

                    }

 

                    Write-host $application.DisplayName " version: "$application.DisplayVersion"`n";

                    $global:output += $application.DisplayName + "`tversion: " + $application.DisplayVersion + "`n";

                    $global:found = $true;           

              }

          } 

          Catch

          {

               Write-host $_.Exception.Message "`n`nFurther search has been stopped. Please correct input.`n`n";

               exit 1;

          }

 

    }

     $global:output += "`n";

}

 

function SearchFQDN($server)

{

 

    

    if($server -ne "")

    {

         try

        {

            if(Test-Connection -count 1 -computer "$server.swissre.com" -Quiet)

            {

                $server = "$server.swissre.com";

                return $server;

            }

            if( Test-Connection -count 1 -computer "$server.itecorp.itegwpnet.com" -Quiet)

            {

                $server = "$server.itecorp.itegwpnet.com";

                return $server;

            }

            if( Test-Connection -count 1 -computer "$server.testcorp.testgwpnet.com" -Quiet)

            {

                $server = "$server.testcorp.testgwpnet.com";

                return $server;

            }

            if( Test-Connection -count 1 -computer "$server.corp.gwpnet.com" -Quiet)

            {

                $server = "$server.corp.gwpnet.com";

                return $server;

            }

            if( Test-Connection -count 1 -computer "$server.exodus.swissre.com" -Quiet)

            {

                $server = "$server.exodus.swissre.com";

                return $server;

            }

            if( Test-Connection -count 1 -computer "$server.devcorp.devgwpnet.com" -Quiet)

            {

                $server = "$server.devcorp.devgwpnet.com";

                return $server;

            }

            if( Test-Connection -count 1 -computer "$server.gwpnet.com" -Quiet)

            {

                $server = "$server.gwpnet.com";

                return $server;

            }

            

        }

        Catch

        {

            Write-Host "Unable to find FQDN for $server";

        }

    }

}

 

 

$searched_application = Read-Host -Prompt "`nPlease enter name of required software. `nIt will be searched on servers included in server_list.txt file";

SearchApplication -app $searched_application;

 

Read-Host -Prompt "Please close this window."

