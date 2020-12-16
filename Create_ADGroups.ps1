#Nahrání assemblerů
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#Předdefinování proměnných pro pozdější použití
$ou_path = "***"
$credentials = Get-Credential
$credentials

#Import PSSession
try {
    $exchange = New-PSSession -ConnectionUri "****" -ConfigurationName Microsoft.Exchange -Credential $credentials -ErrorAction Stop
    Import-PSSession $exchange -ErrorAction Stop
}
catch {
    Write-Warning "Wrong username or password"
	sleep(3)
    break
}

#Založení hlavního okna
$main_frame = New-Object System.Windows.Forms.Form
$main_frame.Text = "ADGroup.Establishment"
$main_frame.Width = 400
$main_frame.Height = 280
$main_frame.AutoSize

#Popisek1
$label1 = New-Object System.Windows.Forms.Label
$label1.Text = "Active Directory GroupType: "
$label1.Location = New-Object System.Drawing.Point(8, 15)
$label1.Font = "Arial,9"
$label1.AutoSize = $true
$main_frame.Controls.Add($label1)

#Stahovácí lišta
$selection_box = New-Object System.Windows.Forms.ComboBox
$selection_box.Location = New-Object System.Drawing.Point(187, 13)
$selection_box.Width = 175
$selection_box.DropDownStyle = "DropDownList"
$selection_box.AutoCompleteSource = "ListItems"
$selection_box.AutoCompleteMode = "Append"
$selection_box.Items.AddRange(@("****"))
$main_frame.Controls.Add($selection_box)

#Popisek2
$label2 = New-Object System.Windows.Forms.Label
$label2.Text = "Active Directory Group name: "
$label2.Location = New-Object System.Drawing.Point(8,50)
$label2.Font = "Arial,9"
$label2.AutoSize = $true
$main_frame.Controls.Add($label2)

#Vytvoření textového pole
$ad_group_name = New-Object System.Windows.Forms.RichTextBox
$ad_group_name.Location = New-Object System.Drawing.Point(187, 48)
$ad_group_name.Width = 175
$ad_group_name.Height = 22
$main_frame.Controls.Add($ad_group_name)

#Popisek3
$label3 = New-Object System.Windows.Forms.Label
$label3.Location = New-Object System.Drawing.Point(8, 85)
$label3.Text = "SR-Ticket number:"
$label3.Font = "Arial,9"
$label3.AutoSize = $true
$main_frame.Controls.Add($label3)

#Vytvoření textového pole
$sr_ticket = New-Object System.Windows.Forms.RichTextBox
$sr_ticket.Location = New-Object System.Drawing.Point(187, 84)
$sr_ticket.Width = 175
$sr_ticket.Height = 22
$main_frame.Controls.Add($sr_ticket)

#Popisek4
$label4 = New-Object System.Windows.Forms.Label
$label4.Location = New-Object System.Drawing.Point(8, 120)
$label4.Text = "Specific OU_Group"
$label4.Font = "Arial,9"
$label4.AutoSize = $true
$main_frame.Controls.Add($label4)

#Vytvoření stahovací lišty
$ou_group_select = New-Object System.Windows.Forms.ComboBox
$ou_group_select.Location = New-Object System.Drawing.Point(187, 118)
$ou_group_select.Width = 175
$ou_group_select.DropDownStyle = "DropDownList"
$ou_group_select.AutoCompleteSource = "ListItems"
$ou_group_select.AutoCompleteMode = "Append"
$ou_group_select.Items.AddRange(@(*****))
$main_frame.Controls.Add($ou_group_select)

#Konstrukce tlačítka "button1"
$button1 = New-Object System.Windows.Forms.Button
$button1.Size = New-Object System.Drawing.Size(55, 30)
$button1.Location = New-Object System.Drawing.Point(20, 170)
$button1.Text = "Create"
$main_frame.Controls.Add($button1)

#Popisek5
$label5 = New-Object System.Windows.Forms.Label
$label5.Location = New-Object System.Drawing.Point(140, 178)
$label5.Text = ""
$label5.Font = "Arial, 9"
$label5.AutoSize = $true
$main_frame.Controls.Add($label5)

#Definice funkce "kliknutí" na vytvoření tlačtko "button1"
$button1.Add_Click(
    {   
        $button1_output = $selection_box.Text
        $text_box_output = $ad_group_name.Text
        $ticket_number = $sr_ticket.Text
        $ou_group = $ou_group_select.Text

        if ($button1_output -eq "" -or $text_box_output -eq "" -or $ticket_number -eq "" -or $ou_group -eq "") {
            $label5.ForeColor = "Red"
            $label5.Text = "Parameter/s missing!"
        }
        else {

            if ($button1_output -eq "****") {

                if ($ou_group -eq "None") {
                    
                    try {                        
                        New-ADGroup -Name "$text_box_output" -DisplayName "$text_box_output" -SamAccountName "$text_box_output" `
                        -GroupCategory Security -GroupScope Global -Path "$ou_path" -Description "$ticket_number" -Credential $credentials
                        $label5.ForeColor = "Green"
                        $label5.text = "AD group created!"
						sleep(2)
						$label5.text = ""
                    }
                    catch {
                        $label5.ForeColor = "Red"
                        $label5.text = "AD Group creation failed!"
                    }
                }    
            
                else {

                    try {                        
                        New-ADGroup -Credential $credentials -Name $text_box_output -DisplayName $text_box_output -SamAccountName $text_box_output `
                        -GroupCategory Security -GroupScope Global -Path "OU=$ou_group,$ou_path" -Description $ticket_number
                        $label5.ForeColor = "Green"
                        $label5.text = "AD group created!"
						sleep(3)
						$label5.text = ""
                    }

                    catch {
                        $label5.ForeColor = "Red"
                        $label5.text = "AD Group creation failed!"
                        Write-Warning $Error[0]
                    }
                }
            }
            else {
                
                try {
                    $dist_name = $text_box_output -replace "***", "***".ToString()
                    New-DistributionGroup -SamAccountName $dist_name -Name $dist_name -DisplayName $text_box_output -PrimarySmtpAddress "$text_box_output" -OrganizationalUnit $ou_path
                    $count = 0
                    while ($count -ne 1) {
                        try {
                            Set-ADGroup $dist_name -Description $ticket_number
                            $count += 1
                            
                        }
                        catch {
                            continue
                        }
                    }

                    $label5.ForeColor = "Green"
                    $label5.Text = "AD Distribution group created!"
					sleep(3)
                    $label5.Text = ""
                }
                
                catch {
                    $label5.ForeColor = "Red"
                    $label5.Text = "AD Distribution group creation failed!"
                    Write-Warning $Error[0]
                }

            }
        }  
    }
)

#Zobrazení celého okna se všemi funkcemi
$main_frame.ShowDialog() | Out-Null