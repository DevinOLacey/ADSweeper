
        # Process users by job title after AD changes
        $usersByTitle = $AdUsers | Group-Object -Property Title
        foreach ($group in $usersByTitle) {
            $title = $group.Name
            $users = $group.Group

            if ($users.Count -eq 1) {
                Write-Host "Only one user with title '$title'. Groups are considered correct." -ForegroundColor Cyan
                continue
            }

            if ($users.Count -eq 2) {
                # Special handling for two users
                $user1Groups = Get-ADUser -Identity $users[0].SamAccountName -Properties MemberOf | Select-Object -ExpandProperty MemberOf
                $user2Groups = Get-ADUser -Identity $users[1].SamAccountName -Properties MemberOf | Select-Object -ExpandProperty MemberOf

                # Assuming "Domain Users" is the default group
                $defaultGroup = "CN=Domain Users,CN=Users,DC=yourdomain,DC=com"

                if ($user1Groups.Count -eq 1 -and $user1Groups -contains $defaultGroup) {
                    $missingGroups = $user2Groups
                    $userToUpdate = $users[0]
                } elseif ($user2Groups.Count -eq 1 -and $user2Groups -contains $defaultGroup) {
                    $missingGroups = $user1Groups
                    $userToUpdate = $users[1]
                } else {
                    $missingGroups = @()
                }

                foreach ($group in $missingGroups) {
                    try {
                        Add-ADGroupMember -Identity $group -Members $userToUpdate.SamAccountName
                        Write-Host "Added $($userToUpdate.SamAccountName) to group $group." -ForegroundColor Green
                    } catch {
                        Write-Host "Failed to add $($userToUpdate.SamAccountName) to group ${group}: $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
                continue
            }

            $commonGroups = Get-CommonGroups -users $users

            foreach ($user in $users) {
                $currentGroups = Get-ADUser -Identity $user.SamAccountName -Properties MemberOf | Select-Object -ExpandProperty MemberOf
                $missingGroups = $commonGroups | Where-Object { $currentGroups -notcontains $_ }

                foreach ($group in $missingGroups) {
                    try {
                        Add-ADGroupMember -Identity $group -Members $user.SamAccountName
                        Write-Host "Added $($user.SamAccountName) to group ${group}." -ForegroundColor Green
                    } catch {
                        Write-Host "Failed to add $($user.SamAccountName) to group ${group}: $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            }
        }
