Feature: Checking whether users are listed in the AD Users list

  As an infrastructure team
  we want to verify that expected users are available in Azure
  So that appropriate control can be maintained over security

  @Security @UserPermissions
Scenario: The correct RM AD users are in the Azure subscription
  Given I want to check azure details using the "Get-AzureRmADUser" command
  When I check Azure using the Azure Resource Management script
  Then I expect the user list to include "Derek" and "Paul" and "Carl"

  @Security @UserPermissions
Scenario: If your name isn't down, you're not coming in
  Given I want to check azure details using the "Get-AzureRmADUSer" command
  When I check Azure using the Azure Resource Management script
  Then I expect the output to not include "Mickey Mouse"

Scenario: Get a different powershell output
  Given I want to check azure details using the "Get-AzureSubscription" command
  When I check Azure using the Azure Resource Management script
  Then I expect the output to include "bab9ed05-2c6e-4631-a0cf-7373c33838cc"

  Scenario: Another test
    Given I want to check azure details using the "Get-AzureSubscription" command
    When I check Azure using the Azure Resource Management script
    Then I expect the output to not include "aaaaa"




