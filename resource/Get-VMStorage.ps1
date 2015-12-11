
$rgname = "pgrg070"
$vms = (Get-AzureResource -ResourceGroupName $rgname -ResourceType "Microsoft.Compute/virtualMachines" -ExpandProperties)

foreach ($vm in $vms){
    $storProps = $vm.Properties["storageProfile"];
    write-host $vm.Name, $storProps

}
