
$subscriptionID="311818f8-d369-419b-bfe1-fdf644de096f"
$resourceGroup="pgrg071"
$apiVersion="2014-04-01-preview"
$deploymentName="pgrg072"
$jsonSource=".\azuredeploy.json"

$content = Get-Content $jsonSource

$body = "{
  'properties': {
    'template': {$content)},
    'mode': 'Incremental',
    'parameters': {
        'siteLocation': {
            'value': 'West Europe'
        },

        'sku': {
            'value': 'Free'
        },

        'administratorLogin': {
            'value': 'greepau'
        },

        'administratorPassword': {
            'value': 'R3dpixie'
        },

        'applicationName': {
            'value': 'Default Application Name'
        },

        'repoUrl': {
            'value': 'na'
        },
        'branch': {
            'value': 'na'
        }      
    }
  }
}"




$uri = "https://management.azure.com/subscriptions/${subscriptionID}/resourcegroups/${resourceGroup}/providers/microsoft.resources/deployments/${deploymentName}/validate?api-version=${apiVersion}"

$result = Invoke-RestMethod -Method Post -Uri $uri -Body $body

Write-Host $result
