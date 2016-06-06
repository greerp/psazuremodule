
$url = "https://graph.facebook.com/v2.6/301847349881903/members?access_token=EAACEdEose0cBAEIkeVwnVqZBtImvj6UWCnfZCqA13kxOJFPRUZAqCi5iEs2AZAtGSHjFUZBghZBfYeIU2t4lJnc7CS0jvxxmetZBNxFBjB0UZAwYSnY8D1hhvmgrgSSwjdLmTnzmlAfat8u25RZC7BxPishUdO26lMzL1bYpGFRTQQwZDZD"
$members = @()
$page = 0

while ( $url ){
    $page++
    $data = Invoke-RestMethod $url
    foreach ( $member in $data.data ) {
        $members+=$member
    }
    $url = $data.paging.next
}



