#requires -Version 2
function Get-PhishingGeoLocationAllIPs
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()] 
        [object[]]$AllIPData,
        [parameter(HelpMessage = 'Please provide a folder path for ouputting generated maps')]
        [ValidateNotNullOrEmpty()]
        [string]$FolderPath
    )

    $OutStartingIPMap = "$FolderPath\PhishReporterPolyline.html"


    $html = @"
<!DOCTYPE html>
<html>
  <head>
    <meta name="viewport" content="initial-scale=1.0, user-scalable=no">
    <meta charset="utf-8">
    <title>Complex icons</title>
    <style>
      html, body {
        height: 100%;
        margin: 0;
        padding: 0;
      }
      #map {
        height: 100%;
      }
    </style>
  </head>
  <body>
    <div id="map"></div>
    <script>



    var markers = [ $(foreach ($item in $AllIPData){ "$($item.marker)," }) ];

    var Colors = [
    "#FF0000", 
    "#00FF00", 
    "#0000FF", 
    "#FFFFFF", 
    "#000000", 
    "#FFFF00", 
    "#00FFFF", 
    "#FF00FF"
    ];

    function initMap() {
        var map = new google.maps.Map(document.getElementById('map'), {
            zoom: 5,
            center: {lat: 37.419, lng: -92.3283}
    });

        for( var i = 0; i < markers.length; i++ ) {
		    var location = markers[i]
			var poly = new google.maps.Polyline({
			  path: location,
			  strokeColor: Colors[i],
			  strokeOpacity: 1.0,
			  strokeWeight: 3,
			  map: map    
			});
        }  
    }
    </script>
    <script async defer
        src="https://maps.googleapis.com/maps/api/js?signed_in=true&callback=initMap"></script>
  </body>
</html>   
"@ | Out-File $OutStartingIPMap
}
