#requires -Version 2
function Get-PhishingGeoLocationStartingIPs
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()] 
        [object[]]$StartingIPData,
        [parameter(HelpMessage = 'Please provide a folder path for ouputting generated maps')]
        [ValidateNotNullOrEmpty()]
        [string]$FolderPath
    )

    $OutStartingIPMap = "$FolderPath\PhishReporterFirstReceivedIPMap.html"


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



    var markers = [ $($StartingIPData.marker -join ',') ];

    function initMap() {
        var map = new google.maps.Map(document.getElementById('map'), {
            zoom: 5,
            center: {lat: 37.419, lng: -92.3283}
    });

	
		var infoWindow = new google.maps.InfoWindow();
		
        for( var i = 0; i < markers.length; i++ ) {
		    var location = markers[i]
			latlng = new google.maps.LatLng(location.lat, location.lng);
            var marker = new google.maps.Marker({
			    position: latlng,
			    map: map,
			    title: location.title
		    });

            (function(marker, location) {
				// Attaching a click event to the current marker
				google.maps.event.addListener(marker, "click", function(e) {
				infoWindow.setContent(location.description);
				infoWindow.open(map, marker);
			});
			})(marker, location);
	    }

    }
    </script>
    <script async defer
        src="https://maps.googleapis.com/maps/api/js?signed_in=true&callback=initMap"></script>
  </body>
</html>   
"@ | Out-File $OutStartingIPMap
}
