#requires -Version 2
function Get-PhishingGeoLocationHeatMap
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()] 
        $HeatMapData,
        [parameter(HelpMessage = 'Please provide a folder path for ouputting generated maps')]
        [ValidateNotNullOrEmpty()]
        [string]$FolderPath
    )

    $OutHeatMap = "$FolderPath\PPRT_HeatMap.html"

    $webdata = @"
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
	  #floating-panel {
  position: absolute;
  top: 10px;
  left: 25%;
  z-index: 5;
  background-color: #fff;
  padding: 5px;
  border: 1px solid #999;
  text-align: center;
  font-family: 'Roboto','sans-serif';
  line-height: 30px;
  padding-left: 10px;
}

      #floating-panel {
        background-color: #fff;
        border: 1px solid #999;
        left: 25%;
        padding: 5px;
        position: absolute;
        top: 10px;
        z-index: 5;
      }
    </style>
  </head>
  <body>
	<div id="floating-panel">
		<button onclick="changeRadius()">Change radius</button>
		<button onclick="changeGradient()">Change gradient</button>
	</div>
    <div id="map"></div>
    <script>

	var map, heatmap;

function initMap() {

    var heatmapData = [ $($HeatMapData.marker -join ',') ];

    	var customMapType = new google.maps.StyledMapType([
	{
		"featureType": "all",
		"elementType": "labels.text.fill",
		"stylers":
		[
			{"saturation":36},
			{"color":"#000000"},
			{"lightness":40}
		]
	},
	{
		"featureType":"all",
		"elementType":"labels.text.stroke",
		"stylers":
		[
			{"visibility":"on"},
			{"color":"#000000"},
			{"lightness":16}
		]
	},
	{
		"featureType":"all",
		"elementType":"labels.icon",
		"stylers":
		[
			{"visibility":"off"}
		]
	},
	{
		"featureType":"administrative",
		"elementType":"geometry.fill",
		"stylers":
		[
			{"color":"#000000"},
			{"lightness":20}
		]
	},
	{
		"featureType":"administrative",
		"elementType":"geometry.stroke",
		"stylers":
		[
			{"color":"#000000"},
			{"lightness":17},
			{"weight":1.2}
		]
	},
	{
		"featureType":"landscape",
		"elementType":"geometry",
		"stylers":
		[
			{"color":"#000000"},
			{"lightness":20}
		]
	},
	{
		"featureType":"poi",
		"elementType":"geometry",
		"stylers":
		[
			{"color":"#000000"},
			{"lightness":21}
		]
	},
	{
		"featureType":"road.highway",
		"elementType":"geometry.fill",
		"stylers":
		[
			{"color":"#000000"},
			{"lightness":17}
		]
	},
	{
		"featureType":"road.highway",
		"elementType":"geometry.stroke",
		"stylers":
		[
			{"color":"#000000"},
			{"lightness":29},
			{"weight":0.2}
		]
	},
	{
		"featureType":"road.arterial",
		"elementType":"geometry",
		"stylers":
		[
			{"color":"#000000"},
			{"lightness":18}
		]
	},
	{
		"featureType":"road.arterial",
		"elementType":"labels.text.fill",
		"stylers":
		[
			{"visibility":"on"},
			{"color":"#ffffff"}
		]
	},
	{
		"featureType":"road.local",
		"elementType":"geometry",
		"stylers":
		[
			{"color":"#000000"},
			{"lightness":16}
		]
	},
	{
		"featureType":"road.local",
		"elementType":"labels.text.fill",
		"stylers":
		[
			{"color":"#ffffff"}
		]
	},
	{
		"featureType":"transit",
		"elementType":"geometry",
		"stylers":
		[
			{"color":"#000000"},
			{"lightness":19}
		]
	},
	{
		"featureType":"water",
		"elementType":"geometry",
		"stylers":
		[
			{"color":"#37424a"},
			{"lightness":17}
		]
	}
	],
	{
      name: 'Custom Style'
	});
  
	var customMapTypeId = 'custom_style';
    var ColumbiaMo = new google.maps.LatLng(37.419, -92.3283);

    map = new google.maps.Map(document.getElementById('map'),
	{
        center: ColumbiaMo,
        zoom: 5,
        mapTypeControlOptions: {
			mapTypeIds: [google.maps.MapTypeId.ROADMAP, customMapTypeId]
		}
    });
	
	map.mapTypes.set(customMapTypeId, customMapType);
	map.setMapTypeId(customMapTypeId);
  
    heatmap = new google.maps.visualization.HeatmapLayer({
        data: heatmapData
    });
	

    heatmap.setOptions({radius: heatmap.get('1'),
						opacity: 1});
    heatmap.setMap(map);

}
    function changeRadius() {
	heatmap.set('radius', heatmap.get('radius') ? null : 40);
	}

	function changeGradient() {
  var gradient = [
    'rgba(0, 255, 255, 0)',
    'rgba(0, 255, 255, 1)',
    'rgba(0, 191, 255, 1)',
    'rgba(0, 127, 255, 1)',
    'rgba(0, 63, 255, 1)',
    'rgba(0, 0, 255, 1)',
    'rgba(0, 0, 223, 1)',
    'rgba(0, 0, 191, 1)',
    'rgba(0, 0, 159, 1)',
    'rgba(0, 0, 127, 1)',
    'rgba(63, 0, 91, 1)',
    'rgba(127, 0, 63, 1)',
    'rgba(191, 0, 31, 1)',
    'rgba(255, 0, 0, 1)'
  ]
  heatmap.set('gradient', heatmap.get('gradient') ? null : gradient);
}

    </script>
    <script async defer
        src="https://maps.googleapis.com/maps/api/js?signed_in=true&callback=initMap&libraries=visualization"></script>
  </body>
</html>  
"@ | Out-File $OutHeatMap
}
