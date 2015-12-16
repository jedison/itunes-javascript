var   iTunesApp = WScript.CreateObject("iTunes.Application");
var   tracks = iTunesApp.LibraryPlaylist.Tracks;
var   fso = new ActiveXObject("Scripting.FileSystemObject");
var   logFile = fso.CreateTextFile("ConvertPurchased.log", true);
var   i, j;
var   TO_CONVERT = "AAC audio file";

function getAllTracksToConvert(tracksCollection)
{
	var   convertFile = fso.CreateTextFile("TracksWithPurchasedDate.txt", true);
	var   k;
	var   convertTracks = [];
	var   convertCount = 0;
	var   ITTrackKindFile	= 1;

	for (k = 1; k <= tracksCollection.Count; k++) {
		var	currTrack = tracksCollection.Item( k );
		// Is this a file track?
		// And type to convert
		// And Not empty location and File exists?
		if ( 
		     ( currTrack.Kind == ITTrackKindFile ) &&
		     ( currTrack.KindAsString == TO_CONVERT) &&
		     ( currTrack.Location != "" ) &&
		     ( fso.FileExists( currTrack.Location ) ) &&
		     ( currTrack.PurchasedDate != null) &&
		     ( currTrack.PurchasedDate.trim() != "") 
		   )
		{
			try {
				convertFile.WriteLine( currTrack.Artist + "," + currTrack.Album + "," + currTrack.Name + "," + currTrack.Location );
			} catch ( exception ) {
				convertFile.WriteLine( "Cannot write some track" );
			}
			convertCount++;
			convertTracks.push( currTrack );
		}
	}
	logFile.WriteLine("Found " + convertCount + " track(s) of kind " + TO_CONVERT + " to convert. " );
	convertFile.Close();
	return convertTracks;
}

var tracksToConvert = getAllTracksToConvert( tracks );


logFile.Close();
