var   iTunesApp = WScript.CreateObject("iTunes.Application");
var   tracks = iTunesApp.LibraryPlaylist.Tracks;
var   fso = new ActiveXObject("Scripting.FileSystemObject");
var   logFile = fso.CreateTextFile("ConvertPurchased.log", true);
var   convertSize = 10;
var   i, j;
var   TO_CONVERT = "Purchased AAC audio file";

function getAllTracksToConvert(tracksCollection)
{
//	var   convertFile = fso.CreateTextFile("Converted.txt", true);
	var   k;
	var   convertTracks = [];
	var   convertCount = 0;
	var   ITTrackKindFile	= 1;

	for (k = 1; k <= tracksCollection.Count; k++) {
		var	currTrack = tracksCollection.Item( k );
		// Is this a file track?
		// And type to convert
		// And Not empty location and File exists?
		if ( ( currTrack.Kind == ITTrackKindFile ) &&
			 ( currTrack.KindAsString == TO_CONVERT) &&
		     ( currTrack.Location != "" ) &&
		     ( fso.FileExists( currTrack.Location ) ) )
		{
			try {
				logFile.WriteLine( currTrack.Artist + "," + currTrack.Album + "," + currTrack.Name + "," + currTrack.Location );
			} catch ( exception ) {
				logFile.WriteLine( "Cannot write some track" );
			}
			convertCount++;
			convertTracks.push( currTrack );
		}
	}
	logFile.WriteLine("Found " + convertCount + " track(s) of kind " + TO_CONVERT + " to convert. " );
//	logFile.Close();
	return convertTracks;
}

var tracksToConvert = getAllTracksToConvert( tracks );
var   convertedCount = 0;

for ( i = 0; i < tracksToConvert.length; i++) {
	var currTrack = tracksToConvert[i];

	try {
		logFile.WriteLine("Found track of kind " + TO_CONVERT + " to convert. " + currTrack.Location);
	} catch ( exception ) {
		logFile.WriteLine("Issues for track with name " + currTrack.Name);
		continue;
	}
	var status = iTunesApp.ConvertFile2( currTrack.Location );
	while ( status.InProgress == true ) {
		// Wait one second
		WScript.Sleep(1000);
	}

	var convertedTracks = status.Tracks;

	if ( convertedTracks.Count == 1 ) {
		var oldLocation = currTrack.Location;

		logFile.WriteLine("Converted " + oldLocation + " to " + convertedTracks.Item(1).Location );
		// Delete file at location that currTrack is stored.
		fso.DeleteFile( oldLocation );
		// Move newly converted file to location where currTrack was stored.
		fso.MoveFile( convertedTracks.Item(1).Location, oldLocation );

		logFile.WriteLine("Moved newly converted file to " + oldLocation );

		// Delete converted track from iTunes so that there is not a duplicate
		convertedTracks.Item(1).Delete();
		// Refresh track information from its associated file, so that track is updated from old kind to newly converted kind
		currTrack.UpdateInfoFromFile();
		convertedCount++;
	} else {
		logFile.WriteLine("iTunes ConvertFile2 did not return one track for " + currTrack.Artist + "," + currTrack.Album + "," + currTrack.Name );
	}
}

logFile.WriteLine("Converted " + convertedCount + " track(s) of kind " + TO_CONVERT + ". " );
logFile.Close();
