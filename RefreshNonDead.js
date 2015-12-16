var   iTunesApp = WScript.CreateObject("iTunes.Application");
var   mainLibrary = iTunesApp.LibraryPlaylist;
var   tracks = mainLibrary.Tracks;
var   numTracks = tracks.Count;
var   i;
var   ITTrackKindFile	= 1;
var   deletedTracks = 0;
var   fso = new ActiveXObject("Scripting.FileSystemObject");
var   deadTracks = fso.CreateTextFile("DeadTracks.txt", true);

for (i = 1; i <= numTracks; i++)
{
	var	currTrack = tracks.Item( i );
	
	// is this a file track?
	if ( ( currTrack != null ) && 
		 ( currTrack.Kind == ITTrackKindFile ) )
	{
		// yes, does it have an empty location?
		if ( ( currTrack.Location == "" ) ||
			 ( ! fso.FileExists( currTrack.Location ) ) )
		{
			// write info about the track to a file
			deadTracks.WriteLine( currTrack.Artist + "," + currTrack.Album + "," + currTrack.Name + "," + currTrack.Location );
			deletedTracks++;
		} 
		else
		{
			if ( fso.FileExists( currTrack.Location ) ) 
			{
				currTrack.UpdateInfoFromFile();
			}
        	}
	}
}

if (deletedTracks > 0)
{
		WScript.Echo("Found " + deletedTracks + " dead track(s).");
}
else
{
	WScript.Echo("No dead tracks were found.");
}
deadTracks.Close();
