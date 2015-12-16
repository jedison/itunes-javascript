/**
  * Try to find artwork from copies of audio files
  */
var   fso = new ActiveXObject("Scripting.FileSystemObject");
var   iTunesApp = WScript.CreateObject("iTunes.Application");
var   logFile = fso.CreateTextFile("FindArtwork.log", true);

var NAME = "cover";
var i;
var   ITTrackKindFile	= 1;

var AlbumArt = new Array();

var imageExtensions = new Array(4);
imageExtensions[1] = 'jpg';
imageExtensions[2] = 'png';
imageExtensions[3] = 'bmp';

var   tracks = iTunesApp.LibraryPlaylist.Tracks;
var libraryPlaylist = iTunesApp.LibraryPlaylist;

function saveArtwork2( currTrack, addedTrack ) {
	var file = null;
	if ( addedTrack.Artwork.Count > 0 ) {

		var artwork = addedTrack.Artwork.item(1);
		var format = artwork.Format;
		if (format != 0) 
		{
			var location = addedTrack.location;
			file = location.substring(0, location.lastIndexOf('\\') + 1) + NAME + '.' + imageExtensions[format];

			if ( addedTrack.Album == "" )
				return;

			AlbumArt[addedTrack.Album] = file;
			artwork.SaveArtworkToFile(file);
			currTrack.AddArtworkFromFile(AlbumArt[addedTrack.Album]);
			logFile.WriteLine( "Saved artwork for this album " + currTrack.Album );
		}
		// Refresh track information from its associated file, so that artwork is updated
		currTrack.UpdateInfoFromFile();
		logFile.WriteLine( "Added artwork for this album " + currTrack.Album );
	} else {
		logFile.WriteLine( "No artwork to save for album " + currTrack.Name );
	}
	// Delete converted track from iTunes so that there is not a duplicate
	addedTrack.Delete();
}	

logFile.WriteLine( "tracks.Count=" + tracks.Count );

for (i = 1; i <= tracks.Count; i++) 
{
	var currTrack = tracks.item(i);
	if ( ( currTrack.Kind == ITTrackKindFile ) &&
		 ( currTrack.KindAsString.indexOf("audio file" ) != -1 ) )
	{
		var artworks = currTrack.Artwork;

		// Missing artwork
		if (artworks.Count == 0) 
		{
			logFile.WriteLine( "Missing artwork on " + currTrack.Name );
			
			// Did we already extract the artwork for this album?
			if ( ( currTrack.Album != "" ) && ( currTrack.Album != "Unknown Album" ) && ( AlbumArt[currTrack.Album] != null ) ) {
				currTrack.AddArtworkFromFile(AlbumArt[currTrack.Album]);
				// Refresh track information from its associated file, so that artwork is updated
				currTrack.UpdateInfoFromFile();
				logFile.WriteLine( "Added artwork for this album " + currTrack.Album );
			} else {

				var location = currTrack.location;
				// ASSUMES that currTrack is on E: and that possible old versions are on C:, and that they are MP3 files!
				var file = "C:" + location.substring(2, location.lastIndexOf('.') ) + '.mp3' ;

				if ( fso.FileExists(file) ) {
					var status = libraryPlaylist.AddFile( file );

					while ( status.InProgress == true ) {
						// Wait one second
						WScript.Sleep(1000);
					}

					var addedTracks = status.Tracks;

					if ( addedTracks.Count == 1 ) {
						var addedTrack = addedTracks.Item(1);

						saveArtwork2( currTrack, addedTrack );
					} else {
						logFile.WriteLine("iTunes addFile did not return one track for " + currTrack.Name );
						logFile.WriteLine("iTunes addFile status.InProgress=" + status.InProgress );
						if ( addedTracks == null ) {
							logFile.WriteLine("iTunes addFile addedTracks == null" );
						} else {
							logFile.WriteLine("iTunes addFile addedTracks " + addedTracks.Count );
						}
					}
				} else {
					 try {
						 logFile.WriteLine("Corresponding file not found at " + file );
					 } catch (exception) {
						 logFile.WriteLine("Corresponding file not found for " + currTrack.Name + "; file name not writable.");
					 }
				}
			}
		}
	}
}
logFile.Close();
