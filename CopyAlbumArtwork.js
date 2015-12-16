/**
  * Try and find artwork from copies of audio files
  */
var   fso = new ActiveXObject("Scripting.FileSystemObject");
var   iTunesApp = WScript.CreateObject("iTunes.Application");
var   logFile = fso.CreateTextFile("CopyAlbumArtwork.log", true);

var NAME = "cover";
var i;
var   ITTrackKindFile	= 1;

var AlbumArt = new Array();

var imageExtensions = new Array(4);
imageExtensions[1] = 'jpg';
imageExtensions[2] = 'png';
imageExtensions[3] = 'bmp';

var libraryPlaylist = iTunesApp.LibraryPlaylist;
var tracks = libraryPlaylist.Tracks;

function saveArtwork( currTrack ) {
	var file = null;
	var location = currTrack.location;
	var artwork = currTrack.Artwork.Item(1);
	var format = artwork.Format;
	file = location.substring(0, location.lastIndexOf('\\') + 1) + NAME + '.' + imageExtensions[format];
	if ( ! fso.FileExists(file) ) {
		artwork.SaveArtworkToFile(file);
		logFile.WriteLine( "Saved artwork for this album " + currTrack.Album );
	} else {
		logFile.WriteLine( "Found previously stored artwork for this album " + currTrack.Album );
	} 
	return file;
}

for (i = 1; i <= tracks.Count; i++) 
{
	var currTrack = tracks.item(i);
	if ( ( currTrack.Kind == ITTrackKindFile ) &&
		 ( currTrack.KindAsString.indexOf("audio file" ) != -1 ) )
	{
		// If we have artwork for this album but not saved in AlbumArt, save it
		if (currTrack.Artwork.Count > 0) {
			// NB: This is a bit risky for Albums with names like "Greatest Hits", therefore check that we don't already have artwork
			if ( ( currTrack.Album != "" ) && ( currTrack.Album != "Unknown Album" ) && ( AlbumArt[currTrack.Album] == null ) ) {
				try {
					AlbumArt[currTrack.Album] = saveArtwork( currTrack );
				} catch (exception) {
					logFile.WriteLine( "Failure on " + currTrack.Album );
				}
			} 
		// Missing artwork
		} else {
			logFile.WriteLine("Missing artwork for album " + currTrack.Album );

			// Did we already extract the artwork for this album?
			if ( ( currTrack.Album != "" ) && ( currTrack.Album != "Unknown Album" ) && ( AlbumArt[currTrack.Album] != null ) ) {
				currTrack.AddArtworkFromFile(AlbumArt[currTrack.Album]);
				// Refresh track information from its associated file, so that artwork is updated
				currTrack.UpdateInfoFromFile();
				logFile.WriteLine( "Added artwork for this album " + currTrack.Album );
			} else {
				var location = currTrack.location;
				// Don't look for imageExtensions[0] as that is not defined
				for ( format = 1; format < imageExtensions.length; format++ ) {
					var file = location.substring(0, location.lastIndexOf('\\') + 1) + NAME + '.' + imageExtensions[format];
	
					if ( fso.FileExists(file) ) {
						logFile.WriteLine("CopyAlbumArtwork found image file for " + currTrack.Artist + ", " + currTrack.Album + ", " + currTrack.Name );
					} else {
						logFile.WriteLine("Missing artwork and nothing found " + currTrack.Artist + ", " + currTrack.Album + ", " + currTrack.Name + ", " + currTrack.Location + " at " + file );
					}
				}
			}
		}
	}
}
