/**
  * This Program goes through all the selected tracks, saves the album art, and links the songs with no artwork to the saved artwork for that album.
  * Improved on version from Brian Anders
  */

var extensions;
var NAME = "Folder";
var t;
var   ITTrackKindFile	= 1;

var AlbumArt = new Array();

var extensions = new Array(4);
imageExtensions[1] = 'jpg';
imageExtensions[2] = 'png';
imageExtensions[3] = 'bmp';

var tracks = WScript.CreateObject("iTunes.Application").LibrarySource.Playlists.ItemByName("Music");
for (t = 1; t <= tracks.Count; t++) 
{
	try 
	{
		var currTrack = tracks.item(t);
		if ( ( currTrack.Kind == ITTrackKindFile ) &&
			 ( currTrack.KindAsString.indexof("audio file" ) != -1 ) )
		{
			var artworks = currTrack.Artwork;
			if (artworks.Count > 0) 
			{
				var artwork = artworks.item(1);
				var format = artwork.Format;
				if (format != 0) 
				{
					var location = currTrack.location;
					var file = location.substring(0, location.lastIndexOf('\\') + 1) + NAME + '.' + extensions[format];
					AlbumArt[currTrack.Album] = file;
					artwork.SaveArtworkToFile(file);
				}
			}
		}
	}
	catch (exception) 
	{
		WScript.echo("Problem with " + currTrack.Name);
	}
}
WScript.Echo("Starting");

for (t = 1; t <= tracks.Count; t++) 
{
	try 
		{
		var currTrack = tracks.Item(t);
		if(currTrack != null)
		{
			if(AlbumArt[currTrack.Album] != null && currTrack.Artwork.Count == 0)
			{
				currTrack.AddArtworkFromFile(AlbumArt[currTrack.Album]);
			}
		}
	}
	catch(exception)
	{
	}
}
WScript.Echo("Done");
