var iTunesApp = WScript.CreateObject("iTunes.Application"); 
var tracks = iTunesApp.LibrarySource.Playlists.ItemByName("Music");
var numTracks = tracks.Count;
var i;
NoArtPlaylist = iTunesApp.CreatePlaylist("No Artwork");

for (i = 1; i <= numTracks; i++) 
{ 	
	var currTrack = tracks.Item(i); 
	if ( currTrack.Artwork.Count == 0 ) 
		NoArtPlaylist.AddTrack(currTrack);
} 
