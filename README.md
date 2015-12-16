# itunes-javascript
Javascripts to be run on Windows to perform iTunes related tasks

Over the past months (nay, years) I have been cleaning up my music library. I have used different tools and services (including paying ones like Tune-Up).  But, I am lazy, my goal has always been to automate things as much as possible. So I was very happy to find various scripts which will, for example, search for "dead" tracks (tracks for which their is no corresponding media file) rather than having to "Get Info" on a track and then sit holding ALT+N (on Windows) to go to the Next track while iTunes opens ever track in my library to see how many "broken" links I have.

I have found a few different sources (some of which are just copies of each other or copies of examples from Apple) to "Manage iTunes on Windows using JavaScript with the iTunes SDK". Obviously, there is also the iTunes COM Windows SDK which includes a few JS samples.
Here are a couple of links to other sources of Javascript for iTunes.

- [OttoDestruct](http://ottodestruct.com/blog/2005/itunes-javascripts/)
- [EverythingiTunes](https://everythingitunes.wordpress.com/scripts/)
- [Brian Ander](https://github.com/briananders/iTunes-JavaScript-JScript)

Now, for me, these are good resources, but what I am missing is some more powerful and tricky scripts. For example, I have recently subscribed to iTunes Match so that I can upgrade my older MP3 ripped music into higher bitrate audio (e.g. 256 kbps AAC).
The problem that I have with this is that the "Matched AAC audio file" that iTunes Match allows me to download has my iTunes account stamped into it. So what if I want to share this audio file with family or friends? What happens (hopefully far in the future) when I die and my iTunes account is no longer active and iTunes Match is turned off?
Well, basically, I’d like to have the audio files in a format which I own, not which Apple controls (e.g. "AAC audio file", not "Matched AAC audio file").

And now, here comes the reason for this posting. I have several playlists of hundreds and thousands of songs. Most of those playlists are built from converted (MP3) and purchased ("Purchased AAC audio file") tracks, so I can’t upgrade from MP3 files or Purchased AAC to Matched AAC audio files and then and delete the MP3 or Purchased version.
So, I started working on somewhat more complex scripts which would help me smoothly upgrade my collection to clean AAC audio files without the traces of purchase or iTunes matching.

These scripts are not perfect and certainly need some upgrading, but as I find them useful you might as well.

1. This first one does the main job of finding all tracks of kind "Matched AAC audio file" and running the iTunes ConvertFile2(). Please note that this assumes that you have the correct audio converter setup in iTunes; I know that it is possible to set the audio converter (and reset it to the original converter when done, but I couldn’t be bothered to implement that yet): ConvertMatched.js
2. This one is a bit more specific to my situation. I followed guides about iTunesMatch upgrading of songs ([MacWorld]( http://www.macworld.com/article/1163620/how_to_upgrade_tracks_to_itunes_match_fast.html)), but then found that some of my upgraded songs (cleaned up with the script from #1 above) did not have artwork. So I restored them from the Recycle Bin (to a different location than my actual iTunes library, so that my active library is on E: and the restored files are on C:). NB: this scripts makes some assumptions, so don’t just use it blindly: FindArtwork.js
3. And that got me thinking, some of these albums already have artwork, just not on the upgraded (via iTunes Match) songs. So, another script was born: CopyAlbumArtwork.js
This is similar to Brian Anders’ SelectedArtworkUnify.js, but I think it’s improved.
Here I add a word of warning, this script should work fine except for albums of different artists with the same name like "Greatest Hits".
4. Here is the simple, classic "NoArtwork.js", but be aware that a Smart Playlist in iTunes is more powerful as it can be defined to update "live" (each time a track gets artwork it is automatically removed from the playlist: no human intervention required).
NoArtwork.js
If you don’t know how to make a Smart Playlist to show all Music tracks without Artwork, then just ask, and I’ll post how that is done.

To use any of these scripts, just download them, save them to a file with a name ending *.js file, then run them like any other program (double click, select and press enter, type the name from a command prompt, etc). If you have the Windows Scripting Host installed on other Windows boxes (installed by default), then the wscript.exe program actually runs the scripts, much like cmd.exe runs batch files. The scripts connect to iTunes as a COM object and use it in that fashion.
