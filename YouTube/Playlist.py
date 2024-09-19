from pytube import Playlist, YouTube
import os

def download_playlist(playlist_url, download_path):
    # Create the download directory if it does not exist
    if not os.path.exists(download_path):
        os.makedirs(download_path)

    # Create a Playlist object
    playlist = Playlist(playlist_url)
    
    # Print the total number of videos in the playlist
    print(f'Downloading playlist: {playlist.title}')
    print(f'Number of videos in playlist: {len(playlist.video_urls)}')
    
    # Iterate through all the video URLs in the playlist
    for url in playlist.video_urls:
        try:
            # Create a YouTube object for each video
            video = YouTube(url)
            print(f'Downloading video: {video.title}')
            
            # Download the highest resolution stream available
            video.streams.get_highest_resolution().download(output_path=download_path)
            print(f'Successfully downloaded: {video.title}')
        except Exception as e:
            print(f'Failed to download video: {url}. Error: {str(e)}')

# Example usage
playlist_url = input("Enter the YouTube playlist URL: ")
download_path = input("Enter the directory to download the videos to: ")
download_playlist(playlist_url, download_path)
