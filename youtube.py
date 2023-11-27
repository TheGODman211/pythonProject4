# import youtube_dl
#
#
# def download_youtube_playlist(playlist_url):
#     ydl_opts = {
#         'format': 'bestvideo[ext=mp4]+bestaudio[ext=m4a]/best[ext=mp4]/best',
#         'outtmpl': '%(title)s.%(ext)s',
#         'ignoreerrors': True,
#     }
#
#     with youtube_dl.YoutubeDL(ydl_opts) as ydl:
#         ydl.download([playlist_url])
#
#
# # Example usage
# playlist_url = 'https://www.youtube.com/watch?v=c2IBhnkMw8o&list=OLAK5uy_mLZErmQxkzqCbaHdrEEz2VmD8vXBaXCfc'
# download_youtube_playlist(playlist_url)


# from pytube import Playlist
#
# playlist = Playlist('https://www.youtube.com/watch?v=c2IBhnkMw8o&list=OLAK5uy_mLZErmQxkzqCbaHdrEEz2VmD8vXBaXCfc')
# print('Number of videos in playlist: %s' % len(playlist.video_urls))
# playlist.download_all()



#this code allows you to download a playlist to your assigned folder

# import re
# from pytube import Playlist
# playlist = Playlist('https://www.youtube.com/watch?v=c2IBhnkMw8o&list=OLAK5uy_mLZErmQxkzqCbaHdrEEz2VmD8vXBaXCfc')
# DOWNLOAD_DIR = 'D:\Video'
# playlist._video_regex = re.compile(r"\"url\":\"(/watch\?v=[\w-]*)")
# print(len(playlist.video_urls))
# for url in playlist.video_urls:
#     print(url)
# for video in playlist.videos:
#     print('downloading : {} with url : {}'.format(video.title, video.watch_url))
#     video.streams.\
#         filter(type='video', progressive=True, file_extension='mp4').\
#         order_by('resolution').\
#         desc().\
#         first().\
#         download(DOWNLOAD_DIR)


import os
import re
from pytube import Playlist

def download_youtube_playlists(playlists):
    DOWNLOAD_DIR = 'D:/Video'  # Change the directory path as per your preference

    for playlist_url in playlists:
        playlist = Playlist(playlist_url)
        playlist_name = playlist.title()

        # Create a folder for the playlist
        playlist_folder = os.path.join(DOWNLOAD_DIR, playlist_name)
        os.makedirs(playlist_folder, exist_ok=True)

        playlist._video_regex = re.compile(r"\"url\":\"(/watch\?v=[\w-]*)")
        print("Total videos in playlist '{}': {}".format(playlist_name, len(playlist.video_urls)))

        for video in playlist.videos:
            print('Downloading video: {} with URL: {}'.format(video.title, video.watch_url))

            # Download the video and save it in the playlist folder
            video.streams.filter(type='video', progressive=True, file_extension='mp4') \
                .order_by('resolution').desc().first() \
                .download(output_path=playlist_folder)

# Example usage
playlists = [
    'https://www.youtube.com/watch?v=OGKKAJalYqc&list=OLAK5uy_nX6ElRPrNZIOfpzyYpiLA8HrOWA37Bnp0',
    'https://www.youtube.com/watch?v=c2IBhnkMw8o&list=OLAK5uy_mLZErmQxkzqCbaHdrEEz2VmD8vXBaXCfc',
    "https://www.youtube.com/watch?v=seCoBkatV1Q&list=OLAK5uy_kwnPeM43nkPGrWQ4Ymk5vNq3_3B0cKun8",
    'https://www.youtube.com/watch?v=TeQWQ2WwPN0&list=OLAK5uy_mxCx4W00OqAVj_hIUnmhg2hBGVS9A4-Mw',
    'https://www.youtube.com/watch?v=6uAvh2K_rnk&list=OLAK5uy_k71xHs7p02IsiH-nQbmnMi94KSCLXXDzE',
    'https://www.youtube.com/watch?v=lzZv37M2IVk&list=OLAK5uy_nOifZnhmLWjSxl2Uz8wiijHx53DglSBII',
    'https://www.youtube.com/watch?v=rT5mKXz6iSw&list=OLAK5uy_n1oavrvTnF7EafGwwlbhDbScNZzxC7v-M',

]

download_youtube_playlists(playlists)
