import moviepy.editor
from pathlib import Path

video_file = Path('C:/Users/Noizz/Desktop/learn/звук из видео/INFLAMES - I Am Above.mp4')
video = moviepy.editor.VideoFileClip(f'{video_file}')
audio = video.audio
audio.write_audiofile(f'C:/Users/Noizz/Desktop/learn/звук из видео/ + {video_file.stem}.mp3')
