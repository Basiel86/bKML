from pydub import AudioSegment


if __name__ == '__main__':
    sound = AudioSegment.from_mp3('Dendy.mp3')
    sound.export('Dendy_2.mp3', format="mp3", bitrate="128k")
