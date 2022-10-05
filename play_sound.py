import os
os.environ['PYGAME_HIDE_SUPPORT_PROMPT'] = "hide"
from pygame import mixer
import random
from resource_path import resource_path


class PlaySound(object):
    def __init__(self):
        self.sounds_folder = resource_path(r'Sounds')
        self.mixer = mixer
        self.mixer.init()

    def random_file(self):
        return random.choice(os.listdir(self.sounds_folder))

    def play_music(self):
        random_track = os.path.join(self.sounds_folder, self.random_file())
        self.mixer.music.load(random_track)
        self.mixer.music.play()

    def stop_music(self):
        self.mixer.music.stop()

if __name__ == '__main__':
    s = PlaySound()

