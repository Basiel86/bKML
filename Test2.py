import gdown

output = "FR.rar"
g_path_compl = 'https://drive.google.com/file/d/1zEbwU3DzrKxLUU8ywP1ER4BZLL75v2-m/view?usp=sharing'
g_path_f = 'https://drive.google.com/drive/folders/1JvSwjtzCVPVaayF0yKpepwRK65wEuEnE?usp=sharing'

if __name__ == '__main__':
    gdown.download(g_path_f, output, fuzzy=True)
