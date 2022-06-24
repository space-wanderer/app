Wanderer, [Чт 23.06.22 22:16]
from http.client import FOUND
from statistics import mode
from gtts import gTTS
from art import tprint
import pdfplumber
from pathlib import Path

def pdf_to_mp3(file_path='test.pdf', language='ru'):
    if Path(file_path).is_file() and Path(file_path).suffix == '.pdf':
        #return 'File exits!'

        print(f'- Original file: {Path(file_path).name}')
        print(f'- Processing...')

        with pdfplumber.PDF(open(file=file_path, mode='rb')) as pdf:
            pages = [page.extract_text() for page in pdf.pages]

        text = ''.join(pages) 
        """
        with open('text.txt', 'w') as file:
            file.write(text)
        """
        text = text.replace('\n', '')    
        """
        with open('text2.txt', 'w') as file:
            file.write(text)    
        """
        my_audio = gTTS(text=text, lang=language) 
        file_name = Path(file_path).stem
        my_audio.save(f'{file_name}.mp3')   

        return f'[+] {file_name}.mp3 saved successfylly!'
    else:
        return 'File not exits, check the file path'

def main():
    tprint('PDF_TO_MP3', font='bulbhead')
    file_path = input('/nВведите путь к файлу:')
    language = input('Вветиде язык: ru или en')
    print(pdf_to_mp3(file_path=file_path, lang = language))     

if __name__ == '__main__':
    main()