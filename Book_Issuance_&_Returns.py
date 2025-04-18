'''Учет выдачи и принятии книг.
Что нужно сделать: программа должна сохранять историю по выдаче и возврату книг посетителей, удалять записи'''
from datetime import datetime
from pathlib import Path
import json

class BookLog:
    def __init__(self, json_file='story_books.json'):
        self.json_file = Path(json_file)
        self.books_titles = self._load_json()
         
    def load_json():
        '''Загрузка из БД'''
        pass
    def save_json():
        '''Сохранение в БД'''
        pass
    def book_history():
        '''История книги: взята или выдана, кому, когда'''
        pass
    def all_books():
        '''Вывод информации'''
        pass