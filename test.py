'''добавлено сохранение и загрузка из файла json'''
from datetime import datetime
from pathlib import Path
import json

class Library:
    def __init__(self, json_file="books.json"):
        self.json_file = Path(json_file)
        self.books_titles = self._load_books()

    def _load_books(self):
        if self.json_file.exists():
            with open(self.json_file, "r", encoding="utf-8") as f:
                return json.load(f)
        return []
    
    def _save_to_json(self):
        with open(self.json_file, "w", encoding="utf-8") as f:
            json.dump(self.books_titles, f, ensure_ascii=False, indent=3)
    
    def сheck_book(self, book_title):
        return any(book["title"] == book_title for book in self.books_titles)

    def add_book(self, book_title, ticket):
        if not self.сheck_book(book_title):
            book = {
                'title': book_title, 
                'ticket': ticket, 
                'dt': datetime.now().strftime("%d.%m.%Y %H:%M:%S")
                }
            self.books_titles.append(book)
            self._save_to_json()
            print('Книга внесена в базу')
        else:
            print('Книга уже в базе')

    def all_books(self, formatted=False):
        if not formatted:
            return self.books_titles
        books_formatted = (
            f"{i+1}. {book['title']} (билет: {book['ticket_num']}, добавлена: {book['date_added']})"
            for i, book in enumerate(self.books_titles)
            )
        result = "\n".join(books_formatted)
        return result

if __name__ == "__main__": # можно и убрать, ничего пока не меняется из за этой строчки
    library = Library()
    book_title = input('Введине название книги: ')
    ticket = input('Введите читательский билет: ')
    library.add_book(book_title, ticket)

    for book in library.all_books():
        print(book)