from datetime import datetime  # хочу, чтобы была история добавление книги

class Library:
    def __init__(self):
        self.books_titles = []
    
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
            print('Книга внесена в базу')
        else:
            print('Книга уже в базе')

    def all_books(self):
        return self.books_titles
    
if __name__ == "__main__":
    library = Library()
    book_title = input('Введине название книги: ')
    ticket = input('Введите читательский билет: ')
    library.add_book(book_title, ticket)

    for book in library.all_books():
        print(book)