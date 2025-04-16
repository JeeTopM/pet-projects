'''всё тоже самое + графический интерфейс (помощь ИИ с интерфейсом)'''
import json
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from pathlib import Path

class LibraryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Закупка: Запрос книг от читателей")
        self.root.geometry("1200x800")
        
        # Загрузка данных
        self.json_file = Path("books.json")
        self.books = self.load_books()
        
        # Создание интерфейса
        self.create_widgets()
        self.update_treeview()
    
    def load_books(self):
        """Загружает книги из JSON."""
        if self.json_file.exists():
            with open(self.json_file, "r", encoding="utf-8") as f:
                return json.load(f)
        return []
    
    def save_books(self):
        """Сохраняет книги в JSON."""
        with open(self.json_file, "w", encoding="utf-8") as f:
            json.dump(self.books, f, ensure_ascii=False, indent=4)
    
    def create_widgets(self):
        """Создает элементы интерфейса."""
        # Фрейм для ввода данных
        input_frame = ttk.Frame(self.root, padding="10")
        input_frame.pack(fill=tk.X)
        
        # Поля ввода
        ttk.Label(input_frame, text="Название книги:").grid(row=0, column=0, sticky=tk.W)
        self.title_entry = ttk.Entry(input_frame, width=40)
        self.title_entry.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=2)
        
        ttk.Label(input_frame, text="Автор:").grid(row=1, column=0, sticky=tk.W)
        self.author_entry = ttk.Entry(input_frame, width=40)
        self.author_entry.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=2)
        
        ttk.Label(input_frame, text="Номер билета:").grid(row=2, column=0, sticky=tk.W)
        self.ticket_entry = ttk.Entry(input_frame, width=40)
        self.ticket_entry.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=2)
        
        # Кнопки
        buttons_frame = ttk.Frame(self.root, padding="10")
        buttons_frame.pack(fill=tk.X)
        
        ttk.Button(
            buttons_frame,
            text="Добавить книгу",
            command=self.add_book
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            buttons_frame,
            text="Очистить поля ввода",
            command=self.clear_entries
        ).pack(side=tk.LEFT, padx=5)
        
        # Таблица для отображения книг
        self.tree = ttk.Treeview(
            self.root,
            columns=("title", "author", "ticket_num", "date_added"),
            show="headings",
            height=20
        )
        self.tree.heading("title", text="Название")
        self.tree.heading("author", text="Автор")
        self.tree.heading("ticket_num", text="Номер билета")
        self.tree.heading("date_added", text="Дата добавления")
        self.tree.pack(fill=tk.BOTH, expand=True)
    
    def add_book(self):
        """Добавляет книгу в список."""
        title = self.title_entry.get().strip()
        author = self.author_entry.get().strip()
        ticket_num = self.ticket_entry.get().strip()
        
        # Проверка на пустые поля
        if not title or not ticket_num:
            messagebox.showerror("Ошибка", "Все поля должны быть заполнены!")
            return
        
        # Проверка уникальности: автор + название
        if any(book["title"] == title and book["author"] == author for book in self.books):
            messagebox.showwarning("Внимание", "Эта книга уже есть в списке!")
            return
        
        # Добавление новой книги
        new_book = {
            "title": title,
            "author": author,
            "ticket_num": ticket_num,
            "date_added": datetime.now().strftime("%d.%m.%Y %H:%M")
        }
        self.books.append(new_book)
        self.save_books()
        self.clear_entries()
        self.update_treeview()
        messagebox.showinfo("Успех", "Книга успешно добавлена!")
    
    def clear_entries(self):
        """Очищает поля ввода."""
        self.title_entry.delete(0, tk.END)
        self.author_entry.delete(0, tk.END)
        self.ticket_entry.delete(0, tk.END)
    
    def update_treeview(self):
        """Обновляет таблицу с книгами."""
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        for book in self.books:
            self.tree.insert(
                "",
                tk.END,
                values=(book["title"], book["author"], book["ticket_num"], book["date_added"])
            )

if __name__ == "__main__":
    root = tk.Tk()
    app = LibraryApp(root)
    root.mainloop()