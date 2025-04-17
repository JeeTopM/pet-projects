from tkinter import *
from tkinter import messagebox
import locale
from datetime import date, timedelta, datetime


def calendar_date():
    try:
        locale.setlocale(locale.LC_ALL, "ru_RU.UTF-8")
        date_str = date_need_tf.get()
        d, m, y = [int(i) for i in date_str.split(".")]
        dt = date(y, m, d)

        data_in = datetime.strftime(dt, "%x")
        week_in = datetime.strftime(dt, "%A")

        while dt.weekday() != 0:
            dt -= timedelta(days=1)

        week_out = []
        for _ in range(7):
            week_out.append(datetime.strftime(dt, "%d %B %Y - %A"))
            dt += timedelta(days=1)

        message = f"Дата: {data_in}\nДень недели: {week_in}\n\nНеделя:\n" + "\n".join(
            week_out
        )
        messagebox.showinfo("Результат", message)

    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")


# ГУИ

window = Tk()
window.title("Какой был день?")
window.geometry("500x500")

frame = Frame(window, padx=10, pady=10)

frame.pack(expand=True)

# Названия пунктов:
# День
date_need_lf = Label(frame, text="Введите дату (дд.мм.гггг): ")
date_need_lf.grid(row=3, column=1)

# Названия окон ввода
# День
date_need_tf = Entry(frame)
date_need_tf.grid(row=3, column=2)

# Кнопка для вывода
res = Button(frame, text="Отчёт", command=calendar_date)
res.grid(row=6, column=2)

window.mainloop()
