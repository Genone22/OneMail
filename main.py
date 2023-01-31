import os
import re
import time
import ctypes
import requests
import webbrowser
from tkinter import *
from bs4 import BeautifulSoup
from tkinter import filedialog
from tkinter import messagebox
from PIL import ImageTk  # pip install pillow
import customtkinter  # pip install customtkinter
import pandas as pd  # pip install pandas
import email_function
from email_marketing_info import info_text
from center_tk_windows import CenterTk


# Modes: "System" (standard), "Dark", "Light"
customtkinter.set_appearance_mode("Dark")
# Themes: "blue" (standard), "green", "dark-blue"
customtkinter.set_default_color_theme("blue")

# This following setting in the ctypes library sets “DPI” awareness.
# DPI stands for Dots per inch, another way of measuring screen resolution
ctypes.windll.shcore.SetProcessDpiAwareness(1)

# Splash screen
w = CenterTk()
# for hiding titlebar
w.overrideredirect(True)

Frame(w, width=427, height=250, bg='#272727').place(x=0, y=0)
label1 = Label(w, text='OneMail', fg='silver', bg='#272727')

# Decorate it
label1.configure(font=("Lab Grotesque", 33, "bold"))
label1.place(x=100, y=70)

label2 = Label(w, text='Загрузка...', fg='silver', bg='#272727')  # decorate it
label2.configure(font=("Roboto", 9))
label2.place(x=10, y=225)

# Making animation
image_a = ImageTk.PhotoImage(file='images/c2.png')
image_b = ImageTk.PhotoImage(file='images/c1.png')

# 2loops
for i in range(2):
    Label(w, image=image_a, border=0, relief=SUNKEN).place(x=180, y=145)
    Label(w, image=image_b, border=0, relief=SUNKEN).place(x=200, y=145)
    Label(w, image=image_b, border=0, relief=SUNKEN).place(x=220, y=145)
    Label(w, image=image_b, border=0, relief=SUNKEN).place(x=240, y=145)
    w.update_idletasks()
    time.sleep(0.2)

    Label(w, image=image_b, border=0, relief=SUNKEN).place(x=180, y=145)
    Label(w, image=image_a, border=0, relief=SUNKEN).place(x=200, y=145)
    Label(w, image=image_b, border=0, relief=SUNKEN).place(x=220, y=145)
    Label(w, image=image_b, border=0, relief=SUNKEN).place(x=240, y=145)
    w.update_idletasks()
    time.sleep(0.2)

    Label(w, image=image_b, border=0, relief=SUNKEN).place(x=180, y=145)
    Label(w, image=image_b, border=0, relief=SUNKEN).place(x=200, y=145)
    Label(w, image=image_a, border=0, relief=SUNKEN).place(x=220, y=145)
    Label(w, image=image_b, border=0, relief=SUNKEN).place(x=240, y=145)
    w.update_idletasks()
    time.sleep(0.2)

    Label(w, image=image_b, border=0, relief=SUNKEN).place(x=180, y=145)
    Label(w, image=image_b, border=0, relief=SUNKEN).place(x=200, y=145)
    Label(w, image=image_b, border=0, relief=SUNKEN).place(x=220, y=145)
    Label(w, image=image_a, border=0, relief=SUNKEN).place(x=240, y=145)
    w.update_idletasks()
    time.sleep(0.2)


# God Object
class BulkEmail(customtkinter.CTk):
    def __init__(self, root):
        self.info = info_text
        self.root = root
        self.root.iconbitmap('images/modern15.ico')
        # configure window
        self.root.title('OneMail')
        self.root.geometry('760x600+490+150')
        self.root.resizable(False, False)
        # Icons&Images
        self.email_icon = ImageTk.PhotoImage(file='images/modern10.png')
        self.setting_icon = ImageTk.PhotoImage(file='images/icons8-menu-60.png')
        self.browse_icon = ImageTk.PhotoImage(file='images/import35.png')
        self.clear_icon = ImageTk.PhotoImage(file='images/clear35.png')
        self.question_icon = ImageTk.PhotoImage(file='images/question-mark.png')
        self.html_view_icon = ImageTk.PhotoImage(file='images/html-35.png')
        self.scrap_icon = ImageTk.PhotoImage(file='images/scraper-35.png')
        # Title
        title = Label(self.root,
                      text='Делайте Email расcылки быстро и легко',
                      image=self.email_icon,
                      compound=LEFT,
                      font=('Lab Grotesque', 20),
                      bg='#0B4283',
                      fg='#242424',
                      anchor='w')
        title.place(x=0,
                    y=0,
                    relwidth=1)

        # Radio buttons
        self.var_choice = StringVar()
        single = customtkinter.CTkRadioButton(self.root,
                                              text='Один адресат',
                                              value='single',
                                              variable=self.var_choice,
                                              command=self.check_single_or_bulk,
                                              font=('Lab Grotesque', 20)
                                              )
        single.place(x=10,
                     y=110)

        bulk = customtkinter.CTkRadioButton(self.root,
                                            text='Рассылка',
                                            value='bulk',
                                            variable=self.var_choice,
                                            command=self.check_single_or_bulk,
                                            font=('Lab Grotesque', 20)
                                            )
        bulk.place(x=195,
                   y=110)
        self.var_choice.set('single')

        # Labels
        to = customtkinter.CTkLabel(self.root,
                                    text='Кому',
                                    font=('Lab Grotesque', 15))
        to.place(x=20,
                 y=202)

        subj = customtkinter.CTkLabel(self.root,
                                      text='Тема',
                                      font=('Lab Grotesque', 15))
        subj.place(x=20,
                   y=252)

        msg = customtkinter.CTkLabel(self.root,
                                     text='Сообщение',
                                     font=('Lab Grotesque', 15))
        msg.place(x=20,
                  y=300)

        # Entries
        self.txt_to = customtkinter.CTkEntry(self.root,
                                             font=('Lab Grotesque', 17),
                                             width=350,
                                             height=35,
                                             border_width=1,
                                             corner_radius=10
                                             )
        self.txt_to.place(x=150,
                          y=205)


        self.txt_subj = customtkinter.CTkEntry(self.root,
                                               font=('Lab Grotesque', 17),
                                               width=350,
                                               height=35,
                                               border_width=1,
                                               corner_radius=10
                                               )
        self.txt_subj.place(x=150,
                            y=255)

        self.txt_msg = customtkinter.CTkTextbox(self.root,
                                                font=('Lab Grotesque', 17),
                                                scrollbar_button_hover_color=
                                                '#0B4283',
                                                border_width=2,
                                                corner_radius=10,
                                                width=580,
                                                height=220)
        self.txt_msg.place(x=150,
                           y=305)

        # Status
        self.lbl_total = customtkinter.CTkLabel(self.root,
                                                text='',
                                                font=('Lab Grotesque', 14))

        self.lbl_total.place(x=20,
                             y=575)

        self.lbl_sent = customtkinter.CTkLabel(self.root,
                                               text='',
                                               font=('Lab Grotesque', 14))
        self.lbl_sent.place(x=150,
                            y=575)

        self.lbl_failed = customtkinter.CTkLabel(self.root,
                                                 text='',
                                                 text_color='#FF2400',
                                                 font=('Lab Grotesque', 14))
        self.lbl_failed.place(x=300,
                              y=575)

        # Buttons
        btn_settings = Button(self.root,
                              image=self.setting_icon,
                              bd=0,
                              bg='#0B4283',
                              activebackground='#0B4283',
                              cursor='hand2',
                              command=self.setting_window)
        btn_settings.place(x=895)

        btn_info = Button(self.root,
                          command=self.info_window,
                          image=self.question_icon,
                          bd=0,
                          bg='#242424',
                          activebackground='#242424',
                          cursor='hand2')
        btn_info.place(x=390,
                       y=130)


        btn_send = customtkinter.CTkButton(self.root,
                                           text='Отправить',
                                           font=('Lab Grotesque', 13, 'bold'),
                                           cursor='hand2',
                                           command=self.send_email)
        btn_send.place(x=610,
                       y=540,
                       width=150,
                       height=50)
        self.check_file_exist()


        self.btn_browse = Button(self.root,
                                 command=self.browse_file,
                                 image=self.browse_icon,
                                 bd=0,
                                 bg='#242424',
                                 activebackground='#242424',
                                 state=DISABLED,
                                 cursor='hand2')
        self.btn_browse.place(x=692,
                              y=682)


        btn_clear = Button(self.root,
                           command=self.broom,
                           image=self.clear_icon,
                           bd=0,
                           bg='#242424',
                           activebackground='#242424',
                           cursor='hand2')
        btn_clear.place(x=632,
                        y=682)


        btn_html_view = Button(self.root,
                               command=self.show_html,
                               image=self.html_view_icon,
                               bd=0,
                               bg='#242424',
                               activebackground='#242424',
                               cursor='hand2')
        btn_html_view.place(x=572,
                            y=682)


        btn_scrape_window = Button(self.root,
                                   command=self.show_scrap_window,
                                   image=self.scrap_icon,
                                   bd=0,
                                   bg='#242424',
                                   activebackground='#242424',
                                   cursor='hand2')
        btn_scrape_window.place(x=512,
                                y=682)


    def show_scrap_window(self):
        scraper = customtkinter.CTkToplevel()
        scraper.title("Веб скрапер")
        scraper.geometry('760x600+490+150')
        scraper.resizable(False, False)
        scraper.iconbitmap('images/modern15.ico')

        # Title
        url_title = Label(scraper,
                          text='Вставьте URL для извлечения email-адресов:',
                          compound=LEFT,
                          font=('Lab Grotesque', 12),
                          bg='#242424',
                          fg='silver',
                          anchor='w')
        url_title.place(x=3,
                        y=3,
                        relwidth=1)

        self.url_text = customtkinter.CTkTextbox(scraper,
                                                 corner_radius=10,
                                                 width=755,
                                                 height=240)
        self.url_text.place(x=3, y=33)

        self.scrape_button = customtkinter.CTkButton(scraper,
                                                     text="Извлечь",
                                                     command=self.scrape_website,
                                                     font=('Lab Grotesque', 13,
                                                           'bold'),
                                                     cursor='hand2'
                                                     )
        self.scrape_button.place(x=240, y=290, width=150, height=50)

        self.save_button = customtkinter.CTkButton(scraper, text="Сохранить",
                                                   command=self.save_to_excel,
                                                   font=(
                                                       'Lab Grotesque', 13,
                                                       'bold'),
                                                   cursor='hand2'
                                                   )
        self.save_button.place(x=400, y=290, width=150, height=50)

        self.emails_text = customtkinter.CTkTextbox(scraper,
                                                    corner_radius=10,
                                                    width=755,
                                                    height=245)
        self.emails_text.place(x=3, y=350)

    def scrape_website(self):
        input_text = self.url_text.get("1.0", 'end-1c')
        urls = re.findall(r'https?://(?:[-\w.]|(?:%[\da-fA-F]{2}))+',
                          input_text)
        emails = set()
        for url in urls:
            try:
                response = requests.get(url)
                soup = BeautifulSoup(response.text, 'html.parser')
                text = soup.get_text()
                email_list = re.findall(r'[\w\.-]+@[\w\.-]+', text)
                if not email_list:
                    raise ValueError("No email addresses found on this website")
                emails.update(email_list)
            except requests.exceptions.RequestException as e:
                print(f"Error connecting to {url}: {e}")
            except ValueError as e:
                print(f"{url}: {e}")
        self.emails_text.delete("1.0", 'end')
        for email in emails:
            self.emails_text.insert('end', email + '\n')

    def save_to_excel(self):
        # Use the filedialog to ask the user to select a file
        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                filetypes=[
                                                    ("Excel files", "*.xlsx")])
        if filepath:
            # Get the email addresses from the Text widget
            emails = self.emails_text.get("1.0", 'end-1c').strip().split('\n')
            # Create a DataFrame with a single column "email"
            df = pd.DataFrame(emails, columns=["Email"])
            # Write the DataFrame to the selected Excel file
            df.to_excel(filepath, index=False)


    def show_html(self):
        filepath = filedialog.askopenfilename(
            filetypes=[("HTML files", "*.html"), ("All files", "*.*")])
        if filepath:
            with open(filepath, 'r', encoding='utf-8') as f:
                html_text = f.read()
            webbrowser.open(filepath)
            self.txt_msg.insert('1.0', html_text)
        else:
            messagebox.showinfo("Отмена", "HTML файл не выбран")

    def info_window(self):
        self.root_3 = customtkinter.CTkToplevel()
        self.root_3.title('Информация')
        self.root_3.geometry('760x600+490+150')
        self.root_3.resizable(False, False)
        self.root_3.iconbitmap('images/modern15.ico')
        
        

        # Title_3
        title_info = Label(self.root_3,
                           text='Как не попасть в спам',
                           padx=10,
                           compound=CENTER,
                           font=('Lab Grotesque', 20, 'bold'),
                           bg='#0B4283',
                           fg='#242424',
                           anchor='n')
        title_info.place(x=0,
                         y=0,
                         relwidth=1)

        self.info = customtkinter.CTkTextbox(self.root_3,
                                             border_width=0,
                                             corner_radius=0,
                                             width=760,
                                             height=570)
        self.info.place(x=5,
                        y=35)

        self.info.insert("0.0", info_text)

        self.info.configure(state=DISABLED,
                            text_color='silver',
                            fg_color='#242424',
                            font=('Lab Grotesque', 18),
                            scrollbar_button_hover_color='#0B4283')

    def browse_file(self):
        op = filedialog.askopenfile(initialdir='/',
                                    title='Выбор файла Excel',
                                    filetypes=(('All files', '*.*'),
                                               ('Excel files', '.xlsx')))
        if op is not None:
            data = pd.read_excel(op.name)

            if 'Email' in data.columns:
                self.emails = list(data['Email'])
                c = []
                for i in self.emails:
                    if pd.isnull(i) == False:
                        c.append(i)
                self.emails = c
                if len(self.emails) > 0:
                    self.txt_to.configure(state=NORMAL)
                    self.txt_to.delete(0, 'end')
                    self.txt_to.insert(0, str(op.name.split('/')[-1]))
                    self.txt_to.configure(state='readonly')
                    self.lbl_total.configure(text='Всего:   ' + str(len(
                        self.emails)))
                    self.lbl_sent.configure(text='Отправлено:   ')
                    # self.lbl_left.configure(text='Пропало: ')
                    self.lbl_failed.configure(text='Не отправлено:   ')
                else:
                    messagebox.showerror('Ошибка!', 'Выбранный файл не '
                                                    'содержит Email адресов',
                                         parent=self.root)

            else:
                messagebox.showerror('Ошибка!', 'Пожалуйста, выберите файл '
                                                'содержащий столбец "Email"',
                                     parent=self.root)

    def send_email(self):
        x = len(self.txt_msg.get(1.0, 'end'))
        if self.txt_to.get() == '' or self.txt_subj.get() == '' or x == 1:
            messagebox.showerror('Ошибка', 'Пожалуйста, заполните все поля',
                                 parent=self.root)
        else:
            if self.var_choice.get() == 'single':
                name = "InterSert"
                status = email_function.emails_send_funct(self.txt_to.get(),
                                                          self.txt_subj.get(),
                                                          self.txt_msg.get(1.0,
                                                                           'end'
                                                                           ),
                                                          self.from_,
                                                          name)
                if status == 's':
                    messagebox.showinfo('Готово',
                                        'Электронное письмо успешно отправлено',
                                        parent=self.root)
                if status == 'f':
                    messagebox.showerror('Ошибка'
                                         'Электронное письмо не отправлено',
                                         parent=self.root)

            if self.var_choice.get() == 'bulk':
                name = "InterSert"
                self.failed = []
                self.s_count = 0
                self.f_count = 0
                for x in self.emails:
                    status = email_function.emails_send_funct(x,
                                                              self.txt_subj.get
                                                              (),
                                                              self.txt_msg.get
                                                              (1.0, 'end'),
                                                              self.from_,
                                                              name)
                    if status == 's':
                        self.s_count += 1
                    if status == 'f':
                        self.f_count += 1
                    self.status_bar()
                    time.sleep(1)

                messagebox.showinfo('Готово',
                                    'Электронные письма успешно отправлены.',
                                    parent=self.root)

    def status_bar(self):
        self.lbl_total.configure(text='Статус: ' + str(len(self.emails)) +
                                      '=>')
        self.lbl_sent.configure(text='Отправлено: ' + str(self.s_count))
        # self.lbl_left.configure(text='Пропало: ' + str(len(self.emails) -
        #                                                (self.s_count +
        #                                                 self.f_count)))
        self.lbl_failed.configure(text='Не отправлено: ' + str(self.f_count))
        self.lbl_total.update()
        self.lbl_sent.update()
        # self.lbl_left.update()
        self.lbl_failed.update()

    def check_single_or_bulk(self):
        if self.var_choice.get() == 'single':
            self.btn_browse.configure(state=DISABLED)
            self.txt_to.configure(state=NORMAL)
            self.txt_to.delete(0, 'end')
            self.broom()
        if self.var_choice.get() == 'bulk':
            self.btn_browse.configure(state=NORMAL)
            self.txt_to.delete(0, 'end')
            self.txt_to.configure(state='readonly')

    def broom(self):
        self.txt_to.configure(state=NORMAL)
        self.txt_to.delete(0, 'end')
        self.txt_subj.delete(0, 'end')
        self.txt_msg.delete(1.0, 'end')
        self.var_choice.set('single')
        self.btn_browse.config(state=DISABLED)
        self.lbl_total.configure(text=' ')
        self.lbl_sent.configure(text=' ')
        # self.lbl_left.configure(text=' ')
        self.lbl_failed.configure(text=' ')

    def setting_window(self):
        self.check_file_exist()
        self.root_2 = customtkinter.CTkToplevel()
        self.root_2.title('Учётная запись')
        self.root_2.geometry('350x200+740+400')
        self.root_2.resizable(False, False)
        self.root_2.focus_force()
        self.root_2.grab_set()
        self.root_2.iconbitmap('images/modern15.ico')

        # Title_2
        title_2 = Label(self.root_2,
                        text='Введите вашу почту и пароль',
                        padx=10,
                        compound=LEFT,
                        font=('Lab Grotesque', 15, 'bold'),
                        bg='#0B4283',
                        fg='#242424',
                        anchor='n')
        title_2.place(x=0,
                      y=0,
                      relwidth=1)

        # Settings Entries
        self.txt_from = customtkinter.CTkEntry(self.root_2,
                                               font=('Lab Grotesque', 17),
                                               placeholder_text="эл. почта",
                                               width=310,
                                               height=35,
                                               border_width=1,
                                               corner_radius=10
                                               )
        self.txt_from.place(x=25,
                            y=55)

        self.txt_pass = customtkinter.CTkEntry(self.root_2,
                                               font=('Lab Grotesque', 17),
                                               placeholder_text="пароль",
                                               width=310,
                                               height=35,
                                               border_width=1,
                                               corner_radius=10,
                                               show='*'
                                               )
        self.txt_pass.place(x=25,
                            y=105)
        # root_2 save button
        btn_save = customtkinter.CTkButton(self.root_2,
                                           command=self.save_setting,
                                           text='Сохранить',
                                           font=('Lab Grotesque', 13, 'bold'),
                                           cursor='hand2')
        btn_save.place(x=122,
                       y=155,
                       width=150,
                       height=40)

        self.txt_from.insert(0, self.from_)
        self.txt_pass.insert(0, self.pass_)


        self.var_show_password = StringVar()
        btn_chkbx = customtkinter.CTkCheckBox(self.root_2,
                                              onvalue='Показать',
                                              text='Показать',
                                              variable=self.var_show_password,
                                              cursor='hand2',
                                              font=('Lab Grotesque', 12),
                                              command=self.show_password)
        btn_chkbx.place(x=25,
                        y=145)
        btn_chkbx.configure(width=10,
                            border_width=1,
                            corner_radius=10,
                            checkbox_height=15,
                            checkbox_width=15)

    def show_password(self):
        if self.var_show_password.get() == 'Показать':
           self.txt_pass.configure(show='')
        else:
            self.txt_pass.configure(show='*')

    def check_file_exist(self):
        if os.path.exists('important.txt') == False:
            f = open('important.txt', 'w')
            f.write(',')
            f.close()
        f2 = open('important.txt', 'r')
        self.credentials = []
        for i in f2:
            self.credentials.append([i.split(',')[0], i.split(',')[1]])
        # print(self.credentials)
        self.from_ = self.credentials[0][0]
        self.pass_ = self.credentials[0][1]
        print(self.from_, self.pass_)

    def save_setting(self):
        if self.txt_from.get() == '' or self.txt_pass.get() == '':
            messagebox.showerror('Ошибка', 'Пожалуйста, заполните все поля',
                                 parent=self.root_2)
        else:
            f = open('important.txt', 'w')
            f.write(self.txt_from.get() + ',' + self.txt_pass.get())
            f.close()
            messagebox.showinfo('Готово', 'Почта успешно сохранена',
                                parent=self.root_2)
            self.check_file_exist()


# Cut / Paste / Copy function
def on_key_release(event):
    ctrl = (event.state & 0x4) != 0
    if event.keycode == 88 and ctrl and event.keysym.lower() != "x":
        event.widget.event_generate("<<Cut>>")

    if event.keycode == 86 and ctrl and event.keysym.lower() != "v":
        event.widget.event_generate("<<Paste>>")

    if event.keycode == 67 and ctrl and event.keysym.lower() != "c":
        event.widget.event_generate("<<Copy>>")

    if event.keycode == 65 and ctrl and event.keysym.lower() != "a":
        event.widget.select_range(0, 'end')


w.destroy()
# Create CTk window
root = customtkinter.CTk()

root.bind_all("<Key>", on_key_release, "+")

root.bind('<Escape>', lambda event=None: root.destroy())

obj = BulkEmail(root)

root.mainloop()
