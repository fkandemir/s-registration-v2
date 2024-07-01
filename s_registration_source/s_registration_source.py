import json
import os
from tkinter import ttk
from tkinter import *
import customtkinter
from PIL import Image
from datetime import datetime
import xlsxwriter
import sqlite3

global state
state = True


class Utilization:

    version = "v1.9 ios edition"

    def __init__(self):
        pass

    def getAttributes(self):
        if customtkinter.get_appearance_mode() == "Light":
            return "#DBDBDB", "black"
        elif customtkinter.get_appearance_mode() == "Dark":
            return "#2B2B2B", "lightgray"

    def dayFinder(self):
        day = datetime.today().weekday()

        if day == 0: day = "Pazartesi"
        elif day == 1: day = "Salı"
        elif day == 2: day = "Çarşamba"
        elif day == 3: day = "Perşembe"
        elif day == 4: day = "Cuma"
        elif day == 5: day = "Cumartesi"
        elif day == 6: day = "Pazar"
        return day.upper()

    def monthFinder(self):
        month = datetime.today().strftime("%m")

        if str(month) == "01": month = "Ocak"
        elif str(month) == "02": month = "Şubat"
        elif str(month) == "03": month = "Mart"
        elif str(month) == "04": month = "Nisan"
        elif str(month) == "05": month = "Mayıs"
        elif str(month) == "06": month = "Haziran"
        elif str(month) == "07": month = "Temmuz"
        elif str(month) == "08": month = "Ağustos"
        elif str(month) == "09": month = "Eylül"
        elif str(month) == "10": month = "Ekim"
        elif str(month) == "11": month = "Kasım"
        elif str(month) == "12": month = "Aralık"
        else: month = None
        return month

    def yearFinder(self):
        year = datetime.today().strftime("%Y")
        return year

    def ensure_directory_exists(self, path):
        if not os.path.exists(path):
            os.makedirs(path)


class Repo:
    UTILIZATION_OBJ = Utilization()
    registrations_json_file = "./json_files/" + datetime.today().strftime("%d.%m.%Y") + ".json"

    def __init__(self):
        pass

    def get_registrations_json_file(self):
        year = self.UTILIZATION_OBJ.yearFinder()
        month = self.UTILIZATION_OBJ.monthFinder()

        if os.path.isdir("./json_files/" + year): pass
        else: os.mkdir("./json_files/" + year)

        if os.path.isdir("./json_files/" + year + "/" + month): pass
        else: os.mkdir("./json_files/" + year + "/" + month)

        file_path = "./json_files/" + year + "/" + month + "/" + datetime.today().strftime("%d.%m.%Y") + ".json"
        if os.path.isfile(file_path):
            return file_path
        else:
            with open("./json_files/template_sales_rep.json", "r", encoding="utf-8") as file:
                data = json.load(file)
            with open(file_path, "w", encoding="utf-8") as file:
                json.dump(data, file, indent=4)
            return file_path

    def get_template_sales_rep_array(self):
        with open("./txt_files/template_sales_rep.txt", "r", encoding="utf-8") as file:
            string = file.readline()
            sales_rep_array = string.split("-")
        return sales_rep_array

    def get_sales_rep_array(self):
        arr = self.get_template_sales_rep_array()
        year = self.UTILIZATION_OBJ.yearFinder()
        month = self.UTILIZATION_OBJ.monthFinder()

        if os.path.isdir("./txt_files/" + year): pass
        else: os.mkdir("./txt_files/" + year)

        if os.path.isdir("./txt_files/" + year + "/" + month): pass
        else: os.mkdir("./txt_files/" + year + "/" + month)

        file_path = "./txt_files/" + year + "/" + month + "/" + datetime.today().strftime("%d.%m.%Y") + ".txt"
        if os.path.isfile(file_path):
            with open(file_path, "r", encoding="utf-8") as file:
                string = file.readline()
                sales_rep_array = string.split("-")
            return sales_rep_array
        else:
            with open(file_path, "w", encoding="utf-8") as file:
                for i in range(len(arr)):
                    file.write(arr[i])
                    if i == len(arr) - 1: pass
                    else: file.write("-")
            return arr

    def addSalesRepresentative(self, sales_rep):
        year = self.UTILIZATION_OBJ.yearFinder()
        month = self.UTILIZATION_OBJ.monthFinder()
        file_path_txt = "./txt_files/" + year + "/" + month + "/" + datetime.today().strftime("%d.%m.%Y") + ".txt"
        with open(file_path_txt, "a", encoding="utf-8") as file:
            file.write("-" + sales_rep)

        json_file = self.get_registrations_json_file()
        with open(json_file, "r", encoding="utf-8") as file:
            data = json.load(file)

        with open(json_file, "w", encoding="utf-8") as file:
            data[sales_rep] = []
            json.dump(data, file, indent=4)

    def removeSalesRepresentative(self, sales_rep):
        old_array = self.get_sales_rep_array()
        year = self.UTILIZATION_OBJ.yearFinder()
        month = self.UTILIZATION_OBJ.monthFinder()

        for i in range(len(old_array)):
            if old_array[i] == sales_rep:
                old_array.pop(i)
                break

        file_path_txt = "./txt_files/" + year + "/" + month + "/" + datetime.today().strftime("%d.%m.%Y") + ".txt"
        with open(file_path_txt, "w", encoding="utf-8") as file:
            for i in range(len(old_array)):
                file.write(old_array[i])
                if i == len(old_array) - 1: pass
                else: file.write("-")

        json_file = self.get_registrations_json_file()
        with open(json_file, "r", encoding="utf-8") as file:
            data = json.load(file)

        with open(json_file, "w", encoding="utf-8") as file:
            dict = {}
            for item in old_array:
                if item == sales_rep: pass
                else: dict[item] = data[item]
            json.dump(dict, file, indent=4)


class MainUI:
    REPO_OBJ = Repo()
    UTILIZATION_OBJ = Utilization()
    window = customtkinter.CTk()

    side_bar_frame = customtkinter.CTkFrame(window, height=780, width=170, corner_radius=0)
    side_bar_frame.place(x=0, y=0)
    scrollable_frame = customtkinter.CTkScrollableFrame(window, height=750, width=1225, orientation="vertical")

    def __init__(self):
        customtkinter.set_appearance_mode("Light")
        customtkinter.set_default_color_theme("dark-blue")  # Themes: "blue" (standard), "green", "dark-blue"
        self.window.title("Firma Adı Günlük Kayıt Sistemi")
        size = "{}x{}+0+0"
        window_height = int(self.window.winfo_screenheight()*0.9//1)
        window_width = int(self.window.winfo_screenwidth())

        self.window.geometry(size.format(window_width, window_height))
        self.side_bar_frame.configure(height=window_height*0.959, width=window_width*0.114)
        self.scrollable_frame.configure(height=window_height*0.93, width=window_width*0.864)

        self.scrollable_frame.place(x=(window_width*0.114)+5, y=10)

    def _initializeButtons(self):

        new_reg_button = customtkinter.CTkButton(self.side_bar_frame, height=30, width=120, corner_radius=10,
                                                 text="Yeni Kayıt", command=self.new_registration)
        new_reg_button.place(x=23, y=70)

        amend_reg_button = customtkinter.CTkButton(self.side_bar_frame, height=30, width=120, corner_radius=10,
                                                   text="Kayıt Düzenle", command=self.update_registration)
        amend_reg_button.place(x=23, y=120)

        delete_reg_button = customtkinter.CTkButton(self.side_bar_frame, height=30, width=120, corner_radius=10,
                                                    text="Kayıt Sil", command=self.delete_registration)
        delete_reg_button.place(x=23, y=170)

        change_reg_button = customtkinter.CTkButton(self.side_bar_frame, height=30, width=120, corner_radius=10,
                                                    text="Kayıt Değiştir", command=self.change_registration)
        change_reg_button.place(x=23, y=220)

        add_rep_button = customtkinter.CTkButton(self.side_bar_frame, height=30, width=120, corner_radius=10,
                                                 text="Temsilci Ekle", command=self.addRepresentative)
        add_rep_button.place(x=23, y=295)

        delete_rep_button = customtkinter.CTkButton(self.side_bar_frame, height=30, width=120, corner_radius=10,
                                                    text="Temsilci Sil", command=self.removeRepresentative)
        delete_rep_button.place(x=23, y=335)

        create_report_button = customtkinter.CTkButton(self.side_bar_frame, height=30, width=120, corner_radius=10,
                                                       text="Rapor Oluştur", command=self.create_report)
        create_report_button.place(x=23, y=410)

        # theme_switch = customtkinter.CTkSwitch(self.side_bar_frame, text="", switch_width=80, switch_height=25,
        #                                        button_color="#3a7ebf", button_hover_color="#325882",
        #                                        command=lambda: self.change_appearance_mode(theme_switch.get()))
        # theme_switch.place(x=70, y=491)

    def _initializeLables(self):
        logo_label = customtkinter.CTkLabel(self.side_bar_frame, text="S&S Motors",
                                            font=customtkinter.CTkFont(size=25, weight="bold"))
        logo_label.place(x=13, y=20)

        # theme_label = customtkinter.CTkLabel(self.side_bar_frame, height=45, text="Tema:")
        # theme_label.configure(font=("Avenir", 17, "bold"))
        # theme_label.place(x=17, y=480)

        currenth = self.side_bar_frame.winfo_height()
        img = customtkinter.CTkImage(Image.open("image_file_name"), size=(144, 80))

        image_label = customtkinter.CTkLabel(self.side_bar_frame, text="", image=img)
        image_label.place(x=10, y=currenth - 165)

        copyright_label = customtkinter.CTkLabel(self.side_bar_frame,
                                                 text="Lisans Sahini Firma Adı" + self.UTILIZATION_OBJ.version,
                                                 font=customtkinter.CTkFont(family="Times New Roman TUR", size=10,
                                                                            weight="bold"))
        copyright_label.place(x=35, y=currenth - 70)

    def startApp(self):
        REPO_OBJ = Repo()
        self._initializeLables()
        self._initializeButtons()
        Registration().list_registrations(REPO_OBJ.get_registrations_json_file())
        # print("screen height: "+ str(self.window.winfo_screenheight()))
        # print("screen width: "+ str(self.window.winfo_screenwidth()))
        # print("---------------------------------------------------------")
        # print("window height: " + str(self.window.winfo_height()))
        # print("window width: " + str(self.window.winfo_width()))
        # print("---------------------------------------------------------")
        # print("sidebar height: "+ str(self.side_bar_frame.winfo_height()))
        # print("sidebar width: "+ str(self.side_bar_frame.winfo_width()))
        # print("---------------------------------------------------------")
        # print("scrollable frame height: "+ str(self.scrollable_frame.winfo_height()))
        # print("scrollable frame width: "+ str(self.scrollable_frame.winfo_width()))

    def closeApp(self):
        self.window.mainloop()

    # def change_appearance_mode(self, new_appearance_mode):
    #     REG_OBJ = Registration()
    #     if new_appearance_mode == 1: new_appearance_mode="Dark"
    #     elif new_appearance_mode == 0: new_appearance_mode = "Light"
    #     customtkinter.set_appearance_mode(new_appearance_mode)
    #     self.clear_frame(self.scrollable_frame)
    #     REG_OBJ.list_registrations()

    def clear_frame(self, frame):
        for widget in frame.winfo_children():
            widget.destroy()

    def close_window(self, window, time=100):
        window.after(time, lambda: window.destroy())

    def addRepresentative(self):
        REPO_OBJ = Repo()
        REGISTRATION_OBJ = Registration()

        def submit():
            if sales_rep_entry.get().find("-") != -1 or sales_rep_entry.get() == "":
                alert = customtkinter.CTkLabel(add_rep_frame, text="Temsilci İsmi (-) Karakteri İçeremez ve Boş Olamaz!",
                                               font=customtkinter.CTkFont(size=15), text_color="red")
                alert.place(x=70, y=80)
                alert.after(2000, lambda: alert.destroy())
                return
            else:
                REPO_OBJ.addSalesRepresentative(sales_rep_entry.get())
                self.clear_frame(self.scrollable_frame)
                REGISTRATION_OBJ.list_registrations(self.REPO_OBJ.get_registrations_json_file())
                self.close_window(add_rep_window)

        add_rep_window = customtkinter.CTk()
        add_rep_window.title("Yeni Temsilci Ekle")
        add_rep_window.geometry("500x200+0+0")

        add_rep_frame = customtkinter.CTkFrame(add_rep_window, height=180, width=480, corner_radius=0)
        add_rep_frame.place(x=10, y=10)

        sales_rep_label = customtkinter.CTkLabel(add_rep_frame, text="Temsilci:", font=customtkinter.CTkFont(size=15))
        sales_rep_label.place(x=100, y=50)
        sales_rep_entry = customtkinter.CTkEntry(add_rep_frame, width=200, font=customtkinter.CTkFont(size=15))
        sales_rep_entry.place(x=170, y=50)

        submit_button = customtkinter.CTkButton(add_rep_frame, height=30, width=137, corner_radius=10, text="Ekle",
                                                font=customtkinter.CTkFont(size=15), command=submit)
        submit_button.place(x=90, y=110)

        cancel_button = customtkinter.CTkButton(add_rep_frame, height=30, width=137, corner_radius=10, text="İptal",
                                                font=customtkinter.CTkFont(size=15), command=lambda: self.close_window(add_rep_window))
        cancel_button.place(x=237, y=110)

        add_rep_window.mainloop()

    def removeRepresentative(self):
        REPO_OBJ = Repo()
        REGISTRATION_OBJ = Registration()

        def submit():
            if sales_rep_opt_menu.get() == "Temsilci Seç":
                alert = customtkinter.CTkLabel(remove_rep_frame, text="Temsilci Seçilmedi!",
                                               font=customtkinter.CTkFont(size=15), text_color="red")
                alert.place(x=30, y=50)
                alert.after(2000, lambda: alert.destroy())
                return
            else:
                REPO_OBJ.removeSalesRepresentative(sales_rep_opt_menu.get())
                self.clear_frame(self.scrollable_frame)
                REGISTRATION_OBJ.list_registrations(self.REPO_OBJ.get_registrations_json_file())
                self.close_window(remove_rep_window)

        remove_rep_window = customtkinter.CTk()
        remove_rep_window.title("Temsilci Sil")
        remove_rep_window.geometry("500x200+0+0")

        remove_rep_frame = customtkinter.CTkFrame(remove_rep_window, height=180, width=480, corner_radius=0)
        remove_rep_frame.place(x=10, y=10)

        sales_rep_label = customtkinter.CTkLabel(remove_rep_frame, text="Temsilci:", font=customtkinter.CTkFont(size=15))
        sales_rep_label.place(x=100, y=50)
        sales_rep_opt_menu = customtkinter.CTkOptionMenu(remove_rep_frame, width=110,
                                                         values=REPO_OBJ.get_sales_rep_array(),
                                                         font=customtkinter.CTkFont(size=15))
        sales_rep_opt_menu.place(x=170, y=50)
        sales_rep_opt_menu.set("Temsilci Seç")

        submit_button = customtkinter.CTkButton(remove_rep_frame, height=30, width=140, corner_radius=10, text="Sil",
                                                font=customtkinter.CTkFont(size=15), command=submit)
        submit_button.place(x=90, y=110)

        cancel_button = customtkinter.CTkButton(remove_rep_frame, height=30, width=140, corner_radius=10, text="İptal",
                                                font=customtkinter.CTkFont(size=15), command=lambda: self.close_window(remove_rep_window))
        cancel_button.place(x=240, y=110)

        remove_rep_window.mainloop()

    def new_registration(self):
        REPO_OBJ = Repo()

        def submit_entry():
            JSON_OBJ = JSONFileManager()
            REGISTRATION_OBJ = Registration()

            if sales_rep_opt_menu.get() == "Temsilci Seç":
                alert = customtkinter.CTkLabel(reg_frame, text="Temsilci Seçilmedi!", font=customtkinter.CTkFont(size=15), text_color="red")
                alert.place(x=50, y=40)
                alert.after(2000, lambda: alert.destroy())
                return

            content = JSON_OBJ.convert_entry_to_json(sales_rep_opt_menu.get(), time_entry.get(), client_name_entry.get(),
                                                   client_amount_opt_menu.get(), detail_entry.get(), description_entry.get(),
                                                   platform_opt_menu.get(), contact_entry.get())
            sales_rep = sales_rep_opt_menu.get()
            REGISTRATION_OBJ.add_registration(content, sales_rep)
            self.clear_frame(self.scrollable_frame)
            REGISTRATION_OBJ.list_registrations(self.REPO_OBJ.get_registrations_json_file())
            self.close_window(reg_window)

        def get_time():
            now = datetime.now()
            current_time = now.strftime("%H:%M")
            if time_entry != None:
                time_entry.delete(0, len(time_entry.get()))
                time_entry.insert(0, current_time)

        def isContact(value):
            if value==1: contact_entry.insert(0,"MEVCUT")
            else: contact_entry.delete(0,len(contact_entry.get()))

        reg_window = customtkinter.CTk()
        reg_window.title("Yeni Kayıt Oluştur")
        reg_window.geometry("600x440+0+0")

        reg_frame = customtkinter.CTkFrame(reg_window, height=420, width=580, corner_radius=0)
        reg_frame.place(x=10, y=10)

        sales_rep_label = customtkinter.CTkLabel(reg_frame, text="Temsilci:", font=customtkinter.CTkFont(size=15))
        sales_rep_label.place(x=100, y=40)
        sales_rep_opt_menu = customtkinter.CTkOptionMenu(reg_frame,
                                                         width=110,
                                                         values=REPO_OBJ.get_sales_rep_array(),
                                                         font=customtkinter.CTkFont(size=15))
        sales_rep_opt_menu.place(x=190, y=40)
        sales_rep_opt_menu.set("Temsilci Seç")

        time_label = customtkinter.CTkLabel(reg_frame, text="Geliş Saati:", font=customtkinter.CTkFont(size=15))
        time_label.place(x=100, y=80)
        time_entry = customtkinter.CTkEntry(reg_frame, width=100, font=customtkinter.CTkFont(size=15))
        time_entry.place(x=190, y=80)
        get_time_button = customtkinter.CTkButton(reg_frame, height=30, width=130, corner_radius=10, text="Aktar",
                                                  font=customtkinter.CTkFont(size=15), command=get_time)
        get_time_button.place(x=300, y=80)

        client_amount_label = customtkinter.CTkLabel(reg_frame, text="Kişi Sayısı:",
                                                     font=customtkinter.CTkFont(size=15))
        client_amount_label.place(x=100, y=120)
        client_amount_opt_menu = customtkinter.CTkOptionMenu(reg_frame, width=70, values=[' ', '1', '2', '3', '4', '5'],
                                                             font=customtkinter.CTkFont(size=15), anchor="center")
        client_amount_opt_menu.place(x=190, y=120)

        client_name_label = customtkinter.CTkLabel(reg_frame, text="Müşteri Adı:", font=customtkinter.CTkFont(size=15))
        client_name_label.place(x=100, y=160)
        client_name_entry = customtkinter.CTkEntry(reg_frame, width=300, font=customtkinter.CTkFont(size=15))
        client_name_entry.place(x=190, y=160)

        detail_label = customtkinter.CTkLabel(reg_frame, text="Detay:", font=customtkinter.CTkFont(size=15))
        detail_label.place(x=100, y=200)
        detail_entry = customtkinter.CTkEntry(reg_frame, width=300, font=customtkinter.CTkFont(size=15))
        detail_entry.place(x=190, y=200)

        description_label = customtkinter.CTkLabel(reg_frame, text="Açıklama:", font=customtkinter.CTkFont(size=15))
        description_label.place(x=100, y=240)
        description_entry = customtkinter.CTkEntry(reg_frame, width=300, font=customtkinter.CTkFont(size=15))
        description_entry.place(x=190, y=240)

        platform_label = customtkinter.CTkLabel(reg_frame, text="Platform:", font=customtkinter.CTkFont(size=15))
        platform_label.place(x=100, y=280)
        platform_opt_menu = customtkinter.CTkOptionMenu(reg_frame, width=110, values=['GÖRÜŞME', 'TELEFON'],
                                                        font=customtkinter.CTkFont(size=15))
        platform_opt_menu.place(x=190, y=280)

        contact_label = customtkinter.CTkLabel(reg_frame, text="İletişim:", font=customtkinter.CTkFont(size=15))
        contact_label.place(x=100, y=320)
        contact_entry = customtkinter.CTkEntry(reg_frame, font=customtkinter.CTkFont(size=15))
        contact_entry.place(x=190, y=320)

        contact_checkbox_button = customtkinter.CTkCheckBox(reg_frame, text="", checkbox_width=15, checkbox_height=15,
                                                         border_color="gray", command=lambda: isContact(contact_checkbox_button.get()))
        contact_checkbox_button.place(x=340, y=322)

        submit_button = customtkinter.CTkButton(reg_frame, height=30, width=190, corner_radius=10, text="Tamamla",
                                                font=customtkinter.CTkFont(size=15), command=submit_entry)
        submit_button.place(x=100, y=360)

        cancel_button = customtkinter.CTkButton(reg_frame, height=30, width=190, corner_radius=10, text="İptal",
                                                font=customtkinter.CTkFont(size=15), command=lambda: self.close_window(reg_window))
        cancel_button.place(x=310, y=360)

        reg_window.mainloop()

    def update_registration(self):
        JSON_OBJ = JSONFileManager()
        REGISTRATION_OBJ = Registration()

        def bring_reg():
            def submit():
                content = JSON_OBJ.convert_entry_to_json(sales_rep_opt_menu.get(), time_entry.get(),
                                                         client_name_entry.get(),
                                                         client_amount_opt_menu.get(), detail_entry.get(), description_entry.get(),
                                                         platform_opt_menu.get(), contact_entry.get())
                REGISTRATION_OBJ.update_registration(content, sales_rep_opt_menu.get(), int(reg_number_entry.get()))
                self.clear_frame(self.scrollable_frame)
                REGISTRATION_OBJ.list_registrations(self.REPO_OBJ.get_registrations_json_file())
                self.close_window(update_window)

            if sales_rep_opt_menu.get() == "Temsilci Seç":
                alert = customtkinter.CTkLabel(update_frame, text="Temsilci Seçilmedi!",
                                               font=customtkinter.CTkFont(size=15), text_color="red")
                alert.place(x=45, y=40)
                alert.after(2000, lambda: alert.destroy())
                return

            data = JSON_OBJ.read_json_content()
            a = len(data[sales_rep_opt_menu.get()])
            isInt = True
            try:
                val = int(reg_number_entry.get())
            except ValueError:
                isInt = False

            if isInt == False or int(reg_number_entry.get()) <= 0 or int(reg_number_entry.get()) > a:
                alert = customtkinter.CTkLabel(update_frame, text="Hatalı Kayıt No!",
                                               font=customtkinter.CTkFont(size=15), text_color="red")
                alert.place(x=57, y=80)
                alert.after(2000, lambda: alert.destroy())
                return
            item = data[sales_rep_opt_menu.get()][int(reg_number_entry.get()) - 1]

            reg_frame = customtkinter.CTkFrame(update_window, height=300, width=500,
                                               corner_radius=0)
            reg_frame.place(x=100, y=130)

            time_label = customtkinter.CTkLabel(reg_frame, text="Geliş Saati:", font=customtkinter.CTkFont(size=15))
            time_label.place(x=50, y=30)
            time_entry = customtkinter.CTkEntry(reg_frame, width=100, font=customtkinter.CTkFont(size=15))
            time_entry.place(x=140, y=30)
            time_entry.insert(0, item['arrival_time'])

            client_amount_label = customtkinter.CTkLabel(reg_frame, text="Kişi Sayısı:",
                                                         font=customtkinter.CTkFont(size=15))
            client_amount_label.place(x=50, y=70)
            client_amount_opt_menu = customtkinter.CTkOptionMenu(reg_frame, width=110,
                                                                 values=[' ', '1', '2', '3', '4', '5'],
                                                                 font=customtkinter.CTkFont(size=15))
            client_amount_opt_menu.place(x=140, y=70)
            client_amount_opt_menu.set(item['client_amount'])

            client_name_label = customtkinter.CTkLabel(reg_frame, text="Müşteri Adı:",
                                                       font=customtkinter.CTkFont(size=15))
            client_name_label.place(x=50, y=110)
            client_name_entry = customtkinter.CTkEntry(reg_frame, width=300, font=customtkinter.CTkFont(size=15))
            client_name_entry.place(x=140, y=110)
            client_name_entry.insert(0, item['client_name'])

            detail_label = customtkinter.CTkLabel(reg_frame, text="Detay:",
                                                       font=customtkinter.CTkFont(size=15))
            detail_label.place(x=50, y=150)
            detail_entry = customtkinter.CTkEntry(reg_frame, width=300, font=customtkinter.CTkFont(size=15))
            detail_entry.place(x=140, y=150)
            detail_entry.insert(0, item['detail'])

            description_label = customtkinter.CTkLabel(reg_frame, text="Açıklama:",
                                                       font=customtkinter.CTkFont(size=15))
            description_label.place(x=50, y=190)
            description_entry = customtkinter.CTkEntry(reg_frame, width=300, font=customtkinter.CTkFont(size=15))
            description_entry.place(x=140, y=190)
            description_entry.insert(0, item['description'])

            platform_label = customtkinter.CTkLabel(reg_frame, text="Platform:",
                                                    font=customtkinter.CTkFont(size=15))
            platform_label.place(x=50, y=230)
            platform_opt_menu = customtkinter.CTkOptionMenu(reg_frame, values=['GÖRÜŞME', 'TELEFON'],
                                                            font=customtkinter.CTkFont(size=15))
            platform_opt_menu.place(x=140, y=230)
            platform_opt_menu.set(item['platform'])

            contact_label = customtkinter.CTkLabel(reg_frame, text="İletişim:", font=customtkinter.CTkFont(size=15))
            contact_label.place(x=50, y=270)
            contact_entry = customtkinter.CTkEntry(reg_frame, font=customtkinter.CTkFont(size=15))
            contact_entry.place(x=140, y=270)
            contact_entry.insert(0, item['contact'])

            submit_button = customtkinter.CTkButton(update_frame, height=30, width=200, corner_radius=10,
                                                    text="Düzenle",
                                                    font=customtkinter.CTkFont(size=15), command=submit)
            submit_button.place(x=130, y=430)

        update_window = customtkinter.CTk()
        update_window.title("Kayıt Düzenle")
        update_window.geometry("700x500+0+0")

        update_frame = customtkinter.CTkFrame(update_window, height=480, width=680, corner_radius=0)
        update_frame.place(x=10, y=10)

        sales_rep_label = customtkinter.CTkLabel(update_frame, text="Temsilci:", font=customtkinter.CTkFont(size=15))
        sales_rep_label.place(x=100, y=40)
        sales_rep_opt_menu = customtkinter.CTkOptionMenu(update_frame, width=110, values=self.REPO_OBJ.get_sales_rep_array(),
                                                         font=customtkinter.CTkFont(size=15))
        sales_rep_opt_menu.place(x=190, y=40)
        sales_rep_opt_menu.set("Temsilci Seç")

        reg_number_label = customtkinter.CTkLabel(update_frame, text="Kayıt No:", font=customtkinter.CTkFont(size=15))
        reg_number_label.place(x=100, y=80)
        reg_number_entry = customtkinter.CTkEntry(update_frame, width=80, font=customtkinter.CTkFont(size=15))
        reg_number_entry.place(x=190, y=80)

        bring_reg_button = customtkinter.CTkButton(update_frame, height=30, width=100, corner_radius=10, text="Kayıt Bul",
                                                font=customtkinter.CTkFont(size=15), command=bring_reg)
        bring_reg_button.place(x=300, y=80)

        cancel_button = customtkinter.CTkButton(update_frame, height=30, width=200, corner_radius=10, text="İptal",
                                                font=customtkinter.CTkFont(size=15),
                                                command=lambda: self.close_window(update_window))
        cancel_button.place(x=340, y=430)

        update_window.mainloop()

    def delete_registration(self):
        REGISTRATION_OBJ = Registration()
        REPO_OBJ = Repo()
        JSON_OBJ = JSONFileManager()

        def submit():
            if sales_rep_opt_menu.get() == "Temsilci Seç":
                alert = customtkinter.CTkLabel(delete_frame, text="Temsilci Seçilmedi!", font=customtkinter.CTkFont(size=15), text_color="red")
                alert.place(x=12, y=40)
                alert.after(2000, lambda: alert.destroy())
                return
            data = JSON_OBJ.read_json_content()
            a = len(data[sales_rep_opt_menu.get()])
            isInt = True
            try:
                val = int(reg_number_entry.get())
            except ValueError:
                isInt = False

            if isInt == False or int(reg_number_entry.get()) <= 0 or int(reg_number_entry.get()) > a:
                alert = customtkinter.CTkLabel(delete_frame, text="Hatalı Kayıt No!",
                                               font=customtkinter.CTkFont(size=15), text_color="red")
                alert.place(x=37, y=80)
                alert.after(2000, lambda: alert.destroy())
                return
            REGISTRATION_OBJ.delete_registration(sales_rep_opt_menu.get(), int(reg_number_entry.get()))
            self.clear_frame(self.scrollable_frame)
            REGISTRATION_OBJ.list_registrations(REPO_OBJ.get_registrations_json_file())
            self.close_window(delete_window)

        delete_window = customtkinter.CTk()
        delete_window.title("Kayıt Sil")
        delete_window.geometry("500x250+0+0")

        delete_frame = customtkinter.CTkFrame(delete_window, height=230, width=480, corner_radius=0)
        delete_frame.place(x=10, y=10)

        sales_rep_label = customtkinter.CTkLabel(delete_frame, text="Temsilci:", font=customtkinter.CTkFont(size=15))
        sales_rep_label.place(x=80, y=40)

        sales_rep_opt_menu = customtkinter.CTkOptionMenu(delete_frame,
                                                         width=110,
                                                         values=REPO_OBJ.get_sales_rep_array(),
                                                         font=customtkinter.CTkFont(size=15))
        sales_rep_opt_menu.place(x=150, y=40)
        sales_rep_opt_menu.set("Temsilci Seç")

        reg_number_label = customtkinter.CTkLabel(delete_frame, text="Kayıt No:", font=customtkinter.CTkFont(size=15))
        reg_number_label.place(x=80, y=80)
        reg_number_entry = customtkinter.CTkEntry(delete_frame, width=100
                                            , font=customtkinter.CTkFont(size=15))
        reg_number_entry.place(x=150, y=80)

        submit_button = customtkinter.CTkButton(delete_frame, height=30, width=150, corner_radius=10, text="Sil",
                                                font=customtkinter.CTkFont(size=15), command=submit)
        submit_button.place(x=80, y=180)

        cancel_button = customtkinter.CTkButton(delete_frame, height=30, width=150, corner_radius=10, text="İptal",
                                                font=customtkinter.CTkFont(size=15), command=lambda: self.close_window(delete_window))
        cancel_button.place(x=240, y=180)

        delete_window.mainloop()

    def change_registration(self):
        JSON_OBJ = JSONFileManager()
        REGISTRATION_OBJ = Registration()
        REPO_OBJ = Repo()

        def submit():
            if from_opt_menu.get() == "Temsilci Seç":
                alert = customtkinter.CTkLabel(change_frame, text="Temsilci Seçilmedi!", font=customtkinter.CTkFont(size=15), text_color="red")
                alert.place(x=10, y=40)
                alert.after(2000, lambda: alert.destroy())
                return
            if to_opt_menu.get() == "Temsilci Seç":
                alert = customtkinter.CTkLabel(change_frame, text="Temsilci Seçilmedi!", font=customtkinter.CTkFont(size=15), text_color="red")
                alert.place(x=10, y=80)
                alert.after(2000, lambda: alert.destroy())
                return
            if to_opt_menu.get() == from_opt_menu.get():
                alert = customtkinter.CTkLabel(change_frame, text="Temsilci Aynı!", font=customtkinter.CTkFont(size=15), text_color="red")
                alert.place(x=40, y=80)
                alert.after(2000, lambda: alert.destroy())
                return
            data = JSON_OBJ.read_json_content()
            a = len(data[from_opt_menu.get()])
            isInt = True
            try:
                val = int(reg_number_entry.get())
            except ValueError:
                isInt = False

            if isInt == False or int(reg_number_entry.get()) <= 0 or int(reg_number_entry.get()) > a:
                alert = customtkinter.CTkLabel(change_frame, text="Hatalı Kayıt No!",
                                               font=customtkinter.CTkFont(size=15), text_color="red")
                alert.place(x=37, y=120)
                alert.after(2000, lambda: alert.destroy())
                return
            from_sales_rep = from_opt_menu.get()
            to_sales_rep = to_opt_menu.get()
            reg_number = int(reg_number_entry.get())
            data[from_sales_rep][reg_number-1]['sales_rep'] = to_sales_rep
            REGISTRATION_OBJ.add_registration(data[from_sales_rep][reg_number-1], to_sales_rep)
            REGISTRATION_OBJ.delete_registration(from_sales_rep, reg_number)
            self.clear_frame(self.scrollable_frame)
            REGISTRATION_OBJ.list_registrations(REPO_OBJ.get_registrations_json_file())
            self.close_window(change_window)

        change_window = customtkinter.CTk()
        change_window.title("Kayıt Değiştir")
        change_window.geometry("500x250+0+0")

        change_frame = customtkinter.CTkFrame(change_window, height=230, width=480, corner_radius=0)
        change_frame.place(x=10, y=10)

        from_label = customtkinter.CTkLabel(change_frame, text="Kimden:", font=customtkinter.CTkFont(size=15))
        from_label.place(x=80, y=40)
        from_opt_menu = customtkinter.CTkOptionMenu(change_frame,
                                                         width=110,
                                                         values=REPO_OBJ.get_sales_rep_array(),
                                                         font=customtkinter.CTkFont(size=15))
        from_opt_menu.place(x=160, y=40)
        from_opt_menu.set("Temsilci Seç")

        to_label = customtkinter.CTkLabel(change_frame, text="Kime:", font=customtkinter.CTkFont(size=15))
        to_label.place(x=80, y=80)
        to_opt_menu = customtkinter.CTkOptionMenu(change_frame,
                                                         width=110,
                                                         values=REPO_OBJ.get_sales_rep_array(),
                                                         font=customtkinter.CTkFont(size=15))
        to_opt_menu.place(x=160, y=80)
        to_opt_menu.set("Temsilci Seç")

        reg_number_label = customtkinter.CTkLabel(change_frame, text="Kayıt No:", font=customtkinter.CTkFont(size=15))
        reg_number_label.place(x=80, y=120)
        reg_number_entry = customtkinter.CTkEntry(change_frame, width=100
                                            , font=customtkinter.CTkFont(size=15))
        reg_number_entry.place(x=160, y=120)

        submit_button = customtkinter.CTkButton(change_frame, height=30, width=150, corner_radius=10, text="Değiştir",
                                                font=customtkinter.CTkFont(size=15), command=submit)
        submit_button.place(x=80, y=180)

        cancel_button = customtkinter.CTkButton(change_frame, height=30, width=150, corner_radius=10, text="İptal",
                                                font=customtkinter.CTkFont(size=15), command=lambda: self.close_window(change_window))
        cancel_button.place(x=240, y=180)

        change_window.mainloop()

    def create_report(self):
        REPORT = ReportManager()
        report_window = customtkinter.CTk()
        report_window.title("Rapor Oluştur")
        report_window.geometry("500x200+0+0")

        report_frame = customtkinter.CTkFrame(report_window, height=180, width=480, corner_radius=0)
        report_frame.place(x=10, y=10)

        label = customtkinter.CTkLabel(report_frame, text="Gün sonu raporu oluşturma işlemine \n devam etmek istiyor musunuz?",
                                       font=customtkinter.CTkFont(size=25))
        label.place(x=43, y=40)

        progressbar = customtkinter.CTkProgressBar(report_frame, mode="determinate",
                                                   determinate_speed=0.60, width=360, height=15) #0.81 olmalı windows için!

        def submit():
            REPORT.create_report()
            label.destroy()
            progressbar.place(x=60, y=60)
            progressbar.set(0)
            progressbar.start()
            progressbar.after(1795, lambda: progressbar.stop())
            self.close_window(report_window, 2000)

        submit_button = customtkinter.CTkButton(report_frame, height=30, width=190, corner_radius=10, text="Evet",
                                                font=customtkinter.CTkFont(size=15),
                                                command=submit)
        submit_button.place(x=45, y=120)

        cancel_button = customtkinter.CTkButton(report_frame, height=30, width=190, corner_radius=10, text="Hayır",
                                                font=customtkinter.CTkFont(size=15),
                                                command=lambda: self.close_window(report_window))
        cancel_button.place(x=250, y=120)

        report_window.mainloop()


class JSONFileManager:
    REPO_OBJ = Repo()

    def __init__(self):
        pass

    def read_json_content(self, filename=REPO_OBJ.get_registrations_json_file()):
        with open(filename, 'r', encoding="utf-8") as file:
            data = json.load(file)
            return data

    def add_json_entry(self, new_data, sales_rep, filename=REPO_OBJ.get_registrations_json_file()):
        data = self.read_json_content()
        with open(filename, 'w', encoding="utf-8") as file:
            data[sales_rep].append(new_data)
            file.seek(0)
            sorted_data = self._sort_entries(data)
            json.dump(sorted_data, file, indent=4)
        if self._check_json_file_length():
            self._save_content()

    def delete_json_entry(self, sales_rep, index, filename=REPO_OBJ.get_registrations_json_file()):
        data = self.read_json_content()
        data[sales_rep].pop(index - 1)
        with open(filename, 'w', encoding="utf-8") as file:
            counter = 1
            for item in data[sales_rep]:
                item['index'] = counter
                counter += 1
            file.seek(0)
            sorted_data = self._sort_entries(data)
            json.dump(sorted_data, file, indent=4)

    def update_json_entry(self, new_data, sales_rep, index, filename=REPO_OBJ.get_registrations_json_file()):
        data = self.read_json_content()
        with open(filename, 'w', encoding="utf-8") as file:
            data[sales_rep][int(index)-1] = new_data
            file.seek(0)
            sorted_data = self._sort_entries(data)
            json.dump(sorted_data, file, indent=4)

    def _check_json_file_length(self):
        data = self.read_json_content()
        entry_amount = 0
        sales_rep_counter = 0
        sales_rep_array = self.REPO_OBJ.get_sales_rep_array()
        while True:
            if sales_rep_counter >= len(sales_rep_array): break
            entry_amount = entry_amount + len(data[sales_rep_array[sales_rep_counter]])
            sales_rep_counter += 1
        if entry_amount != 0 and entry_amount%3 == 0: return True
        else: return False

    def convert_entry_to_json(self, sales_rep, time, clientname, clientamount, detail, description, platform, contact):
        index = len(self.read_json_content()[sales_rep]) + 1
        json_entry = {
            "index": index,
            "sales_rep": sales_rep,
            "arrival_time": time,
            "client_name": clientname,
            "client_amount": clientamount,
            "detail": detail,
            "description": description,
            "platform": platform,
            "contact": contact
        }
        return json_entry

    def _save_content(self):
        ReportManager().create_backup()

    def _sort_entries(self, data):
        sales_rep_counter = 0
        REPO_OBJ = Repo()

        while True:
            if sales_rep_counter >= len(self.REPO_OBJ.get_sales_rep_array()): break
            arr = data[REPO_OBJ.get_sales_rep_array()[sales_rep_counter]]
            n = len(arr)
            for i in range(n):
                for j in range(0, n - i - 1):
                    if arr[j]['arrival_time'] > arr[j + 1]['arrival_time']:
                        arr[j], arr[j + 1] = arr[j + 1], arr[j]
            sales_rep_counter += 1

        sales_rep_counter = 0
        counter = 1
        while True:
            if sales_rep_counter >= len(REPO_OBJ.get_sales_rep_array()): break
            for item in data[REPO_OBJ.get_sales_rep_array()[sales_rep_counter]]:
                item['index'] = counter
                counter += 1
            counter = 1
            sales_rep_counter += 1
        return data


class EntryWidget:
    UTILIZATION_OBJ = Utilization()

    def __init__(self):
        pass

    def filterByName(self, sales_rep):
        REGISTRATION_OBJ = Registration()
        REGISTRATION_OBJ.list_registrations_by_name(sales_rep)

    def addSalesRepWidget(self, frame, sales_rep, record_amount):

        bg, fc = self.UTILIZATION_OBJ.getAttributes()
        style = ttk.Style()
        #style.theme_use("default") "for windows"
        style.configure("Treeview",
                        background=bg,
                        foreground=fc,
                        fieldbackground=bg,
                        rowheight=30)
        style.configure("Treeview.Heading",
                        foreground=fc,
                        font=("Arial", 14, "bold"))
        sales_rep_label = customtkinter.CTkButton(frame, width=140, text=sales_rep,
                                                  font=customtkinter.CTkFont(size=20), command=lambda: self.filterByName(sales_rep))
        sales_rep_label.pack()
        list_screen = ttk.Treeview(frame, height=record_amount)
        list_screen['columns'] = ("İndex", "Geliş Saati", "Müşteri İsmi", "M.Sayısı",
                                  "Detay", "Açıklama", "Platform", "İletişim Bilgisi")

        list_screen.column("#0", width=0, stretch=NO)
        list_screen.column("İndex", anchor=CENTER, width=40, minwidth=30)
        list_screen.column("Geliş Saati", anchor=CENTER, width=100, minwidth=70)
        list_screen.column("Müşteri İsmi", anchor=W, width=250, minwidth=210)
        list_screen.column("M.Sayısı", anchor=CENTER, width=65, minwidth=65)
        list_screen.column("Detay", anchor=W, width=180, minwidth=140)
        list_screen.column("Açıklama", anchor=W, width=280, minwidth=240)
        list_screen.column("Platform", anchor=CENTER, width=100, minwidth=80)
        list_screen.column("İletişim Bilgisi", anchor=CENTER, width=150, minwidth=120)

        list_screen.heading("#0", text="", anchor=W)
        list_screen.heading("İndex", text="#", anchor=CENTER)
        list_screen.heading("Geliş Saati", text="Geliş Saati", anchor=CENTER)
        list_screen.heading("Müşteri İsmi", text="Müşteri İsmi", anchor=CENTER)
        list_screen.heading("M.Sayısı", text="M.Sayısı", anchor=CENTER)
        list_screen.heading("Detay", text="Detay", anchor=CENTER)
        list_screen.heading("Açıklama", text="Açıklama", anchor=CENTER)
        list_screen.heading("Platform", text="Platform", anchor=CENTER)
        list_screen.heading("İletişim Bilgisi", text="İletişim Bilgisi", anchor=CENTER)
        list_screen.pack(pady=5)
        return list_screen

    def addEntryWidget(self, content, counter, table):
        table.insert(parent="", index="end", iid=counter, text="",
                           values=(content["index"], content["arrival_time"], content["client_name"],
                                   content["client_amount"], content["detail"], content["description"],
                                   content["platform"], content["contact"]))


class Registration:
    JSON_OBJ = JSONFileManager()
    REPO_OBJ = Repo()
    MAIN_UI_OBJ = MainUI()

    def __init__(self):
        pass

    def add_registration(self, new_data, sales_rep, file=REPO_OBJ.get_registrations_json_file()):
        self.JSON_OBJ.add_json_entry(new_data, sales_rep, file)

    def update_registration(self, new_data, sales_rep, index):
        self.JSON_OBJ.update_json_entry(new_data, sales_rep, index)

    def delete_registration(self, sales_rep, index):
        self.JSON_OBJ.delete_json_entry(sales_rep, index)

    def list_registrations(self, fileName=REPO_OBJ.get_registrations_json_file()):
        global state
        state = True
        data = self.JSON_OBJ.read_json_content(fileName)
        sales_rep_array = self.REPO_OBJ.get_sales_rep_array()
        WIDGET = EntryWidget()

        sales_rep_counter = 0
        counter = 0
        while True:
            if sales_rep_counter >= len(sales_rep_array):
                break
            record_amount = len(data[sales_rep_array[sales_rep_counter]])
            table_of_sales_rep = WIDGET.addSalesRepWidget(self.MAIN_UI_OBJ.scrollable_frame, sales_rep_array[sales_rep_counter], record_amount)
            if record_amount == 0: pass
            else:
                for item in data[sales_rep_array[sales_rep_counter]]:
                    WIDGET.addEntryWidget(item, counter, table_of_sales_rep)
                    counter += 1
            sales_rep_counter += 1
    def list_registrations_by_name(self, sales_rep, fileName=REPO_OBJ.get_registrations_json_file()):
        self.MAIN_UI_OBJ.clear_frame(self.MAIN_UI_OBJ.scrollable_frame)
        global state
        if state == True:
            data = self.JSON_OBJ.read_json_content(fileName)
            WIDGET = EntryWidget()

            counter = 0

            record_amount = len(data[sales_rep])
            table_of_sales_rep = WIDGET.addSalesRepWidget(self.MAIN_UI_OBJ.scrollable_frame, sales_rep, record_amount)
            if record_amount == 0: pass
            else:
                for item in data[sales_rep]:
                    WIDGET.addEntryWidget(item, counter, table_of_sales_rep)
                    counter += 1
            state = False
        else:
            self.list_registrations()


class ExcelHandler:
    workbook = None
    DATE_FORMAT = {  # DATE FORMAT ON TOP
        'font_color': 'red',
        'font_size': 16,
        # 'align': 'center'
    }

    SALES_REPRESENTATIVE_FORMAT = {  # SALES_REP MERGED CELLS SALES REPRESENTATIVE FORMAT
        'border': 1,
        'font_color': 'white',
        'font_size': 16,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': 'black'
    }

    INDEX_NUMBER_FORMAT = {  # INDEX NUMBERS FORMAT
        'border': 1,
        'font_color': 'red',
        'font_size': 16,
        'align': 'center',
        'fg_color': 'silver'
    }

    TIME_FORMAT = {  # TIME FORMAT
        'border': 1,
        # 'font_color': 'black',
        'font_size': 16,
        'align': 'center',
        'fg_color': 'silver'
    }

    CLIENT_NAME_FORMAT = { # CLIENT NAME FORMAT
        'border': 1,
        'font_size': 16,
        'fg_color': 'silver'
    }

    CLIENT_AMOUNT_FORMAT = {  # CLIENT AMOUNT FORMAT
        'border': 1,
        'font_size': 16,
        'align': 'center',
        'fg_color': 'silver'
    }

    DETAILS_FORMAT = { # DETAILS FORMAT
        'border': 1,
        'font_size': 16,
        'fg_color': 'silver'
    }

    PLATFORM_FORMAT = { # PLATFORM FORMAT
        'border': 1,
        'font_size': 16,
        'align': 'center',
        'fg_color': 'silver'
    }

    CONTACT_FORMAT = { # CONTACT FORMAT
        'border': 1,
        'font_size': 16,
        'align': 'center',
        'fg_color': 'silver'
    }

    def __init__(self):
        pass

    def createNewExcelFile(self):
        self.workbook = xlsxwriter.Workbook(datetime.today().strftime("%d.%m.%Y") + " Günlük Rapor" + ".xlsx")
        return self.workbook

    def createNewSheet(self, workbook):
        worksheet = workbook.add_worksheet(datetime.today().strftime("%d.%m.%Y"))
        return worksheet

    def writeToSheet(self, worksheet, row, column, content, format):
        worksheet.write(row, column, content, format)


class ReportManager:
    EXCELHANDLER_OBJ = ExcelHandler()
    JSON_OBJ = JSONFileManager()
    REPO_OBJ = Repo()
    UTILIZATION_OBJ = Utilization()

    def __init__(self):
        pass

    def _write_to_sheet(self, worksheet, row, column, content, format):
        self.EXCELHANDLER_OBJ.writeToSheet(worksheet, row, column, content, format)

    def create_report(self):
        sales_rep_array = self.REPO_OBJ.get_sales_rep_array()
        entries = self.JSON_OBJ.read_json_content()
        os.chdir("..")

        year = self.UTILIZATION_OBJ.yearFinder()
        month = self.UTILIZATION_OBJ.monthFinder()

        if os.path.isdir("./" + year + " Günlük Rapor"): pass
        else: os.mkdir(year + " Günlük Rapor")
        os.chdir("./" + year + " Günlük Rapor")

        if os.path.isdir(month + " Günlük Rapor"): pass
        else: os.mkdir(month + " Günlük Rapor")
        os.chdir("./" + month + " Günlük Rapor")

        excel_file = self.EXCELHANDLER_OBJ.createNewExcelFile()
        excel_sheet = self.EXCELHANDLER_OBJ.createNewSheet(excel_file)

        self._write_to_sheet(excel_sheet, 0, 2, datetime.today().strftime("%d.%m.%Y") +
                             " " + self.UTILIZATION_OBJ.dayFinder(), excel_file.add_format(self.EXCELHANDLER_OBJ.DATE_FORMAT))
        excel_sheet.set_column("C:C", 30)
        excel_sheet.set_column("A:A", 4)
        excel_sheet.set_column("D:D", 4)
        excel_sheet.set_column("E:E", 30)
        excel_sheet.set_column("F:F", 65)
        excel_sheet.set_column("G:G", 20)
        excel_sheet.set_column("H:H", 20)

        sales_rep_counter = 0
        row = 1
        column = 0

        while True:
            if sales_rep_counter >= len(sales_rep_array): break
            excel_sheet.merge_range(row, 0, row, 7, sales_rep_array[sales_rep_counter],
                                    excel_file.add_format(self.EXCELHANDLER_OBJ.SALES_REPRESENTATIVE_FORMAT))
            row += 1
            for item in entries[sales_rep_array[sales_rep_counter]]:
                self._write_to_sheet(excel_sheet, row, column, item['index'], excel_file.add_format(self.EXCELHANDLER_OBJ.INDEX_NUMBER_FORMAT))
                self._write_to_sheet(excel_sheet, row, column+1, item['arrival_time'], excel_file.add_format(self.EXCELHANDLER_OBJ.TIME_FORMAT))
                self._write_to_sheet(excel_sheet, row, column+2, item['client_name'], excel_file.add_format(self.EXCELHANDLER_OBJ.CLIENT_NAME_FORMAT))
                self._write_to_sheet(excel_sheet, row, column+3, item['client_amount'], excel_file.add_format(self.EXCELHANDLER_OBJ.CLIENT_AMOUNT_FORMAT))
                self._write_to_sheet(excel_sheet, row, column+4, item['detail'], excel_file.add_format(self.EXCELHANDLER_OBJ.CLIENT_NAME_FORMAT))
                self._write_to_sheet(excel_sheet, row, column+5, item['description'], excel_file.add_format(self.EXCELHANDLER_OBJ.DETAILS_FORMAT))
                self._write_to_sheet(excel_sheet, row, column+6, item['platform'], excel_file.add_format(self.EXCELHANDLER_OBJ.PLATFORM_FORMAT))
                self._write_to_sheet(excel_sheet, row, column+7, item['contact'], excel_file.add_format(self.EXCELHANDLER_OBJ.CONTACT_FORMAT))
                row += 1
            row += 1
            sales_rep_counter += 1
        excel_file.close()
        os.chdir("..")
        os.chdir("..")
        os.chdir("./s_registration_source")

    def create_backup(self):
        sales_rep_array = self.REPO_OBJ.get_sales_rep_array()
        entries = self.JSON_OBJ.read_json_content()
        os.chdir("..")

        if os.path.isdir("./Yedek"): pass
        else: os.mkdir("Yedek")
        os.chdir("./Yedek")

        excel_file = self.EXCELHANDLER_OBJ.createNewExcelFile()
        excel_sheet = self.EXCELHANDLER_OBJ.createNewSheet(excel_file)

        self._write_to_sheet(excel_sheet, 0, 2, datetime.today().strftime("%d.%m.%Y") +
                             " " + self.UTILIZATION_OBJ.dayFinder(),
                             excel_file.add_format(self.EXCELHANDLER_OBJ.DATE_FORMAT))
        excel_sheet.set_column("C:C", 30)
        excel_sheet.set_column("A:A", 4)
        excel_sheet.set_column("D:D", 4)
        excel_sheet.set_column("E:E", 30)
        excel_sheet.set_column("F:F", 65)
        excel_sheet.set_column("G:G", 20)
        excel_sheet.set_column("H:H", 20)

        sales_rep_counter = 0
        row = 1
        column = 0

        while True:
            if sales_rep_counter >= len(sales_rep_array): break
            excel_sheet.merge_range(row, 0, row, 7, sales_rep_array[sales_rep_counter],
                                    excel_file.add_format(self.EXCELHANDLER_OBJ.SALES_REPRESENTATIVE_FORMAT))
            row += 1
            for item in entries[sales_rep_array[sales_rep_counter]]:
                self._write_to_sheet(excel_sheet, row, column, item['index'],
                                     excel_file.add_format(self.EXCELHANDLER_OBJ.INDEX_NUMBER_FORMAT))
                self._write_to_sheet(excel_sheet, row, column + 1, item['arrival_time'],
                                     excel_file.add_format(self.EXCELHANDLER_OBJ.TIME_FORMAT))
                self._write_to_sheet(excel_sheet, row, column + 2, item['client_name'],
                                     excel_file.add_format(self.EXCELHANDLER_OBJ.CLIENT_NAME_FORMAT))
                self._write_to_sheet(excel_sheet, row, column + 3, item['client_amount'],
                                     excel_file.add_format(self.EXCELHANDLER_OBJ.CLIENT_AMOUNT_FORMAT))
                self._write_to_sheet(excel_sheet, row, column + 4, item['detail'],
                                     excel_file.add_format(self.EXCELHANDLER_OBJ.CLIENT_NAME_FORMAT))
                self._write_to_sheet(excel_sheet, row, column + 5, item['description'],
                                     excel_file.add_format(self.EXCELHANDLER_OBJ.DETAILS_FORMAT))
                self._write_to_sheet(excel_sheet, row, column + 6, item['platform'],
                                     excel_file.add_format(self.EXCELHANDLER_OBJ.PLATFORM_FORMAT))
                self._write_to_sheet(excel_sheet, row, column + 7, item['contact'],
                                     excel_file.add_format(self.EXCELHANDLER_OBJ.CONTACT_FORMAT))
                row += 1
            row += 1
            sales_rep_counter += 1
        excel_file.close()
        os.chdir("..")
        os.chdir("./s_registration_source")


app = MainUI()
app.startApp()
app.closeApp()

#TODO: MainUI init() is called 2 times