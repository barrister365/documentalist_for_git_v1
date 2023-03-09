
import tkinter
from tkinter.ttk import Radiobutton
import tkinter as tk
import webbrowser
from docxtpl import DocxTemplate
from tkinter import Tk, Menu, scrolledtext, messagebox, ttk
import pyperclip

def spravkaoprogramme():
    spravkawindow = tk.Toplevel(window)
    spravkawindow.title("Справка о программе")
    spravkawindow.geometry('400x250')
    filespr = open("Справка документалист.txt", 'r', encoding='utf-8')
    asd = tkinter.Label(spravkawindow, text=filespr.read())
    asd.grid(column=0, row=0, sticky=tk.N)


def save_doc(source, a, b, c, d, e, f, g, h, i, saving_path):
    doc = DocxTemplate(source)
    context = {'dateofcontracttext': a, 'cityofcontract': b, 'seller': c,
               'buyer': d, 'canceldate': e, 'price': f,
               'detailsseller': g, 'detailsbuyer': h, 'tzsell': i}
    doc.render(context)
    doc.save(saving_path)
    messagebox.showinfo('Документалист v. 1.0.', 'Файл сохранён в указнной директории')


def sale():
    tab2 = ttk.Frame(tab_control)
    tab_control.add(tab2, text='Договор купли-продажи')
    tab2.bind('<Control-v>', pyperclip.paste())
    dateofcontract = tkinter.Label(tab2, text="Введите дату договора:")
    dateofcontract.grid(column=0, row=1, sticky=tk.E)
    date = tk.StringVar()
    dateofcontracttext = tkinter.Entry(tab2, width=45, textvariable=date)
    dateofcontracttext.grid(column=1, row=1, sticky=tk.W)
    dateofcontracttext.focus()
    dateofcontracthelp = tkinter.Label(tab2, text="Например: 01.01.2022 или 1 января 2022 года")
    dateofcontracthelp.grid(column=2, row=1, sticky=tk.W)
    cityofcontract = tkinter.Label(tab2, text="Введите место составления договора:")
    cityofcontract.grid(column=0, row=2, sticky=tk.E)
    city = tk.StringVar()
    cityofcontracttext = tkinter.Entry(tab2, width=45, textvariable=city)
    cityofcontracttext.grid(column=1, row=2, sticky=tk.W)
    cityofcontracttext.focus()
    cityofcontracthelp = tkinter.Label(tab2, text="Например: Москва")
    cityofcontracthelp.grid(column=2, row=2, sticky=tk.W)
    seller = tkinter.Label(tab2, text="Продавец (лицо, действующее от его имени на основании документа):")
    seller.grid(column=0, row=3, sticky=tk.E)
    sellerto = tk.StringVar()
    sellertext = tkinter.Entry(tab2, width=45, textvariable=sellerto)
    sellertext.grid(column=1, row=3, sticky=tk.W)
    sellertext.focus()
    sellerhelp = tkinter.Label(tab2, text="Например: ООО Ромашка, в лице директора ...")
    sellerhelp.grid(column=2, row=3, sticky=tk.W)
    buyer = tkinter.Label(tab2, text="Покупатель (лицо, действующее от его имени на основании документа):")
    buyer.grid(column=0, row=4, sticky=tk.E)
    buyerto = tk.StringVar()
    buyertext = tkinter.Entry(tab2, width=45, textvariable=buyerto)
    buyertext.grid(column=1, row=4, sticky=tk.W)
    buyertext.focus()
    price = tkinter.Label(tab2, text="Введите цену договора,"
                             "\nвключая параметры применения НДС:")
    price.grid(column=0, row=5, sticky=tk.E)
    priceto = tk.StringVar()
    pricetext = tkinter.Entry(tab2, width=45, textvariable=priceto)
    pricetext.grid(column=1, row=5, sticky=tk.W)
    pricetext.focus()
    pricehelp= tkinter.Label(tab2, text="Например: 100 000 рублей 00 коп. или 100 000 руб.\n"
                                "Цена обязательно указывается с НДС или без НДС,\n"
                                "также необходимо указать на каком основании\n"
                                "НДС не применяется.")
    pricehelp.grid(column=2, row=5, sticky=tk.W)
    canceldate = tkinter.Label(tab2, text="Введите дату окончания срока действия договора:")
    canceldate.grid(column=0, row=6, sticky=tk.E)
    canceldateto = tk.StringVar
    canceldatetext = tkinter.Entry(tab2, width=45, textvariable=canceldateto)
    canceldatetext.grid(column=1, row=6, sticky=tk.W)
    canceldatetext.focus()
    detailsseller = tkinter.Label(tab2, text="Банковские реквизиты продавца:")
    detailsseller.grid(column=0, row=7, sticky=tk.E)
    detailssellertxt = scrolledtext.ScrolledText(tab2, width=35, height=6)
    detailssellertxt.grid(column=1, row=7, sticky=tk.W)
    detailssellerhelp = tkinter.Label(tab2, text="Наименование, ИНН, ОГРН, БИК Банка, р\с ...")
    detailssellerhelp.grid(column=2, row=7, sticky=tk.W)
    detailsbuyer = tkinter.Label(tab2, text="Банковские реквизиты покупателя:")
    detailsbuyer.grid(column=0, row=8, sticky=tk.E)
    detailsbuyertxt = scrolledtext.ScrolledText(tab2, width=35, height=6)
    detailsbuyertxt.grid(column=1, row=8, sticky=tk.W)
    tzsell = tkinter.Label(tab2, text="Описание объекта покупки:")
    tzsell.grid(column=0, row=9, sticky=tk.E)
    tzselltxt = scrolledtext.ScrolledText(tab2, width=38, height=14)
    tzselltxt.grid(column=1, row=9, sticky=tk.W)
    tzsellhelp = tkinter.Label(tab2, text="Указываются спецификация товара, порядок его поставки,\n"
                                  "упаковки, идентификационный номер.\n"
                                  "В случаях, если товар является ценной бумагой,\n "
                                  "валютной ценностью, указываются особенности его\n"
                                  "купли-продажи в соответствии с законодательством\n"
                                  "Российской Федерации. В случаях сделок розничной\n"
                                  "купли-продажи, поставки товаров, контрактации,\n"
                                  "продажи недвижимости, продажи предприятия\n"
                                  "также указываются их особенности в соответствии\n"
                                  "с законодательством Российской Федерации.")
    tzsellhelp.grid(column=2, row=9, sticky=tk.W)
    btnhelp = tk.Button(tab2, text="Судебная практика по типу документа",
                        command=lambda: webbrowser.open("https://barrister365-my.sharepoint.com/:f:/g/personal/yan_barrister365_onmicrosoft_com/EvydKT8ZUPRBnDamDBJuDcEBmBb-A7bNZHtha8AKcFcHhA?e=bPeDSP"))
    btnhelp.grid(column=0, row=11)
    dirforsave = tkinter.Label(tab2, text="Директория для сохранения файла:")
    dirforsave.grid(column=0, row=10, sticky=tk.E)
    standartdirectory = tk.StringVar()
    standartdirectory.set(r"C:\Users\yandv\OneDrive - индивидуальный предприниматель Дворянкин Ян Александрович\Приложение Документалист\Договор1.docx")
    dirforsavetext = tkinter.Entry(tab2, width=45, textvariable = standartdirectory, state='normal')
    dirforsavetext.grid(column=1, row=10, sticky=tk.W)
    dirforsavetext.focus()
    btnsave = tk.Button(tab2, text="Сохранить готовый документ",
                        command=(lambda:
                                 save_doc(source='Договор_купли_продажи.docx',
                                          a=dateofcontracttext.get(),
                                          b=cityofcontracttext.get(),
                                          c=sellertext.get(),
                                          d=buyertext.get(),
                                          e=pricetext.get(),
                                          f=canceldatetext.get(),
                                          g=detailssellertxt.get("1.0", tk.END),
                                          h=detailsbuyertxt.get("1.0", tk.END),
                                          i=tzselltxt.get("1.0", tk.END),
                                          saving_path=dirforsavetext.get())
                                 )
                        )
    btnsave.grid(column=1, row=11)
    btnexitsale = tk.Button(tab2, text="Закрыть вкладку", command=lambda: tab_control.forget(tab2))
    btnexitsale.grid(column=2, row=11)


def barter():
    tab3 = ttk.Frame(tab_control)
    tab_control.add(tab3, text='Договор мены')
    tab3.bind('<Control-v>', pyperclip.paste())
    barter_date = tkinter.Label(tab3, text="Введите дату договора:")
    barter_date.grid(column=0, row=1, sticky=tk.E)
    barter_date_entry = tk.StringVar()
    barter_date_entry_text = tkinter.Entry(tab3, width=45, textvariable=barter_date_entry)
    barter_date_entry_text.grid(column=1, row=1, sticky=tk.W)
    barter_date_entry_text.focus()
    barter_date_entry_help = tkinter.Label(tab3, text="Например: 01.01.2022 или 1 января 2022 года")
    barter_date_entry_help.grid(column=2, row=1, sticky=tk.W)
    place_of_barter = tkinter.Label(tab3, text="Введите место составления договора:")
    place_of_barter.grid(column=0, row=2, sticky=tk.E)
    place_of_barter_entry = tk.StringVar()
    place_of_barter_text = tkinter.Entry(tab3, width=45, textvariable=place_of_barter_entry)
    place_of_barter_text.grid(column=1, row=2, sticky=tk.W)
    place_of_barter_text.focus()
    place_of_barter_entry_help = tkinter.Label(tab3, text="Например: Москва")
    place_of_barter_entry_help.grid(column=2, row=2, sticky=tk.W)
    first_seller = tkinter.Label(tab3, text="Продавец-1 (лицо, действующее от его имени на основании документа):")
    first_seller.grid(column=0, row=3, sticky=tk.E)
    first_seller_entry = tk.StringVar()
    first_seller_text = tkinter.Entry(tab3, width=45, textvariable=first_seller_entry)
    first_seller_text.grid(column=1, row=3, sticky=tk.W)
    first_seller_text.focus()
    first_seller_entry_help = tkinter.Label(tab3, text="Например: ООО Ромашка, в лице директора ...")
    first_seller_entry_help.grid(column=2, row=3, sticky=tk.W)
    second_seller = tkinter.Label(tab3, text="Продавец-2 (лицо, действующее от его имени на основании документа):")
    second_seller.grid(column=0, row=4, sticky=tk.E)
    second_seller_entry = tk.StringVar()
    second_seller_text = tkinter.Entry(tab3, width=45, textvariable=second_seller_entry)
    second_seller_text.grid(column=1, row=4, sticky=tk.W)
    second_seller_text.focus()
    barter_price = tkinter.Label(tab3, text="Введите сумму, которую сторона\n"
                                    "должна доплатить при мене:")
    barter_price.grid(column=0, row=5, sticky=tk.E)
    barter_price_entry = tk.StringVar()
    barter_price_text = tkinter.Entry(tab3, width=45, textvariable=barter_price_entry)
    barter_price_text.grid(column=1, row=5, sticky=tk.W)
    barter_price_text.focus()
    barter_price_help= tkinter.Label(tab3, text="Например: 100 000 рублей 00 коп. или 100 000 руб.\n"
                                "Цена обязательно указывается с НДС или без НДС,\n"
                                "также необходимо указать на каком основании\n"
                                "НДС не применяется.")
    barter_price_help.grid(column=2, row=5, sticky=tk.W)
    expire_date_of_barter = tkinter.Label(tab3, text="Введите дату окончания срока действия договора:")
    expire_date_of_barter.grid(column=0, row=6, sticky=tk.E)
    expire_date_of_barter_entry = tk.StringVar
    expire_date_of_barter_text = tkinter.Entry(tab3, width=45, textvariable=expire_date_of_barter_entry)
    expire_date_of_barter_text.grid(column=1, row=6, sticky=tk.W)
    expire_date_of_barter_text.focus()
    detail_of_first_seller = tkinter.Label(tab3, text="Банковские реквизиты покупателя-1:")
    detail_of_first_seller.grid(column=0, row=7, sticky=tk.E)
    detail_of_first_seller_text = scrolledtext.ScrolledText(tab3, width=35, height=6)
    detail_of_first_seller_text.grid(column=1, row=7, sticky=tk.W)
    detail_of_first_seller_help = tkinter.Label(tab3, text="Наименование, ИНН, ОГРН, БИК Банка, р\с ...")
    detail_of_first_seller_help.grid(column=2, row=7, sticky=tk.W)
    detail_of_second_seller = tkinter.Label(tab3, text="Банковские реквизиты покупателя-2:")
    detail_of_second_seller.grid(column=0, row=8, sticky=tk.E)
    detail_of_second_seller_text = scrolledtext.ScrolledText(tab3, width=35, height=6)
    detail_of_second_seller_text.grid(column=1, row=8, sticky=tk.W)
    object_of_barter = tkinter.Label(tab3, text="Описание объекта мены:")
    object_of_barter.grid(column=0, row=9, sticky=tk.E)
    object_of_barter_text = scrolledtext.ScrolledText(tab3, width=38, height=14)
    object_of_barter_text.grid(column=1, row=9, sticky=tk.W)
    object_of_barter_help = tkinter.Label(tab3, text="Указываются спецификация товара,\n"
                                  "кто из покупателей должен доплатить и т.д.")
    object_of_barter_help.grid(column=2, row=9, sticky=tk.W)
    barter_btn_help = tk.Button(tab3, text="Судебная практика по типу документа",
                        command=lambda: webbrowser.open("https://barrister365-my.sharepoint.com/:f:/g/personal/yan_barrister365_onmicrosoft_com/EgNaGWKtEwJDvyJWX3wE4A8BNL-GnqEqHRFjU_8iZOxDPQ?e=g7MgS1"))
    barter_btn_help.grid(column=0, row=11)
    dir_for_barter = tkinter.Label(tab3, text="Директория для сохранения файла:")
    dir_for_barter.grid(column=0, row=10, sticky=tk.E)
    stnd_for_barter_dir = tk.StringVar()
    stnd_for_barter_dir.set(r"D:\Договор мены1.docx")
    dir_for_barter_text = tkinter.Entry(tab3, width=45, textvariable = stnd_for_barter_dir, state='normal')
    dir_for_barter_text.grid(column=1, row=10, sticky=tk.W)
    dir_for_barter_text.focus()
    btnsave_barter = tk.Button(tab3, text="Сохранить готовый документ",
                        command=(lambda:
                                 save_doc(source='Договор_мены.docx',
                                          a=barter_date_entry_text.get(),
                                          b=place_of_barter_text.get(),
                                          c=first_seller_text.get(),
                                          d=second_seller_text.get(),
                                          e=barter_price_text.get(),
                                          f=expire_date_of_barter_text.get(),
                                          g=detail_of_first_seller_text.get("1.0", tk.END),
                                          h=detail_of_second_seller_text.get("1.0", tk.END),
                                          i=object_of_barter_text.get("1.0", tk.END),
                                          saving_path=dir_for_barter_text.get())
                                 )
                        )
    btnsave_barter.grid(column=1, row=11)
    btnexitsale_barter = tk.Button(tab3, text="Закрыть вкладку", command=lambda: tab_control.forget(tab3))
    btnexitsale_barter.grid(column=2, row=11)


def donations():
    tab4 = ttk.Frame(tab_control)
    tab_control.add(tab4, text='Договор дарения')
    tab4.bind('<Control-v>', pyperclip.paste())
    donations_date = tkinter.Label(tab4, text="Введите дату договора:")
    donations_date.grid(column=0, row=1, sticky=tk.E)
    donations_date_entry = tk.StringVar()
    donations_date_entry_text = tkinter.Entry(tab4, width=45, textvariable=donations_date_entry)
    donations_date_entry_text.grid(column=1, row=1, sticky=tk.W)
    donations_date_entry_text.focus()
    donations_date_entry_help = tkinter.Label(tab4, text="Например: 01.01.2022 или 1 января 2022 года")
    donations_date_entry_help.grid(column=2, row=1, sticky=tk.W)
    place_of_donations = tkinter.Label(tab4, text="Введите место составления договора:")
    place_of_donations.grid(column=0, row=2, sticky=tk.E)
    place_of_donations_entry = tk.StringVar()
    place_of_donations_text = tkinter.Entry(tab4, width=45, textvariable=place_of_donations_entry)
    place_of_donations_text.grid(column=1, row=2, sticky=tk.W)
    place_of_donations_text.focus()
    place_of_donations_entry_help = tkinter.Label(tab4, text="Например: Москва")
    place_of_donations_entry_help.grid(column=2, row=2, sticky=tk.W)
    gifted = tkinter.Label(tab4, text="Одаряемый (лицо, действующее от его имени на основании документа):")
    gifted.grid(column=0, row=3, sticky=tk.E)
    gifted_entry = tk.StringVar()
    gifted_text = tkinter.Entry(tab4, width=45, textvariable=gifted_entry)
    gifted_text.grid(column=1, row=3, sticky=tk.W)
    gifted_text.focus()
    gifted_entry_help = tkinter.Label(tab4, text="Например: ООО Ромашка, в лице директора ...")
    gifted_entry_help.grid(column=2, row=3, sticky=tk.W)
    gifter = tkinter.Label(tab4, text="Даритель (лицо, действующее от его имени на основании документа):")
    gifter.grid(column=0, row=4, sticky=tk.E)
    gifter_entry_to = tk.StringVar()
    gifter_text = tkinter.Entry(tab4, width=45, textvariable=gifter_entry_to)
    gifter_text.grid(column=1, row=4, sticky=tk.W)
    gifter_text.focus()
    expire_date_of_donations = tkinter.Label(tab4, text="Введите дату окончания срока действия договора:")
    expire_date_of_donations.grid(column=0, row=5, sticky=tk.E)
    expire_date_of_donations_entry = tk.StringVar
    expire_date_of_donations_text = tkinter.Entry(tab4, width=45, textvariable=expire_date_of_donations_entry)
    expire_date_of_donations_text.grid(column=1, row=5, sticky=tk.W)
    expire_date_of_donations_text.focus()
    detail_of_gifted = tkinter.Label(tab4, text="Банковские реквизиты одаряемого:")
    detail_of_gifted.grid(column=0, row=6, sticky=tk.E)
    detail_of_gifted_text = scrolledtext.ScrolledText(tab4, width=35, height=6)
    detail_of_gifted_text.grid(column=1, row=6, sticky=tk.W)
    detail_of_gifted_help = tkinter.Label(tab4, text="Наименование, ИНН, ОГРН, БИК Банка, р\с ...")
    detail_of_gifted_help.grid(column=2, row=6, sticky=tk.W)
    detail_of_gifter = tkinter.Label(tab4, text="Банковские реквизиты дарителя:")
    detail_of_gifter.grid(column=0, row=7, sticky=tk.E)
    detail_of_gifter_text = scrolledtext.ScrolledText(tab4, width=35, height=6)
    detail_of_gifter_text.grid(column=1, row=7, sticky=tk.W)
    object_of_donations = tkinter.Label(tab4, text="Описание дара:")
    object_of_donations.grid(column=0, row=8, sticky=tk.E)
    object_of_donations_text = scrolledtext.ScrolledText(tab4, width=38, height=14)
    object_of_donations_text.grid(column=1, row=8, sticky=tk.W)
    object_of_donations_help = tkinter.Label(tab4, text="Указываются характеристики дара")
    object_of_donations_help.grid(column=2, row=8, sticky=tk.W)
    donations_btn_help = tk.Button(tab4, text="Судебная практика по типу документа",
                        command=lambda: webbrowser.open("https://barrister365-my.sharepoint.com/:f:/g/personal/yan_barrister365_onmicrosoft_com/EqwQOgHfKvZFq-9kbVIC2v4B2noG-oGEzQU7z-wf6lFswQ?e=BPcbSl"))
    donations_btn_help.grid(column=0, row=10)
    dir_for_donations = tkinter.Label(tab4, text="Директория для сохранения файла:")
    dir_for_donations.grid(column=0, row=9, sticky=tk.E)
    stnd_for_donations_dir = tk.StringVar()
    stnd_for_donations_dir.set(r"D:\Договор дарения1.docx")
    dir_for_donations_text = tkinter.Entry(tab4, width=45, textvariable = stnd_for_donations_dir, state='normal')
    dir_for_donations_text.grid(column=1, row=9, sticky=tk.W)
    dir_for_donations_text.focus()
    btnsave_donations = tk.Button(tab4, text="Сохранить готовый документ",
                        command=(lambda:
                                 save_doc(source='Договор_дарения.docx',
                                          a=donations_date_entry_text.get(),
                                          b=place_of_donations_text.get(),
                                          c=gifted_text.get(),
                                          d=gifter_text.get(),
                                          e=None,
                                          f=expire_date_of_donations_text.get(),
                                          g=detail_of_gifted_text.get("1.0", tk.END),
                                          h=detail_of_gifter_text.get("1.0", tk.END),
                                          i=object_of_donations_text.get("1.0", tk.END),
                                          saving_path=dir_for_donations_text.get())
                                 )
                        )
    btnsave_donations.grid(column=1, row=10)
    btnexit_donations = tk.Button(tab4, text="Закрыть вкладку", command=lambda: tab_control.forget(tab4))
    btnexit_donations.grid(column=2, row=10)


def lease():
    tab5 = ttk.Frame(tab_control)
    tab_control.add(tab5, text='Договор аренды нежилого помещения')
    tab5.bind('<Control-v>', pyperclip.paste())
    lease_date = tkinter.Label(tab5, text="Введите дату договора:")
    lease_date.grid(column=0, row=1, sticky=tk.E)
    lease_date_entry = tk.StringVar()
    lease_date_entry_text = tkinter.Entry(tab5, width=45, textvariable=lease_date_entry)
    lease_date_entry_text.grid(column=1, row=1, sticky=tk.W)
    lease_date_entry_text.focus()
    lease_date_entry_help = tkinter.Label(tab5, text="Например: 01.01.2022 или 1 января 2022 года")
    lease_date_entry_help.grid(column=2, row=1, sticky=tk.W)
    place_of_lease = tkinter.Label(tab5, text="Введите место составления договора:")
    place_of_lease.grid(column=0, row=2, sticky=tk.E)
    place_of_lease_entry = tk.StringVar()
    place_of_lease_entry_text = tkinter.Entry(tab5, width=45, textvariable=place_of_lease_entry)
    place_of_lease_entry_text.grid(column=1, row=2, sticky=tk.W)
    place_of_lease_entry_text.focus()
    place_of_lease_entry_help = tkinter.Label(tab5, text="Например: Москва")
    place_of_lease_entry_help.grid(column=2, row=2, sticky=tk.W)
    tenant = tkinter.Label(tab5, text="Арендатор (лицо, действующее от его имени на основании документа):")
    tenant.grid(column=0, row=3, sticky=tk.E)
    tenant_entry = tk.StringVar()
    tenant_text = tkinter.Entry(tab5, width=45, textvariable=tenant_entry)
    tenant_text.grid(column=1, row=3, sticky=tk.W)
    tenant_text.focus()
    tenant_entry_help = tkinter.Label(tab5, text="Например: ООО Ромашка, в лице директора ...")
    tenant_entry_help.grid(column=2, row=3, sticky=tk.W)
    landlord = tkinter.Label(tab5, text="Арендодатель (лицо, действующее от его имени на основании документа):")
    landlord.grid(column=0, row=4, sticky=tk.E)
    landlord_entry = tk.StringVar()
    landlord_text = tkinter.Entry(tab5, width=45, textvariable=landlord_entry)
    landlord_text.grid(column=1, row=4, sticky=tk.W)
    landlord_text.focus()
    lease_price = tkinter.Label(tab5, text="Введите стоимость аренды:")
    lease_price.grid(column=0, row=5, sticky=tk.E)
    lease_price_entry = tk.StringVar()
    lease_price_text = tkinter.Entry(tab5, width=45, textvariable=lease_price_entry)
    lease_price_text.grid(column=1, row=5, sticky=tk.W)
    lease_price_text.focus()
    lease_price_help= tkinter.Label(tab5, text="Например: 100 000 рублей 00 коп. или 100 000 руб.\n"
                                "Цена обязательно указывается с НДС или без НДС,\n"
                                "также необходимо указать на каком основании\n"
                                "НДС не применяется.")
    lease_price_help.grid(column=2, row=5, sticky=tk.W)
    expire_date_of_lease = tkinter.Label(tab5, text="Введите дату окончания срока действия договора:")
    expire_date_of_lease.grid(column=0, row=6, sticky=tk.E)
    expire_date_of_lease_entry = tk.StringVar
    expire_date_of_lease_text = tkinter.Entry(tab5, width=45, textvariable=expire_date_of_lease_entry)
    expire_date_of_lease_text.grid(column=1, row=6, sticky=tk.W)
    expire_date_of_lease_text.focus()
    detail_of_tenant = tkinter.Label(tab5, text="Банковские реквизиты арендатора:")
    detail_of_tenant.grid(column=0, row=7, sticky=tk.E)
    detail_of_tenant_text = scrolledtext.ScrolledText(tab5, width=35, height=6)
    detail_of_tenant_text.grid(column=1, row=7, sticky=tk.W)
    detail_of_tenant_help = tkinter.Label(tab5, text="Наименование, ИНН, ОГРН, БИК Банка, р\с ...")
    detail_of_tenant_help.grid(column=2, row=7, sticky=tk.W)
    detail_of_landlord = tkinter.Label(tab5, text="Банковские реквизиты арендодателя:")
    detail_of_landlord.grid(column=0, row=8, sticky=tk.E)
    detail_of_landlord_text = scrolledtext.ScrolledText(tab5, width=35, height=6)
    detail_of_landlord_text.grid(column=1, row=8, sticky=tk.W)
    object_of_lease = tkinter.Label(tab5, text="Описание объекта мены:")
    object_of_lease.grid(column=0, row=9, sticky=tk.E)
    object_of_lease_text = scrolledtext.ScrolledText(tab5, width=38, height=14)
    object_of_lease_text.grid(column=1, row=9, sticky=tk.W)
    object_of_barter_help = tkinter.Label(tab5, text="Указываются адрес помещения, кадастровый номер,\n"
                                             "техническое описание, порядок пользования\n"
                                             "инженерными сетями и т.д.")
    object_of_barter_help.grid(column=2, row=9, sticky=tk.W)
    lease_btn_help = tk.Button(tab5, text="Судебная практика по типу документа",
                        command=lambda: webbrowser.open("https://barrister365-my.sharepoint.com/:f:/g/personal/yan_barrister365_onmicrosoft_com/EvNNqPxgTKtBiPNqyJgwtTIBlOVnfV42Sc50S4yABd_JsA?e=Bw7bAq"))
    lease_btn_help.grid(column=0, row=11)
    dir_for_lease = tkinter.Label(tab5, text="Директория для сохранения файла:")
    dir_for_lease.grid(column=0, row=10, sticky=tk.E)
    stnd_for_lease_dir = tk.StringVar()
    stnd_for_lease_dir.set(r"D:\Договор аренды нежилого помещения1.docx")
    dir_for_lease_text = tkinter.Entry(tab5, width=45, textvariable = stnd_for_lease_dir, state='normal')
    dir_for_lease_text.grid(column=1, row=10, sticky=tk.W)
    dir_for_lease_text.focus()
    btnsave_lease = tk.Button(tab5, text="Сохранить готовый документ",
                        command=(lambda:
                                 save_doc(source='Договор_Аренды.docx',
                                          a=lease_date_entry_text.get(),
                                          b=place_of_lease_entry_text.get(),
                                          c=tenant_text.get(),
                                          d=landlord_text.get(),
                                          e=lease_price_text.get(),
                                          f=expire_date_of_lease_text.get(),
                                          g=detail_of_tenant_text.get("1.0", tk.END),
                                          h=detail_of_landlord_text.get("1.0", tk.END),
                                          i=object_of_lease_text.get("1.0", tk.END),
                                          saving_path=dir_for_lease_text.get())
                                 )
                        )
    btnsave_lease.grid(column=1, row=11)
    btnexitsale_lease = tk.Button(tab5, text="Закрыть вкладку", command=lambda: tab_control.forget(tab5))
    btnexitsale_lease.grid(column=2, row=11)


def podryad():
    tab6 = ttk.Frame(tab_control)
    tab_control.add(tab6, text='Договор подряда')
    tab6.bind('<Control-v>', pyperclip.paste())
    podryad_date = tkinter.Label(tab6, text="Введите дату договора:")
    podryad_date.grid(column=0, row=1, sticky=tk.E)
    podryad_date_entry = tk.StringVar()
    podryad_date_entry_text = tkinter.Entry(tab6, width=45, textvariable=podryad_date_entry)
    podryad_date_entry_text.grid(column=1, row=1, sticky=tk.W)
    podryad_date_entry_text.focus()
    podryad_date_entry_help = tkinter.Label(tab6, text="Например: 01.01.2022 или 1 января 2022 года")
    podryad_date_entry_help.grid(column=2, row=1, sticky=tk.W)
    place_of_podryad = tkinter.Label(tab6, text="Введите место составления договора:")
    place_of_podryad.grid(column=0, row=2, sticky=tk.E)
    place_of_podryad_entry = tk.StringVar()
    place_of_podryad_entry_text = tkinter.Entry(tab6, width=45, textvariable=place_of_podryad_entry)
    place_of_podryad_entry_text.grid(column=1, row=2, sticky=tk.W)
    place_of_podryad_entry_text.focus()
    place_of_podryad_entry_help = tkinter.Label(tab6, text="Например: Москва")
    place_of_podryad_entry_help.grid(column=2, row=2, sticky=tk.W)
    zakazchik = tkinter.Label(tab6, text="Заказчик (лицо, действующее от его имени на основании документа):")
    zakazchik.grid(column=0, row=3, sticky=tk.E)
    zakazchik_entry = tk.StringVar()
    zakazchik_text = tkinter.Entry(tab6, width=45, textvariable=zakazchik_entry)
    zakazchik_text.grid(column=1, row=3, sticky=tk.W)
    zakazchik_text.focus()
    zakazchik_entry_help = tkinter.Label(tab6, text="Например: ООО Ромашка, в лице директора ...")
    zakazchik_entry_help.grid(column=2, row=3, sticky=tk.W)
    podryadchik = tkinter.Label(tab6, text="Подрядчик (лицо, действующее от его имени на основании документа):")
    podryadchik.grid(column=0, row=4, sticky=tk.E)
    podryadchik_entry = tk.StringVar()
    podryadchik_text = tkinter.Entry(tab6, width=45, textvariable=podryadchik_entry)
    podryadchik_text.grid(column=1, row=4, sticky=tk.W)
    podryadchik_text.focus()
    podryad_price = tkinter.Label(tab6, text="Введите цену договора:")
    podryad_price.grid(column=0, row=5, sticky=tk.E)
    podryad_price_entry = tk.StringVar()
    podryad_price_text = tkinter.Entry(tab6, width=45, textvariable=podryad_price_entry)
    podryad_price_text.grid(column=1, row=5, sticky=tk.W)
    podryad_price_text.focus()
    podryad_price_help= tkinter.Label(tab6, text="Например: 100 000 рублей 00 коп. или 100 000 руб.\n"
                                "Цена обязательно указывается с НДС или без НДС,\n"
                                "также необходимо указать на каком основании\n"
                                "НДС не применяется.")
    podryad_price_help.grid(column=2, row=5, sticky=tk.W)
    expire_date_of_podryad = tkinter.Label(tab6, text="Введите дату окончания срока действия договора:")
    expire_date_of_podryad.grid(column=0, row=6, sticky=tk.E)
    expire_date_of_podryad_entry = tk.StringVar
    expire_date_of_podryad_text = tkinter.Entry(tab6, width=45, textvariable=expire_date_of_podryad_entry)
    expire_date_of_podryad_text.grid(column=1, row=6, sticky=tk.W)
    expire_date_of_podryad_text.focus()
    detail_of_zakazchik = tkinter.Label(tab6, text="Банковские реквизиты заказчика:")
    detail_of_zakazchik.grid(column=0, row=7, sticky=tk.E)
    detail_of_zakazchik_text = scrolledtext.ScrolledText(tab6, width=35, height=6)
    detail_of_zakazchik_text.grid(column=1, row=7, sticky=tk.W)
    detail_of_zakazchik_help = tkinter.Label(tab6, text="Наименование, ИНН, ОГРН, БИК Банка, р\с ...")
    detail_of_zakazchik_help.grid(column=2, row=7, sticky=tk.W)
    detail_of_podryadchik = tkinter.Label(tab6, text="Банковские реквизиты подрядчика:")
    detail_of_podryadchik.grid(column=0, row=8, sticky=tk.E)
    detail_of_podryadchik_text = scrolledtext.ScrolledText(tab6, width=35, height=6)
    detail_of_podryadchik_text.grid(column=1, row=8, sticky=tk.W)
    object_of_podryad = tkinter.Label(tab6, text="Техническое задание:")
    object_of_podryad.grid(column=0, row=9, sticky=tk.E)
    object_of_podryad_text = scrolledtext.ScrolledText(tab6, width=38, height=14)
    object_of_podryad_text.grid(column=1, row=9, sticky=tk.W)
    object_of_podryad_help = tkinter.Label(tab6, text="Указываются спецификации объекта подряда,\n"
                                              "подробное описание: что конкретно необходимо\n"
                                              "сделать, этапы и т.д.")
    object_of_podryad_help.grid(column=2, row=9, sticky=tk.W)
    podryad_btn_help = tk.Button(tab6, text="Судебная практика по типу документа",
                        command=lambda: webbrowser.open("https://barrister365-my.sharepoint.com/:f:/g/personal/yan_barrister365_onmicrosoft_com/EtOVeNUj9khCvlz2Inuyc18B4Sk52Dy-HnECxS-TmO79Vg?e=eAMtZt"))
    podryad_btn_help.grid(column=0, row=11)
    dir_for_podryad = tkinter.Label(tab6, text="Директория для сохранения файла:")
    dir_for_podryad.grid(column=0, row=10, sticky=tk.E)
    stnd_for_podryad_dir = tk.StringVar()
    stnd_for_podryad_dir.set(r"D:\Договор подряда1.docx")
    dir_for_podryad_text = tkinter.Entry(tab6, width=45, textvariable = stnd_for_podryad_dir, state='normal')
    dir_for_podryad_text.grid(column=1, row=10, sticky=tk.W)
    dir_for_podryad_text.focus()
    btnsave_podryad = tk.Button(tab6, text="Сохранить готовый документ",
                        command=(lambda:
                                 save_doc(source='Договор_подряда.docx',
                                          a= podryad_date_entry_text.get(),
                                          b=place_of_podryad_entry_text.get(),
                                          c=zakazchik_text.get(),
                                          d=podryadchik_text.get(),
                                          e=podryad_price_text.get(),
                                          f=expire_date_of_podryad_text.get(),
                                          g=detail_of_zakazchik_text.get("1.0", tk.END),
                                          h=detail_of_podryadchik_text.get("1.0", tk.END),
                                          i=object_of_podryad_text.get("1.0", tk.END),
                                          saving_path=dir_for_podryad_text.get())
                                 )
                        )
    btnsave_podryad.grid(column=1, row=11)
    btnexitsale_podryad = tk.Button(tab6, text="Закрыть вкладку", command=lambda: tab_control.forget(tab6))
    btnexitsale_podryad.grid(column=2, row=11)


def service():
    tab7 = ttk.Frame(tab_control)
    tab_control.add(tab7, text='Договор возмездного оказания услуг')
    tab7.bind('<Control-v>', pyperclip.paste())
    service_date = tkinter.Label(tab7, text="Введите дату договора:")
    service_date.grid(column=0, row=1, sticky=tk.E)
    service_date_entry = tk.StringVar()
    service_date_entry_text = tkinter.Entry(tab7, width=45, textvariable=service_date_entry)
    service_date_entry_text.grid(column=1, row=1, sticky=tk.W)
    service_date_entry_text.focus()
    service_date_entry_help = tkinter.Label(tab7, text="Например: 01.01.2022 или 1 января 2022 года")
    service_date_entry_help.grid(column=2, row=1, sticky=tk.W)
    place_of_service = tkinter.Label(tab7, text="Введите место составления договора:")
    place_of_service.grid(column=0, row=2, sticky=tk.E)
    place_of_service_entry = tk.StringVar()
    place_of_service_entry_text = tkinter.Entry(tab7, width=45, textvariable=place_of_service_entry)
    place_of_service_entry_text.grid(column=1, row=2, sticky=tk.W)
    place_of_service_entry_text.focus()
    place_of_service_entry_help = tkinter.Label(tab7, text="Например: Москва")
    place_of_service_entry_help.grid(column=2, row=2, sticky=tk.W)
    customer = tkinter.Label(tab7, text="Заказчик (лицо, действующее от его имени на основании документа):")
    customer.grid(column=0, row=3, sticky=tk.E)
    customer_entry = tk.StringVar()
    customer_text = tkinter.Entry(tab7, width=45, textvariable=customer_entry)
    customer_text.grid(column=1, row=3, sticky=tk.W)
    customer_text.focus()
    customer_text_entry_help = tkinter.Label(tab7, text="Например: ООО Ромашка, в лице директора ...")
    customer_text_entry_help.grid(column=2, row=3, sticky=tk.W)
    executor = tkinter.Label(tab7, text="Исполнитель (лицо, действующее от его имени на основании документа):")
    executor.grid(column=0, row=4, sticky=tk.E)
    executor_entry = tk.StringVar()
    executor_text = tkinter.Entry(tab7, width=45, textvariable=executor_entry)
    executor_text.grid(column=1, row=4, sticky=tk.W)
    executor_text.focus()
    service_price = tkinter.Label(tab7, text="Введите цену договора:")
    service_price.grid(column=0, row=5, sticky=tk.E)
    service_price_entry = tk.StringVar()
    service_price_text = tkinter.Entry(tab7, width=45, textvariable=service_price_entry)
    service_price_text.grid(column=1, row=5, sticky=tk.W)
    service_price_text.focus()
    service_price_help= tkinter.Label(tab7, text="Например: 100 000 рублей 00 коп. или 100 000 руб.\n"
                                "Цена обязательно указывается с НДС или без НДС,\n"
                                "также необходимо указать на каком основании\n"
                                "НДС не применяется.")
    service_price_help.grid(column=2, row=5, sticky=tk.W)
    expire_date_of_service = tkinter.Label(tab7, text="Введите дату окончания срока действия договора:")
    expire_date_of_service.grid(column=0, row=6, sticky=tk.E)
    expire_date_of_service_entry = tk.StringVar
    expire_date_of_service_entry_text = tkinter.Entry(tab7, width=45, textvariable=expire_date_of_service_entry)
    expire_date_of_service_entry_text.grid(column=1, row=6, sticky=tk.W)
    expire_date_of_service_entry_text.focus()
    detail_of_customer = tkinter.Label(tab7, text="Банковские реквизиты заказчика:")
    detail_of_customer.grid(column=0, row=7, sticky=tk.E)
    detail_of_customer_text = scrolledtext.ScrolledText(tab7, width=35, height=6)
    detail_of_customer_text.grid(column=1, row=7, sticky=tk.W)
    detail_of_customer_text_help = tkinter.Label(tab7, text="Наименование, ИНН, ОГРН, БИК Банка, р\с ...")
    detail_of_customer_text_help.grid(column=2, row=7, sticky=tk.W)
    detail_of_executor = tkinter.Label(tab7, text="Банковские реквизиты исполнителя:")
    detail_of_executor.grid(column=0, row=8, sticky=tk.E)
    detail_of_executor_text = scrolledtext.ScrolledText(tab7, width=35, height=6)
    detail_of_executor_text.grid(column=1, row=8, sticky=tk.W)
    object_of_service = tkinter.Label(tab7, text="Техническое задание:")
    object_of_service.grid(column=0, row=9, sticky=tk.E)
    object_of_service_text = scrolledtext.ScrolledText(tab7, width=38, height=14)
    object_of_service_text.grid(column=1, row=9, sticky=tk.W)
    object_of_service_help = tkinter.Label(tab7, text="Указываются спецификации заказа,подробное\n"
                                              "описание: что конкретно необходимо\n"
                                              "сделать, этапы и т.д.")
    object_of_service_help.grid(column=2, row=9, sticky=tk.W)
    service_btn_help = tk.Button(tab7, text="Судебная практика по типу документа",
                        command=lambda: webbrowser.open("https://barrister365-my.sharepoint.com/:f:/g/personal/yan_barrister365_onmicrosoft_com/EvT2f3QaklBDn8np-a5NgKIBnXFqQOpyr06xdPodddW2jw?e=iumvCn"))
    service_btn_help.grid(column=0, row=11)
    dir_for_service = tkinter.Label(tab7, text="Директория для сохранения файла:")
    dir_for_service.grid(column=0, row=10, sticky=tk.E)
    stnd_for_service_dir = tk.StringVar()
    stnd_for_service_dir.set(r"D:\Договор возмездного оказания услуг1.docx")
    dir_for_service_text = tkinter.Entry(tab7, width=45, textvariable = stnd_for_service_dir, state='normal')
    dir_for_service_text.grid(column=1, row=10, sticky=tk.W)
    dir_for_service_text.focus()
    btnsave_service = tk.Button(tab7, text="Сохранить готовый документ",
                        command=(lambda:
                                 save_doc(source='Договор_возмездного_оказания_услуг.docx',
                                          a=service_date_entry_text.get(),
                                          b=place_of_service_entry_text.get(),
                                          c=customer_text.get(),
                                          d=executor_text.get(),
                                          e=service_price_text.get(),
                                          f=expire_date_of_service_entry_text.get(),
                                          g=detail_of_customer_text.get("1.0", tk.END),
                                          h=detail_of_executor_text.get("1.0", tk.END),
                                          i=object_of_service_text.get("1.0", tk.END),
                                          saving_path=dir_for_service_text.get())
                                 )
                        )
    btnsave_service.grid(column=1, row=11)
    btnexitsale_service = tk.Button(tab7, text="Закрыть вкладку", command=lambda: tab_control.forget(tab7))
    btnexitsale_service.grid(column=2, row=11)


def loan():
    tab8 = ttk.Frame(tab_control)
    tab_control.add(tab8, text='Договор займа')
    tab8.bind('<Control-v>', pyperclip.paste())
    loan_date = tkinter.Label(tab8, text="Введите дату договора:")
    loan_date.grid(column=0, row=1, sticky=tk.E)
    loan_date_entry = tk.StringVar()
    loan_date_entry_text = tkinter.Entry(tab8, width=45, textvariable=loan_date_entry)
    loan_date_entry_text.grid(column=1, row=1, sticky=tk.W)
    loan_date_entry_text.focus()
    loan_date_entry_help = tkinter.Label(tab8, text="Например: 01.01.2022 или 1 января 2022 года")
    loan_date_entry_help.grid(column=2, row=1, sticky=tk.W)
    place_of_loan = tkinter.Label(tab8, text="Введите место составления договора:")
    place_of_loan.grid(column=0, row=2, sticky=tk.E)
    place_of_loan_entry = tk.StringVar()
    place_of_loan_entry_text = tkinter.Entry(tab8, width=45, textvariable=place_of_loan_entry)
    place_of_loan_entry_text.grid(column=1, row=2, sticky=tk.W)
    place_of_loan_entry_text.focus()
    place_of_loan_entry_help = tkinter.Label(tab8, text="Например: Москва")
    place_of_loan_entry_help.grid(column=2, row=2, sticky=tk.W)
    borrower = tkinter.Label(tab8, text="Заёмщик (лицо, действующее от его имени на основании документа):")
    borrower.grid(column=0, row=3, sticky=tk.E)
    borrower_entry = tk.StringVar()
    borrower_text = tkinter.Entry(tab8, width=45, textvariable=borrower_entry)
    borrower_text.grid(column=1, row=3, sticky=tk.W)
    borrower_text.focus()
    borrower_text_entry_help = tkinter.Label(tab8, text="Например: ООО Ромашка, в лице директора ...")
    borrower_text_entry_help.grid(column=2, row=3, sticky=tk.W)
    lender = tkinter.Label(tab8, text="Займодавец (лицо, действующее от его имени на основании документа):")
    lender.grid(column=0, row=4, sticky=tk.E)
    lender_entry = tk.StringVar()
    lender_text = tkinter.Entry(tab8, width=45, textvariable=lender_entry)
    lender_text.grid(column=1, row=4, sticky=tk.W)
    lender_text.focus()
    loan_sum = tkinter.Label(tab8, text="Введите сумму займа:")
    loan_sum.grid(column=0, row=5, sticky=tk.E)
    loan_sum_entry = tk.StringVar()
    loan_sum_text = tkinter.Entry(tab8, width=45, textvariable=loan_sum_entry)
    loan_sum_text.grid(column=1, row=5, sticky=tk.W)
    loan_sum_text.focus()
    loan_sum_help = tkinter.Label(tab8, text="Например: 100 000 рублей 00 коп. или 100 000 руб.")
    loan_sum_help.grid(column=2, row=5, sticky=tk.W)
    expire_date_of_loan = tkinter.Label(tab8, text="Введите дату окончания срока действия договора:")
    expire_date_of_loan.grid(column=0, row=6, sticky=tk.E)
    expire_date_of_loan_entry = tk.StringVar
    expire_date_of_loan_entry_text = tkinter.Entry(tab8, width=45, textvariable=expire_date_of_loan_entry)
    expire_date_of_loan_entry_text.grid(column=1, row=6, sticky=tk.W)
    expire_date_of_loan_entry_text.focus()
    detail_of_borrower = tkinter.Label(tab8, text="Банковские реквизиты заёмщика:")
    detail_of_borrower.grid(column=0, row=7, sticky=tk.E)
    detail_of_borrower_text = scrolledtext.ScrolledText(tab8, width=35, height=6)
    detail_of_borrower_text.grid(column=1, row=7, sticky=tk.W)
    detail_of_borrower_text_help = tkinter.Label(tab8, text="Наименование, ИНН, ОГРН, БИК Банка, р\с ...")
    detail_of_borrower_text_help.grid(column=2, row=7, sticky=tk.W)
    detail_of_lender = tkinter.Label(tab8, text="Банковские реквизиты займодавца:")
    detail_of_lender.grid(column=0, row=8, sticky=tk.E)
    detail_of_lender_text = scrolledtext.ScrolledText(tab8, width=35, height=6)
    detail_of_lender_text.grid(column=1, row=8, sticky=tk.W)

    loam_btn_help = tk.Button(tab8, text="Судебная практика по типу документа",
                                 command=lambda: webbrowser.open("https://barrister365-my.sharepoint.com/:f:/g/personal/yan_barrister365_onmicrosoft_com/Eje4axCOLPlLuQIM3JHbySsB0dnthk3j53KdNMbiCc7loA?e=wd7yGh"))
    loam_btn_help.grid(column=0, row=11)
    dir_for_loan = tkinter.Label(tab8, text="Директория для сохранения файла:")
    dir_for_loan.grid(column=0, row=10, sticky=tk.E)
    stnd_for_loan_dir = tk.StringVar()
    stnd_for_loan_dir.set(r"D:\Договор займа1.docx")
    dir_for_loan_text = tkinter.Entry(tab8, width=45, textvariable=stnd_for_loan_dir, state='normal')
    dir_for_loan_text.grid(column=1, row=10, sticky=tk.W)
    dir_for_loan_text.focus()
    btnsave_loan = tk.Button(tab8, text="Сохранить готовый документ",
                                command=(lambda:
                                         save_doc(source='Договор_займа.docx',
                                                  a=loan_date_entry_text.get(),
                                                  b=place_of_loan_entry_text.get(),
                                                  c=borrower_text.get(),
                                                  d=lender_text.get(),
                                                  e=loan_sum_text.get(),
                                                  f=expire_date_of_loan_entry_text.get(),
                                                  g=detail_of_borrower_text.get("1.0", tk.END),
                                                  h=detail_of_lender_text.get("1.0", tk.END),
                                                  i=None,
                                                  saving_path=dir_for_loan_text.get())
                                         )
                                )
    btnsave_loan.grid(column=1, row=11)
    btnexitsale_loan = tk.Button(tab8, text="Закрыть вкладку", command=lambda: tab_control.forget(tab8))
    btnexitsale_loan.grid(column=2, row=11)


def commission():
    tab9 = ttk.Frame(tab_control)
    tab_control.add(tab9, text='Договор комиссии')
    tab9.bind('<Control-v>', pyperclip.paste())
    commission_date = tkinter.Label(tab9, text="Введите дату договора:")
    commission_date.grid(column=0, row=1, sticky=tk.E)
    commission_date_entry = tk.StringVar()
    commission_date_entry_text = tkinter.Entry(tab9, width=45, textvariable=commission_date_entry)
    commission_date_entry_text.grid(column=1, row=1, sticky=tk.W)
    commission_date_entry_text.focus()
    commission_date_entry_help = tkinter.Label(tab9, text="Например: 01.01.2022 или 1 января 2022 года")
    commission_date_entry_help.grid(column=2, row=1, sticky=tk.W)
    place_of_commission = tkinter.Label(tab9, text="Введите место составления договора:")
    place_of_commission.grid(column=0, row=2, sticky=tk.E)
    place_of_commission_entry = tk.StringVar()
    place_of_commission_entry_text = tkinter.Entry(tab9, width=45, textvariable=place_of_commission_entry)
    place_of_commission_entry_text.grid(column=1, row=2, sticky=tk.W)
    place_of_commission_entry_text.focus()
    place_of_commission_entry_help = tkinter.Label(tab9, text="Например: Москва")
    place_of_commission_entry_help.grid(column=2, row=2, sticky=tk.W)
    committee = tkinter.Label(tab9, text="Комитент (заказчик) (лицо, действующее от его имени на основании документа):")
    committee.grid(column=0, row=3, sticky=tk.E)
    committee_entry = tk.StringVar()
    committee_text = tkinter.Entry(tab9, width=45, textvariable=committee_entry)
    committee_text.grid(column=1, row=3, sticky=tk.W)
    committee_text.focus()
    committee_text_entry_help = tkinter.Label(tab9, text="Например: ООО Ромашка, в лице директора ...")
    committee_text_entry_help.grid(column=2, row=3, sticky=tk.W)
    commission_agent = tkinter.Label(tab9, text="Комиссионер (лицо, действующее от его имени на основании документа):")
    commission_agent.grid(column=0, row=4, sticky=tk.E)
    commission_agent_entry = tk.StringVar()
    commission_agent_text = tkinter.Entry(tab9, width=45, textvariable=commission_agent_entry)
    commission_agent_text.grid(column=1, row=4, sticky=tk.W)
    commission_agent_text.focus()
    commission_price = tkinter.Label(tab9, text="Введите цену договора:")
    commission_price.grid(column=0, row=5, sticky=tk.E)
    commission_price_entry = tk.StringVar()
    commission_price_text = tkinter.Entry(tab9, width=45, textvariable=commission_price_entry)
    commission_price_text.grid(column=1, row=5, sticky=tk.W)
    commission_price_text.focus()
    commission_price_text_help= tkinter.Label(tab9, text="Например: 100 000 рублей 00 коп. или 100 000 руб.\n"
                                "Цена обязательно указывается с НДС или без НДС,\n"
                                "также необходимо указать на каком основании\n"
                                "НДС не применяется.")
    commission_price_text_help.grid(column=2, row=5, sticky=tk.W)
    expire_date_of_commission = tkinter.Label(tab9, text="Введите дату окончания срока действия договора:")
    expire_date_of_commission.grid(column=0, row=6, sticky=tk.E)
    expire_date_of_commission_entry = tk.StringVar
    expire_date_of_commission_entry_text = tkinter.Entry(tab9, width=45, textvariable=expire_date_of_commission_entry)
    expire_date_of_commission_entry_text.grid(column=1, row=6, sticky=tk.W)
    expire_date_of_commission_entry_text.focus()
    detail_of_committee = tkinter.Label(tab9, text="Банковские реквизиты комитента:")
    detail_of_committee.grid(column=0, row=7, sticky=tk.E)
    detail_of_committee_text = scrolledtext.ScrolledText(tab9, width=35, height=6)
    detail_of_committee_text.grid(column=1, row=7, sticky=tk.W)
    detail_of_committee_text_help = tkinter.Label(tab9, text="Наименование, ИНН, ОГРН, БИК Банка, р\с ...")
    detail_of_committee_text_help.grid(column=2, row=7, sticky=tk.W)
    detail_of_commission_agent = tkinter.Label(tab9, text="Банковские реквизиты комиссионера:")
    detail_of_commission_agent.grid(column=0, row=8, sticky=tk.E)
    detail_of_commission_agent_text = scrolledtext.ScrolledText(tab9, width=35, height=6)
    detail_of_commission_agent_text.grid(column=1, row=8, sticky=tk.W)
    object_of_commission = tkinter.Label(tab9, text="Техническое задание:")
    object_of_commission.grid(column=0, row=9, sticky=tk.E)
    object_of_commission_text = scrolledtext.ScrolledText(tab9, width=38, height=14)
    object_of_commission_text.grid(column=1, row=9, sticky=tk.W)
    object_of_commission_help = tkinter.Label(tab9, text="Указываются спецификации заказа,подробное\n"
                                              "описание: что конкретно необходимо\n"
                                              "сделать, какие сделки необходимо заключить и т.д.")
    object_of_commission_help.grid(column=2, row=9, sticky=tk.W)
    commission_btn_help = tk.Button(tab9, text="Судебная практика по типу документа",
                        command=lambda: webbrowser.open("https://barrister365-my.sharepoint.com/:f:/g/personal/yan_barrister365_onmicrosoft_com/ElNvMeDMNIhIo6QMubJFnYgBTKFlud8RiZq6AT89VzHMuw?e=IP0sCl"))
    commission_btn_help.grid(column=0, row=11)
    dir_for_commission = tkinter.Label(tab9, text="Директория для сохранения файла:")
    dir_for_commission.grid(column=0, row=10, sticky=tk.E)
    stnd_for_commission_dir = tk.StringVar()
    stnd_for_commission_dir.set(r"D:\Договор комиссии1.docx")
    dir_for_commission_text = tkinter.Entry(tab9, width=45, textvariable = stnd_for_commission_dir, state='normal')
    dir_for_commission_text.grid(column=1, row=10, sticky=tk.W)
    dir_for_commission_text.focus()
    btnsave_commission = tk.Button(tab9, text="Сохранить готовый документ",
                        command=(lambda:
                                 save_doc(source='Договор_комиссии.docx',
                                          a=commission_date_entry_text.get(),
                                          b=place_of_commission_entry_text.get(),
                                          c=committee_text.get(),
                                          d=commission_agent_text.get(),
                                          e=commission_price_text.get(),
                                          f=expire_date_of_commission_entry_text.get(),
                                          g=detail_of_committee_text.get("1.0", tk.END),
                                          h=detail_of_commission_agent_text.get("1.0", tk.END),
                                          i=object_of_commission_text.get("1.0", tk.END),
                                          saving_path=dir_for_commission_text.get())
                                 )
                        )
    btnsave_commission.grid(column=1, row=11)
    btnexitsale_commission = tk.Button(tab9, text="Закрыть вкладку", command=lambda: tab_control.forget(tab9))
    btnexitsale_commission.grid(column=2, row=11)


def agency():
    tab10 = ttk.Frame(tab_control)
    tab_control.add(tab10, text='Агентский договор')
    tab10.bind('<Control-v>', pyperclip.paste())
    agency_date = tkinter.Label(tab10, text="Введите дату договора:")
    agency_date.grid(column=0, row=1, sticky=tk.E)
    agency_date_entry = tk.StringVar()
    agency_date_entry_text = tkinter.Entry(tab10, width=45, textvariable=agency_date_entry)
    agency_date_entry_text.grid(column=1, row=1, sticky=tk.W)
    agency_date_entry_text.focus()
    agency_date_entry_help = tkinter.Label(tab10, text="Например: 01.01.2022 или 1 января 2022 года")
    agency_date_entry_help.grid(column=2, row=1, sticky=tk.W)
    place_of_agency = tkinter.Label(tab10, text="Введите место составления договора:")
    place_of_agency.grid(column=0, row=2, sticky=tk.E)
    place_of_agency_entry = tk.StringVar()
    place_of_agency_entry_text = tkinter.Entry(tab10, width=45, textvariable=place_of_agency_entry)
    place_of_agency_entry_text.grid(column=1, row=2, sticky=tk.W)
    place_of_agency_entry_text.focus()
    place_of_agency_entry_help = tkinter.Label(tab10, text="Например: Москва")
    place_of_agency_entry_help.grid(column=2, row=2, sticky=tk.W)
    principal = tkinter.Label(tab10, text="Принципал (заказчик) (лицо, действующее от его имени на основании документа):")
    principal.grid(column=0, row=3, sticky=tk.E)
    principal_entry = tk.StringVar()
    principal_text = tkinter.Entry(tab10, width=45, textvariable=principal_entry)
    principal_text.grid(column=1, row=3, sticky=tk.W)
    principal_text.focus()
    principal_text_entry_help = tkinter.Label(tab10, text="Например: ООО Ромашка, в лице директора ...")
    principal_text_entry_help.grid(column=2, row=3, sticky=tk.W)
    agency_agent = tkinter.Label(tab10, text="Агент (лицо, действующее от его имени на основании документа):")
    agency_agent.grid(column=0, row=4, sticky=tk.E)
    agency_agent_entry = tk.StringVar()
    agency_agent_text = tkinter.Entry(tab10, width=45, textvariable=agency_agent_entry)
    agency_agent_text.grid(column=1, row=4, sticky=tk.W)
    agency_agent_text.focus()
    agency_price = tkinter.Label(tab10, text="Введите цену договора:")
    agency_price.grid(column=0, row=5, sticky=tk.E)
    agency_price_entry = tk.StringVar()
    agency_price_text = tkinter.Entry(tab10, width=45, textvariable=agency_price_entry)
    agency_price_text.grid(column=1, row=5, sticky=tk.W)
    agency_price_text.focus()
    agency_price_text_help= tkinter.Label(tab10, text="Например: 100 000 рублей 00 коп. или 100 000 руб.\n"
                                "Цена обязательно указывается с НДС или без НДС,\n"
                                "также необходимо указать на каком основании\n"
                                "НДС не применяется.")
    agency_price_text_help.grid(column=2, row=5, sticky=tk.W)
    expire_date_of_agency = tkinter.Label(tab10, text="Введите дату окончания срока действия договора:")
    expire_date_of_agency.grid(column=0, row=6, sticky=tk.E)
    expire_date_of_agency_entry = tk.StringVar
    expire_date_of_agency_entry_text = tkinter.Entry(tab10, width=45, textvariable=expire_date_of_agency_entry)
    expire_date_of_agency_entry_text.grid(column=1, row=6, sticky=tk.W)
    expire_date_of_agency_entry_text.focus()
    detail_of_principal = tkinter.Label(tab10, text="Банковские реквизиты принципала:")
    detail_of_principal.grid(column=0, row=7, sticky=tk.E)
    detail_of_principal_text = scrolledtext.ScrolledText(tab10, width=35, height=6)
    detail_of_principal_text.grid(column=1, row=7, sticky=tk.W)
    detail_of_principal_text_help = tkinter.Label(tab10, text="Наименование, ИНН, ОГРН, БИК Банка, р\с ...")
    detail_of_principal_text_help.grid(column=2, row=7, sticky=tk.W)
    detail_of_agency_agent = tkinter.Label(tab10, text="Банковские реквизиты агента:")
    detail_of_agency_agent.grid(column=0, row=8, sticky=tk.E)
    detail_of_agency_agent_text = scrolledtext.ScrolledText(tab10, width=35, height=6)
    detail_of_agency_agent_text.grid(column=1, row=8, sticky=tk.W)
    object_of_agency = tkinter.Label(tab10, text="Техническое задание:")
    object_of_agency.grid(column=0, row=9, sticky=tk.E)
    object_of_agency_text = scrolledtext.ScrolledText(tab10, width=38, height=14)
    object_of_agency_text.grid(column=1, row=9, sticky=tk.W)
    object_of_agency_help = tkinter.Label(tab10, text="Указываются спецификации заказа,подробное\n"
                                              "описание: что конкретно необходимо\n"
                                              "сделать, какие сделки необходимо заключить и т.д.")
    object_of_agency_help.grid(column=2, row=9, sticky=tk.W)
    agency_btn_help = tk.Button(tab10, text="Судебная практика по типу документа",
                        command=lambda: webbrowser.open("https://barrister365-my.sharepoint.com/:f:/g/personal/yan_barrister365_onmicrosoft_com/EvVHFDRBv15Hlis6zSaDSg4BL3ojkSi7Tg_TtyPbOrxbsA?e=rurf0U"))
    agency_btn_help.grid(column=0, row=11)
    dir_for_agency = tkinter.Label(tab10, text="Директория для сохранения файла:")
    dir_for_agency.grid(column=0, row=10, sticky=tk.E)
    stnd_for_agency_dir = tk.StringVar()
    stnd_for_agency_dir.set(r"D:\Агентский договор1.docx")
    dir_for_agency_text = tkinter.Entry(tab10, width=45, textvariable = stnd_for_agency_dir, state='normal')
    dir_for_agency_text.grid(column=1, row=10, sticky=tk.W)
    dir_for_agency_text.focus()
    btnsave_agency = tk.Button(tab10, text="Сохранить готовый документ",
                        command=(lambda:
                                 save_doc(source='Агентский договор.docx',
                                          a=agency_date_entry_text.get(),
                                          b=place_of_agency_entry_text.get(),
                                          c=principal_text.get(),
                                          d=agency_agent_text.get(),
                                          e=agency_price_text.get(),
                                          f=expire_date_of_agency_entry_text.get(),
                                          g=detail_of_principal_text.get("1.0", tk.END),
                                          h=detail_of_agency_agent_text.get("1.0", tk.END),
                                          i=object_of_agency_text.get("1.0", tk.END),
                                          saving_path=dir_for_agency_text.get())
                                 )
                        )
    btnsave_agency.grid(column=1, row=11)
    btnexitsale_agency = tk.Button(tab10, text="Закрыть вкладку", command=lambda: tab_control.forget(tab10))
    btnexitsale_agency.grid(column=2, row=11)


def licenseagr():
    tab11 = ttk.Frame(tab_control)
    tab_control.add(tab11, text='Лицензионный договор')
    tab11.bind('<Control-v>', pyperclip.paste())
    licenseagr_date = tkinter.Label(tab11, text="Введите дату договора:")
    licenseagr_date.grid(column=0, row=1, sticky=tk.E)
    licenseagr_date_entry = tk.StringVar()
    licenseagr_date_entry_text = tkinter.Entry(tab11, width=45, textvariable=licenseagr_date_entry)
    licenseagr_date_entry_text.grid(column=1, row=1, sticky=tk.W)
    licenseagr_date_entry_text.focus()
    licenseagr_date_entry_help = tkinter.Label(tab11, text="Например: 01.01.2022 или 1 января 2022 года")
    licenseagr_date_entry_help.grid(column=2, row=1, sticky=tk.W)
    place_of_licenseagr = tkinter.Label(tab11, text="Введите место составления договора:")
    place_of_licenseagr.grid(column=0, row=2, sticky=tk.E)
    place_of_licenseagr_entry = tk.StringVar()
    place_of_licenseagr_entry_text = tkinter.Entry(tab11, width=45, textvariable=place_of_licenseagr_entry)
    place_of_licenseagr_entry_text.grid(column=1, row=2, sticky=tk.W)
    place_of_licenseagr_entry_text.focus()
    place_of_licenseagr_entry_help = tkinter.Label(tab11, text="Например: Москва")
    place_of_licenseagr_entry_help.grid(column=2, row=2, sticky=tk.W)
    licensee = tkinter.Label(tab11, text="Лицензиат (заказчик) (лицо, действующее от его имени на основании документа):")
    licensee.grid(column=0, row=3, sticky=tk.E)
    licensee_entry = tk.StringVar()
    licensee_text = tkinter.Entry(tab11, width=45, textvariable=licensee_entry)
    licensee_text.grid(column=1, row=3, sticky=tk.W)
    licensee_text.focus()
    licensee_text_entry_help = tkinter.Label(tab11, text="Например: ООО Ромашка, в лице директора ...")
    licensee_text_entry_help.grid(column=2, row=3, sticky=tk.W)
    licensor = tkinter.Label(tab11, text="Лицензиар (лицо, действующее от его имени на основании документа):")
    licensor.grid(column=0, row=4, sticky=tk.E)
    licensor_entry = tk.StringVar()
    licensor_text = tkinter.Entry(tab11, width=45, textvariable=licensor_entry)
    licensor_text.grid(column=1, row=4, sticky=tk.W)
    licensor_text.focus()
    licenseagr_price = tkinter.Label(tab11, text="Введите цену договора:")
    licenseagr_price.grid(column=0, row=5, sticky=tk.E)
    licenseagr_price_entry = tk.StringVar()
    licenseagr_price_text = tkinter.Entry(tab11, width=45, textvariable=licenseagr_price_entry)
    licenseagr_price_text.grid(column=1, row=5, sticky=tk.W)
    licenseagr_price_text.focus()
    licenseagr_price_help= tkinter.Label(tab11, text="Например: 100 000 рублей 00 коп. или 100 000 руб.\n"
                                "Цена обязательно указывается с НДС или без НДС,\n"
                                "также необходимо указать на каком основании\n"
                                "НДС не применяется.")
    licenseagr_price_help.grid(column=2, row=5, sticky=tk.W)
    expire_date_of_licenseagr = tkinter.Label(tab11, text="Введите дату окончания срока действия договора:")
    expire_date_of_licenseagr.grid(column=0, row=6, sticky=tk.E)
    expire_date_of_licenseagr_entry = tk.StringVar
    expire_date_of_licenseagr_entry_text = tkinter.Entry(tab11, width=45, textvariable=expire_date_of_licenseagr_entry)
    expire_date_of_licenseagr_entry_text.grid(column=1, row=6, sticky=tk.W)
    expire_date_of_licenseagr_entry_text.focus()
    detail_of_licensee = tkinter.Label(tab11, text="Банковские реквизиты лицензиата:")
    detail_of_licensee.grid(column=0, row=7, sticky=tk.E)
    detail_of_licensee_text = scrolledtext.ScrolledText(tab11, width=35, height=6)
    detail_of_licensee_text.grid(column=1, row=7, sticky=tk.W)
    detail_of_licensee_text_help = tkinter.Label(tab11, text="Наименование, ИНН, ОГРН, БИК Банка, р\с ...")
    detail_of_licensee_text_help.grid(column=2, row=7, sticky=tk.W)
    detail_of_licensor = tkinter.Label(tab11, text="Банковские реквизиты лицензиара:")
    detail_of_licensor.grid(column=0, row=8, sticky=tk.E)
    detail_of_licensor_text = scrolledtext.ScrolledText(tab11, width=35, height=6)
    detail_of_licensor_text.grid(column=1, row=8, sticky=tk.W)
    object_of_licenseagr = tkinter.Label(tab11, text="Техническое описание РИД:")
    object_of_licenseagr.grid(column=0, row=9, sticky=tk.E)
    object_of_licenseagr_text = scrolledtext.ScrolledText(tab11, width=38, height=14)
    object_of_licenseagr_text.grid(column=1, row=9, sticky=tk.W)
    object_of_licenseagr_help = tkinter.Label(tab11, text="Указывается, что конкретно передаётся\n"
                                                  "в пользование по лицензионному договору,\n"
                                                  "спецификации и т.д.")
    object_of_licenseagr_help.grid(column=2, row=9, sticky=tk.W)
    licenseagr_btn_help = tk.Button(tab11, text="Судебная практика по типу документа",
                        command=lambda: webbrowser.open("https://barrister365-my.sharepoint.com/:f:/g/personal/yan_barrister365_onmicrosoft_com/EtTyU5oVAWtBhji-q463mqoBiElochb90khHjmqxFgOVVA?e=oH3lEg"))
    licenseagr_btn_help.grid(column=0, row=11)
    dir_for_licenseagr = tkinter.Label(tab11, text="Директория для сохранения файла:")
    dir_for_licenseagr.grid(column=0, row=10, sticky=tk.E)
    stnd_for_licenseagr_dir = tk.StringVar()
    stnd_for_licenseagr_dir.set(r"D:\Лицензионный договор1.docx")
    dir_for_licenseagr_text = tkinter.Entry(tab11, width=45, textvariable = stnd_for_licenseagr_dir, state='normal')
    dir_for_licenseagr_text.grid(column=1, row=10, sticky=tk.W)
    dir_for_licenseagr_text.focus()
    btnsave_licenseagr = tk.Button(tab11, text="Сохранить готовый документ",
                        command=(lambda:
                                 save_doc(source='Лицензионный_договор.docx',
                                          a=licenseagr_date_entry_text.get(),
                                          b=place_of_licenseagr_entry_text.get(),
                                          c=licensee_text.get(),
                                          d=licensor_text.get(),
                                          e=licenseagr_price_text.get(),
                                          f=expire_date_of_licenseagr_entry_text.get(),
                                          g=detail_of_licensee_text.get("1.0", tk.END),
                                          h=detail_of_licensor_text.get("1.0", tk.END),
                                          i=object_of_licenseagr_text.get("1.0", tk.END),
                                          saving_path=dir_for_licenseagr_text.get())
                                 )
                        )
    btnsave_licenseagr.grid(column=1, row=11)
    btnexitsale_licenseagr = tk.Button(tab11, text="Закрыть вкладку", command=lambda: tab_control.forget(tab11))
    btnexitsale_licenseagr.grid(column=2, row=11)


def court_order():
    tab12 = ttk.Frame(tab_control)
    tab_control.add(tab12, text='Заявление на выдачу судебного приказа')
    tab12.bind('<Control-v>', pyperclip.paste())
    order_date = tkinter.Label(tab12, text="Введите дату договора, на котором основаны требования:")
    order_date.grid(column=0, row=1, sticky=tk.E)
    order_date_entry = tk.StringVar()
    order_date_entry_text = tkinter.Entry(tab12, width=45, textvariable=order_date_entry)
    order_date_entry_text.grid(column=1, row=1, sticky=tk.W)
    order_date_entry_text.focus()
    order_date_entry_help = tkinter.Label(tab12, text="Например: 01.01.2022 или 1 января 2022 года")
    order_date_entry_help.grid(column=2, row=1, sticky=tk.W)
    place_of_order = tkinter.Label(tab12, text="Суд, в который обращается взыскатель:")
    place_of_order.grid(column=0, row=2, sticky=tk.E)
    place_of_order_entry = tk.StringVar()
    place_of_order_entry_text = tkinter.Entry(tab12, width=45, textvariable=place_of_order_entry)
    place_of_order_entry_text.grid(column=1, row=2, sticky=tk.W)
    place_of_order_entry_text.focus()
    place_of_order_entry_help = tkinter.Label(tab12, text="Например: Москва")
    place_of_order_entry_help.grid(column=2, row=2, sticky=tk.W)
    recoverer = tkinter.Label(tab12, text="Взыскатель (лицо, действующее от его имени на основании документа):")
    recoverer.grid(column=0, row=3, sticky=tk.E)
    recoverer_entry = tk.StringVar()
    recoverer_text = tkinter.Entry(tab12, width=45, textvariable=recoverer_entry)
    recoverer_text.grid(column=1, row=3, sticky=tk.W)
    recoverer_text.focus()
    recoverer_text_entry_help = tkinter.Label(tab12, text="Например: ООО Ромашка, в лице директора ...")
    recoverer_text_entry_help.grid(column=2, row=3, sticky=tk.W)
    debtor = tkinter.Label(tab12, text="Должник (лицо, действующее от его имени на основании документа):")
    debtor.grid(column=0, row=4, sticky=tk.E)
    debtor_entry = tk.StringVar()
    debtor_text = tkinter.Entry(tab12, width=45, textvariable=debtor_entry)
    debtor_text.grid(column=1, row=4, sticky=tk.W)
    debtor_text.focus()
    order_price = tkinter.Label(tab12, text="Введите цену договора (размер задолженности):")
    order_price.grid(column=0, row=5, sticky=tk.E)
    order_price_entry = tk.StringVar()
    order_price_text = tkinter.Entry(tab12, width=45, textvariable=order_price_entry)
    order_price_text.grid(column=1, row=5, sticky=tk.W)
    order_price_text.focus()
    order_price_help= tkinter.Label(tab12, text="Например: 100 000 рублей 00 коп. или 100 000 руб.\n"
                                "В этой же строке можно указать проценты на сумму долга\n")
    order_price_help.grid(column=2, row=5, sticky=tk.W)
    expire_date_of_order = tkinter.Label(tab12, text="Введите дату подписания заявления:")
    expire_date_of_order.grid(column=0, row=6, sticky=tk.E)
    expire_date_of_order_entry = tk.StringVar
    expire_date_of_order_entry_text = tkinter.Entry(tab12, width=45, textvariable=expire_date_of_order_entry)
    expire_date_of_order_entry_text.grid(column=1, row=6, sticky=tk.W)
    expire_date_of_order_entry_text.focus()
    detail_of_recoverer = tkinter.Label(tab12, text="Реквизиты взыскателя:")
    detail_of_recoverer.grid(column=0, row=7, sticky=tk.E)
    detail_of_recoverer_text = scrolledtext.ScrolledText(tab12, width=35, height=6)
    detail_of_recoverer_text.grid(column=1, row=7, sticky=tk.W)
    detail_of_recoverer_text_help = tkinter.Label(tab12, text="Наименование, ИНН, ОГРН, БИК Банка, р\с ...")
    detail_of_recoverer_text_help.grid(column=2, row=7, sticky=tk.W)
    detail_of_debtor = tkinter.Label(tab12, text="Реквизиты должника:")
    detail_of_debtor.grid(column=0, row=8, sticky=tk.E)
    detail_of_debtor_text = scrolledtext.ScrolledText(tab12, width=35, height=6)
    detail_of_debtor_text.grid(column=1, row=8, sticky=tk.W)
    object_of_order = tkinter.Label(tab12, text="Основание требования:")
    object_of_order.grid(column=0, row=9, sticky=tk.E)
    object_of_order_text = scrolledtext.ScrolledText(tab12, width=38, height=14)
    object_of_order_text.grid(column=1, row=9, sticky=tk.W)
    object_of_order_help = tkinter.Label(tab12, text="Описывается ситуация, которая\n"
                                             "привела к нарушению условий договора и\n"
                                             "повлекла задолженность, пункт договора,\n"
                                             "по которому взыскатель впарве обратится в суд\n"
                                             "и другие сведения, доказывающие возможность взыскания")
    object_of_order_help.grid(column=2, row=9, sticky=tk.W)
    order_btn_help = tk.Button(tab12, text="Судебная практика по типу документа",
                        command=lambda: webbrowser.open("https://barrister365-my.sharepoint.com/:f:/g/personal/yan_barrister365_onmicrosoft_com/Ela38_fB9IFJg3pRZoNk6l4BIEwoO4MO09nYqMPFWy21jQ?e=VoLVIM"))
    order_btn_help.grid(column=0, row=11)
    dir_for_order = tkinter.Label(tab12, text="Директория для сохранения файла:")
    dir_for_order.grid(column=0, row=10, sticky=tk.E)
    stnd_for_order_dir = tk.StringVar()
    stnd_for_order_dir.set(r"D:\Заявление судебный приказ1.docx")
    dir_for_order_text = tkinter.Entry(tab12, width=45, textvariable = stnd_for_order_dir, state='normal')
    dir_for_order_text.grid(column=1, row=10, sticky=tk.W)
    dir_for_order_text.focus()
    btnsave_order = tk.Button(tab12, text="Сохранить готовый документ",
                        command=(lambda:
                                 save_doc(source='Заявление_о_выдаче_судебного_приказа_по_сделке.docx',
                                          a=order_date_entry_text.get(),
                                          b=place_of_order_entry_text.get(),
                                          c=recoverer_text.get(),
                                          d=debtor_text.get(),
                                          e=order_price_text.get(),
                                          f=expire_date_of_order_entry_text.get(),
                                          g=detail_of_recoverer_text.get("1.0", tk.END),
                                          h=detail_of_debtor_text.get("1.0", tk.END),
                                          i=object_of_order_text.get("1.0", tk.END),
                                          saving_path=dir_for_order_text.get())
                                 )
                        )
    btnsave_order.grid(column=1, row=11)
    btnexitsale_order = tk.Button(tab12, text="Закрыть вкладку", command=lambda: tab_control.forget(tab12))
    btnexitsale_order.grid(column=2, row=11)


def complaint_letter():
    tab13 = ttk.Frame(tab_control)
    tab_control.add(tab13, text='Письмо претензия')
    tab13.bind('<Control-v>', pyperclip.paste())
    complaint_letter_date = tkinter.Label(tab13, text="Введите дату письма:")
    complaint_letter_date.grid(column=0, row=1, sticky=tk.E)
    complaint_letter_date_entry = tk.StringVar()
    complaint_letter_date_entry_text = tkinter.Entry(tab13, width=45, textvariable=complaint_letter_date_entry)
    complaint_letter_date_entry_text.grid(column=1, row=1, sticky=tk.W)
    complaint_letter_date_entry_text.focus()
    complaint_letter_date_entry_help = tkinter.Label(tab13, text="Например: 01.01.2022 или 1 января 2022 года")
    complaint_letter_date_entry_help.grid(column=2, row=1, sticky=tk.W)
    correspondent = tkinter.Label(tab13, text="Должность отправителя:")
    correspondent.grid(column=0, row=2, sticky=tk.E)
    correspondent_entry = tk.StringVar()
    correspondent_text = tkinter.Entry(tab13, width=45, textvariable=correspondent_entry)
    correspondent_text.grid(column=1, row=2, sticky=tk.W)
    correspondent_text.focus()
    correspondent_text_entry_help = tkinter.Label(tab13, text="Например: Генеральный директор")
    correspondent_text_entry_help.grid(column=2, row=2, sticky=tk.W)
    addressee = tkinter.Label(tab13, text="Имя и Отчество адресата (руководителя адресата)")
    addressee.grid(column=0, row=3, sticky=tk.E)
    addressee_entry = tk.StringVar()
    addressee_text = tkinter.Entry(tab13, width=45, textvariable=addressee_entry)
    addressee_text.grid(column=1, row=3, sticky=tk.W)
    addressee_text.focus()
    complaint_letter_price = tkinter.Label(tab13, text="Тема письма:")
    complaint_letter_price.grid(column=0, row=4, sticky=tk.E)
    complaint_letter_price_entry = tk.StringVar()
    complaint_letter_price_text = tkinter.Entry(tab13, width=45, textvariable=complaint_letter_price_entry)
    complaint_letter_price_text.grid(column=1, row=4, sticky=tk.W)
    complaint_letter_price_text.focus()
    complaint_letter_price_text_help= tkinter.Label(tab13, text="Например: О возникшей задолженности по контракту")
    complaint_letter_price_text_help.grid(column=2, row=4, sticky=tk.W)
    detail_of_correspondent = tkinter.Label(tab13, text="Реквизиты отправителя:")
    detail_of_correspondent.grid(column=0, row=5, sticky=tk.E)
    detail_of_correspondent_text = scrolledtext.ScrolledText(tab13, width=35, height=6)
    detail_of_correspondent_text.grid(column=1, row=5, sticky=tk.W)
    detail_of_correspondent_text_help = tkinter.Label(tab13, text="Наименование, ИНН, ОГРН, БИК Банка, р\с ...")
    detail_of_correspondent_text_help.grid(column=2, row=5, sticky=tk.W)
    detail_of_addressee = tkinter.Label(tab13, text="Реквизиты адресата:")
    detail_of_addressee.grid(column=0, row=6, sticky=tk.E)
    detail_of_addressee_text = scrolledtext.ScrolledText(tab13, width=35, height=6)
    detail_of_addressee_text.grid(column=1, row=6, sticky=tk.W)
    object_of_complaint_letter = tkinter.Label(tab13, text="Содержание письма:")
    object_of_complaint_letter.grid(column=0, row=7, sticky=tk.E)
    object_of_complaint_letter_text = scrolledtext.ScrolledText(tab13, width=38, height=14)
    object_of_complaint_letter_text.grid(column=1, row=7, sticky=tk.W)
    object_of_complaint_letter_help = tkinter.Label(tab13, text="Письмо претензия готовится, как\n"
                                                        "правило, в виде меморанудума: история\n"
                                                        "возникновения отношений (01.01.2020 между\n"
                                                        "нами был заключен контракт ...), что было\n"
                                                        "сделано Вами, и что не сделано (сделано)\n"
                                                        "адресатом, со ссылками на отчётные документы\n"
                                                        "и пункты контракта, какие обязательства у\n"
                                                        "строны в связи с этим возникли,\n"
                                                        "размер обязательств и сроки их исполнения")
    object_of_complaint_letter_help.grid(column=2, row=7, sticky=tk.W)
    complaint_letter_btn_help = tk.Button(tab13, text="Судебная практика по типу документа",
                        command=lambda: webbrowser.open("https://barrister365-my.sharepoint.com/:f:/g/personal/yan_barrister365_onmicrosoft_com/Elkx2HmmvR5LnMkuxRrgu7MBW-XwAv-JHPIyahUQ_MaLwA?e=1MfbQb"))
    complaint_letter_btn_help.grid(column=0, row=10)
    dir_for_complaint_letter = tkinter.Label(tab13, text="Директория для сохранения файла:")
    dir_for_complaint_letter.grid(column=0, row=8, sticky=tk.E)
    stnd_for_complaint_letter_dir = tk.StringVar()
    stnd_for_complaint_letter_dir.set(r"D:\Письмо претензия1.docx")
    dir_for_complaint_letter_text = tkinter.Entry(tab13, width=45, textvariable = stnd_for_complaint_letter_dir, state='normal')
    dir_for_complaint_letter_text.grid(column=1, row=8, sticky=tk.W)
    dir_for_complaint_letter_text.focus()
    btnsave_complaint_letter = tk.Button(tab13, text="Сохранить готовый документ",
                        command=(lambda:
                                 save_doc(source='Письмо-претензия.docx',
                                          a=complaint_letter_date_entry_text.get(),
                                          b=None,
                                          c=correspondent_text.get(),
                                          d=addressee_text.get(),
                                          e=complaint_letter_price_text.get(),
                                          f=None,
                                          g=detail_of_correspondent_text.get("1.0", tk.END),
                                          h=detail_of_addressee_text.get("1.0", tk.END),
                                          i=object_of_complaint_letter_text.get("1.0", tk.END),
                                          saving_path=dir_for_complaint_letter_text.get())
                                 )
                        )
    btnsave_complaint_letter.grid(column=1, row=10)
    btnexitsale_complaint_letter = tk.Button(tab13, text="Закрыть вкладку", command=lambda: tab_control.forget(tab13))
    btnexitsale_complaint_letter.grid(column=2, row=10)


def acceptance():
    tab14 = ttk.Frame(tab_control)
    tab_control.add(tab14, text='Акт приёмки')
    tab14.bind('<Control-v>', pyperclip.paste())
    acceptance_date = tkinter.Label(tab14, text="Введите дату договора:")
    acceptance_date.grid(column=0, row=1, sticky=tk.E)
    acceptance_date_entry = tk.StringVar()
    acceptance_date_entry_text = tkinter.Entry(tab14, width=45, textvariable=acceptance_date_entry)
    acceptance_date_entry_text.grid(column=1, row=1, sticky=tk.W)
    acceptance_date_entry_text.focus()
    acceptance_date_entry_help = tkinter.Label(tab14, text="Например: 01.01.2022 или 1 января 2022 года")
    acceptance_date_entry_help.grid(column=2, row=1, sticky=tk.W)
    place_of_acceptance = tkinter.Label(tab14, text="Введите место составления договора:")
    place_of_acceptance.grid(column=0, row=2, sticky=tk.E)
    place_of_acceptance_entry = tk.StringVar()
    place_of_acceptance_entry_text = tkinter.Entry(tab14, width=45, textvariable=place_of_acceptance_entry)
    place_of_acceptance_entry_text.grid(column=1, row=2, sticky=tk.W)
    place_of_acceptance_entry_text.focus()
    place_of_acceptance_entry_help = tkinter.Label(tab14, text="Например: Москва")
    place_of_acceptance_entry_help.grid(column=2, row=2, sticky=tk.W)
    cust_accepted = tkinter.Label(tab14, text=" Заказчик (лицо, действующее от его имени на основании документа):")
    cust_accepted.grid(column=0, row=3, sticky=tk.E)
    cust_accepted_entry = tk.StringVar()
    cust_accepted_text = tkinter.Entry(tab14, width=45, textvariable=cust_accepted_entry)
    cust_accepted_text.grid(column=1, row=3, sticky=tk.W)
    cust_accepted_text.focus()
    cust_accepted_text_entry_help = tkinter.Label(tab14, text="Например: ООО Ромашка, в лице директора ...")
    cust_accepted_text_entry_help.grid(column=2, row=3, sticky=tk.W)
    executer_accepted = tkinter.Label(tab14, text="Исполнитель (лицо, действующее от его имени на основании документа):")
    executer_accepted.grid(column=0, row=4, sticky=tk.E)
    executer_accepted_entry = tk.StringVar()
    executer_accepted_text = tkinter.Entry(tab14, width=45, textvariable=executer_accepted_entry)
    executer_accepted_text.grid(column=1, row=4, sticky=tk.W)
    executer_accepted_text.focus()
    acceptance_price = tkinter.Label(tab14, text="Введите цену договора:")
    acceptance_price.grid(column=0, row=5, sticky=tk.E)
    acceptance_price_entry = tk.StringVar()
    acceptance_price_text = tkinter.Entry(tab14, width=45, textvariable=acceptance_price_entry)
    acceptance_price_text.grid(column=1, row=5, sticky=tk.W)
    acceptance_price_text.focus()
    acceptance_price_help= tkinter.Label(tab14, text="Например: 100 000 рублей 00 коп. или 100 000 руб.\n"
                                "Цена обязательно указывается с НДС или без НДС,\n"
                                "также необходимо указать на каком основании\n"
                                "НДС не применяется.")
    acceptance_price_help.grid(column=2, row=5, sticky=tk.W)
    expire_date_of_acceptance = tkinter.Label(tab14, text="Введите дату приёмки:")
    expire_date_of_acceptance.grid(column=0, row=6, sticky=tk.E)
    expire_date_of_acceptance_entry = tk.StringVar
    expire_date_of_acceptance_entry_text = tkinter.Entry(tab14, width=45, textvariable=expire_date_of_acceptance_entry)
    expire_date_of_acceptance_entry_text.grid(column=1, row=6, sticky=tk.W)
    expire_date_of_acceptance_entry_text.focus()
    detail_of_cust_accepted = tkinter.Label(tab14, text="Банковские реквизиты заказчика:")
    detail_of_cust_accepted.grid(column=0, row=7, sticky=tk.E)
    detail_of_cust_accepted_text = scrolledtext.ScrolledText(tab14, width=35, height=6)
    detail_of_cust_accepted_text.grid(column=1, row=7, sticky=tk.W)
    detail_of_cust_accepted_text_help = tkinter.Label(tab14, text="Наименование, ИНН, ОГРН, БИК Банка, р\с ...")
    detail_of_cust_accepted_text_help.grid(column=2, row=7, sticky=tk.W)
    detail_of_executer_accepted = tkinter.Label(tab14, text="Банковские реквизиты исполнителя:")
    detail_of_executer_accepted.grid(column=0, row=8, sticky=tk.E)
    detail_of_executer_accepted_text = scrolledtext.ScrolledText(tab14, width=35, height=6)
    detail_of_executer_accepted_text.grid(column=1, row=8, sticky=tk.W)
    object_of_acceptance = tkinter.Label(tab14, text="Что принималось по акту:")
    object_of_acceptance.grid(column=0, row=9, sticky=tk.E)
    object_of_acceptance_text = scrolledtext.ScrolledText(tab14, width=38, height=14)
    object_of_acceptance_text.grid(column=1, row=9, sticky=tk.W)
    object_of_acceptance_help = tkinter.Label(tab14, text="Указывается, что конкретно принимается,\n"
                                                  "по количеству, номерам, виду работ и т.д.")
    object_of_acceptance_help.grid(column=2, row=9, sticky=tk.W)
    acceptance_btn_help = tk.Button(tab14, text="Судебная практика по типу документа",
                        command=lambda: webbrowser.open("https://barrister365-my.sharepoint.com/:f:/g/personal/yan_barrister365_onmicrosoft_com/EoV6jRX9VH9NhHX6--192sQBknP136tvYYBJpPAzPwqChg?e=n9nH99"))
    acceptance_btn_help.grid(column=0, row=11)
    dir_for_acceptance = tkinter.Label(tab14, text="Директория для сохранения файла:")
    dir_for_acceptance.grid(column=0, row=10, sticky=tk.E)
    stnd_for_acceptance_dir = tk.StringVar()
    stnd_for_acceptance_dir.set(r"D:\Акт приёмки1.docx")
    dir_for_acceptance_text = tkinter.Entry(tab14, width=45, textvariable = stnd_for_acceptance_dir, state='normal')
    dir_for_acceptance_text.grid(column=1, row=10, sticky=tk.W)
    dir_for_acceptance_text.focus()
    btnsave_acceptance = tk.Button(tab14, text="Сохранить готовый документ",
                        command=(lambda:
                                 save_doc(source='Акт_сдачи_приёмки.docx',
                                          a=acceptance_date_entry_text.get(),
                                          b=place_of_acceptance_entry_text.get(),
                                          c=cust_accepted_text.get(),
                                          d=executer_accepted_text.get(),
                                          e=acceptance_price_text.get(),
                                          f=expire_date_of_acceptance_entry_text.get(),
                                          g=detail_of_cust_accepted_text.get("1.0", tk.END),
                                          h=detail_of_executer_accepted_text.get("1.0", tk.END),
                                          i=object_of_acceptance_text.get("1.0", tk.END),
                                          saving_path=dir_for_acceptance_text.get())
                                 )
                        )
    btnsave_acceptance.grid(column=1, row=11)
    btnexitsale_acceptance = tk.Button(tab14, text="Закрыть вкладку", command=lambda: tab_control.forget(tab14))
    btnexitsale_acceptance.grid(column=2, row=11)


def organization_order():
    tab15 = ttk.Frame(tab_control)
    tab_control.add(tab15, text='Приказ')
    tab15.bind('<Control-v>', pyperclip.paste())
    organization_order_date = tkinter.Label(tab15, text="Введите дату приказа:")
    organization_order_date.grid(column=0, row=1, sticky=tk.E)
    organization_order_date_entry = tk.StringVar()
    organization_order_date_entry_text = tkinter.Entry(tab15, width=45, textvariable=organization_order_date_entry)
    organization_order_date_entry_text.grid(column=1, row=1, sticky=tk.W)
    organization_order_date_entry_text.focus()
    organization_order_date_entry_help = tkinter.Label(tab15, text="Например: 01.01.2022 или 1 января 2022 года")
    organization_order_date_entry_help.grid(column=2, row=1, sticky=tk.W)
    post = tkinter.Label(tab15, text="Введите должность лица, издающего приказ:")
    post.grid(column=0, row=2, sticky=tk.E)
    post_entry = tk.StringVar()
    post_text = tkinter.Entry(tab15, width=45, textvariable=post_entry)
    post_text.grid(column=1, row=2, sticky=tk.W)
    post_text.focus()
    post_text_entry_help = tkinter.Label(tab15, text="Например: ГЕНЕРАЛЬНЫЙ ДИРЕКТОР ...")
    post_text_entry_help.grid(column=2, row=2, sticky=tk.W)
    detail_of_post = tkinter.Label(tab15, text="Реквизиты организации - издателя приказа:")
    detail_of_post.grid(column=0, row=3, sticky=tk.E)
    detail_of_post_text = scrolledtext.ScrolledText(tab15, width=35, height=6)
    detail_of_post_text.grid(column=1, row=3, sticky=tk.W)
    detail_of_post_text_help = tkinter.Label(tab15, text="Наименование, ИНН, ОГРН, БИК Банка, р\с ...")
    detail_of_post_text_help.grid(column=2, row=3, sticky=tk.W)
    object_of_organization_order = tkinter.Label(tab15, text="Содержание приказа")
    object_of_organization_order.grid(column=0, row=4, sticky=tk.E)
    object_of_organization_order_text = scrolledtext.ScrolledText(tab15, width=38, height=14)
    object_of_organization_order_text.grid(column=1, row=4, sticky=tk.W)
    object_of_organization_order_help = tkinter.Label(tab15, text="Например: "
                                                          "1.Утвердить Методику Определения трудозатрат\n"
                                                          "при выполнении научно-исследовательских и\n"
                                                          "опытно конструкторских работ.\n"
                                                          "2. Коммерческому директору в срок до трёх\n"
                                                          "рабочих дней ознакомить с настоящим приказом\n"
                                                          "всех заинтересованных лиц.\n"
                                                          "3. Приказ вступает в силу с момента его подписания.\n"
                                                          "4. Контроль за исполнением приказа оставляю за собой.")
    object_of_organization_order_help.grid(column=2, row=4, sticky=tk.W)
    organization_order_btn_help = tk.Button(tab15, text="Судебная практика по типу документа",
                        command=lambda: webbrowser.open("https://barrister365-my.sharepoint.com/:f:/g/personal/yan_barrister365_onmicrosoft_com/EhDmVxTqxsdAoeCQ5CC81isB4HHrTQgE4JqYW0Gyen7jCg?e=Zyt5ZR"))
    organization_order_btn_help.grid(column=0, row=7)
    dir_for_organization_order = tkinter.Label(tab15, text="Директория для сохранения файла:")
    dir_for_organization_order.grid(column=0, row=6, sticky=tk.E)
    stnd_for_organization_order_dir = tk.StringVar()
    stnd_for_organization_order_dir.set(r"D:\Приказ1.docx")
    dir_for_organization_order_text = tkinter.Entry(tab15, width=45, textvariable = stnd_for_organization_order_dir, state='normal')
    dir_for_organization_order_text.grid(column=1, row=6, sticky=tk.W)
    dir_for_organization_order_text.focus()
    btnsave_organization_order = tk.Button(tab15, text="Сохранить готовый документ",
                        command=(lambda:
                                 save_doc(source='Приказ_по_организации.docx',
                                          a=organization_order_date_entry_text.get(),
                                          b=None,
                                          c=post_text.get(),
                                          d=None,
                                          e=None,
                                          f=None,
                                          g=detail_of_post_text.get("1.0", tk.END),
                                          h=None,
                                          i=object_of_organization_order_text.get("1.0", tk.END),
                                          saving_path=dir_for_organization_order_text.get())
                                 )
                        )
    btnsave_organization_order.grid(column=1, row=7)
    btnexitsale_organization_order = tk.Button(tab15, text="Закрыть вкладку", command=lambda: tab_control.forget(tab15))
    btnexitsale_organization_order.grid(column=2, row=7)






window = Tk()
window.title("Документалист V.1.0.")
window.attributes('-fullscreen', True)
menu = Menu(window)
new_item = Menu(menu)
new_item.add_command(label='Справка о программе', command=lambda: spravkaoprogramme())
new_item.add_command(label='Закрыть приложение', command=window.quit)
menu.add_cascade(label='Файл', menu=new_item)
window.config(menu=menu)
tab_control = ttk.Notebook(window)
tab1 = ttk.Frame(tab_control)
tab_control.add(tab1, text='Тип документа')
tab_control.pack(expand=1, fill='both')
doctype = tkinter.Label(tab1, text="Выберете тип документа:")
doctype.grid(column=0, row=0, sticky=tk.W)
rad1 = Radiobutton(tab1,
                   text='Договор купли-продажи',
                   value=1,
                   command=lambda: sale())
rad2 = Radiobutton(tab1, text='Договор мены',
                   value=2,
                   command=lambda: barter())
rad3 = Radiobutton(tab1, text='Договор дарения',
                   value=3,
                   command=lambda: donations())
rad4 = Radiobutton(tab1, text='Договор аренды нежилого помещения',
                   value=4, command=lambda: lease())
rad7 = Radiobutton(tab1, text='Договор подряда',
                   value=7, command=lambda: podryad())
rad9 = Radiobutton(tab1, text='Договор возмездного оказания услуг',
                   value=9, command=lambda: service())
rad12 = Radiobutton(tab1, text='Договор займа',
                    value=12, command=lambda: loan())
rad20 = Radiobutton(tab1, text='Договор комиссии',
                    value=20, command=lambda: commission())
rad21 = Radiobutton(tab1, text='Агентский договор',
                    value=21, command=lambda: agency())
rad27 = Radiobutton(tab1, text='Лицензионный договор',
                    value=27, command=lambda: licenseagr())
rad28 = Radiobutton(tab1, text='Заявление о выдаче судебного приказа',
                    value=28, command=lambda: court_order())
rad29 = Radiobutton(tab1, text='Письмо-претензия',
                    value=29, command=lambda: complaint_letter())
rad30 = Radiobutton(tab1, text='Акт приёмки выполненных работ (оказаннных услуг)',
                    value=30, command=lambda: acceptance())
rad31 = Radiobutton(tab1, text='Приказ по организации',
                    value=31, command=lambda: organization_order())

rad1.grid(column=0, row=1, sticky=tk.W)
rad2.grid(column=0, row=2, sticky=tk.W)
rad3.grid(column=0, row=3, sticky=tk.W)
rad4.grid(column=0, row=4, sticky=tk.W)
rad7.grid(column=0, row=5, sticky=tk.W)
rad9.grid(column=0, row=7, sticky=tk.W)
rad12.grid(column=0, row=8, sticky=tk.W)
rad20.grid(column=0, row=10, sticky=tk.W)
rad21.grid(column=0, row=11, sticky=tk.W)
rad27.grid(column=0, row=12, sticky=tk.W)
rad28.grid(column=1, row=1, sticky=tk.W)
rad29.grid(column=1, row=2, sticky=tk.W)
rad30.grid(column=1, row=3, sticky=tk.W)
rad31.grid(column=1, row=4, sticky=tk.W)

window.mainloop()
