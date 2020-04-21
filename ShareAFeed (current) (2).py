from tkinter import ttk
import tkinter as tk
import pandas as pd
import numpy as np
from tkinter import filedialog
import re
import math

window = tk.Tk()

window.title("ShareASale Datafeed Converter")
window.iconbitmap('p2.ico')

window.geometry('650x300')
window.resizable(width=False, height=False)


def browse():
    filename = filedialog.askopenfilename(initialdir="/", title="Choose File to Convert",
                                          filetypes=(("csv files", "*.csv"), ("all files", "*.*")))
    if not filename:
        return
    file_box.delete(0, tk.END)
    file_box.insert(tk.END, filename)


def clicked():
    try:
        status['text'] = 'Running conversion...'
        status['fg'] = 'black'
        bar["value"] = 0

        mid = mid_box.get()
        if not mid:
            bar["value"] = 0
            status['text'] = 'Error! Please Enter a Merchant ID'
            status['fg'] = 'red'
            return
        url = URL_box.get()
        if not url:
            bar["value"] = 0
            status['text'] = 'Error! Please Enter Your Website'
            status['fg'] = 'red'
            return
        cat0 = cat_box.get()
        cat1 = re.search("(?=\d{1,2} -)\d{1,2}", cat0)
        cat = cat1.group()
        sub_cat0 = sub_cat_box.get()
        sub_cat1 = re.search("(?=\d{1,3} -)\d{1,3}", sub_cat0)
        sub_cat = sub_cat1.group()
        reward = RS_box.get()
        stores = SC_box.get()
        file = file_box.get()
        if not file:
            bar["value"] = 0
            status['text'] = 'Error! Please Choose a Source File'
            status['fg'] = 'red'
            return
        default_name = mid + '_Datafeed'

        status['text'] = 'Removing Non-ShareASale Columns'
        try:
            df = pd.read_csv(file)
        except UnicodeDecodeError:
            df = pd.read_csv(file, encoding = "ISO-8859-1")
        df = df.drop(df.columns.difference(['Variant SKU','Handle', 'Body (HTML)', 'Variant Price', 'Image Src',
                                            'Type', 'Tags', 'Title']), axis=1)

        bar.step(20)

        status['text'] = 'Renaming Columns'
        df = df.rename(
            columns={'Variant SKU': 'SKU', 'Handle': 'URL', 'Body (HTML)': 'Description', 'Variant Price': 'Price',
                     'Image Src': 'FullImage', 'Type': 'MerchantCategory', 'Tags': 'SearchTerms', 'Title': 'Name'})
        df['ThumbnailImage'] = df['FullImage']
        status['text'] = 'Removing Items with Missing SKUs, Price or URL. Removing Duplicate SKUs'
        df = df.dropna(axis=0, subset=['SKU'])
        df = df.dropna(axis=0, subset=['URL'])
        df = df.dropna(axis=0, subset=['Price'])
        df = df.drop_duplicates(subset=['SKU'])
        bar.step(10)

        status['text'] = 'Adding Remaining ShareASale Columns, adding Merchant ID, Category and Subcategory Values'
        df = df.assign(
            **{'RetailPrice': np.nan, 'Commission': np.nan, 'Category': cat, 'Subcategory': sub_cat, 'Status': np.nan,
               'MerchantID': mid, 'Custom1': np.nan, 'Custom2': np.nan, 'Custom3': np.nan, 'Custom4': np.nan,
               'Custom5': np.nan, 'Manufacturer': np.nan,
               'PartNumber': np.nan, 'MerchantSubcategory': np.nan, 'ShortDescription': np.nan, 'ISBN': np.nan,
               'UPC': np.nan, 'CrossSell': np.nan,
               'MerchantGroup': np.nan, 'MerchantSubgroup': np.nan, 'CompatibleWith': np.nan, 'CompareTo': np.nan,
               'QuantityDiscount': np.nan, 'Bestseller': np.nan,
               'AddToCartURL': np.nan, 'ReviewsRSSURL': np.nan, 'Option1': np.nan, 'Option2': np.nan, 'Option3': np.nan,
               'Option4': np.nan, 'Option5': np.nan,
               'customCommissions': np.nan, 'customCommissionisFlatRate': np.nan,
               'customCommissionNewCustomerMultiplier': np.nan, 'mobileURL': np.nan,
               'mobileImage': np.nan, 'mobileThumbnail': np.nan, 'ReservedForFutureUse1': np.nan,
               'ReservedForFutureUse2': np.nan, 'ReservedForFutureUse3': np.nan, 'ReservedForFutureUse4': np.nan})

        if reward == 'Yes':
            currency = RS_cur.get()
            country = RS_cc.get()
            df = df.assign(**{'Custom4': currency, 'Custom5': country})

        status['text'] = 'Organizing Columns'
        df = df[['SKU', 'Name', 'URL', 'Price', 'RetailPrice', 'FullImage', 'ThumbnailImage', 'Commission', 'Category',
                 'Subcategory', 'Description', 'SearchTerms',
                 'Status', 'MerchantID', 'Custom1', 'Custom2', 'Custom3', 'Custom4', 'Custom5', 'Manufacturer',
                 'PartNumber', 'MerchantCategory', 'MerchantSubcategory',
                 'ShortDescription', 'ISBN', 'UPC', 'CrossSell', 'MerchantGroup', 'MerchantSubgroup', 'CompatibleWith',
                 'CompareTo', 'QuantityDiscount', 'Bestseller',
                 'AddToCartURL', 'ReviewsRSSURL', 'Option1', 'Option2', 'Option3', 'Option4', 'Option5',
                 'customCommissions', 'customCommissionisFlatRate',
                 'customCommissionNewCustomerMultiplier', 'mobileURL', 'mobileImage', 'mobileThumbnail',
                 'ReservedForFutureUse1', 'ReservedForFutureUse2',
                 'ReservedForFutureUse3', 'ReservedForFutureUse4']]
        bar.step(20)

        if stores == 'Yes':
            storeID = SNUM_box.get()
            if not storeID:
                bar["value"] = 0
                status['text'] = 'Error! Please Enter a Store ID.\nIf you are not using StoresConnect, check NO above.'
                status['fg'] = 'red'
                return
            df = df.assign(**{'StoreID': storeID})
            df = df[
                ['SKU', 'Name', 'URL', 'Price', 'RetailPrice', 'FullImage', 'ThumbnailImage', 'Commission', 'Category',
                 'Subcategory', 'Description', 'SearchTerms',
                 'Status', 'MerchantID', 'Custom1', 'Custom2', 'Custom3', 'Custom4', 'Custom5', 'StoreID',
                 'Manufacturer',
                 'PartNumber', 'MerchantCategory', 'MerchantSubcategory',
                 'ShortDescription', 'ISBN', 'UPC', 'CrossSell', 'MerchantGroup', 'MerchantSubgroup', 'CompatibleWith',
                 'CompareTo', 'QuantityDiscount', 'Bestseller',
                 'AddToCartURL', 'ReviewsRSSURL', 'Option1', 'Option2', 'Option3', 'Option4', 'Option5',
                 'customCommissions', 'customCommissionisFlatRate',
                 'customCommissionNewCustomerMultiplier', 'mobileURL', 'mobileImage', 'mobileThumbnail',
                 'ReservedForFutureUse1', 'ReservedForFutureUse2',
                 'ReservedForFutureUse3', 'ReservedForFutureUse4']]

        status['text'] = 'Executing Validation'
        df["URL"] = "https://" + url + "/products/" + df["URL"]

        df['Description'] = df['Description'].str.encode('ascii', 'ignore').str.decode('ascii')
        df = df.replace({'\<.*?\>': '', 'https\:\/\/https\:\/\/': 'https://', 'https\:\/\/http\:\/\/': 'http://',
                         '\/\/products': '/products', '\&amp\;': '&', '\&lt\;': '<', '\&gt\;': '>'}, regex=True)
        df['Description'] = df['Description'].replace('\s+', ' ', regex=True)  # remove extra whitespace
        df['Description'] = df['Description'].replace('\n', ' ', regex=True)
        df['Description'] = df['Description'].str.strip()
        bar.step(20)
        status['text'] = 'Choosing File Save Location'
        newfile = filedialog.asksaveasfilename(initialdir="/", initialfile=default_name, title="Create New File",
                                               filetypes=(("csv files", "*.csv"),))
        if not newfile:
            bar["value"] = 0
            status['text'] = 'Conversion Canceled'
            status['fg'] = 'black'
            return
        newfile = newfile + '.csv'
        if '.csv.csv' in newfile:
            newfile = newfile[0:-4]
            bar.step(10)

        status['text'] = 'Finalizing'
        df.to_csv(newfile, index=False)
        bar["value"] = 100
        status['text'] = 'Your ShareASale datafeed was created successfully!'
    except:
        status['text'] = 'Unsupported feed. Please select a new file.'
        status['fg'] = 'red'
        bar["value"] = 0
        return


def show_codes():
    reward = RS_box.get()
    sub_cat_lbl.focus()
    if reward == 'Yes':
        RS_cc.bind("<<ComboboxSelected>>", lambda e: cat_lbl.focus())
        RS_cc.grid(column=1, row=8, sticky='N')
        RS_cc_label.grid(column=1, row=7, sticky='N')
        RS_cur.bind("<<ComboboxSelected>>", lambda e: cat_lbl.focus())
        RS_cur.grid(column=2, row=8)
        RS_cur_label.grid(column=2, row=7)
    else:
        RS_cc.grid_forget()
        RS_cc_label.grid_forget()
        RS_cur.grid_forget()
        RS_cur_label.grid_forget()


def show_stores():
    stores = SC_box.get()
    sub_cat_lbl.focus()
    if stores == 'Yes':
        SNUM_lbl.grid(column=1, row=9, sticky='E')
        SNUM_box.grid(column=2, row=9, sticky='W')
    else:
        SNUM_lbl.grid_forget()
        SNUM_box.grid_forget()


def list_sub_cats():
    cat = cat_box.get()
    cat_lbl.focus()
    if cat == "1 - Art/Media/Performance":
        sub_cat_box['values'] = (
        "1 - Art", "2 - Photography", "3 - Posters/Prints", "4 - Music", "5 - Music Instruments",
        "187 - Art Supplies")
        sub_cat_box.current(0)
    elif cat == "2 - Auto/Boat/Plane":
        sub_cat_box['values'] = (
        "6 - Accessories", "7 - Car Audio", "8 - Cleaning/Care", "9 - Motorcycles", "10 - Misc.",
        "11 - Repair", "12 - Parts")
        sub_cat_box.current(0)
    elif cat == "3 - Books/Reading":
        sub_cat_box['values'] = ("13 - Art", "14 - Careers", "15 - Business", "16 - Childrens", "17 - Computers",
                                 "18 - Crafts/Hobbies", "19 - Education", "20 - Engineering", "21 - Gifts",
                                 "22 - Health",
                                 "23 - History", "24 - Fiction", "25 - Law", "26 - Magazines", "27 - Financial",
                                 "28 - Medical", "29 - Office", "30 - Real Estate", "31 - Misc.", "164 - Religious",
                                 "173 - Science/Nature")
        sub_cat_box.current(0)
    elif cat == "4 - Business/Services":
        sub_cat_box['values'] = ("32 - Advertising", "33 - Motivational", "34 - Coupons/Freebies", "35 - Financial",
                                 "36 - Loans", "37 - Office", "38 - Careers", "39 - Mis.", "179 - Education")
        sub_cat_box.current(0)
    elif cat == "5 - Computer":
        sub_cat_box['values'] = ("40 - Hardware", "41 - Software", "42 - Instruction", "43 - Handheld/Wireless",
                                 "162 - Web Hosting")
        sub_cat_box.current(0)
    elif cat == "6 - Electronics":
        sub_cat_box['values'] = ("44 - Audio", "45 - Video", "46 - Camera", "47 - Wireless")
        sub_cat_box.current(0)
    elif cat == "7 - Entertainment":
        sub_cat_box['values'] = ("48 - Audio", "49 - Video", "50 - DVD", "51 - Laser Disc", "52 - Sheet Music",
                                 "53 - Crafts/Hobbies", "184 - Tickets")
        sub_cat_box.current(0)
    elif cat == "8 - Fashion":
        sub_cat_box['values'] = ("54 - Boys", "55 - Clearance", "56 - Vintage", "57 - Girls", "58 - Men", "59 - Women",
                                 "60 - Maternity", "61 - Footware", "62 - Accessories", "63 - Baby/Infant",
                                 "64 - Jewelry",
                                 "65 - Lingerie", "66 - Plus-Size", "67 - Athletic", "161 - T-Shirts",
                                 "166 - Big And Tall",
                                 "168 - Petite", "169 - Unisex", "172 - Costume")
        sub_cat_box.current(0)
    elif cat == "9 - Food/Beverage":
        sub_cat_box['values'] = ("68 - Baked Goods", "69 - Beverages", "70 - Chocolate", "71 - Cheese/Condiments",
                                 "72 - Coupons", "73 - Diet", "74 - International", "75 - Gifts/Gift Baskets",
                                 "76 - Nuts",
                                 "77 - Cookies/Desserts", "78 - Organic", "163 - Tobacco", "176 - Gourmet",
                                 "177 - Meals/Complete Dishes", "180 - Appetizers", "181 - Soups")
        sub_cat_box.current(0)
    elif cat == "10 - Gifts/Specialty":
        sub_cat_box['values'] = ("79 - Anniversary", "80 - Birthday", "81 - Misc. Holiday", "82 - Collectibles",
                                 "83 - Coupons", "84 - Executive Gifts", "85 - Flowers", "86 - Baskets",
                                 "87 - Greeting Card", "88 - Baby/Infant", "89 - Party", "90 - Religious",
                                 "91 - Sympathy",
                                 "92 - Valentine's Day", "93 - Wedding", "170 - Personalized")
        sub_cat_box.current(0)
    elif cat == "11 - Home/Family":
        sub_cat_box['values'] = (
        "94 - Bed/Bath", "95 - Garden", "96 - Cleaning/Care", "97 - Furniture", "98 - Home Decor",
        "99 - Home Improvement", "100 - Kitchen", "101 - Pets")
        sub_cat_box.current(0)
    elif cat == "12 - Personal Care":
        sub_cat_box['values'] = ("102 - Cosmetics", "103 - Exercise/Wellness", "104 - Safety", "183 - Medical")
        sub_cat_box.current(0)
    elif cat == "13 - Sports/Outdoors":
        sub_cat_box['values'] = ("105 - Accessories", "106 - Auto", "107 - Outdoors/Camping",
                                 "108 - Parlor/Backyard Games", "109 - Baseball/Softball", "110 - Cricket",
                                 "111 - Billiards", "112 - Boating", "113 - Body Building/Fitness", "114 - Bowling",
                                 "115 - Boxing", "116 - Canoeing", "117 - Climbing/Mountaineering", "118 - Cycling",
                                 "119 - Diving", "120 - Field Hockey", "121 - Skating", "122 - Fishing",
                                 "123 - Football",
                                 "124 - Frisbee", "125 - Golf", "126 - Gymnastics", "127 - Hockey", "128 - Horses",
                                 "129 - Hunting", "130 - In-line Skating", "131 - Kayaking", "132 - Lacrosse",
                                 "133 - Martial Arts", "134 - Racquetball", "135 - Running", "136 - Skateboards",
                                 "137 - Ski/Snowboard", "138 - Soccer", "139 - Surfing", "140 - Tennis",
                                 "141 - Teamware / Logo", "142 - Volleyball", "143 - Wrestling", "165 - Birding",
                                 "174 - Prospecting/Treasure Hunting", "175 - Swimming", "178 - Basketball")
        sub_cat_box.current(0)
    elif cat == "14 - Toys/Games":
        sub_cat_box['values'] = ("144 - Action", "145 - Animals", "146 - Baby/Infant", "147 - Board Games",
                                 "148 - Card/Casino", "149 - Electronic", "150 - Educational", "151 - Magic",
                                 "152 - Misc.",
                                 "153 - Musical", "154 - Outdoor", "155 - Video")
        sub_cat_box.current(0)
    elif cat == "15 - Travel":
        sub_cat_box['values'] = ("156 - Coupons", "157 - Maps", "158 - References / Guides", "159 - Vacation Travel",
                                 "185 - Luggage", "186 - Accessories")
        sub_cat_box.current(0)
    elif cat == "16 - Metaphysical":
        sub_cat_box['values'] = ("160 - Metaphysical",)
        sub_cat_box.current(0)
    elif cat == "17 - Parts/Equipment":
        sub_cat_box['values'] = ("167 - HVAC (Heating and Air Conditioning)", "171 - Medical", "182 - Military")
        sub_cat_box.current(0)

def save_settings():
    file = file_box.get()
    mid = mid_box.get()
    url = URL_box.get()
    cat = cat_box.get()
    sub_cat = sub_cat_box.get()
    rs = RS_box.get()
    country = RS_cc.get()
    currency = RS_cur.get()
    stores = SC_box.get()
    storeID = SNUM_box.get()

    data = [{'file': file, 'mid': mid, 'website': url, 'catg': cat, 'subcat': sub_cat, 'reward': rs,
               'country': country, 'currency': currency, 'stores': stores, 'storeID': storeID}]

    set_frame = pd.DataFrame(data)
    try:
        set_frame.to_csv('settings.csv', index=False)
        status['text'] = 'Settings Saved!'
        status['fg'] = 'black'
        bar["value"] = 0

    except PermissionError:
        status['text'] = 'Hey! Close your settings file!'
        status['fg'] = 'red'
        bar["value"] = 0

saved_file = ""
saved_mid = ""
saved_url = ""
saved_cat = "None"
saved_subcat = "None"
saved_reward = "No"
saved_country = "None"
saved_currency = "None"
saved_stores = "No"
saved_storeID = ""
saved_user = ""
saved_pass = ""

try:
    settings = pd.read_csv("settings.csv")
    saved_file = settings.iloc[0].file
    if isinstance(saved_file, int) and math.isnan(saved_file):
        saved_file = ""
    saved_mid = settings.iloc[0].mid
    if isinstance(saved_mid, int) and math.isnan(saved_mid):
        saved_mid = ""
    saved_url = settings.iloc[0].website
    if isinstance(saved_url, int) and math.isnan(saved_url):
        saved_url = ""
    saved_cat = settings.iloc[0].catg
    if isinstance(saved_cat, int) and math.isnan(saved_cat):
        saved_cat = "None"
    saved_subcat = settings.iloc[0].subcat
    if isinstance(saved_subcat, int) and math.isnan(saved_subcat):
        saved_subcat = "None"
    saved_reward = settings.iloc[0].reward
    if isinstance(saved_reward, int) and math.isnan(saved_reward):
        saved_reward = "No"
    saved_country = settings.iloc[0].country
    if isinstance(saved_country, int) and math.isnan(saved_country):
        saved_country = "None"
    saved_currency = settings.iloc[0].currency
    if isinstance(saved_currency, int) and math.isnan(saved_currency):
        saved_currency = "None"
    saved_stores = settings.iloc[0].stores
    if isinstance(saved_stores, int) and math.isnan(saved_stores):
        saved_stores = "No"
    saved_storeID = settings.iloc[0].storeID
    if isinstance(saved_storeID, int) and math.isnan(saved_stores):
        saved_storeID = ""

except (FileNotFoundError, IndexError):
    print ('No settings saved')

file_lbl = tk.Label(window, text="Import Your Product Feed")
file_lbl.grid(column=0, row=0, sticky='W')
file_box = tk.Entry(window, width=50)
file_box.grid(column=1, row=0)
file_btn = tk.Button(window, text="Browse", command=browse)
file_btn.grid(column=2, row=0, sticky='W')
file_box.insert(0, saved_file)

mid_lbl = tk.Label(window, text="Enter Your ShareASale Merchant ID")
mid_lbl.grid(column=0, row=1, sticky='W')
mid_box = tk.Entry(window, width=50)
mid_box.grid(column=1, row=1, sticky='W')
mid_box.insert(0, saved_mid)

URL_lbl = tk.Label(window, text="Enter Your Website URL")
URL_lbl.grid(column=0, row=2, sticky='W')
URL_box = tk.Entry(window, width=50)
URL_box.grid(column=1, row=2, sticky='W')
URL_box.insert(0, saved_url)

cat_lbl = tk.Label(window, text="Choose a Default Product Category")
cat_lbl.grid(column=0, row=3, sticky='W')
cat_box = ttk.Combobox(window, state="readonly", width="47")
cat_box['values'] = (
"1 - Art/Media/Performance", "2 - Auto/Boat/Plane", "3 - Books/Reading", "4 - Business/Services", "5 - Computer",
"6 - Electronics", "7 - Entertainment", "8 - Fashion", "9 - Food/Beverage", "10 - Gifts/Specialty", "11 - Home/Family",
"12 - Personal Care", "13 - Sports/Outdoors", "14 - Toys/Games", "15 - Travel", "16 - Metaphysical",
"17 - Parts/Equipment")
cat_box.current(0)
cat_box.bind("<<ComboboxSelected>>", lambda e: list_sub_cats())
cat_box.grid(column=1, row=3, sticky='W')

sub_cat_lbl = tk.Label(window, text="Choose a Default Product Sub Category")
sub_cat_lbl.grid(column=0, row=4, sticky='W')
sub_cat_box = ttk.Combobox(window, state="readonly", width="47")
list_sub_cats()
sub_cat_box.current(0)
sub_cat_box.bind("<<ComboboxSelected>>", lambda e: sub_cat_lbl.focus())
sub_cat_box.grid(column=1, row=4, sticky='W')

optn_lbl = tk.Label(window, text='Optional Fields', font='Helvetica 10 bold underline')
optn_lbl.grid(column=0, row=7)

RS_label = tk.Label(window, text="Are you working with RewardStyle?")
RS_label.grid(column=0, row=8, sticky='W')
RS_box = ttk.Combobox(window, state="readonly", width=4)
RS_box['values'] = ('No', 'Yes')
RS_box.current(0)
# RS_box.bind("<<ComboboxSelected>>", lambda e: cat_lbl.focus())
RS_box.grid(column=1, row=8, sticky='W')
RS_box.bind("<<ComboboxSelected>>", lambda _: show_codes())
if saved_reward == 'Yes':
    RS_box.current(1)
elif saved_reward == 'No':
    RS_box.current(0)

RS_cc = ttk.Combobox(window, width=3)
RS_cc['values'] = ('US', 'CA', 'GB', 'FR', 'IT', 'CN', 'HK')
if saved_country == "None":
    RS_cc.current(0)
else:
    country_index = RS_cc['values'].index(saved_country)
    RS_cc.current(country_index)
RS_cc_label = tk.Label(window, text="Country Code")

RS_cur = ttk.Combobox(window, width=4)
RS_cur['values'] = ('USD', 'CAD', 'GBP', 'EUR', 'CNY', 'HKD')
if saved_currency == "None":
    RS_cur.current(0)
else:
    currency_index = RS_cur['values'].index(saved_currency)
    RS_cur.current(currency_index)
RS_cur_label = tk.Label(window, text="Currency Code")

SC_label = tk.Label(window, text="Are you using StoresConnect?")
SC_label.grid(column=0, row=9, sticky='W')
SC_box = ttk.Combobox(window, state="readonly", width=4)
SC_box['values'] = ('No', 'Yes')
if saved_stores == 'Yes':
    SC_box.current(1)
elif saved_stores == 'No':
    SC_box.current(0)
SC_box.bind("<<ComboboxSelected>>", lambda _: show_stores())
SC_box.grid(column=1, row=9, sticky='W')

SNUM_lbl = tk.Label(window, text='Store Number:')
SNUM_box = tk.Entry(window, width=4)

spacer1 = tk.Label(window, text='')
spacer1.grid(column=0, row=6, sticky='W')

btn = tk.Button(window, text="Begin Conversion", command=clicked)
btn.grid(column=0, row=25)

save_btn = tk.Button(window, text="Save Settings", command=save_settings)
save_btn.grid(column=1, row=25)

res = tk.Label(window, text="Progress")
res.grid(column=0, row=26)

bar = ttk.Progressbar(window, length=200)
bar.grid(column=0, row=27)

status = tk.Label(window, text='', fg='black')
status.grid(column=1, row=27, sticky='W')

if saved_reward == "Yes":
    show_codes()
if saved_stores == "Yes":
    show_stores()
    SNUM_box.insert(0, saved_storeID)

if saved_cat == "None":
    print ("No Category Loaded")
else:
    cat_index = cat_box['values'].index(saved_cat)
    cat_box.current(cat_index)
    list_sub_cats()

if saved_subcat == "None":
    print ("No SubCategory Loaded")
else:
    sub_cat_index = sub_cat_box['values'].index(saved_subcat)
    sub_cat_box.current(sub_cat_index)

window.mainloop()
