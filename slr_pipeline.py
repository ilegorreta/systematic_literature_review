#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
from tkinter.constants import EXTENDED
from tkinter.font import BOLD
import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter.filedialog import askopenfilename, asksaveasfile
from functools import partial
from matplotlib import pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)

class Data:
    def __init__(self):
        self.scopus_path = None
        self.wos_path = None
        self.df = None

    def print_properties(self):
        print(f"Scopus: {self.scopus_path}")
        print(f"WoS: {self.wos_path}")



def get_country_wos(address):
    country = address.split("]")[1].split(";")[0].split(",")[-1].strip()
    if "USA" in country:
        return "United States"
    elif "China" in country:
        return "China"
    else:
        return country

def get_q(title, df_q):
    try:
        cuartil = df_q.loc[df_q["Title"] == title, "SJR Best Quartile"].values[0]
        if cuartil == "-":
            return "No rank"
        else:
            return cuartil
    except:
        #print(f"Revista: {title} no encontrada en Scimagojr")
        return "No rank"

def main(data):
    #df_scopus = pd.read_csv('/home/ilegorreta/Downloads/scopus.csv')
    df_scopus = pd.read_csv(data.scopus_path)
    df_scopus["Times Cited"] = df_scopus["Cited by"].apply(lambda x: int(x) if pd.notnull(x) else "No data")
    df_scopus["Affiliations"] = df_scopus["Affiliations"].apply(lambda x: x.split(",")[-1].strip() if pd.notnull(x) else "No data")
    df_scopus = df_scopus[['Authors', 'Title', 'Year', 'Affiliations', "Times Cited", 'DOI', 'Source title', 'Abstract', 'Author Keywords', 'Document Type', 'Source']]
    df_scopus.rename(columns={'Title': 'Article Title', 'Year': 'Publication Year', 'Source title': 'Source Title'}, inplace=True)
    #df_wos = pd.read_excel('/home/ilegorreta/Downloads/wos_data.xls')
    df_wos = pd.read_excel(data.wos_path)
    df_wos["Affiliations"] = df_wos["Addresses"].apply(lambda x: get_country_wos(x))
    df_wos = df_wos[['Authors', 'Article Title', 'Publication Year', 'Affiliations', 'Cited Reference Count', 'DOI', 'Source Title', 'Abstract', 'Author Keywords', 'Document Type']]
    df_wos.rename(columns={"Cited Reference Count": "Times Cited"}, inplace=True)
    df_wos['Source'] = "WoS"
    df = pd.concat([df_scopus, df_wos])
    df.reset_index(inplace=True, drop=True)
    df['Index'] = df.index.values
    df['Index'] += 1
    df = df[['Index', 'Authors', 'Article Title', 'Publication Year', 'Affiliations', 'Times Cited', 'DOI', 'Source Title', 'Abstract', 'Author Keywords', 'Document Type', 'Source']]
    df.loc[df["Source"] == "Scopus", "Authors"] = df["Authors"].apply(lambda x: x.split(",")[0])
    df.loc[df["Source"] == "WoS", "Authors"] = df["Authors"].apply(lambda x: x.split(";")[0])
    df.fillna("No data", inplace=True)
    df["Article Title"] = df["Article Title"].apply(lambda x: x.title())
    df.loc[df.duplicated(subset='DOI', keep=False), "Source"] = "Scopus/WoS"
    df.loc[df.duplicated(subset='Article Title', keep=False), "Source"] = "Scopus/WoS"
    df.loc[df.duplicated(subset='Abstract', keep=False), "Source"] = "Scopus/WoS"
    df.drop_duplicates(subset='DOI', inplace=True)
    df.drop_duplicates(subset='Article Title', inplace=True)
    df.drop_duplicates(subset='Abstract', inplace=True)
    df_q = pd.read_csv('/home/ilegorreta/Downloads/scimagojr_2020.csv', sep=';')
    df["Cuartil"] = df["Source Title"].apply(lambda x: get_q(x, df_q))
    print(f"Final DF Columns: {df.columns}")
    data.df = df

def printTable(data, table):
    df = data.df
    cols = list(df.columns)
    newWindow = tk.Toplevel(window)
    tree = ttk.Treeview(newWindow)
    if table == "full":
        tree.pack(side="top")
    elif table == "esc":
        tree.pack(side="right")
    tree["columns"] = cols
    for i in cols:
        tree.column(i, anchor="w")
        tree.heading(i, text=i, anchor='w')
    for index, row in df.iterrows():
        tree.insert("",0,text=index,values=list(row))

def get_scopus_path(data):
    scopus_data_path = askopenfilename(initialdir = os.getcwd(), title = "Seleccione archivo SCOPUS")
    data.scopus_path = scopus_data_path

def get_wos_path(data):
    wos_data_path = askopenfilename(initialdir = os.getcwd(), title = "Seleccione archivo Web of Science")
    data.wos_path = wos_data_path

def export_as_excel(data):
    df = data.df
    try:
        with asksaveasfile(mode='w', defaultextension=".xlsx") as file:
            df.to_excel(file.name)
            #df.to_csv(file.name, index=False)
    except Exception as e:
        print("The user cancelled save")
        print(e)

def count_kw(item, kw_count):
    for kw in item:
        if kw in kw_count.keys():
            kw_count[kw] += 1
        else:
            kw_count[kw] = 1

def plot_keywords(data):
    df = data.df
    df["Author Keywords 2"] = df['Author Keywords'].str.lower().str.split(";").apply(lambda x: [kw.strip() for kw in x])
    kw_count = {}
    df["Author Keywords 2"].apply(lambda x: count_kw(x, kw_count))
    df.drop("Author Keywords 2", axis=1, inplace=True)
    try:
        del kw_count["no data"]
    except:
        pass
    keywords_df = pd.DataFrame(kw_count.items(), columns=['Keywords', 'Frequency'])
    keywords_df.set_index('Keywords', inplace=True)
    newWindow = tk.Toplevel(window)
    figure = plt.Figure(figsize=(7,7), dpi=100, tight_layout=True)
    ax = figure.add_subplot(111)
    chart_type = FigureCanvasTkAgg(figure, newWindow)
    chart_type.get_tk_widget().pack()
    keywords_df.sort_values(by=['Frequency'], ascending=False).iloc[:15].plot(kind="bar", legend=False, fontsize=14, color="Orange", ax=ax)
    ax.set_title("Keywords Más Frecuentes", fontsize=24, fontweight='bold')
    ax.set_xlabel("Author Keywords", fontsize=20, fontweight='bold')
    ax.set_ylabel("Frecuencia", fontsize=20, fontweight='bold')

def save_keywords(data):
    df = data.df
    df["Author Keywords 2"] = df['Author Keywords'].str.lower().str.split(";").apply(lambda x: [kw.strip() for kw in x])
    kw_count = {}
    df["Author Keywords 2"].apply(lambda x: count_kw(x, kw_count))
    df.drop("Author Keywords 2", axis=1, inplace=True)
    try:
        del kw_count["no data"]
    except:
        pass
    keywords_df = pd.DataFrame(kw_count.items(), columns=['Keywords', 'Frequency'])
    keywords_df.set_index('Keywords', inplace=True)
    figure = plt.Figure(figsize=(7,7), dpi=100, tight_layout=True)
    ax = figure.add_subplot(111)
    keywords_df.sort_values(by=['Frequency'], ascending=False).iloc[:15].plot(kind="bar", legend=False, fontsize=14, color="Orange", ax=ax)
    ax.set_title("Most Cited Articles (By index)", fontsize=24, fontweight='bold')
    ax.set_xlabel("Author Keywords", fontsize=20, fontweight='bold')
    ax.set_ylabel("Frequency", fontsize=20, fontweight='bold')
    try:
        with asksaveasfile(mode='w', defaultextension=".png") as file:
            figure.savefig(file.name)
    except Exception as e:
        print("The user cancelled save")
        print(e)

def plot_citas(data):
    df = data.df
    df_cited = pd.DataFrame(df[['Index', 'Times Cited']], columns=['Index', 'Times Cited'])
    newWindow = tk.Toplevel(window)
    figure = plt.Figure(figsize=(14,14), dpi=100, tight_layout=True)
    ax = figure.add_subplot(111)
    chart_type = FigureCanvasTkAgg(figure, newWindow)
    chart_type.get_tk_widget().pack()
    df_cited.loc[df_cited['Times Cited'] != "No data", :].sort_values(by=['Times Cited'], ascending=False).iloc[:15].plot(kind="bar", legend=False, fontsize=14, color="Orange", ax=ax, x="Index", y="Times Cited")
    ax.set_title("Most Cited Articles (By index)", fontsize=24, fontweight='bold')
    ax.set_xlabel("Article Index", fontsize=20, fontweight='bold')
    ax.set_ylabel("Times Cited", fontsize=20, fontweight='bold')

def save_cites(data):
    df = data.df
    df_cited = pd.DataFrame(df[['Index', 'Times Cited']], columns=['Index', 'Times Cited'])
    #newWindow = tk.Toplevel(window)
    figure = plt.Figure(figsize=(14,14), dpi=100, tight_layout=True)
    ax = figure.add_subplot(111)
    #chart_type = FigureCanvasTkAgg(figure, newWindow)
    #chart_type.get_tk_widget().pack()
    df_cited.loc[df_cited['Times Cited'] != "No data", :].sort_values(by=['Times Cited'], ascending=False).iloc[:15].plot(kind="bar", legend=False, fontsize=14, color="Orange", ax=ax, x="Index", y="Times Cited")
    ax.set_title("Keywords Más Frecuentes", fontsize=24, fontweight='bold')
    ax.set_xlabel("Article Index", fontsize=20, fontweight='bold')
    ax.set_ylabel("Times Cited", fontsize=20, fontweight='bold')
    try:
        with asksaveasfile(mode='w', defaultextension=".png") as file:
            figure.savefig(file.name)
    except Exception as e:
        print("The user cancelled save")
        print(e)

def set_master_window():
    window = tk.Tk()
    window.title('File Explorer')
    #window.geometry("500x500")
    #window.config(background = "yellow")
    greeting = "Hola, este es un programa que realiza un SLR automatico!\n\
                Desarrollado por: Iván Legorreta González\n\
                Asesor: Dr. Edgar Omar López Caudana\n\
                Coordinador: Dr. Raúl Crespo"
    frame = tk.Frame(master=window)
    label = tk.Label(master=frame, text=greeting, fg="white", bg="black")
    label.pack()
    b_scopus = tk.Button(frame, text ="1. Obtener Scopus Path", command = partial(get_scopus_path, data))
    b_wos = tk.Button(frame, text ="2. Obtener WoS Path", command = partial(get_wos_path, data))
    slr = tk.Button(frame, text ="3. Systematic Literature Review", command = partial(main, data), bg="green", fg="white")
    df_table = tk.Button(frame, text ="Desplegar data procesada", command = partial(printTable, data, "full"))
    to_xlsx = tk.Button(frame, text ="Exportar data procesada en formato Excel", command = partial(export_as_excel, data))
    kw_plot = tk.Button(frame, text ="Desplegar Grafica de Keywords", command = partial(plot_keywords, data))
    cites_plot = tk.Button(frame, text ="Desplegar Grafica de Citas", command = partial(plot_citas, data))
    kw_save = tk.Button(frame, text ="Guardar Grafica de Keywords", command = partial(save_keywords, data))
    cites_save = tk.Button(frame, text ="Guardar Grafica de Citas", command = partial(save_cites, data))
    exit = tk.Button(frame, text ="Salir", command = window.destroy, bg="red", fg="white")
    b_scopus.pack()
    b_wos.pack()
    slr.pack()
    df_table.pack()
    to_xlsx.pack()
    kw_plot.pack()
    kw_save.pack()
    cites_plot.pack()
    cites_save.pack()
    exit.pack()
    frame.pack()
    return window

if __name__ == '__main__':
    data = Data()
    window = set_master_window()
    window.mainloop()