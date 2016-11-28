#!/usr/bin/python3
import pyexcel as pe

OS = "Ubuntu 16.04"

class Laptop:
    """Class for laptop object holding specs and price"""
    def __init__(self, id, price, brand, cpu, ram, hdd, optical, notes, battery_life,other_items, screen):
        self.id = id
        self.price = str(price).strip()
        self.brand = str(brand).strip()
        self.cpu = str(cpu).strip()
        self.ram = str(ram).strip()
        self.hdd = str(hdd).strip()
        self.optical = str(optical).strip()
        self.notes = str(notes).strip()
        self.battery_life = str(battery_life).strip()
        self.other_items = str(other_items).strip()
        self.screen = str(screen).strip()
        self.os = OS

    def __repr__(self):
        return "ID: {:5} \
        Brand: {:>40} \
        Price: {:6} \
        CPU: {:20} \
        RAM: {:6} \
        HDD: {:10} \
        Optical: {:20} \
        Battery Life: {:40} \
        Screen Size: {:10} \
        Operarting System: {:40} \
        other_items: {:40} \
        Notes: {:40}" \
        .format(str(self.id),self.brand,self.price,self.cpu,self.ram,
        self.hdd, self.optical,self.battery_life,self.screen,self.os,
        self.other_items,self.notes)

    def toArray(self):
        return [self.id,self.price,self.brand,self.cpu,self.ram,self.hdd,
        self.optical,self.battery_life,self.screen,self.os,
        self.notes,self.other_items]


class Desktop:
    """Class for desktop object holding specs and price"""
    def __init__(self, id, market_price, conc_price, cpu, ram, hdd, specs, notes):
        self.id = id
        self.market_price = str(market_price).strip()
        self.conc_price = str(conc_price).strip()
        self.cpu = str(cpu).strip()
        self.ram = str(ram).strip()
        self.hdd = str(hdd).strip()
        self.specs = str(specs).strip()
        self.notes = str(notes).strip()
        self.os = OS

    def __repr__(self):
        return "ID: {:5}\
        Market Price: {:8}\
        Conc Price: {:8}\
        CPU: {:20}\
        RAM: {:6}\
        HDD: {:10}\
        Operating System: {:40}\
        Specs: {:80}\
        Notes: {:40}" \
        .format(str(self.id), self.market_price, self.conc_price, self.cpu, self.ram,
        self.hdd, self.os, self.specs, self.notes)

    def toArray(self):
        return [self.id,self.market_price,self.conc_price,self.cpu,self.ram,
        self.hdd,self.os,self.specs,self.notes]


def getSheets(filename):
    "Reads the Stock and Laptop Stock from given file"
    book = pe.get_book(file_name=filename)
    return (book["Stock"], book["Laptop Stock"])

def getLaptops(laptop_sheet):
    """Deserializes laptop objects from sheet
    Returns a list of Laptops"""
    done = False
    current_row = 2
    laptop_stock = []
    while not done:
        row = laptop_sheet.row[current_row]
        if(row[1]):
            laptop_stock.append(Laptop(row[1],row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[9],row[10],row[11]))
        else:
            done = True
        current_row+=1
    return laptop_stock

def getDesktops(desktop_sheet):
    """Deserializes desktop objects from sheet
    Returns a list of Desktops"""
    done = False
    current_row = 2
    desktop_stock = []
    while not done:
        row = desktop_sheet.row[current_row]
        if(row[1]):
            desktop_stock.append(Desktop(row[1],row[2],row[3],row[4],row[5],row[6],row[7],row[8]))
        else:
            done = True
        current_row+=1
    return desktop_stock
