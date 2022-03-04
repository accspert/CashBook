# -*- coding: utf-8 -*-
"""
Created on Fri May  7 10:05:04 2021

@author: Egon
"""
import sqlite3
from ErrorLogger import *

class SqlHelper:
    
    def __init__(self,name=None):
        self.conn = None
        self.cursor = None
        
        if name:
            self.open(name)
            
    def open(self, name):
        try:
            self.conn = sqlite3.connect(name)
            self.cursor = self.conn.cursor()
            print(sqlite3.version)
        except Exception as e:
                    ErrorLogger.WriteError(traceback.format_exc())
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))       
    def edit (self,query,updates):  # Update
        c = self.cursor
        c.execute(query,updates)
        self.conn.commit()  
    def delete(self,query):         # Delete
        c = self.cursor
        c.execute(query)
        self.conn.commit()
    def insert(self,query,inserts): # Insert
        c = self.cursor
        c.execute(query,inserts)
        self.conn.commit()
    def insert_many(self, query, inserts):  # Insert many
        c = self.cursor
        c.executemany(query, inserts) 
        self.conn.commit()        
    def select (self,query):        # Select *
        c = self.cursor
        c.execute(query)
        return c.fetchall()   
    def select_one (self,query):        # Select * one
        c = self.cursor
        c.execute(query)
        return c.fetchone()[0]      
    def select_para (self,query,selects): # Select with parameter
        c = self.cursor
        c.execute(query,selects)
        return c.fetchone()
    def select_para_all (self,query,selects): # Select with parameter
        c = self.cursor
        c.execute(query,selects)
        return c.fetchall()        

