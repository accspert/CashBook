from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib.units import inch, cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import Image

from datetime import datetime
from datetime import date
import os
from decimal import *
getcontext().prec = 20

class reportObj:
    def __init__(self, language, table, date_from = 0, date_to = 0):
        
        self.summe_einnahme = 0.0
        self.summe_ausgabe = 0.0
        self.summe_entMwSt = 0.0
        self.data = table
        self.date_from = date_from
        self.date_to = date_to
        self.language = language
        
        if(self.language == "de"):
            headers = ("BuchId", "Datum", "Einnahme", "Ausgabe", "Text", "Beleg", "MwSt", "Enth MwSt")
        elif(self.language == "en"):
            headers = ("Id", "Date", "Income", "Expense", "Text", "ReceiptNr", "Tax", "Incl Tax")
        elif(self.language == "es"):
            headers = ("Id", "Fecha", "Ingreso", "Gasto", "Texto", "No recibo", "IVA", "IVA incor")
        table.insert(0, headers)
        # append the tax incl of income
        for i in range(1, len(table)):
            if table[i][6]:
                einKomma = table[i][6]/100.0
                nullKomma = (table[i][6] / 100.0) + 1
                if table[i][2]:
                    entMwSt = table[i][2] / nullKomma * einKomma
                    self.summe_einnahme = self.summe_einnahme + table[i][2]
                if table[i][3]:
                    entMwSt = table[i][3] / nullKomma * einKomma
                    self.summe_ausgabe = self.summe_ausgabe + table[i][3]
            self.summe_entMwSt = self.summe_entMwSt + entMwSt    
            table[i] = table[i] + (str('{0:.2f}'.format(entMwSt)),)
            
    def generate_journal_report(self, file_name):
        styles = getSampleStyleSheet()
        styleN = styles['Normal']
        styleH = styles['Heading1']
        story = []

        doc = SimpleDocTemplate(
            file_name,
            pagesize=landscape(A4),
            bottomMargin=.4 * inch,
            topMargin=.6 * inch,
            rightMargin=.4 * inch,
            leftMargin=.4 * inch)

        if self.language == 'de': 
            text_journal = "Journal"
            text_from_to = "Vom " + str(self.date_from) + " Bis " + str(self.date_to)
            text_sum_income = "Summe Einnahme: " + str(self.summe_einnahme)
            text_sum_expensis = "Summe Ausgabe: " + str(self.summe_ausgabe)
            text_tax = "Summe MwSt: " + str('{0:.2f}'.format(self.summe_entMwSt))
            
        elif self.language == 'en':
            text_journal = "Journal"
            text_from_to = "From " + str(self.date_from) + " To " + str(self.date_to)
            text_sum_income = "Sum Income: " + str(self.summe_einnahme)
            text_sum_expensis = "Sum expense: " + str(self.summe_ausgabe)
            text_tax = "Sum tax: " + str('{0:.2f}'.format(self.summe_entMwSt))
        elif self.language == 'es':
            text_journal = 'Diario'
            text_from_to = "A " + str(self.date_from) + " de " + str(self.date_to)
            text_sum_income = "Ingresos totales: " + str(self.summe_einnahme)
            text_sum_expensis = "Gastos totales: " + str(self.summe_ausgabe)            
            text_tax = "IVA total: " + str('{0:.2f}'.format(self.summe_entMwSt))
        
        # P = Paragraph(text_journal, styleH)
        # story.append(P)
        
        # P = Paragraph(text_from_to, styleN)
        # story.append(Spacer(1, 0.3*cm))
        # story.append(P)
        
        ElemWidth = 300
        logo = Image('logo.png')
        logo.drawWidth = 30
        logo.drawHeight = 30
        
        titleTable = Table([
            [text_journal, text_from_to, logo]
        ], ElemWidth)
        
        titleTableStyle = TableStyle([
            ('ALIGN',(0,0),(-1,-1),'CENTER'),
            ('FONTSIZE', (0,0), (-1,-1), 12),
            ('FONTNAME', (0,0), (-1,-1), 
                'Helvetica-Oblique'
                ), 
    
            ('TOPPADDING',(0,0),(-1,-1), 0),
            ('BOTTOMPADDING',(0,0),(-1,-1), 0), 
        ])
        titleTable.setStyle(titleTableStyle)
        story.append(Spacer(1, 0.5*cm))
        story.append(titleTable)
        
        story.append(Spacer(1, 1.0*cm))

        t=Table(self.data,style=[('GRID',(1,1),(-2,-2),1,colors.black),
        ('GRID',(0,0),(-1,-1),0.5,colors.black)])
        story.append(Spacer(1, 0.5*cm))
        story.append(t)

        P = Paragraph(text_sum_income, styleN)
        story.append(Spacer(1, 0.5*cm))
        story.append(P)
        
        P = Paragraph(text_sum_expensis, styleN)
        story.append(Spacer(1, 0.5*cm))
        story.append(P)
        
        P = Paragraph(text_tax, styleN)
        story.append(Spacer(1, 0.5*cm))
        story.append(P)

        doc.build(
            story
        )

        os.startfile(file_name)
