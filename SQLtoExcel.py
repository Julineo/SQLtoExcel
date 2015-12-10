import pyodbc
import xlsxwriter

def AFLbyFac():
    cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER=cna-eis-01;DATABASE=eIS;Trusted_Connection=yes')
    #reading filenames
    cursor = cnxn.cursor()
    cursor.execute("SELECT DISTINCT [COL21_EMPLOYER] FROM [eIS].[dbo].[CNAtmpAFL_NNU-201506] WHERE [COL21_EMPLOYER] IS NOT NULL ORDER BY 1")
    SQLrows = cursor.fetchall()
    FileNames = []
    for SQLrow in SQLrows:
        FileNames.append(str(SQLrow.COL21_EMPLOYER).replace("'","''"))
    
    i = 1
    
    for files in FileNames:
              
        workbook = xlsxwriter.Workbook(str(i) + '.xlsx')
        worksheet = workbook.add_worksheet()
        format = workbook.add_format() #Creating format for Excel headers
        format.set_bold()
        worksheet.set_row(0, None, format) #Setting headers format
        worksheet.set_column('B:E', 18) #Changing columns widths
        worksheet.set_column('I:I', 18)
        worksheet.set_column('J:K', 30)
        worksheet.set_column('L:L', 18)
        worksheet.write('A1', 'AFF_ID')
        worksheet.write('B1', 'FIRST_NAME')
        worksheet.write('C1', 'LAST_NAME')
        worksheet.write('D1', 'ADDRESS')
        worksheet.write('E1', 'CITY')
        worksheet.write('F1', 'STATE')
        worksheet.write('G1', 'ZIP')
        worksheet.write('H1', 'GENDER')
        worksheet.write('I1', 'HOME_PHONE')
        worksheet.write('J1', 'EMAIL')
        worksheet.write('K1', 'EMPLOYER')
        worksheet.write('L1', 'PARTYAFFILIATION')
    
        query = "SELECT [COL1_AFFILIATE_ID],[COL32_FIRST_NAME],[COL34_LAST_NAME],[COL3_ADDRESS],[COL4_CITY],[COL5_STATE],[COL6_ZIP],[COL10_GENDER],[COL11_HOME_PHONE],[COL13_EMAIL],[COL21_EMPLOYER],[PARTYAFFILIATION] FROM [eIS].[dbo].[CNAtmpAFL_NNU-201506] WHERE COL21_EMPLOYER = '" + FileNames[i-1] + "' ORDER BY 1"
        #INSERTING DATA
        cursor = cnxn.cursor()
        cursor.execute(query)
        SQLrows = cursor.fetchall()
        
        row = 1
        for SQLrow in SQLrows:
            #print row.COL1_AFFILIATE_ID, row.COL32_FIRST_NAME, row.COL34_LAST_NAME, row.COL3_ADDRESS, row.COL4_CITY, row.COL5_STATE, row.COL6_ZIP, row.COL10_GENDER, row.COL11_HOME_PHONE, row.COL13_EMAIL, row.COL21_EMPLOYER, row.PARTYAFFILIATION
            worksheet.write(row, 0, SQLrow.COL1_AFFILIATE_ID)
            worksheet.write(row, 1, SQLrow.COL32_FIRST_NAME)
            worksheet.write(row, 2, SQLrow.COL34_LAST_NAME)
            worksheet.write(row, 3, SQLrow.COL3_ADDRESS)
            worksheet.write(row, 4, SQLrow.COL4_CITY)
            worksheet.write(row, 5, SQLrow.COL5_STATE)
            worksheet.write(row, 6, SQLrow.COL6_ZIP)
            worksheet.write(row, 7, SQLrow.COL10_GENDER)
            worksheet.write(row, 8, SQLrow.COL11_HOME_PHONE)
            worksheet.write(row, 9, SQLrow.COL13_EMAIL)
            worksheet.write(row, 10, SQLrow.COL21_EMPLOYER)
            worksheet.write(row, 11, SQLrow.PARTYAFFILIATION)
            row = row + 1
        i = i + 1
        workbook.close()
    return 5
