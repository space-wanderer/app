import pyodbc
from datetime import datetime
from openpyxl import load_workbook

sch_id = 13583
id_value, spisok_value = [], []
value = {}

class Sql:
    def __init__(self, database, server="11-01"):
        self.cnxn = pyodbc.connect("Driver={SQL Server Native Client 11.0};"
                                   "Server="+server+";"
                                   "Database="+database+";"
                                   "Trusted_Connection=yes;")
        self.query = "-- {}\n\n-- Made in Python".format(datetime.now()
                                                         .strftime("%d/%m/%Y"))
def del_whitespace(value):
    temp = value.strip()
    if len(temp) == 3 or len(temp) == 5:
        return temp
    elif len(temp) == 6:
        temp_split = temp.split(' ')
        return temp_split[0] + temp_split[1]  

# дописать коды услуги
def name_CODE_USL(value):
    temp = value.strip()
    if temp == 'EGFR':
        return  'A.26.20.026'
    elif temp == 'KRAS':
        return 'A.26.20.026'
    elif temp == 'BRAF':
        return 'A.26.20.026'
    elif temp == 'NRAS':
        return 'A.26.20.026'
    elif temp == 'BRCA 1,2':
        return 'A.26.20.026'
    else:
        return 'код'                                                                     

def date_sql(value):
    return value.strftime("%d/%m/%Y")
def Quere(value):
    query_sql = f"""
        DECLARE @zsl_id UNIQUEIDENTIFIER 
        DECLARE @ID_PAC UNIQUEIDENTIFIER
        DECLARE @SL_ID UNIQUEIDENTIFIER 
        DECLARE  @D3_USLGID UNIQUEIDENTIFIER
        DECLARE  @D3_USLGID_2 UNIQUEIDENTIFIER
        DECLARE  @D3_USLGID_3 UNIQUEIDENTIFIER
        DECLARE  @D3_USLGID_4 UNIQUEIDENTIFIER

        DECLARE @D3_PID INT
        declare @D3_ZSLID INT
        DECLARE @D3_SLID INT

        SET @zsl_id = NEWID() 
        SET @ID_PAC = NEWID()
        SET @SL_ID = NEWID()
        SET @D3_USLGID = NEWID()


        INSERT INTO D3_PACIENT_OMS ( D3_SCID, ID_PAC, FAM, IM, OT, W, DR,  SNILS,  VPOLIS,  NPOLIS,  SMO, NOVOR)
        VALUES('{sch_id}', @ID_PAC, '{_fam}', '{_im}', '{_ot}', 1, '{_dr}',  '{_snils}',  3,  '1234567891234567',  '46003', '0')

        SET @D3_PID = (SELECT id FROM D3_PACIENT_OMS WHERE ID_PAC = @ID_PAC)

        INSERT INTO D3_ZSL_OMS ( ZSL_ID, D3_PID, D3_PGID, D3_SCID, DATE_Z_1, DATE_Z_2, NPR_DATE, MTR, USERID)
                 VALUES(@zsl_id, @D3_PID, @ID_PAC, '{sch_id}', '{_date_napr}', '{_date_napr}', '{_date_napr}', 0, 1)                
              
        SET @D3_ZSLID = (SELECT id FROM d3_zsl_oms WHERE ZSL_ID = @zsl_id)                
                
        INSERT INTO D3_SL_OMS ( D3_ZSLID, D3_ZSLGID, SL_ID, DATE_1, DATE_2, DS1, VERS_SPEC, WEI)
        VALUES(@D3_ZSLID, @zsl_id, @SL_ID, '{_date_napr}', '{_date_napr}','C00.3', 'V021', 0)

        SET @D3_SLID = (SELECT id FROM d3_sl_oms WHERE SL_ID = @SL_ID)

        INSERT INTO D3_USL_OMS (D3_SLID, 	D3_ZSLID, 	D3_SLGID, 	DATE_IN, 	DATE_OUT, DS, VERS_SPEC, 	D3_USLGID, CODE_USL,  VID_VME)
        VALUES (@D3_SLID, @D3_ZSLID, @SL_ID, '{_date_napr}', '{_date_napr}', 'C00.3', 'V021', @D3_USLGID, '{_CODE_USL_1}', '{_CODE_USL_1}') 
        """
    CODE_USL_2 = """
        SET @D3_USLGID_2 = NEWID()
        INSERT INTO D3_USL_OMS (D3_SLID, 	D3_ZSLID, 	D3_SLGID, 	DATE_IN, 	DATE_OUT, DS, VERS_SPEC, 	D3_USLGID, CODE_USL, VID_VME)
        VALUES (@D3_SLID, @D3_ZSLID, @SL_ID, '{_date_napr}', '{_date_napr}', 'C00.3', 'V021', @D3_USLGID_2, '{_CODE_USL_2}', '{_CODE_USL_2}')
    """
    CODE_USL_3 = """
        SET @D3_USLGID_3 = NEWID()
        INSERT INTO D3_USL_OMS (D3_SLID, 	D3_ZSLID, 	D3_SLGID, 	DATE_IN, 	DATE_OUT, DS, VERS_SPEC, 	D3_USLGID, CODE_USL, VID_VME)
        VALUES (@D3_SLID, @D3_ZSLID, @SL_ID, '{_date_napr}', '{_date_napr}', 'C00.3', 'V021', @D3_USLGID_3, '{_CODE_USL_3}', '{_CODE_USL_3}')    
    """
    CODE_USL_4 = """
    SET @D3_USLGID_2 = NEWID()
        INSERT INTO D3_USL_OMS (D3_SLID, 	D3_ZSLID, 	D3_SLGID, 	DATE_IN, 	DATE_OUT, DS, VERS_SPEC, 	D3_USLGID, CODE_USL, VID_VME)
        VALUES (@D3_SLID, @D3_ZSLID, @SL_ID, '{_date_napr}', '{_date_napr}', 'C00.3', 'V021', @D3_USLGID_4, '{_CODE_USL_4}', '{_CODE_USL_4}')
    """ 
    if value == 'query_sql':
        return query_sql
    elif value == 'CODE_USL_1':
        return query_sql + CODE_USL_2
    elif value == 'CODE_USL_3':
        return query_sql + CODE_USL_2 + CODE_USL_3
    elif value == 'CODE_USL_4':
        return query_sql + CODE_USL_2 + CODE_USL_3 + CODE_USL_4   


wb = load_workbook('./Белгород.xlsx')

sheet = wb.active

for row in sheet['A']:  
    if row.value:
        id_value.insert(0, row.value)     

for row in sheet:
    for cell in row: 
        # for val in cell:
        # print(cell.value, end = ' ')
        # spisok.append(cell.value)
        if cell.value != None:
            spisok_value.append(cell.value)    

for id in id_value: 
    for index, k in enumerate(spisok_value):
        if spisok_value[index] == id:
            #print(index, k)
            #id_spisok.append(index)
            temp = spisok_value[index + 1:]
            del spisok_value[index:]
            value[id] = temp
            # print(spisok_value[index], end=' ')



for key  in value:
    if len(value[key]) == 9: 
        _date_napr = value[key][0].strftime("%d/%m/%Y") 
        _fam = str(value[key][1])
        _im = str(value[key][2])
        _ot = str(value[key][3])
        _dr = value[key][4].strftime("%d/%m/%Y") 
        _snils = str(value[key][5])
        _enp = str(value[key][6])
        _ds1 = str(del_whitespace(value[key][7])) 
        _CODE_USL_1 = str(name_CODE_USL(value[key][8]))
               
        sql = Sql('elmedicine')

        cursor = sql.cnxn.cursor()

        cursor.execute(Quere('query_sql'))
        cursor.commit()
        cursor.close()
        sql.cnxn.close()
        
    elif len(value[key]) == 10:  
        _date_napr = value[key][0].strftime("%d/%m/%Y") 
        _fam = str(value[key][1])
        _im = str(value[key][2])
        _ot = str(value[key][3])
        _dr = value[key][4].strftime("%d/%m/%Y") 
        _snils = str(value[key][5])
        _enp = str(value[key][6])
        _ds1 = str(del_whitespace(value[key][7])) 
        _CODE_USL_1 = str(name_CODE_USL(value[key][8]))
        _CODE_USL_2 = str(name_CODE_USL(value[key][9]))

        query_sql = ()
        sql = Sql('elmedicine')
        cursor = sql.cnxn.cursor()
        cursor.execute(Quere('CODE_USL_2'))
        cursor.commit()
        cursor.close()
        sql.cnxn.close()

    elif len(value[key]) == 11:  
        _date_napr = value[key][0].strftime("%d/%m/%Y") 
        _fam = str(value[key][1])
        _im = str(value[key][2])
        _ot = str(value[key][3])
        _dr = value[key][4].strftime("%d/%m/%Y") 
        _snils = str(value[key][5])
        _enp = str(value[key][6])
        _ds1 = str(del_whitespace(value[key][7])) 
        _CODE_USL_1 = str(name_CODE_USL(value[key][8]))
        _CODE_USL_2 = str(name_CODE_USL(value[key][9]))
        _CODE_USL_3 = str(name_CODE_USL(value[key][9]))

        query_sql = ()
        sql = Sql('elmedicine')
        cursor = sql.cnxn.cursor()
        cursor.execute(Quere('CODE_USL_3'))
        cursor.commit()
        cursor.close()
        sql.cnxn.close()

    elif len(value[key]) == 12:  
        _date_napr = value[key][0].strftime("%d/%m/%Y") 
        _fam = str(value[key][1])
        _im = str(value[key][2])
        _ot = str(value[key][3])
        _dr = value[key][4].strftime("%d/%m/%Y") 
        _snils = str(value[key][5])
        _enp = str(value[key][6])
        _ds1 = str(del_whitespace(value[key][7])) 
        _CODE_USL_1 = str(name_CODE_USL(value[key][8]))
        _CODE_USL_2 = str(name_CODE_USL(value[key][9]))
        _CODE_USL_3 = str(name_CODE_USL(value[key][9]))
        _CODE_USL_4 = str(name_CODE_USL(value[key][9]))

        query_sql = ()
        sql = Sql('elmedicine')
        cursor = sql.cnxn.cursor()
        cursor.execute(Quere('CODE_USL_4'))
    else:
        print(value[key])    
