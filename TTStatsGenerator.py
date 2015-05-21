'''
Created on 18/05/2015

@author: Silvio
'''
import csv
import os
import xlsxwriter
from datetime import datetime, date

def main():
    #reseteamos el fichero de estadisticas
    try:
        os.remove("./stats.xlsx")
    except:
        pass
    try:
        for path, directories, files in os.walk("./"):
            for fil in files:
                # ignore files without extension (can have the same name as the ext)
                file_ext = fil.split('.')[-1] if len(fil.split('.')) > 1 else None
                # ignore dots in given extensions
                extensions = [ext.replace('.', '') for ext in ["csv",]]
                if file_ext in extensions:
                    process(path, fil)
    except:
        pass
    f = open("./TimeTravel.log", "r")
    dict = {'foo' : 1,}
    for linea in f:
        encuentra = encuentraStatus(linea)
        if encuentra != -1:
            anno, mes, dia, hora, minuto = parseaFecha(linea)
            seccion, status, timetravel = parseaTimeTravel(linea)
            guardaSeccion(dia,mes,anno,hora,minuto,seccion,status,timetravel)
            dict[seccion] = 1
            
    del dict['foo']
    
    workbook = xlsxwriter.Workbook('stats.xlsx')
    
    for sec in dict.keys():
        csvToexcel(sec, workbook)
        
    workbook.close()
    #volvemos a borrar los csv
    try:
        for path, directories, files in os.walk("./"):
            for fil in files:
                # ignore files without extension (can have the same name as the ext)
                file_ext = fil.split('.')[-1] if len(fil.split('.')) > 1 else None
                # ignore dots in given extensions
                extensions = [ext.replace('.', '') for ext in ["csv",]]
                if file_ext in extensions:
                    process(path, fil)
    except:
        pass


def process(the_path, the_file):
    processed = 0
    src_file = os.path.join(the_path, the_file)
    
    os.remove(src_file)
    processed = 1
    return processed

def csvToexcel(seccion, workbook):
    f = open('./'+seccion+'.csv', "rU")

    csv.register_dialect('blank', delimiter=' ')

    reader = csv.reader(f, dialect='blank')
    
    worksheet = workbook.add_worksheet(seccion)

    marca_default = False
    max_row = 0
    for row_index, row in enumerate(reader):
        max_row = max_row +1
        for col_index, col in enumerate(row):
            if(col_index == 4):
                worksheet.write_number(row_index, col_index, int(col))
                if marca_default ==True:
                    worksheet.write_number(row_index, col_index+20, int(col))
                else:
                    worksheet.write(row_index, col_index+20, '')
            else:
                 worksheet.write(row_index, col_index, col)
                 if(col_index==3):
                    if col=="default":
                        marca_default = True
                    else:
                        marca_default = False


    # Create a new chart object.
    chart = workbook.add_chart({'type': 'line'})
    chart.set_title({'name': 'Time Travel (secs)'})
    chart.set_legend({'none': True})

    chart.set_x_axis({'num_font':  {'rotation': 45}})

    # Add a series to the chart.
    chart.add_series({'categories' :'='+str(seccion)+'!$B$1:$B$'+str(max_row),
                      'values': '='+str(seccion)+'!$E$1:$E$'+str(max_row),
                      'marker': {
                      'type': 'automatic',
                      'size': 6,
                      'border': {'color': 'black'},
                      'fill':   {'color': 'blue'},
                        },
                      })
    
    chart2 = workbook.add_chart({'type': 'scatter'})
    
    chart2.add_series({'categories' :'='+str(seccion)+'!$B$1:$B$'+str(max_row),
                      'values': '='+str(seccion)+'!$Y$1:$Y$'+str(max_row),
                      'marker': {
                      'fill':   {'color': 'red'},
                      'type': 'diamond',
                      'size': 6,
                        },
                      })
    
    chart.combine(chart2)
        
    # Insert the chart into the worksheet.
    worksheet.insert_chart('C1', chart)
    #worksheet.insert_chart('C2', chart2)


def guardaSeccion(dia,mes,anno,hora,minuto,seccion,status,timetravel):
    f = open(seccion+".csv",'a')
    f.write(dia+"/" +mes+"/" +anno+ " " +hora+ ":"+ minuto+" "+seccion+" "+status+" "+ timetravel+"\n")
    f.close()        

def parseaTimeTravel(linea):
    seccion =  ((linea.split('|')[1]).split(':')[0])[1:8]
    if linea.count(':') ==4:
      status = "ok"
      timetravel = ((linea.split('|')[1]).split('(')[0]).split(':')[1]
    else:
        try:
            status = "default"
            timetravel = ((linea.split('|')[1]).split(')')[1]).split(':')[1]
        except:
            timetravel = ""
            status = ""
    return seccion.strip(), status.strip(), timetravel.strip()

def parseaFecha(linea):
    fechaAnno = linea.split()[0]
    fechaHora = linea.split()[1]
    
    dia = fechaAnno.split('/')[0]
    mes = fechaAnno.split('/')[1]
    anno = fechaAnno.split('/')[2]

    hora = fechaHora.split(':')[0]
    minuto = fechaHora.split(':')[1]
    
    return anno, mes, dia, hora, minuto

def encuentraStatus(linea):
    line = linea.split('|')[1]
    try:
        if line[9] == 't' and line[8]!='n':
            return linea
        else:
            return -1
    except:
        return -1
    
if __name__ == '__main__':
    main()