'''
Created on 18/05/2015

@author: Silvio
'''
import os
import xlsxwriter
import collections

class Seccion():
    
    def __init__(self):
        self.name = None
        self.status = None
        self.timeTravel = None
        self.dia = None
        self.mes = None
        self.anno = None
        self.hora = None
        self.minuto = None
        self.matches = None
        self.worksheet = None
        self.max_row = 0

    def __repr__(self):
        if self.name == None:
            str = "Name: "+"None"
        else:
            str = "Name: "+self.name
        if self.status == None:
            str += "\nStatus: "+"None"
        else:
            str += "\nStatus: "+self.status
        if self.timeTravel == None:
            str += "\nTimeTravel: "+"None"
        else:
            str += "\nTimeTravel: "+self.timeTravel
        str += "\nDate: "+self.dia+"/"+self.mes+"/"+self.anno+" "+self.hora+":"+self.minuto
        if self.matches == None:
            str += "\nMatches: "+"None"
        else:
            str += "\nMatches: "+self.matches
        return str

def main():
    #reseteamos el fichero de estadisticas
    try:
        os.remove("./stats.xlsx")
    except:
        pass
    
    #vamos parseano y guardando status
    f = open("./TimeTravel.log", "r")
    dict = collections.OrderedDict()
    for linea in f:
        encuentra = encuentraStatus(linea)
        if encuentra != -1:
            s = Seccion()
            s.anno, s.mes, s.dia, s.hora, s.minuto = parseaFecha(linea)
            s.name, s.status, s.timeTravel = parseaTimeTravel(linea)
            #print s.name, s.anno, s.mes, s.dia, s.hora, s.minuto
            if dict.has_key(s.name+s.anno+s.mes+s.dia+s.hora+s.minuto)==False:
                dict[s.name+s.anno+s.mes+s.dia+s.hora+s.minuto] = s
            dict[s.name+s.anno+s.mes+s.dia+s.hora+s.minuto].anno = s.anno
            dict[s.name+s.anno+s.mes+s.dia+s.hora+s.minuto].mes = s.mes
            dict[s.name+s.anno+s.mes+s.dia+s.hora+s.minuto].dia = s.dia
            dict[s.name+s.anno+s.mes+s.dia+s.hora+s.minuto].hora = s.hora
            dict[s.name+s.anno+s.mes+s.dia+s.hora+s.minuto].minuto = s.minuto
            dict[s.name+s.anno+s.mes+s.dia+s.hora+s.minuto].name = s.name
            dict[s.name+s.anno+s.mes+s.dia+s.hora+s.minuto].status= s.status
            dict[s.name+s.anno+s.mes+s.dia+s.hora+s.minuto].timeTravel = s.timeTravel
        matchSection = encuentraMatches(linea)
        if matchSection != -1:
            s = Seccion()
            s.anno, s.mes, s.dia, s.hora, s.minuto = parseaFecha(linea)
            s.name = matchSection[0]
            s.matches = matchSection[1]
            dict[s.name+s.anno+s.mes+s.dia+s.hora+s.minuto] = s
    f.close()       
        
    workbook = xlsxwriter.Workbook('stats.xlsx')
    
    count = 1
    #guardamos las secciones de forma temporal en csv (se podria omitir y trabajar in memory)
    for sec in dict.values():
        #si no existe matches es porque al principio del fichero esta incompleto quitamos
        if sec.matches == None:
            del dict[sec.name+sec.anno+sec.mes+sec.dia+sec.hora+sec.minuto]
        else:
            print "Bucle dict.values() seccion: ",sec.name," ",count," de ",len(dict)
            count = count +1
            if sec.anno != None:
                csvToexcel(sec, workbook, dict)

    #creamos graficos
    for sec in dict.values():
        if sec.max_row != 0:
            createGraph(sec, workbook, dict)
        
    print "Saving excel file"
    workbook.close()

def encuentraMatches(linea):
    line = linea.split('|')[1]
    try:
        if line.split()[0] == "Matching":
            #buscamos que sea la misma seccion
            seccion = ((line.split("for")[1]).split("are")[0]).split()[0]
            return [seccion,line.split(':')[1]]
        else:
            return -1
    except:
        return -1
    return -1

def process(the_path, the_file):
    processed = 0
    src_file = os.path.join(the_path, the_file)
    
    os.remove(src_file)
    processed = 1
    return processed

def createGraph(seccion, workbook, dict):
    worksheet = dict[seccion.name].worksheet
    max_row = seccion.max_row
    # Create a new chart object.
    chart = workbook.add_chart({'type': 'line'})
    chart.set_title({'name': 'Time Travel (secs)'})
    chart.set_legend({'none': True})

    chart.set_x_axis({'num_font':  {'rotation': 45}})
    chart.set_size({'width': 560, 'height': 275})

    # Add a series to the chart.
    chart.add_series({'categories' :'='+str(seccion.name)+'!$B$1:$B$'+str(max_row),
                      'values': '='+str(seccion.name)+'!$E$1:$E$'+str(max_row),
                      'marker': {
                      'type': 'automatic',
                      'size': 6,
                      'border': {'color': 'black'},
                      'fill':   {'color': 'blue'},
                        },
                      })
    
    chart2 = workbook.add_chart({'type': 'scatter'})
    
    chart2.add_series({'categories' :'='+str(seccion.name)+'!$B$1:$B$'+str(max_row),
                      'values': '='+str(seccion.name)+'!$Y$1:$Y$'+str(max_row),
                      'marker': {
                      'fill':   {'color': 'red'},
                      'type': 'diamond',
                      'size': 6,
                        },
                      })
    
    chart.combine(chart2)

    chart3 = workbook.add_chart({'type': 'line'})
    chart3.set_title({'name': 'Matches (absolute number)'})
    chart3.set_legend({'none': True})

    chart3.set_x_axis({'num_font':  {'rotation': 45}})
    chart3.set_size({'width': 560, 'height': 275})

    # Add a series to the chart.
    chart3.add_series({'categories' :'='+str(seccion.name)+'!$B$1:$B$'+str(max_row),
                      'values': '='+str(seccion.name)+'!$F$1:$F$'+str(max_row),
                      'marker': {
                      'type': 'triangle',
                      'size': 6,
                      'border': {'color': 'black'},
                      'fill':   {'color': 'green'},
                        },
                      })
        
    # Insert the chart into the worksheet.
    worksheet.insert_chart('J1', chart)
    worksheet.insert_chart('J16', chart3)




def csvToexcel(seccion, workbook, dict):

    #si no existe, creamos
    if dict.has_key(seccion.name) == False:
        worksheet = workbook.add_worksheet(seccion.name)
        #solo guardamos la sheet
        dict[seccion.name] = Seccion()
        dict[seccion.name].name = seccion.name
        dict[seccion.name].worksheet = worksheet
    else:
        worksheet = dict[seccion.name].worksheet

    try:
        max_row = dict[seccion.name].max_row
        #escribimos fecha
        worksheet.write(max_row, 0, seccion.dia+"/"+seccion.mes+"/"+seccion.anno)
        worksheet.write(int(max_row), 1, seccion.hora+":"+seccion.minuto)
        worksheet.write(int(max_row), 2, seccion.name)
        worksheet.write(max_row, 3, seccion.status)
        worksheet.write_number(max_row, 4, int(seccion.timeTravel))
        worksheet.write_number(max_row, 5, int(seccion.matches))
        if seccion.status =="default":
            worksheet.write_number(max_row, 24, int(seccion.timeTravel))
        else:
            worksheet.write(max_row, 24, '')
        #incrementamos el max_row
        dict[seccion.name].max_row = dict[seccion.name].max_row+1
    except Exception as e:
        print "Exception writing cell:  ",e
      

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