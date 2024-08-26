import openpyxl

def main():
    tipos_de_piezas =["ASS","LCD", "LCD GALAXY TAB", "LCD NO ANDROID", 'LCD TAB 7"', 'LCD TAB 9"', 'LCD TAB 10"', 'TOUCH AND', 'TOUCH GTAB', 'TOUCH IPAD', 'TOUCH NO ANDROID', 'TOUCH TAB 8"', 'TOUCH TAB 7.85"', 'TOUCH TAB 7"', 'TOUCH TAB 9"', 'TOUCH TAB 10"', 'ANILLOS PERS', 'BAT DET', 'BAT', 'BAT INTERNA', 'CD', 'CARGADOR PF', 'CARGADOR UNIV', 'COV PERS', 'COV', 'COV DET', 'GEVEY', 'ML BT', 'ML CABLE', 'MEM USB 16GB', 'MCP AAA', 'MCP FIBRA CARBON', 'MCP LARGE', 'MCP MEDIUM', 'MCP R', 'MCP XLARGE', 'MCP XSMALL', 'MC', 'MSD 8GB', 'MSD 16GB', 'MSD 32GB', 'MSD 64GB', 'OTROS', 'PTA 1A', 'PTA 2A', 'RELOJ', 'TARJETA INT', 'TELEF', 'CAMARAS', 'BOCINAS', 'BOTONES, JOY, BSIM', 'CRISTALES', 'FLEX POWER, VOL, ETC', 'FLEX', 'PC CONECTOR 1', 'PC CONECTOR 2', 'PC CONECTOR 3', 'PC CONECTOR 4', 'PSIM', 'TAPAS TRASERAS', 'PC FLEX', 'PLACAS', 'BAT ADAP', 'ANILLOS UNIV', 'TOUCH REC', 'LCD REC', 'ASS SIN CRISTAL', 'MC SIN CLASIF', 'CD MAGNETICO', 'MSD 128GB', 'MSD 256GB', 'MEM USB 32GB', 'MEM USB 64GB', 'MEM USB 128GB', 'COV SIN CLASIF', 'MICROFONOS', 'COV LENTO MOV', 'CRISTALES CAM', 'BAT COMO ADAP', 'BAT DET COMO ADAP', 'MHG']

    # Cargar el archivo Excel
    workbook = openpyxl.load_workbook('test.xlsx')

    # Obtener las primeras 7 hojas
    hojas = workbook.worksheets[:7]

    result={}
    # Iterar sobre cada hoja y realizar alguna acción (por ejemplo, imprimir el nombre)
    for hoja in hojas:
        for i in range(3,151):
            # Obtener el valor de la celda (por ejemplo, celda A1)
            pieza = hoja[f'B{i}'].value
            if pieza in tipos_de_piezas:
                if pieza not in result.keys():
                    result[pieza]=[]
                result[pieza].append(hoja[f'C{i}'].value)
    print(result)
    for key in result.keys():
        print(f'{key}:')
        print(result[key])
        print('')
    
main()